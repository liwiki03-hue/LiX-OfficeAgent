import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time
from urllib.parse import urlsplit
import uuid

import httpx2

from ..update_core import UpdateError, MAX_PACKAGE, channel, check_installed, digest, local_info, read_json, safe_path, unpack, version, write_json


def fetch(url, limit, *, transport=None):
    started = time.monotonic()
    hosts = {'api.github.com', 'github.com', 'release-assets.githubusercontent.com', 'objects.githubusercontent.com', 'github-releases.githubusercontent.com'}
    with httpx2.Client(transport=transport or httpx2.HTTPTransport(retries=0, trust_env=False), trust_env=False, follow_redirects=False, timeout=15) as client:
        for _ in range(4):
            parsed = urlsplit(url)
            if parsed.scheme != 'https' or parsed.hostname not in hosts or parsed.port not in (None, 443) or parsed.username or parsed.password or parsed.fragment:
                raise UpdateError('update_url_refused')
            with client.stream('GET', url, headers={'User-Agent': 'OfficeAgent-Updater', 'Accept': 'application/vnd.github+json' if parsed.hostname == 'api.github.com' else 'application/octet-stream', 'Accept-Encoding': 'identity'}) as response:
                if response.status_code in (301, 302, 303, 307, 308):
                    url = response.headers.get('location', '')
                    continue
                if response.status_code == 404:
                    raise UpdateError('update_release_missing')
                if response.status_code in (403, 429):
                    raise UpdateError('update_rate_limited')
                if response.status_code != 200 or response.headers.get('content-encoding', 'identity') != 'identity':
                    raise UpdateError('update_network_failed')
                body = bytearray()
                for chunk in response.iter_bytes():
                    if time.monotonic() - started > 30:
                        raise UpdateError('update_timeout')
                    if len(body) + len(chunk) > limit:
                        raise UpdateError('update_capacity')
                    body.extend(chunk)
                return bytes(body)
    raise UpdateError('update_url_refused')


class Updates:
    def __init__(self, lock, scope, root, handoff, *, transport=None):
        self.lock, self.scope, self.root, self.handoff = lock, scope, Path(root).absolute(), handoff
        self.transport = transport
        self.epoch = uuid.uuid4().hex
        self.working = False
        self.release = self.prepared = None
        self.queued = False

    def invalidate(self):
        with self.lock:
            self.epoch = uuid.uuid4().hex
            self.release = self.prepared = None

    def guard(self):
        authenticated, busy = self.scope()
        if not authenticated:
            raise UpdateError('not_authenticated')
        if busy or self.queued:
            raise UpdateError('busy')

    def run(self, operation):
        try:
            return {'ok': True, 'result': operation()}
        except UpdateError as error:
            return {'ok': False, 'reason': str(error)}
        except httpx2.TimeoutException:
            return {'ok': False, 'reason': 'update_timeout'}
        except httpx2.HTTPError:
            return {'ok': False, 'reason': 'update_network_failed'}
        except Exception:
            return {'ok': False, 'reason': 'update_failed'}

    def info(self):
        def operation():
            with self.lock:
                self.guard()
                value = local_info(self.root)
                source = channel(self.root)
                state = safe_path(self.root, '.office-agent-updates', internal=True)
                recent = []
                if state.exists():
                    for path in sorted(state.glob('*-audit.json'), key=lambda path: path.stat().st_mtime, reverse=True)[:3]:
                        checked = safe_path(self.root, path.relative_to(self.root).as_posix(), internal=True)
                        report = read_json(checked, 16000)
                        recent.append({key: report.get(key) for key in ('status', 'from', 'to')})
                return {**value, 'configured': bool(source['repository']), 'repository': source['repository'], 'install_history': recent}
        return self.run(operation)

    def network(self, operation):
        def execute():
            with self.lock:
                self.guard()
                if self.working:
                    raise UpdateError('busy')
                self.working = True
                epoch = self.epoch
            try:
                result = operation()
                with self.lock:
                    self.guard()
                    if epoch != self.epoch:
                        raise UpdateError('update_session_changed')
                    return result()
            finally:
                with self.lock:
                    self.working = False
        return self.run(execute)

    def check(self):
        def operation():
            source = channel(self.root)
            if not source['repository']:
                raise UpdateError('update_not_configured')
            data = json.loads(fetch('https://api.github.com/repos/' + source['repository'] + '/releases/latest', 200000, transport=self.transport))
            target = data.get('tag_name', '').removeprefix('v')
            version(target)
            if data.get('draft') is not False or data.get('prerelease') is not False:
                raise UpdateError('update_release_invalid')
            available = version(target) > version(local_info(self.root)['version'])
            asset = None
            if available:
                candidates = [row for row in data.get('assets', []) if row.get('name') == f'office-agent-{target}.zip']
                if len(candidates) != 1:
                    raise UpdateError('update_asset_missing')
                asset = candidates[0]
                expected_url = f'https://github.com/{source["repository"]}/releases/download/{data["tag_name"]}/{asset["name"]}'
                if (asset.get('browser_download_url') != expected_url or type(asset.get('size')) is not int or not 0 < asset['size'] <= MAX_PACKAGE or
                        not re.fullmatch('sha256:[a-f0-9]{64}', asset.get('digest', ''))):
                    raise UpdateError('update_release_invalid')
            result = {'status': 'available' if available else 'current', 'version': target, 'notes': str(data.get('body') or '')[:12000]}
            def commit():
                self.release = {'public': result, 'asset': asset, 'channel': source, 'created': time.monotonic()}
                self.prepared = None
                return result
            return commit
        return self.network(operation)

    def download(self):
        def operation():
            with self.lock:
                release = self.release
                if not release or not release['asset'] or time.monotonic() - release['created'] > 300:
                    raise UpdateError('update_plan_expired')
            check_installed(self.root)
            asset = release['asset']
            body = fetch(asset['browser_download_url'], MAX_PACKAGE, transport=self.transport)
            fingerprint = hashlib.sha256(body).hexdigest()
            if len(body) != asset['size'] or 'sha256:' + fingerprint != asset['digest']:
                raise UpdateError('update_package_invalid')
            manifest, _ = unpack(body, release['channel']['public_key'], release['public']['version'])
            if manifest['dependencies'] != check_installed(self.root)['dependencies']:
                raise UpdateError('update_dependencies_changed')
            def commit():
                state = safe_path(self.root, '.office-agent-updates', internal=True)
                state.mkdir(exist_ok=True)
                token = uuid.uuid4().hex
                target = state / (token + '.zip')
                with target.open('xb') as stream:
                    stream.write(body)
                    stream.flush()
                    os.fsync(stream.fileno())
                self.prepared = {'id': token, 'version': manifest['version'], 'sha256': fingerprint, 'created': time.time()}
                return {'status': 'ready', 'plan_id': token, 'version': manifest['version'], 'expires_in_seconds': 300,
                        'changed_files': len(manifest['files']), 'restart_required': True}
            return commit
        return self.network(operation)

    def confirm(self, token, approved=False):
        def operation():
            with self.lock:
                self.guard()
                if self.working:
                    raise UpdateError('busy')
                plan, self.prepared = self.prepared, None
                if not plan or type(token) is not str or token != plan['id'] or time.time() - plan['created'] > 300:
                    raise UpdateError('update_plan_expired')
                if approved is not True:
                    return {'status': 'cancelled'}
                check_installed(self.root)
                state = safe_path(self.root, '.office-agent-updates', internal=True)
                if (state / 'install.lock').exists():
                    raise UpdateError('update_install_locked')
                if digest(state / (token + '.zip')) != plan['sha256']:
                    raise UpdateError('update_package_invalid')
                ticket = state / (token + '-ticket.json')
                write_json(ticket, plan)
                helper = state / (token + '-installer.py')
                with helper.open('xb') as stream:
                    stream.write((self.root / 'office_agent/update_core.py').read_bytes())
                write_json(state / (token + '-audit.json'), {'status': 'queued', 'from': local_info(self.root)['version'], 'to': plan['version']})
                try:
                    self.handoff(helper, ticket)
                except UpdateError as error:
                    write_json(state / (token + '-audit.json'), {'status': 'refused', 'from': local_info(self.root)['version'], 'to': plan['version'], 'code': str(error)})
                    raise
                self.queued = True
                return {'status': 'queued', 'version': plan['version']}
        return self.run(operation)
