import base64
import hashlib
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import time
import zipfile

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
import httpx2
import pytest

from office_agent.update_core import UpdateError, channel, check_installed, dependency_identity, install, inventory, local_info, safe_path, unpack, version, write_json
from office_agent.webui.updates import Updates, fetch


def fixture(root):
    root.mkdir(exist_ok=True)
    (root / 'office_agent').mkdir()
    (root / 'office_agent/sample.py').write_text('value = 1\n')
    (root / 'office_agent/update_core.py').write_bytes((Path(__file__).resolve().parents[1] / 'office_agent/update_core.py').read_bytes())
    (root / 'pyproject.toml').write_text('[project]\nname = "agent-quickstart"\nversion = "0.2.0"\ndependencies = []\n')
    (root / 'uv.lock').write_text('[[package]]\nname = "agent-quickstart"\nversion = "0.2.0"\n')
    (root / '.env').write_text('OWNED_TEST_FIXTURE=unchanged')
    current = json.loads((Path(__file__).resolve().parents[1] / 'office_agent/release_info.json').read_bytes())
    current['version'] = '0.2.0'
    write_json(root / 'office_agent/release_info.json', current)
    key = Ed25519PrivateKey.generate()
    public = base64.b64encode(key.public_key().public_bytes_raw()).decode()
    write_json(root / 'office_agent/update_channel.json', {'repository': 'owner/office-agent', 'public_key': public})
    write_json(root / 'office_agent/update_inventory.json', inventory(root))
    return key


def package(root, key, *, target='0.2.1', extra=None, dependencies=False):
    files = {name: (root / name).read_bytes() for name in inventory(root)['files']}
    files['office_agent/sample.py'] = b'value = 2\n'
    files['office_agent/release_info.json'] = json.dumps({'version': target, 'history': []}).encode()
    for name in ('pyproject.toml', 'uv.lock'):
        files[name] = files[name].replace(b'0.2.0', target.encode())
    if dependencies:
        files['pyproject.toml'] += b'new_dependency = true\n'
    files.update(extra or {})
    manifest = {'format': 'office-agent-update-v1', 'version': target, 'dependencies': dependency_identity(files),
                'files': {name: hashlib.sha256(value).hexdigest() for name, value in files.items()}}
    metadata = json.dumps(manifest).encode()
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('manifest.json', metadata)
        archive.writestr('manifest.sig', key.sign(metadata))
        for name, value in files.items():
            archive.writestr('payload/' + name, value)
    return stream.getvalue()


def ticket(root, body):
    state = root / '.office-agent-updates'
    state.mkdir(exist_ok=True)
    token = 'a' * 32
    (state / (token + '.zip')).write_bytes(body)
    path = state / (token + '-ticket.json')
    write_json(path, {'id': token, 'version': '0.2.1', 'sha256': hashlib.sha256(body).hexdigest(), 'created': time.time()})
    return path


def service(root, body):
    state = {'authenticated': True, 'busy': False, 'calls': [], 'handoffs': []}
    def handler(request):
        state['calls'].append(str(request.url))
        if request.url.host == 'api.github.com':
            return httpx2.Response(200, json={'tag_name': 'v0.2.1', 'draft': False, 'prerelease': False, 'body': '<script>untrusted notes</script>',
                'assets': [{'name': 'office-agent-0.2.1.zip', 'browser_download_url': 'https://github.com/owner/office-agent/releases/download/v0.2.1/office-agent-0.2.1.zip',
                            'size': len(body), 'digest': 'sha256:' + hashlib.sha256(body).hexdigest()}]})
        return httpx2.Response(200, content=body)
    updates = Updates(threading.RLock(), lambda: (state['authenticated'], state['busy']), root,
                      lambda *args: state['handoffs'].append(args), transport=httpx2.MockTransport(handler))
    return updates, state


def test_version_metadata_and_channel_no_network_when_unconfigured(tmp_path):
    fixture(tmp_path)
    write_json(tmp_path / 'office_agent/update_channel.json', {'repository': '', 'public_key': ''})
    updates, state = service(tmp_path, b'')
    assert updates.info()['result']['version'] == '0.2.0'
    assert updates.check()['reason'] == 'update_not_configured' and not state['calls']
    assert version('0.2.10') > version('0.2.9')
    for value in ('v1.0.0', '01.0.0', '1.0.0-beta', '../1.0.0'):
        with pytest.raises(UpdateError):
            version(value)


def test_signed_update_install_backup_and_preserve_configuration(tmp_path):
    key = fixture(tmp_path)
    before = (tmp_path / 'office_agent/sample.py').read_bytes()
    path = ticket(tmp_path, package(tmp_path, key))
    assert install(tmp_path, path) == 'installed'
    assert local_info(tmp_path)['version'] == '0.2.1' and check_installed(tmp_path)
    assert (tmp_path / '.env').read_text() == 'OWNED_TEST_FIXTURE=unchanged'
    assert (path.parent / ('a' * 32 + '-backup/office_agent/sample.py')).read_bytes() == before
    assert not path.exists() and not (path.parent / 'install.lock').exists()


def test_partial_write_failure_restores_old_version_and_records_rollback(tmp_path):
    key = fixture(tmp_path)
    path = ticket(tmp_path, package(tmp_path, key))
    old = inventory(tmp_path)
    count = [0]
    def failed(source, target):
        count[0] += 1
        if count[0] == 2:
            raise OSError('private system detail')
        os.replace(source, target)
    with pytest.raises(UpdateError, match='update_install_failed_rolled_back'):
        install(tmp_path, path, replace=failed)
    assert inventory(tmp_path) == old and check_installed(tmp_path)
    report = json.loads((path.parent / ('a' * 32 + '-audit.json')).read_bytes())
    assert report['status'] == 'rolled_back' and 'private' not in json.dumps(report)


@pytest.mark.parametrize('change', ['local', 'dependency', 'stale', 'tampered', 'downgrade'])
def test_conflicts_and_invalid_plans_never_replace_program(tmp_path, change):
    key = fixture(tmp_path)
    body = package(tmp_path, key, dependencies=change == 'dependency')
    path = ticket(tmp_path, body)
    if change == 'local':
        (tmp_path / 'office_agent/sample.py').write_text('my local change')
    elif change == 'stale':
        data = json.loads(path.read_bytes()); data['created'] -= 301; write_json(path, data)
    elif change == 'tampered':
        (path.parent / ('a' * 32 + '.zip')).write_bytes(b'tampered')
    elif change == 'downgrade':
        data = json.loads(path.read_bytes()); data['version'] = '0.1.0'; write_json(path, data)
    before = (tmp_path / 'office_agent/sample.py').read_bytes()
    with pytest.raises(UpdateError):
        install(tmp_path, path)
    assert (tmp_path / 'office_agent/sample.py').read_bytes() == before


def test_wrong_signing_key_and_zip_escape_refused(tmp_path):
    key = fixture(tmp_path)
    with pytest.raises(UpdateError, match='signature'):
        unpack(package(tmp_path, Ed25519PrivateKey.generate()), channel(tmp_path)['public_key'], '0.2.1')
    for name in ('../evil.py', 'office_agent/../../evil.py', 'office_agent/update_channel.json', 'office_agent/CON.py', 'office_agent/evil:stream.py', '.env'):
        with pytest.raises(UpdateError):
            unpack(package(tmp_path, key, extra={name: b'unsafe'}), channel(tmp_path)['public_key'], '0.2.1')


def test_download_cancel_auth_and_one_shot_confirm(tmp_path):
    key = fixture(tmp_path)
    updates, state = service(tmp_path, package(tmp_path, key))
    assert not state['calls'] and updates.check()['result']['status'] == 'available'
    ready = updates.download()['result']
    assert updates.confirm(ready['plan_id'], False)['result']['status'] == 'cancelled' and not state['handoffs']
    assert updates.confirm(ready['plan_id'], True)['reason'] == 'update_plan_expired'
    ready = updates.download()['result']
    state['authenticated'] = False
    updates.invalidate()
    assert updates.confirm(ready['plan_id'], True)['reason'] == 'not_authenticated'
    state['authenticated'] = True
    ready = updates.download()
    assert ready['reason'] == 'update_plan_expired'
    updates.check(); ready = updates.download()['result']
    assert updates.confirm(ready['plan_id'], True)['result']['status'] == 'queued'
    assert len(state['handoffs']) == 1 and local_info(tmp_path)['version'] == '0.2.0'


def test_late_network_response_and_repeated_click_refused(tmp_path):
    key = fixture(tmp_path)
    updates, state = service(tmp_path, package(tmp_path, key))
    original = updates.transport.handler
    def handler(request):
        assert updates.check()['reason'] == 'busy'
        updates.invalidate()
        return original(request)
    updates.transport = httpx2.MockTransport(handler)
    assert updates.check()['reason'] == 'update_session_changed' and updates.release is None


@pytest.mark.parametrize('mode,reason', [('auth', 'update_rate_limited'), ('timeout', 'update_timeout'), ('redirect', 'update_url_refused'), ('capacity', 'update_capacity')])
def test_network_errors_are_bounded_and_do_not_echo_service_details(mode, reason):
    calls = []
    def handler(request):
        calls.append(request)
        if mode == 'timeout':
            raise httpx2.ReadTimeout('private header data')
        if mode == 'redirect':
            return httpx2.Response(302, headers={'Location': 'https://untrusted.example/private'})
        if mode == 'auth':
            return httpx2.Response(403, content=b'private service detail')
        return httpx2.Response(200, content=b'too big')
    try:
        fetch('https://api.github.com/repos/owner/project/releases/latest', 3, transport=httpx2.MockTransport(handler))
    except httpx2.TimeoutException:
        assert mode == 'timeout'
    except UpdateError as error:
        assert str(error) == reason
    else:
        raise AssertionError('must refuse')
    assert len(calls) == 1


def test_release_metadata_matches_project_version():
    root = Path(__file__).resolve().parents[1]
    current = local_info(root)['version']
    assert f'version = "{current}"' in (root / 'pyproject.toml').read_text(encoding='utf8')


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows process wait and installer acceptance requires Windows')
def test_real_owned_installer_waits_for_parent_and_exits_cleanly(tmp_path):
    key = fixture(tmp_path)
    path = ticket(tmp_path, package(tmp_path, key))
    write_json(path.parent / ('a' * 32 + '-audit.json'), {'status': 'queued', 'from': '0.2.0', 'to': '0.2.1'})
    helper = path.parent / 'owned-installer.py'
    helper.write_bytes((tmp_path / 'office_agent/update_core.py').read_bytes())
    parent = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(2)'], creationflags=subprocess.CREATE_NO_WINDOW)
    child = None
    try:
        child = subprocess.Popen([sys.executable, str(helper), str(tmp_path), str(path), str(parent.pid), '--no-restart'],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
        assert child.wait(timeout=20) == 0
        assert parent.wait(timeout=3) == 0
        assert local_info(tmp_path)['version'] == '0.2.1' and check_installed(tmp_path)
    finally:
        for process in (child, parent):
            if process is not None and process.poll() is None:
                process.kill()
                process.wait(timeout=5)


def test_publisher_cli_builds_a_verifiable_release_without_shipping_private_key(tmp_path, monkeypatch, capsys):
    from tools import build_update
    root, publisher = tmp_path / 'app', tmp_path / 'publisher'
    fixture(root)
    publisher.mkdir()
    monkeypatch.setattr(build_update, 'ROOT', root)
    private = publisher / 'signing.pem'
    def command(*arguments):
        monkeypatch.setattr(sys, 'argv', ['build_update.py', *map(str, arguments)])
        build_update.main()
        return capsys.readouterr().out.strip()
    public = command('keygen', '--private-key', private)
    assert len(base64.b64decode(public)) == 32 and 'PRIVATE KEY' not in public
    command('configure', '--repository', 'owner/project', '--public-key', public)
    command('bump', '--version', '0.2.1', '--date', '2026-10-18', '--title', 'Owned release', '--item', 'Owned change')
    lock = root / 'uv.lock'
    lock.write_text(lock.read_text().replace('0.2.0', '0.2.1'))
    command('inventory')
    command('build', '--private-key', private, '--output', publisher)
    manifest, contents = unpack((publisher / 'office-agent-0.2.1.zip').read_bytes(), public, '0.2.1')
    assert manifest['version'] == '0.2.1' and 'pyproject.toml' in contents
    assert '.env' not in contents and all('signing.pem' not in name and 'update_channel.json' not in name for name in contents)
    with pytest.raises(SystemExit):
        command('keygen', '--private-key', root / 'office_agent/../unsafe.pem')
    assert not (root / 'unsafe.pem').exists()


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows batch launch verification requires Windows')
def test_restart_path_is_data_not_interpolated_shell_syntax(tmp_path):
    from office_agent.update_core import restart_application
    root = tmp_path / 'Owned & %TEMP% ! folder'
    root.mkdir()
    launcher = root / 'Start-OfficeAgent.cmd'
    launcher.write_bytes(b'@echo off\r\n> "%~dp0restart-marker.txt" echo owned\r\nexit /b 0\r\n')
    process = restart_application(root)
    try:
        assert process.wait(timeout=10) == 0
        assert (root / 'restart-marker.txt').read_text().strip() == 'owned'
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=5)
