import copy
import hashlib
import json
import re
import secrets
import threading
import time
from collections import OrderedDict
from datetime import datetime, timezone

import httpx2


SEARCH_URL = 'https://skills.sh/api/search'
MAX_BYTES = 512 * 1024
MAX_RESULTS = 24
PART = re.compile(r'[A-Za-z0-9][A-Za-z0-9_.-]{0,99}\Z')
SOURCE_ERRORS = frozenset({'invalid_github_source','fixed_commit_required','candidate_expired','source_rate_limited',
    'source_unavailable','source_budget','source_timeout','source_tree_budget','unsupported_source_path',
    'source_path_collision','skill_path_ambiguous','skill_path_missing','package_budget','unsupported_file_type','source_format',
    'source_integrity','unsupported_frontmatter','metadata_budget','body_budget','skill_too_large','invalid_source_version',
    'unsupported_reference_path','unsupported_reference_link','reference_budget','reference_missing','reference_cycle','reference_depth',
    'plugin_manifest_unsupported','plugin_path_missing','plugin_skill_not_declared','plugin_source_unsupported','workflow_asset_invalid'})


def source_reason(exc):
    if isinstance(exc, httpx2.TimeoutException):
        return 'source_timeout'
    reason = str(exc) if isinstance(exc, ValueError) else 'source_unavailable'
    return reason if reason in SOURCE_ERRORS else 'source_unavailable'


class DiscoveryError(Exception):
    pass


def fetch(query):
    started = time.monotonic()
    with httpx2.Client(timeout=8, follow_redirects=False, trust_env=False) as client:
        with client.stream('GET', SEARCH_URL, params={'q': query, 'limit': MAX_RESULTS},
                           headers={'Accept': 'application/json', 'User-Agent': 'LiX-SkillDiscovery/1'}) as response:
            if response.status_code == 429:
                raise DiscoveryError('rate_limited')
            if response.status_code != 200:
                raise DiscoveryError('service_error')
            data = bytearray()
            for chunk in response.iter_bytes():
                if time.monotonic() - started > 12:
                    raise DiscoveryError('timeout')
                data.extend(chunk)
                if len(data) > MAX_BYTES:
                    raise DiscoveryError('response_too_large')
            try:
                return json.loads(data)
            except (ValueError, UnicodeError):
                raise DiscoveryError('invalid_response') from None


def normalize(payload):
    if not isinstance(payload, dict) or not isinstance(payload.get('skills'), list):
        raise DiscoveryError('invalid_response')
    results, seen = [], set()
    for item in payload['skills'][:MAX_RESULTS]:
        if not isinstance(item, dict):
            raise DiscoveryError('invalid_response')
        source, skill, name = item.get('source'), item.get('skillId'), item.get('name')
        if not isinstance(source, str) or len(source.split('/')) != 2 or not all(PART.fullmatch(p) for p in source.split('/')):
            raise DiscoveryError('invalid_response')
        if not isinstance(skill, str) or not PART.fullmatch(skill) or not isinstance(name, str) or not 1 <= len(name) <= 160 or any(ord(c) < 32 for c in name):
            raise DiscoveryError('invalid_response')
        canonical = source + '/' + skill
        if item.get('id') != canonical:
            raise DiscoveryError('invalid_response')
        if canonical in seen:
            continue
        seen.add(canonical)
        results.append({'candidate_id': hashlib.sha256(('skills.sh:' + canonical).encode()).hexdigest()[:32],
                        'name': name, 'source': source, 'skill_id': skill,
                        'url': 'https://skills.sh/' + canonical, 'provider': 'skills.sh',
                        'compatibility': 'not_checked', 'license': None, 'commit': None})
    return results


class SkillDiscovery:
    def __init__(self, fetcher=fetch):
        self._fetch = fetcher
        self._slots = threading.BoundedSemaphore(2)
        self._lock = threading.Lock()
        self._cache = OrderedDict()
        self._resolved = OrderedDict()

    def featured(self):
        from .skill_catalog import featured
        return featured()

    def workflows(self, query=''):
        from .office_workflows import catalog
        if not isinstance(query,str) or len(query)>120 or any(ord(char)<32 for char in query):
            return {'ok':False,'reason':'invalid_query','items':[]}
        try:
            return catalog(query.strip())
        except (OSError,ValueError,KeyError,TypeError):
            return {'ok':False,'reason':'workflow_asset_invalid','items':[]}

    def resolve(self, candidate_id=None, *, github_url=None, reader=None):
        from .public_skill_resolver import github_source, resolve
        if not self._slots.acquire(blocking=False):
            return {'ok':False,'reason':'busy'}
        try:
            local = next((item for item in self.workflows()['items'] if item['candidate_id']==candidate_id),None) if github_url is None else None
            if local:
                from .office_workflows import candidate
                result = candidate()
                if hashlib.sha256(result['_content'].encode()).hexdigest()!=local['package_sha256']:
                    raise ValueError('workflow_asset_invalid')
                return self._remember(result)
            if github_url is not None:
                if not isinstance(github_url, str) or len(github_url)>700:
                    raise ValueError('invalid_github_source')
                repository, commit, path = github_source(github_url)
                skill = path.split('/')[-1]
                fixed = (commit, path)
            else:
                featured = [item for item in self.featured()['items'] if item['candidate_id']==candidate_id]
                with self._lock:
                    matches = [item for result in self._cache.values() for item in result['items'] if item['candidate_id']==candidate_id]
                if featured:
                    item = featured[0]
                    repository, skill, fixed = item['source'],item['skill_id'],(item['commit'],item['path'])
                elif matches:
                    item = matches[0]
                    repository, skill = item['source'],item['skill_id']
                    fixed = (item['commit'],item['path']) if item.get('provider')=='openai/plugins' else None
                else:
                    raise ValueError('candidate_expired')
            options = {'reader':reader} if reader is not None else {}
            result = resolve(repository, skill, fixed=fixed, **options)
            if github_url is None and featured and result['origin']['source_sha256']!=featured[0]['source_sha256']:
                raise ValueError('source_integrity')
            return self._remember(result)
        except (OSError, ValueError, KeyError, TypeError, httpx2.HTTPError) as exc:
            return {'ok':False,'reason':source_reason(exc)}
        finally:
            self._slots.release()

    def _remember(self, result):
        identifier = secrets.token_hex(16)
        with self._lock:
            self._resolved[identifier] = (time.monotonic(), result)
            while len(self._resolved)>32:
                self._resolved.popitem(last=False)
        return {'ok':True,'resolution_id':identifier,**{key:value for key,value in result.items() if not key.startswith('_')}}

    def installation(self, identifier):
        if not isinstance(identifier, str):
            return None
        with self._lock:
            saved = self._resolved.get(identifier)
            return copy.deepcopy(saved[1]) if saved and time.monotonic()-saved[0]<300 else None

    def progressive(self, identifier):
        result = self.installation(identifier)
        if result is None or not result.get('progressive_available') or '_update_snapshot' in result:
            return {'ok':False,'reason':'candidate_expired'}
        result['_content'] = result.pop('_progressive_content')
        result['compatible'], result['blockers'], result['progressive_available'] = True, [], False
        result['reference_loading'] = 'task-matched'
        result['compatibility_note'] = '按需文字参考：完整保存资料，仅加载与要求匹配且能完整放入预算的文件；未加载项明确列出。'
        key = secrets.token_hex(16)
        with self._lock:
            self._resolved[key] = (time.monotonic(),result)
            while len(self._resolved)>32:
                self._resolved.popitem(last=False)
        return {'ok':True,'resolution_id':key,**{name:value for name,value in result.items() if not name.startswith('_')}}

    def resolve_update(self, snapshot, *, reader=None):
        from .public_skill_resolver import resolve
        if not self._slots.acquire(blocking=False):
            return {'ok':False, 'reason':'busy'}
        try:
            origin = snapshot['origin']
            if origin.get('provider')=='lix':
                return {'ok':False,'reason':'workflow_updates_with_app','stage':'update_source_resolution'}
            options = {'reader':reader} if reader is not None else {}
            result = resolve(origin['repository'], origin['path'].split('/')[-1], folder_path=origin['path'], **options)
            if origin.get('reference_loading')=='task-matched' and result.get('progressive_available'):
                result['_content'] = result.pop('_progressive_content')
                result['compatible'], result['blockers'] = True, []
            previous = origin.get('manifest', {'SKILL.md':origin['source_sha256']})
            following = result['manifest']
            files = [{'path':name, 'status':'新增' if name not in previous else '移除' if name not in following else '修改',
                      'before':previous.get(name), 'after':following.get(name)}
                     for name in sorted(previous.keys() | following.keys()) if previous.get(name) != following.get(name)]
            if origin['license_text'] != result['license_text']:
                files.append({'path':'许可文本', 'status':'修改'})
            if origin['local_modified']:
                result['blockers'].append('本地指令或元数据已修改，保留当前版本；先另存改动后再处理更新')
            if origin['commit'] == result['origin']['commit'] and not files:
                result['blockers'].append('当前已是相同固定版本')
            result['compatible'] = not result['blockers']
            result['update'] = {'current':{key:origin[key] for key in ('repository','path','commit','source_sha256','license')},
                                'files':files, 'local_modified':origin['local_modified'],
                                'permissions':'文字参考范围不变；不增加工具、网络或文件权限'}
            result['_update_snapshot'] = copy.deepcopy(snapshot)
            identifier = secrets.token_hex(16)
            with self._lock:
                self._resolved[identifier] = (time.monotonic(), result)
                while len(self._resolved)>32:
                    self._resolved.popitem(last=False)
            return {'ok':True, 'resolution_id':identifier, **{key:value for key,value in result.items() if not key.startswith('_')}}
        except (ValueError, KeyError, TypeError, httpx2.HTTPError) as exc:
            return {'ok':False, 'reason':source_reason(exc), 'stage':'update_source_resolution'}
        finally:
            self._slots.release()

    def search(self, query):
        if not isinstance(query, str) or not 2 <= len(query.strip()) <= 120 or any(ord(c) < 32 for c in query):
            return {'ok': False, 'reason': 'invalid_query', 'items': []}
        query = query.strip()
        if not self._slots.acquire(blocking=False):
            return {'ok': False, 'reason': 'busy', 'items': []}
        started = time.monotonic()
        try:
            items = normalize(self._fetch(query))
            result = {'ok': True, 'query': query, 'items': items, 'provider': 'skills.sh',
                      'fetched_at': datetime.now(timezone.utc).isoformat(),
                      'elapsed_ms': round((time.monotonic() - started) * 1000),
                      'limit': MAX_RESULTS, 'cached': False}
            with self._lock:
                self._cache[query] = copy.deepcopy(result)
                self._cache.move_to_end(query)
                while len(self._cache) > 32:
                    self._cache.popitem(last=False)
            return result
        except (DiscoveryError, httpx2.TimeoutException, httpx2.HTTPError) as exc:
            reason = str(exc) if isinstance(exc, DiscoveryError) else 'timeout' if isinstance(exc, httpx2.TimeoutException) else 'network_error'
            with self._lock:
                cached = copy.deepcopy(self._cache.get(query))
            if cached:
                return {**cached, 'ok': False, 'reason': reason, 'cached': True}
            return {'ok': False, 'reason': reason, 'items': [], 'query': query}
        finally:
            self._slots.release()

    def search_plugins(self, query, *, reader=None, commit=None):
        from .public_skill_resolver import read_api
        from .plugin_sources import search
        if not isinstance(query,str) or not 2<=len(query.strip())<=120 or any(ord(c)<32 for c in query):
            return {'ok':False,'reason':'invalid_query','items':[]}
        query = query.strip()
        key = ('openai/plugins',query)
        if not self._slots.acquire(blocking=False):
            return {'ok':False,'reason':'busy','items':[]}
        try:
            result = search(query,reader or read_api,commit=commit)
            result['fetched_at'] = datetime.now(timezone.utc).isoformat()
            if result['ok'] or result['items']:
                with self._lock:
                    self._cache[key] = copy.deepcopy(result)
                    self._cache.move_to_end(key)
                    while len(self._cache)>32:
                        self._cache.popitem(last=False)
            return result
        except (ValueError,KeyError,TypeError,httpx2.HTTPError) as exc:
            reason = source_reason(exc)
            with self._lock:
                cached = copy.deepcopy(self._cache.get(key))
            return {**cached,'ok':False,'reason':reason,'cached':True} if cached else {'ok':False,'reason':reason,'items':[]}
        finally:
            self._slots.release()
