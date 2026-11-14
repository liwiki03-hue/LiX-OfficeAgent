import hashlib
import io
import json
from pathlib import Path

import pytest

from office_agent import downloads, network
from office_agent.diagnostics import Diagnostics
from tools import installed_preflight


class Response(io.BytesIO):
    status = 206

    def __init__(self, data, start=0, total=None, etag='fixed'):
        super().__init__(data)
        self.headers = {'Content-Range': f'bytes {start}-{start + len(data) - 1}/{total or len(data)}', 'ETag': etag}

    def geturl(self):
        return 'https://example.invalid/archive'


def test_download_retries_resumes_and_keeps_one_deadline(tmp_path):
    data = b'abcdefgh'
    calls, progress = [], []
    class Interrupted(Response):
        def read1(self, size):
            if self.tell():
                raise TimeoutError()
            return self.read(3)
    def opened(request, timeout):
        calls.append(request.get_header('Range'))
        return Interrupted(data) if len(calls) == 1 else Response(data[3:], 3, len(data))
    path = downloads.fetch('https://example.invalid/archive', tmp_path / 'archive', hashlib.sha256(data).hexdigest(),
                           size=len(data), opener=opened, progress=lambda **row: progress.append(row), seconds=10)
    assert path.read_bytes() == data and calls == ['bytes=0-7', 'bytes=3-7']
    assert any(row['state'] == 'retry' and row['attempt'] == 1 for row in progress)
    assert not list(tmp_path.glob('*.partial*'))
    def forbidden(*args, **kwargs):
        raise AssertionError('network must not run')
    assert downloads.fetch('https://example.invalid/archive', path, hashlib.sha256(data).hexdigest(), size=len(data), opener=forbidden) == path
    path.write_bytes(b'bad-cache')
    with pytest.raises(downloads.DownloadError, match='cache_changed'):
        downloads.fetch('https://example.invalid/archive', path, hashlib.sha256(data).hexdigest(), size=len(data), opener=forbidden)
    assert path.read_bytes() == b'bad-cache'
    with pytest.raises(downloads.DownloadError, match='quarantine_denied'):
        downloads.quarantine(path, hashlib.sha256(data).hexdigest(), False)
    assert downloads.quarantine(path, hashlib.sha256(data).hexdigest(), True).read_bytes() == b'bad-cache'


def test_download_cancel_and_changed_range_do_not_append(tmp_path):
    path = tmp_path / 'archive'
    partial = tmp_path / 'archive.partial'
    partial.write_bytes(b'abc')
    def opened(request, timeout):
        return Response(b'bad', 0, 8)
    with pytest.raises(downloads.DownloadError, match='cancelled'):
        downloads.fetch('https://example.invalid/archive', path, '0' * 64, size=8, opener=opened, cancelled=lambda: True)
    with pytest.raises(downloads.DownloadError, match='range'):
        downloads.fetch('https://example.invalid/archive', path, '0' * 64, size=8, opener=opened)
    assert partial.read_bytes() == b'abc' and not path.exists()


def test_preflight_missing_assets_is_readonly_and_actionable(tmp_path, monkeypatch):
    from tools import install_frontend_assets
    monkeypatch.setattr(install_frontend_assets, 'prepare', lambda root: pytest.fail('preflight downloaded'))
    result = installed_preflight.inspect(tmp_path)
    assert not result['ok'] and not result['checks']['frontend'] and not result['checks']['renderer_executable']
    assert result['network_requests'] == 0 and not result['preparation_performed']
    assert result['details']['frontend']['next_step'] == '重新准备第 2 阶段'
    assert not list(tmp_path.iterdir())


def test_diagnostics_rotates_without_content_or_config(tmp_path):
    logger = Diagnostics(tmp_path, limit=512, backups=2)
    for _ in range(20):
        assert logger.record(module='network', action='download', result='failed', code='timeout',
                             api_key='fixture-sensitive', document='private-body', headers={'Authorization': 'private'})
    assert logger.close()
    files = list(logger.path.parent.iterdir())
    assert 1 <= len(files) <= 3
    content = ''.join(path.read_text('utf-8') for path in files)
    assert 'fixture-sensitive' not in content and 'private-body' not in content and 'Authorization' not in content
    preview = logger.preview()
    assert not preview['code'] and preview['events'] and not preview['includes_documents']


def test_network_modes_bypass_and_no_ambient_tls(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(network, 'getproxies', lambda: {'https': 'http://proxy.invalid:8888'})
    assert network.NetworkOptions('auto').route('https://example.invalid') == 'http://proxy.invalid:8888'
    assert network.NetworkOptions('direct').route('https://example.invalid') is None
    configured = network.NetworkOptions('manual', 'http://user:fixture@proxy.invalid:8080', '.internal.invalid,10.0.0.0/8')
    assert configured.route('https://example.invalid') == 'http://user:fixture@proxy.invalid:8080'
    assert configured.route('http://localhost:11434') is None
    assert configured.route('https://service.internal.invalid') is None
    assert configured.route('http://10.2.3.4') is None
    assert 'fixture' not in repr(configured)


def test_actual_controlled_proxy_route_and_cache_quarantine_scope(tmp_path, monkeypatch):
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    import threading
    import httpx2
    from office_agent.download_cache import DownloadCache
    requests = []
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            requests.append((self.path, bool(self.headers.get('Proxy-Authorization'))))
            self.send_response(200); self.send_header('Content-Length', '2'); self.end_headers(); self.wfile.write(b'ok')
        def log_message(self, *args): pass
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    worker = threading.Thread(target=server.serve_forever); worker.start()
    try:
        (tmp_path / '.env').write_text(f'OFFICE_NETWORK_MODE=manual\nOFFICE_NETWORK_PROXY=http://fixture:synthetic@127.0.0.1:{server.server_port}\n', 'utf8')
        with httpx2.Client(transport=network.transport('http://destination.example/resource', root=tmp_path), trust_env=False, timeout=3) as client:
            assert client.get('http://destination.example/resource').text == 'ok'
        assert requests == [('http://destination.example/resource', True)]
    finally:
        server.shutdown(); server.server_close(); worker.join(5); assert not worker.is_alive()
    directory = tmp_path / '.runtime/downloads'; directory.mkdir(parents=True)
    bad = directory / 'superdoc-1.47.2.tgz'; bad.write_bytes(b'bad-cache')
    unrelated = directory / 'keep.docx'; unrelated.write_bytes(b'unrelated')
    cache = DownloadCache(tmp_path); plan = cache.preview('frontend')
    assert plan['repairable'] and cache.quarantine(plan['id'], False)['cancelled'] and bad.exists()
    plan = cache.preview('frontend'); bad.write_bytes(b'changed')
    with pytest.raises(downloads.DownloadError, match='changed'): cache.quarantine(plan['id'], True)
    plan = cache.preview('frontend'); result = cache.quarantine(plan['id'], True)
    assert result['network_requests'] == 0 and not bad.exists() and unrelated.read_bytes() == b'unrelated'
    assert (directory / result['files'][0]).read_bytes() == b'changed'


def test_draft_model_discovery_cache_and_changed_address_key_scope(tmp_path, monkeypatch):
    from tests.test_model_connections import fixture_router, save, payload, connection
    router, authenticated, events = fixture_router(tmp_path)
    save(router, payload([connection()]))
    original = (tmp_path / '.env').read_bytes()
    calls = []
    def probe(prepared, **kwargs):
        calls.append(prepared)
        return {'status': 'reachable', 'models': ['candidate-model'], 'model_listed': False}
    monkeypatch.setattr('office_agent.settings_probe.probe_connection', probe)
    draft = {'id': 'official', 'protocol': 'openai', 'base_url': 'https://official.example/v1', 'key_action': 'keep'}
    assert not router.probe_draft(draft)['cached']
    assert router.probe_draft(draft)['cached']
    assert not router.probe_draft(draft, True)['cached'] and len(calls) == 2
    wrong = {**draft, 'base_url': 'https://other.example/v1'}
    assert not router.probe_draft(wrong)['ok'] and len(calls) == 2
    result = router.probe_draft({**wrong, 'key_action': 'replace', 'api_key': 'fixture-for-other'})
    assert result['ok'] and calls[-1][1] == wrong['base_url'] and calls[-1][3] == 'fixture-for-other'
    assert 'fixture-for-other' not in json.dumps(result)
    assert (tmp_path / '.env').read_bytes() == original


def test_intro_preference_survives_private_browser_and_stale_bounds(tmp_path):
    from office_agent.webui.window_preferences import WindowPreferences
    preferences = WindowPreferences(tmp_path)
    assert preferences.read() == {}
    stale = {'topmost': False}
    preferences.set_intro_seen()
    preferences.write(stale)
    restored = WindowPreferences(tmp_path).read()
    assert restored == {'intro_seen': True, 'topmost': False}
    assert not any(name in preferences.path.read_text('utf8') for name in ('grant', 'approval', 'document', 'api_key'))
