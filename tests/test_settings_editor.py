import json
import threading
from pathlib import Path

import httpx2
import pytest

from office_agent.settings_store import SettingsStore, SettingsError
from office_agent.settings_probe import prepare_probe, probe_connection, MAX_RESPONSE
from office_agent.webui.settings import Settings


FIXTURE = b'MODEL_PROVIDER=deepseek\nDEEPSEEK_MODEL=alpha\nDEEPSEEK_BASE_URL=https://deepseek.invalid/v1\nDEEPSEEK_API_KEY=fixture-private-one\nDASHSCOPE_MODEL=beta\nDASHSCOPE_BASE_URL=https://dashscope.invalid/v1\nDASHSCOPE_API_KEY=fixture-private-two\nOFFICE_EMBEDDING_CHANNEL=local\n'


def store_at(path):
    (path / '.env').write_bytes(FIXTURE)
    return SettingsStore(path)


def test_profiles_diff_noop_and_snapshot_conflict(tmp_path):
    store = store_at(tmp_path)
    view = store.snapshot()
    assert view['profiles']['dashscope']['model'] == 'beta'
    assert view['profiles']['deepseek']['api_key_configured']
    assert 'fixture-private' not in json.dumps(view)
    with pytest.raises(SettingsError, match='settings_no_changes'):
        store.plan_edit({'provider': 'deepseek', 'model': 'alpha'}, 'session', view['revision'])
    plan = store.plan_edit({'provider': 'dashscope', 'model': 'beta-new'}, 'session', view['revision'])
    assert plan['differences'] == [{'field': 'model', 'before': 'beta', 'after': 'beta-new'}, {'field': 'provider', 'before': 'deepseek', 'after': 'dashscope'}]
    assert store.confirm_update(plan['plan_id'], 'session', False)['wrote_nothing']
    assert len(list(tmp_path.iterdir())) == 1
    (tmp_path / '.env').write_bytes(FIXTURE + b'EXTERNAL=1\n')
    with pytest.raises(SettingsError, match='settings_source_changed'):
        store.plan_edit({'model': 'new'}, 'session', view['revision'])


def test_explicit_empty_and_equivalent_numeric_values(tmp_path):
    store = store_at(tmp_path)
    (tmp_path / '.env').write_bytes(b'MODEL_PROVIDER=deepseek\nDEEPSEEK_MODEL=\nOFFICE_EMBEDDING_TIMEOUT=2.50\n')
    view = store.snapshot()
    assert view['profiles']['deepseek']['model_present']
    assert not view['profiles']['deepseek']['base_url_present']
    assert view['disk_provider_explicit']
    with pytest.raises(SettingsError, match='settings_no_changes'):
        store.plan_edit({'provider': 'deepseek', 'embedding_timeout': 2.5}, 'session', view['revision'])
    assert len(list(tmp_path.iterdir())) == 1


@pytest.mark.parametrize('action', ['keep', 'replace', 'clear'])
def test_provider_secret_actions_and_rollback(tmp_path, action):
    store = store_at(tmp_path)
    changes = {'provider': 'dashscope', 'model': 'beta-new'}
    if action != 'keep':
        changes['api_key'] = 'fixture-new-private' if action == 'replace' else ''
    plan = store.plan_edit(changes, 'session', store.snapshot()['revision'])
    assert 'fixture-' not in json.dumps(plan)
    result = store.confirm_update(plan['plan_id'], 'session', True)
    values = store._values(store._read(tmp_path / '.env'))
    assert values['DEEPSEEK_API_KEY'] == 'fixture-private-one'
    assert values['DASHSCOPE_API_KEY'] == {'keep': 'fixture-private-two', 'replace': 'fixture-new-private', 'clear': ''}[action]
    rollback = store.plan_rollback(result['save_id'], 'session')
    assert store.confirm_rollback(rollback['rollback_id'], 'session', True)['restored']
    assert (tmp_path / '.env').read_bytes() == FIXTURE


@pytest.mark.parametrize('code,payload,expected', [
    (200, {'data': [{'id': 'alpha'}]}, 'reachable'),
    (200, {'unsafe': 'fixture-private-one'}, 'models_unverified'),
    (401, {'error': 'fixture-private-one'}, 'authentication_failed'),
    (403, {}, 'authentication_failed'), (404, {}, 'models_unsupported'),
    (405, {}, 'models_unsupported'), (302, {}, 'redirect_refused'),
    (500, {'error': 'fixture-private-one'}, 'service_failed'),
])
def test_probe_bounded_single_request(tmp_path, code, payload, expected):
    store = store_at(tmp_path)
    calls = []
    def handle(request):
        calls.append(request)
        return httpx2.Response(code, stream=httpx2.ByteStream(json.dumps(payload).encode()), headers={'Location': 'https://other.invalid/models'})
    prepared = prepare_probe(store, {'provider': 'deepseek', 'model': 'alpha', 'base_url': 'https://example.invalid/v1', 'api_key': 'fixture-private-one'})
    result = probe_connection(prepared, httpx2.MockTransport(handle))
    assert result['status'] == expected
    assert len(calls) == 1 and calls[0].method == 'GET'
    assert str(calls[0].url) == 'https://example.invalid/v1/models'
    assert calls[0].headers['Authorization'] == 'Bearer fixture-private-one'
    assert 'fixture-private' not in json.dumps(result)
    assert (tmp_path / '.env').read_bytes() == FIXTURE
    assert len(list(tmp_path.iterdir())) == 1


def test_probe_timeout_size_loopback_and_url_guard(tmp_path):
    store = store_at(tmp_path)
    prepared = prepare_probe(store, {'provider': 'ollama', 'base_url': 'http://127.0.0.1:11434', 'model': 'local-model'})
    def timeout(request):
        assert 'Authorization' not in request.headers
        assert request.url.path == '/api/tags'
        raise httpx2.ReadTimeout('fixture-private-one')
    assert probe_connection(prepared, httpx2.MockTransport(timeout)) == {'status': 'timeout'}
    def huge(request):
        return httpx2.Response(200, stream=httpx2.ByteStream(b' ' * (MAX_RESPONSE + 1)))
    assert probe_connection(prepared, httpx2.MockTransport(huge)) == {'status': 'response_too_large'}
    for url in ['http://external.invalid', 'https://user:pass@host.invalid', 'https://host.invalid?key=private', 'https://host.invalid#private']:
        with pytest.raises(SettingsError):
            prepare_probe(store, {'provider': 'deepseek', 'base_url': url})
    assert probe_connection(prepared, httpx2.MockTransport(lambda request: httpx2.Response(200, stream=httpx2.ByteStream(b'{"models":[{"name":"local-model"}]}'))))['model_listed']


def test_probe_logout_and_duplicate_fence(tmp_path, monkeypatch):
    from office_agent import settings_probe
    store_at(tmp_path)
    authenticated = [True]
    settings = Settings(threading.RLock(), lambda: (authenticated[0], False), tmp_path)
    entered, release = threading.Event(), threading.Event()
    def probe(prepared):
        entered.set()
        assert release.wait(5)
        return {'status': 'reachable'}
    monkeypatch.setattr(settings_probe, 'probe_connection', probe)
    results = []
    worker = threading.Thread(target=lambda: results.append(settings.test_connection({'provider': 'deepseek'})))
    worker.start()
    try:
        assert entered.wait(5)
        assert settings.test_connection({'provider': 'deepseek'})['reason'] == 'busy'
        authenticated[0] = False
        settings.invalidate()
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive()
    assert results == [{'ok': False, 'reason': 'not_authenticated'}]


def test_probe_changed_session_does_not_claim_logout(tmp_path, monkeypatch):
    from office_agent import settings_probe
    store_at(tmp_path)
    settings = Settings(threading.RLock(), lambda: (True, False), tmp_path)
    def probe(prepared):
        settings.invalidate()
        return {'status': 'reachable'}
    monkeypatch.setattr(settings_probe, 'probe_connection', probe)
    assert settings.test_connection({'provider': 'deepseek'}) == {'ok': False, 'reason': 'settings_session_invalid'}


from office_agent.browser_runtime import find_edge

EDGE = find_edge() or Path('__LiX_Edge_not_installed__')
STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'


def _editor(tmp_path, monkeypatch, initial=FIXTURE):
    if not EDGE.is_file():
        pytest.skip('Settings browser interaction requires the existing Windows Edge; no runtime is installed')
    from playwright.sync_api import sync_playwright
    monkeypatch.setenv('TEMP', str(tmp_path))
    monkeypatch.setenv('TMP', str(tmp_path))
    if initial is not None:
        (tmp_path / '.env').write_bytes(initial)
    facade = Settings(threading.RLock(), lambda: (True, False), tmp_path)
    calls = []
    methods = {'get_settings': facade.snapshot, 'plan_settings_edit': facade.plan_edit,
               'confirm_settings_update': facade.confirm_update, 'plan_settings_rollback': facade.plan_rollback,
               'confirm_settings_rollback': facade.confirm_rollback,
               'test_settings_connection': lambda draft: {'ok': True, 'result': {'status': 'models_unsupported'}}}
    def host_call(method, args):
        calls.append(method)
        return methods[method](*args)
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 800, 'height': 800})
            page.route('**/*', lambda route: route.abort())
            page.expose_function('settingsHost', host_call)
            html = (STATIC / 'index.html').read_text(encoding='utf-8')
            import re
            html = re.sub(r'<script\b[^>]*>.*?</script>', '', html, flags=re.S)
            page.set_content(html)
            page.add_style_tag(content=(STATIC / 'design_tokens.css').read_text(encoding='utf-8'))
            page.add_style_tag(content=(STATIC / 'theme.css').read_text(encoding='utf-8'))
            page.evaluate('''() => {
              document.getElementById('workspace').hidden = false;
              document.getElementById('auth-screen').hidden = true;
              window.authenticated = true; window.approved = true; window.modalTexts = [];
              window.confirm = text => { window.modalTexts.push(text); return window.approved; };
              window.OfficeAuth = {isAuthenticated: () => window.authenticated, ensure: async () => window.authenticated, wrapApi: api => api};
              window.pywebview = {api: new Proxy({}, {get: (_, method) => (...args) => window.settingsHost(method, args)})};
            }''')
            page.add_script_tag(path=str(STATIC / 'confirm_dialog.js'))
            from tests.ui_confirmation import drive_legacy_confirmations
            drive_legacy_confirmations(page, monkeypatch)
            page.add_script_tag(path=str(STATIC / 'settings_dialog.js'))
            page.add_script_tag(path=str(STATIC / 'settings_editor.js'))
            page.add_script_tag(path=str(STATIC / 'settings.js'))
            page.evaluate('window.OfficeSettings.show()')
            page.wait_for_function("document.getElementById('settings-status').textContent.includes('已读取')")
            yield page, facade, calls
        finally:
            browser.close()


@pytest.fixture
def editor(tmp_path, monkeypatch):
    yield from _editor(tmp_path, monkeypatch)


def test_browser_provider_drafts_close_discard_and_secret(editor):
    page, facade, calls = editor
    assert page.locator('#settings-save').is_disabled()
    page.locator('#settings-model').fill('alpha-draft')
    page.locator('#settings-provider').select_option('dashscope')
    assert page.locator('#settings-model').input_value() == 'beta'
    assert page.locator('#settings-url').input_value() == 'https://dashscope.invalid/v1'
    page.locator('#settings-model').fill('beta-draft')
    page.locator('#settings-provider').select_option('deepseek')
    assert page.locator('#settings-model').input_value() == 'alpha-draft'
    page.locator('#settings-key').fill('fixture-new-private')
    page.evaluate('window.approved = false')
    page.locator('#settings-provider').select_option('dashscope')
    from tests.ui_confirmation import respond_to_confirmation
    respond_to_confirmation(page, approved=False)
    assert page.locator('#settings-provider').input_value() == 'deepseek'
    page.keyboard.press('Escape')
    respond_to_confirmation(page, approved=False)
    assert page.locator('#settings-dialog').is_visible()
    page.locator('#settings-load').click()
    respond_to_confirmation(page, approved=False)
    assert calls == ['get_settings']
    page.locator('#settings-discard').click()
    assert page.locator('#settings-model').input_value() == 'alpha'
    assert page.locator('#settings-key').input_value() == ''
    assert page.locator('#settings-save').is_disabled()
    page.locator('#settings-model').fill('alpha-discard')
    page.evaluate('window.approved = true')
    page.keyboard.press('Escape')
    respond_to_confirmation(page, approved=True)
    assert page.locator('#settings-dialog').is_hidden()
    page.evaluate('window.OfficeSettings.show()')
    assert page.locator('#settings-model').input_value() == 'alpha'
    assert calls == ['get_settings']


def test_browser_save_diff_cancel_reread_rollback_and_probe(editor):
    page, facade, calls = editor
    page.locator('#settings-provider').select_option('dashscope')
    page.locator('#settings-model').fill('beta-new')
    page.locator('#settings-key').fill('fixture-new-private')
    page.evaluate('window.approved = false')
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'denied'")
    assert facade._store._values(facade._store._read(facade._root / '.env'))['MODEL_PROVIDER'] == 'deepseek'
    assert page.locator('#settings-key').input_value() == ''
    page.evaluate('window.approved = true')
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'saved_restart'")
    assert page.locator('#settings-save').is_disabled()
    assert page.locator('#settings-rollback').is_enabled()
    page.evaluate('async()=>{await window.OfficeSettings.close();await window.OfficeSettings.show();}')
    assert page.locator('#settings-status').get_attribute('data-state') == 'saved_restart'
    assert calls.count('get_settings') == 2
    modals = page.evaluate('window.modalTexts')
    assert 'beta → beta-new' in modals[0]
    assert '替换' in modals[0] and 'fixture-new-private' not in json.dumps(modals)
    values = facade._store._values(facade._store._read(facade._root / '.env'))
    assert values['DEEPSEEK_MODEL'] == 'alpha'
    assert values['DASHSCOPE_MODEL'] == 'beta-new'
    assert values['DASHSCOPE_API_KEY'] == 'fixture-private-two'
    page.locator('#settings-test').click()
    page.wait_for_function("document.getElementById('settings-test-status').textContent.includes('无法通过模型列表验证')")
    assert calls.count('test_settings_connection') == 1
    page.locator('#settings-rollback').click()
    page.wait_for_function("document.getElementById('settings-status').textContent.includes('已恢复')")
    assert (facade._root / '.env').read_bytes() == FIXTURE


def test_browser_validation_focus_and_late_response(editor):
    page, facade, calls = editor
    page.locator('#settings-url').fill('https://user:private@bad.invalid')
    assert page.locator('#settings-url').get_attribute('aria-invalid') == 'true'
    page.locator('#settings-save').click()
    assert calls == ['get_settings']
    page.locator('#settings-discard').click()
    page.evaluate('''() => {
      const original = window.pywebview.api;
      window.pywebview.api = new Proxy(original, {get: (target, method) => method === 'plan_settings_edit' ?
        (...args) => new Promise(resolve => {window.releasePlan = async () => resolve(await target[method](...args));}) : target[method]});
    }''')
    page.locator('#settings-model').fill('late-model')
    page.locator('#settings-save').click()
    page.wait_for_function('Boolean(window.releasePlan)')
    page.evaluate("document.getElementById('settings-save').click()")
    page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    page.evaluate('window.releasePlan()')
    assert 'confirm_settings_update' not in calls
    assert page.locator('#settings-key').input_value() == ''
    assert page.locator('#settings-model').input_value() == ''
    assert page.locator('#settings-save').is_disabled()


def test_browser_validation_parity_and_keyboard(editor):
    page, facade, calls = editor
    samples = [('model', 'bad model', False), ('base_url', 'http://cloud.invalid', False),
               ('base_url', 'https:/bad.invalid', False), ('base_url', 'https://host.invalid?key=x', False),
               ('base_url', 'https://host.invalid/v1', True), ('context_window', '1023', False),
               ('context_window', '32768', True), ('embedding_dimension', '1.5', False),
               ('embedding_dimension', '4097', False), ('embedding_dimension', '1024', True),
               ('embedding_timeout', '20.1', False), ('embedding_budget', '60.1', False)]
    ids = {'model': 'settings-model', 'base_url': 'settings-url', 'context_window': 'settings-context',
           'embedding_dimension': 'settings-embedding-dimension', 'embedding_timeout': 'settings-embedding-timeout',
           'embedding_budget': 'settings-embedding-budget'}
    for field, value, valid in samples:
        page.locator('#settings-discard').evaluate('(element) => element.click()')
        page.evaluate('''({id, value}) => {const node = document.getElementById(id); node.value = value; node.dispatchEvent(new Event('input'));}''', {'id': ids[field], 'value': value})
        assert (page.locator('#' + ids[field]).get_attribute('aria-invalid') == 'false') is valid
        change = {'provider': 'deepseek', field: float(value) if field in ('embedding_timeout', 'embedding_budget') else int(value) if field in ('context_window', 'embedding_dimension') and '.' not in value else float(value) if field == 'embedding_dimension' else value}
        response = facade.plan_edit(change, facade.snapshot()['settings']['revision'])
        assert response['ok'] is valid
        if valid:
            facade.confirm_update(response['plan']['plan_id'], False)
        else:
            assert response['reason'] == 'settings_invalid' and response['field'] == field
    page.locator('#settings-discard').click()
    page.locator('#settings-close').focus()
    page.keyboard.press('Shift+Tab')
    assert page.evaluate("document.getElementById('settings-dialog').contains(document.activeElement)")
    page.keyboard.press('Tab')
    assert page.locator('#settings-close').evaluate('(node) => node === document.activeElement')
    page.locator('#nav-settings').evaluate('(node) => node.focus()')
    page.evaluate('async()=>{await window.OfficeSettings.close();await window.OfficeSettings.show();}')
    page.locator('#tab-machine').click()
    assert page.locator('#tab-machine').get_attribute('aria-current') == 'true'
    assert page.locator('#settings-pane-model').is_hidden()
    assert page.locator('#settings-theme-choice').locator('option').count() == 2
    assert page.locator('#settings-theme-choice option').evaluate_all('(options) => options.map(option => option.value)') == ['dark', 'light']


def test_real_loopback_probe(tmp_path):
    from http.server import BaseHTTPRequestHandler, HTTPServer
    requests = []
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            requests.append((self.path, self.headers.get('Authorization')))
            body = b'{"models":[{"name":"local-model"}]}'
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        def log_message(self, *args):
            pass
    server = HTTPServer(('127.0.0.1', 0), Handler)
    worker = threading.Thread(target=server.serve_forever, kwargs={'poll_interval': .05})
    worker.start()
    try:
        store = store_at(tmp_path)
        prepared = prepare_probe(store, {'provider': 'ollama', 'model': 'local-model', 'base_url': f'http://127.0.0.1:{server.server_port}'})
        assert probe_connection(prepared) == {'status': 'reachable', 'model_listed': True}
    finally:
        server.shutdown()
        server.server_close()
        worker.join(5)
    assert not worker.is_alive()
    assert requests == [('/api/tags', None)]
    assert (tmp_path / '.env').read_bytes() == FIXTURE


def test_browser_duplicate_stale_plan_and_unknown_write(editor):
    page, facade, calls = editor
    page.evaluate('''() => {
      const original = window.pywebview.api;
      window.originalSettingsApi = original;
      window.pywebview.api = new Proxy(original, {get: (target, method) => method === 'plan_settings_edit' ?
        (...args) => new Promise(resolve => {window.releasePlan = async () => resolve(await target[method](...args));}) : target[method]});
    }''')
    page.locator('#settings-model').fill('stale-model')
    page.locator('#settings-save').click()
    page.wait_for_function('Boolean(window.releasePlan)')
    page.evaluate("document.getElementById('settings-save').click(); document.getElementById('settings-model').dispatchEvent(new Event('input'))")
    page.evaluate('window.releasePlan()')
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'denied'")
    assert calls.count('plan_settings_edit') == 1
    assert (facade._root / '.env').read_bytes() == FIXTURE
    page.evaluate('''() => {
      window.ragUncertain = false;
      window.addEventListener('officerag:invalidate', event => {window.ragUncertain ||= event.ragUncertain;});
      const original = window.originalSettingsApi;
      window.pywebview.api = new Proxy(original, {get: (target, method) => method === 'confirm_settings_update' ?
        async (...args) => {await target[method](...args); throw Error('private fixture');} : target[method]});
    }''')
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'uncertain'")
    assert '可能已经写入' in page.locator('#settings-status').inner_text()
    assert page.evaluate('window.ragUncertain')
    assert page.locator('#settings-save').is_disabled()
    assert page.locator('#settings-rollback').is_disabled()
    assert b'DEEPSEEK_MODEL=stale-model' in (facade._root / '.env').read_bytes()
