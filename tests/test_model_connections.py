import json
import threading
from types import SimpleNamespace

import pytest

from office_agent.model_connections import from_values, plan_catalog, validate_catalog, ConnectionState
from office_agent.settings_store import SettingsStore, SettingsError
from office_agent.webui.settings import Settings
from office_agent.webui.model_connections import ModelConnections


def connection(cid='official', key='fixture-private-one'):
    return {'id': cid, 'name': cid, 'protocol': 'openai', 'base_url': f'https://{cid}.example/v1',
            'api_key': key, 'enabled': True, 'context_window': 8192,
            'models': [{'id': 'same-model', 'capabilities': ['chat', 'agent']}, {'id': 'image-model', 'capabilities': ['image']}]}


def payload(rows, defaults=None):
    return {'connections': [{**r, 'key_action': 'replace' if r.get('api_key') else 'keep'} for r in rows], 'defaults': defaults or {}}


def fixture_router(tmp_path):
    auth = [True]
    events = []
    app = SimpleNamespace(busy=False, _push=events.append, path=None, editor=None)
    app.emit = lambda kind, data: (setattr(app, 'busy', False) if kind == 'idle' else events.append({'type': kind, 'data': data}))
    settings = Settings(threading.RLock(), lambda: (auth[0], app.busy), tmp_path)
    router = ModelConnections(settings, app, lambda: auth[0])
    return router, auth, events


def save(router, value):
    snap = router.snapshot()
    result = router.plan(value, snap['revision'])
    assert result['ok'], result
    receipt = router.settings.confirm_update(result['plan']['plan_id'], True)
    assert receipt['result']['status'] == 'saved', receipt
    return receipt['result']['save_id']


def select(router, cid, cap='chat', model='same-model'):
    snap = router.snapshot()
    row = next(r for r in snap['catalog']['connections'] if r['id'] == cid)
    result = router.select({'connection_id': cid, 'model_id': model, 'capability': cap, 'revision': row['revision']})
    assert result['ok'], result
    return router.capture(cap)


def test_two_same_name_models_route_credentials_and_freeze(tmp_path, monkeypatch):
    import httpx2 as httpx
    router, _, _ = fixture_router(tmp_path)
    save(router, payload([connection(), connection('custom', 'fixture-private-two')]))
    first = select(router, 'official')
    router.app.busy = True
    second = select(router, 'custom')
    calls = []

    def handle(request):
        calls.append((str(request.url), request.headers['authorization'], json.loads(request.content)['model']))
        return httpx.Response(200, json={'id': 'fixture', 'object': 'chat.completion', 'created': 1,
            'model': 'same-model', 'choices': [{'index': 0, 'message': {'role': 'assistant', 'content': 'ok'}, 'finish_reason': 'stop'}]})

    monkeypatch.setattr('office_agent.webui.model_connections._http_client', lambda: httpx.Client(transport=httpx.MockTransport(handle), follow_redirects=False, trust_env=False))
    one, two = router.build_model(first), router.build_model(second)
    try:
        assert one.invoke('fixture').content == 'ok'
        assert two.invoke('fixture').content == 'ok'
    finally:
        one.http_client.close()
        two.http_client.close()
    assert calls == [('https://official.example/v1/chat/completions', 'Bearer fixture-private-one', 'same-model'),
                     ('https://custom.example/v1/chat/completions', 'Bearer fixture-private-two', 'same-model')]
    assert router.capture('chat').connection_id == 'custom'
    public = json.dumps(router.snapshot())
    assert 'fixture-private-' not in public
    assert router.snapshot()['catalog']['connections'][0]['models'][0]['states']['chat'] == 'verified'


def test_mode_defaults_restart_disable_and_configuration_change(tmp_path):
    router, _, _ = fixture_router(tmp_path)
    defaults = {'chat': {'connection_id': 'official', 'model_id': 'same-model', 'capability': 'chat'}}
    save(router, payload([connection(), connection('custom')], defaults))
    chat = select(router, 'official')
    image = select(router, 'custom', 'image', 'image-model')
    assert router.capture('chat') == chat and router.capture('image') == image
    restarted, _, _ = fixture_router(tmp_path)
    assert restarted.capture('chat').connection_id == 'official'
    with pytest.raises(SettingsError, match='model_selection_required'):
        restarted.capture('image')
    edited = connection(); edited['base_url'] = 'https://updated.example/v1'
    save(router, payload([edited, connection('custom')], defaults))
    snap = router.snapshot()
    assert snap['catalog']['current']['chat']['reason'] == 'model_configuration_changed'
    assert snap['catalog']['connections'][0]['models'][0]['states']['chat'] == 'changed'
    select(router, 'official')
    edited['enabled'] = False
    save(router, payload([edited, connection('custom')]))
    assert router.snapshot()['catalog']['current']['chat']['reason'] == 'model_selection_invalid'
    save(router, payload([connection('custom')]))
    assert router.snapshot()['catalog']['current']['chat']['reason'] == 'model_selection_invalid'
    save(router, payload([]))
    assert router.active()
    with pytest.raises(SettingsError, match='model_selection_invalid'):
        router.capture('chat')


def test_transaction_cancel_conflict_expiry_rollback_and_key_actions(tmp_path, monkeypatch):
    router, _, _ = fixture_router(tmp_path)
    first = save(router, payload([connection()]))
    original = (tmp_path / '.env').read_bytes()
    snap = router.snapshot()
    rows = snap['catalog']['connections']
    public = payload([{k: v for k, v in rows[0].items() if k not in ('revision', 'api_key_configured', 'model_list_status')}])
    public['connections'][0]['models'] = [{'id': 'manual', 'capabilities': ['chat']}]
    result = router.plan(public, snap['revision'])
    assert result['ok']
    assert router.settings.confirm_update(result['plan']['plan_id'], False)['result']['code'] == 'settings_denied'
    assert (tmp_path / '.env').read_bytes() == original
    save(router, public)
    assert from_values(router.settings._store._values(router.settings._store._read(tmp_path / '.env')))['connections'][0]['api_key'] == 'fixture-private-one'
    public['connections'][0]['key_action'] = 'clear'
    save(router, public)
    assert not router.snapshot()['catalog']['connections'][0]['api_key_configured']
    snap = router.snapshot()
    candidate = payload([connection('new')])
    planned = router.plan(candidate, snap['revision'])['plan']
    (tmp_path / '.env').write_bytes((tmp_path / '.env').read_bytes() + b'EXTERNAL=1\n')
    assert router.settings.confirm_update(planned['plan_id'], True)['result']['code'] == 'settings_source_changed'
    snap = router.snapshot()
    planned = router.plan(candidate, snap['revision'])['plan']
    router.settings._store._pending[planned['plan_id']].created -= 301
    assert router.settings.confirm_update(planned['plan_id'], True)['result']['code'] == 'settings_plan_expired'
    last = save(router, candidate)
    rollback = router.settings.plan_rollback(last)['plan']['rollback_id']
    assert router.settings.confirm_rollback(rollback, True)['result']['status'] == 'rolled_back'
    assert router.snapshot()['catalog']['connections'][0]['models'][0]['id'] == 'manual'


@pytest.mark.parametrize('mutation', [
    lambda r: r.update(base_url='https://user:secret@example.test/v1'),
    lambda r: r.update(base_url='https://example.test/v1?key=secret'),
    lambda r: r.update(base_url='http://example.test/v1'),
    lambda r: r.update(protocol='ollama', base_url='http://192.168.1.2:11434'),
    lambda r: r.update(name='fixture-private-one'),
    lambda r: r.update(models=[{'id': 'bad name', 'capabilities': ['chat']}]),
    lambda r: r.update(models=[{'id': 'same-model', 'capabilities': ['shell']}]),
])
def test_invalid_catalog_rejected_without_secret_feedback(mutation):
    row = connection(); mutation(row)
    with pytest.raises(SettingsError):
        validate_catalog({'connections': [row], 'defaults': {}})


def test_auth_late_result_never_restores_selection(tmp_path):
    router, auth, _ = fixture_router(tmp_path)
    save(router, payload([connection()]))
    route = select(router, 'official')
    router.invalidate(); auth[0] = False
    router.record(route, 'verified')
    assert not router.state.outcomes and not router.state.selected
    assert router.snapshot() == {'ok': False, 'reason': 'not_authenticated'}


@pytest.mark.parametrize('status,expected', [(200, 'reachable'), (401, 'authentication_failed'), (404, 'models_unsupported'), (302, 'redirect_refused')])
def test_bounded_probe_does_not_verify_models(tmp_path, monkeypatch, status, expected):
    import httpx2
    from office_agent import settings_probe
    router, _, _ = fixture_router(tmp_path)
    save(router, payload([connection()]))
    row = router.snapshot()['catalog']['connections'][0]
    requests = []

    def handler(request):
        requests.append(request)
        return httpx2.Response(status, stream=httpx2.ByteStream(json.dumps({'data': [{'id': 'discovered'}, {'id': 'fixture-private-one'}]}).encode()), headers={'Location': 'https://other.example/'})

    original = settings_probe.probe_connection
    monkeypatch.setattr(settings_probe, 'probe_connection', lambda prepared, **kw: original(prepared, httpx2.MockTransport(handler), **kw))
    result = router.probe(row['id'], row['revision'])
    assert result['ok'] and result['result']['status'] == expected
    assert len(requests) == 1 and requests[0].headers['authorization'] == 'Bearer fixture-private-one'
    if status == 200:
        assert result['result']['models'] == ['discovered']
    assert router.snapshot()['catalog']['connections'][0]['models'][0]['states']['chat'] == 'unverified'


def test_chat_duplicate_submission_and_inflight_selection(tmp_path, monkeypatch):
    router, _, events = fixture_router(tmp_path)
    save(router, payload([connection(), connection('custom')]))
    select(router, 'official')
    started, release, finished = threading.Event(), threading.Event(), threading.Event()
    used = []

    def build(route):
        used.append(route)

        def invoke(question):
            started.set()
            assert release.wait(5)
            return SimpleNamespace(content='answer')
        return SimpleNamespace(invoke=invoke)

    monkeypatch.setattr(router, 'build_model', build)
    router.app.emit = lambda kind, data: (setattr(router.app, 'busy', False), finished.set())
    assert router.chat('fixture')['ok']
    assert started.wait(5)
    assert router.chat('fixture')['reason'] == 'busy'
    select(router, 'custom')
    release.set(); assert finished.wait(5)
    assert len(used) == 1 and used[0].connection_id == 'official'
    assert router.capture('chat').connection_id == 'custom'
    assert any('official' in e.get('text', '') and e.get('tag') == 'meta' for e in events)


def test_image_inflight_switch_is_nonblocking_and_confirmation_is_one_shot(tmp_path):
    import io
    from PIL import Image
    from office_agent.webui.bridge import Api, WebviewApp
    from office_agent.writes import Editor, digest
    app = WebviewApp()
    app.path = tmp_path / 'owned.docx'
    app.path.write_bytes(b'Owned image transaction fixture')
    app.editor = Editor(tmp_path)
    api = Api(app, settings_root=tmp_path, auth=SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': True}))
    router = api._connections
    save(router, payload([connection(), connection('custom', 'fixture-private-two')]))
    first = select(router, 'official', 'image', 'image-model')
    app.pending_id = 'old-document-plan'
    app._image_pending_ids.add('old-image-plan')
    assert api.ask('An owned blue square')['generation_not_started']
    assert app.pending_id is None and len(app._image_pending_ids) == 1
    assert 'old-image-plan' not in app._image_pending_ids
    plan_id = next(iter(app._image_pending_ids))
    service = app._bound_image_session()[1]
    started, release = threading.Event(), threading.Event()
    calls, results = [], []
    stream = io.BytesIO(); Image.new('RGB', (12, 8), 'blue').save(stream, format='PNG')

    class Provider:
        def generate(self, request, config):
            calls.append(config)
            started.set()
            assert release.wait(5)
            return stream.getvalue()

    service._provider_factory = Provider
    before = digest(app.path)
    worker = threading.Thread(target=lambda: results.append(api.confirm_image_generation(plan_id, True)))
    worker.start()
    try:
        assert started.wait(5)
        assert app.busy
        second = select(router, 'custom', 'image', 'image-model')
        assert second.connection_id == 'custom'
        assert len(calls) == 1 and calls[0].api_key == first.api_key
        assert calls[0].base_url == 'https://official.example/v1'
    finally:
        release.set(); worker.join(5)
    assert not worker.is_alive()
    assert results[0]['ok'] and results[0]['status'] == 'generated'
    assert not api.confirm_image_generation(plan_id, True)['ok']
    assert len(calls) == 1 and digest(app.path) == before and not list(tmp_path.glob('*.png'))
    assert router.snapshot()['catalog']['connections'][0]['models'][1]['states']['image'] == 'verified'


def test_probe_timeout_is_redacted_and_manual_models_remain(tmp_path, monkeypatch):
    import httpx2
    from office_agent import settings_probe
    router, _, _ = fixture_router(tmp_path)
    save(router, payload([connection()]))
    row = router.snapshot()['catalog']['connections'][0]

    def timeout(request):
        raise httpx2.ReadTimeout('fixture-private-one')

    original = settings_probe.probe_connection
    monkeypatch.setattr(settings_probe, 'probe_connection', lambda prepared, **kw: original(prepared, httpx2.MockTransport(timeout), **kw))
    result = router.probe(row['id'], row['revision'])
    assert result == {'ok': True, 'result': {'status': 'timeout'}}
    assert 'fixture-private' not in json.dumps(result)
    current = router.snapshot()['catalog']['connections'][0]
    assert current['models'][0]['id'] == 'same-model' and current['model_list_status'] == 'timeout'


def test_main_agent_task_captures_route_before_worker_runs(tmp_path, monkeypatch):
    from office_agent.webui.bridge import Api, WebviewApp
    from office_agent.writes import Editor
    app = WebviewApp()
    app.path = tmp_path / 'owned.docx'; app.path.write_bytes(b'owned')
    app.editor = Editor(tmp_path)
    api = Api(app, settings_root=tmp_path, auth=SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': True}))
    save(api._connections, payload([connection(), connection('custom')]))
    first = select(api._connections, 'official', 'agent')
    captured = []

    class Deferred:
        def __init__(self, target, args, daemon):
            captured.extend(args)

        def start(self):
            pass

    monkeypatch.setattr('office_agent.webui.bridge.threading.Thread', Deferred)
    app._start_task('fixture')
    select(api._connections, 'custom', 'agent')
    assert captured[-2] == first and api._connections.capture('agent').connection_id == 'custom'


def test_legacy_default_ollama_and_duplicate_names(tmp_path):
    catalog = from_values({'MODEL_PROVIDER': 'ollama'})
    assert catalog['defaults']['agent']['connection_id'] == 'legacy-ollama'
    ambiguous = from_values({'DEEPSEEK_API_KEY': 'fixture-private-one', 'DASHSCOPE_API_KEY': 'fixture-private-two'})
    assert not ambiguous['defaults']
    row = connection('other'); row['name'] = 'OFFICIAL'
    with pytest.raises(SettingsError):
        validate_catalog({'connections': [connection(), row], 'defaults': {}})


def test_invalid_legacy_numeric_value_never_echoes_configuration():
    with pytest.raises(SettingsError) as error:
        from_values({'MODEL_PROVIDER': 'deepseek', 'OFFICE_CONTEXT_WINDOW': 'fixture-private-one'})
    assert str(error.value) == 'settings_file_invalid'
    row = connection(); row['name'] = 'other-private-secret'
    with pytest.raises(SettingsError, match='settings_file_invalid'):
        from_values({'OTHER_API_KEY': 'other-private-secret', 'OFFICE_MODEL_CONNECTIONS': json.dumps({'connections': [row], 'defaults': {}})})


@pytest.mark.parametrize('configured_key', ['', 'fixture-private-one'])
def test_ollama_uses_only_selected_connection_credentials(tmp_path, monkeypatch, configured_key):
    import asyncio
    import httpx2
    monkeypatch.setenv('OLLAMA_API_KEY', 'ambient-private-key')
    router, _, _ = fixture_router(tmp_path)
    row = connection(key=configured_key)
    row.update(protocol='ollama', base_url='http://127.0.0.1:11434', models=[{'id': 'same-model', 'capabilities': ['chat', 'agent']}])
    save(router, payload([row]))
    route = select(router, 'official', 'agent')
    model = router.build_model(route)
    expected = 'Bearer ' + configured_key if configured_key else None
    try:
        assert model._client._client.headers.get('authorization') == expected
        assert model._async_client._client.headers.get('authorization') == expected
        assert 'ambient-private-key' not in repr(model)
        if configured_key:
            assert configured_key not in repr(model)
    finally:
        model._client._client.close()
        asyncio.run(model._async_client._client.aclose())
    used = []

    def respond(request):
        used.append(request.headers.get('authorization'))
        return httpx2.Response(200, stream=httpx2.ByteStream(b'{"message":{"content":"{}"}}'))

    research = router.research_model()
    research.transport = httpx2.MockTransport(respond)
    assert research._request('fixture', {}, lambda: False, 128, {}) == {}
    assert used == [expected]


def test_image_config_revalidation_does_not_wait_for_host_lock(tmp_path):
    router, _, _ = fixture_router(tmp_path)
    save(router, payload([connection()]))
    route = select(router, 'official', 'image', 'image-model')
    done = threading.Event()
    outcomes = []
    service_lock = threading.RLock()

    def reload_config():
        with service_lock:
            outcomes.append(router.image_config(route).model)
            done.set()

    with router.lock:
        worker = threading.Thread(target=reload_config)
        worker.start()
        independent = done.wait(2)
    worker.join(3)
    assert not worker.is_alive()
    assert independent and outcomes == ['image-model']
    router.invalidate()
    from office_agent.imagegen import ImageGenError
    with pytest.raises(ImageGenError, match='image_configuration_changed'):
        router.image_config(route)
