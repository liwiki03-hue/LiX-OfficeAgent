import io
import json
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from office_agent.webui.bridge import Api, WebviewApp
from office_agent.webui.conversations import Conversations
from tests.test_model_connections import connection, payload, save, select


def setup(tmp_path, enabled=False):
    app = WebviewApp()
    active = [True]
    auth = SimpleNamespace(root=tmp_path / 'local-data', status=lambda: {'authenticated': active[0]}, account_key=lambda: 'a' * 64 if active[0] else None)
    api = Api(app, auth=auth, settings_root=tmp_path)
    save(api._connections, payload([connection(), connection('custom')]))
    if enabled:
        plan = api.plan_settings_update({'chat_history': True})
        assert plan['ok'], plan
        assert api.confirm_settings_update(plan['plan']['plan_id'], True)['ok']
    select(api._connections, 'official')
    return app, api, active


def finish(app, seconds=5):
    deadline = time.monotonic() + seconds
    result = []
    while time.monotonic() < deadline:
        result += app.poll()
        if not app.busy:
            return result
        time.sleep(.01)
    raise AssertionError('owned chat worker did not finish')


def test_multiturn_context_and_restart_without_authority(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    calls = []
    def invoke(messages):
        calls.append(messages)
        return SimpleNamespace(content='answer-' + str(len(calls)))
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
    for index in range(6):
        assert api.ask('question-' + str(index), {'stream': False})['ok']
        finish(app)
    assert len(calls[-1]) == 11
    assert [message['role'] for message in calls[-1]] == ['user', 'assistant'] * 5 + ['user']
    snapshot = api.conversation_snapshot()['result']
    assert len(snapshot['sessions'][0]['turns']) == 6
    assert Path(snapshot['location']).is_file()
    auth = api._auth
    restored_app = WebviewApp()
    restored = Api(restored_app, auth=auth, settings_root=tmp_path)
    monkeypatch.setattr(restored._connections, 'build_model', lambda *args: pytest.fail('restore invoked API'))
    state = restored.conversation_snapshot()['result']
    assert state['sessions'] == snapshot['sessions']
    assert restored_app.path is None and restored_app.editor is None
    assert restored_app.pending_id is None and restored_app._read_authorized_path is None
    assert not restored_app.busy and restored_app._image_session is None
    body = Path(snapshot['location']).read_text(encoding='utf8')
    assert 'fixture-private-' not in body and 'api_key' not in body and 'data:image' not in body


def test_disabled_history_and_explicit_preference_cancel(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=lambda messages: SimpleNamespace(content='ok')))
    before = (tmp_path / '.env').read_bytes()
    plan = api.plan_settings_update({'chat_history': True})
    assert api.confirm_settings_update(plan['plan']['plan_id'], False)['ok']
    assert (tmp_path / '.env').read_bytes() == before
    assert api.ask('unsaved')['ok']
    finish(app)
    assert not Path(api.conversation_snapshot()['result']['location']).exists()


def test_transfer_confirmation_and_failed_retry_context(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    calls = []
    fail = [False]
    def invoke(messages):
        calls.append(messages)
        if fail[0]:
            error = RuntimeError('fixture-private-one sensitive server content')
            error.status_code = 429
            raise error
        return SimpleNamespace(content='ok')
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
    assert api.ask('first')['ok']; finish(app)
    select(api._connections, 'custom')
    denied = api.ask('second')
    assert denied['reason'] == 'chat_transfer_required' and len(calls) == 1
    fail[0] = True
    assert api.ask('second', {'transfer_to': denied['route']})['ok']
    events = finish(app)
    assert next(e for e in events if e['type'] == 'chat_error')['code'] == 'rate_limited'
    assert 'sensitive server content' not in json.dumps(events)
    fail[0] = False
    transfer = api.ask('second')
    assert api.ask('second', {'transfer_to': transfer['route']})['ok']; finish(app)
    assert [m['content'] for m in calls[-1]] == ['first', 'ok', 'second']
    assert len(api.conversation_snapshot()['result']['sessions'][0]['turns']) == 3


def test_stream_stop_switch_and_late_completion(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    started, release = threading.Event(), threading.Event()
    calls = []
    def stream(messages):
        calls.append(messages)
        yield SimpleNamespace(content='partial')
        started.set()
        assert release.wait(5)
        yield SimpleNamespace(content='late answer')
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(stream=stream))
    first = api.conversation_snapshot()['result']['current']
    assert api.ask('slow')['ok'] and started.wait(5)
    try:
        assert api.ask('duplicate')['reason'] == 'busy'
        stopped = api.cancel_task()
        assert stopped['local_only'] and stopped['remote_cancelled'] is False
        second = api.new_conversation()['result']['current']
        assert second != first
        assert api.ask('new')['reason'] == 'busy'
    finally:
        release.set()
    events = finish(app)
    assert not any(e.get('tag') == 'answer' for e in events)
    assert api.conversation_snapshot()['result']['current'] == second
    first_row = next(r for r in api.conversation_snapshot()['result']['sessions'] if r['id'] == first)
    assert first_row['turns'][0]['status'] == 'stopped'
    assert len(calls) == 1 and not app.busy


def test_stream_secret_split_and_history_write_failure(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    def stream(messages):
        for piece in ['hello fixture-', 'private-one', ' world']:
            yield SimpleNamespace(content=piece)
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(stream=stream))
    import office_agent.webui.conversations as module
    monkeypatch.setattr(module.os, 'replace', lambda *args: (_ for _ in ()).throw(PermissionError('sensitive')))
    assert api.ask('fixture')['ok']
    events = finish(app)
    assert 'fixture-private-one' not in json.dumps(events)
    state = api.conversation_snapshot()['result']
    assert state['warning'] == 'chat_history_write_failed'
    assert state['sessions'][0]['turns'][0]['answer'] == 'hello [已隐藏密钥] world'
    assert not list((tmp_path / 'local-data').rglob('.pending-*'))


def test_corrupt_history_preserved_until_explicit_clear(tmp_path):
    key = lambda: 'b' * 64
    book = Conversations(tmp_path, key)
    path = book._path(); path.parent.mkdir(parents=True)
    path.write_text('{broken', encoding='utf8')
    state = book.public(True)
    assert state['warning'] == 'chat_history_corrupt'
    book.persist(True)
    assert path.read_text(encoding='utf8') == '{broken'
    with pytest.raises(ValueError):
        book.clear_saved(False)
    book.clear_saved(True)
    assert not path.exists()


def test_standalone_image_plan_preview_choose_save_and_rollback(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    select(api._connections, 'official', 'image', 'image-model')
    result = api.ask('画一张蓝色示意图', {'size': '1024x1536', 'style': 'natural'})
    assert result['ok'] and app.path is None and app.editor is None
    session, service = app._image_session
    plan_id = next(iter(service._plans))
    assert service._plans[plan_id].public['source'] is None
    data = io.BytesIO(); Image.new('RGB', (32, 32), 'blue').save(data, format='PNG')
    from office_agent.imagegen import GeneratedImage
    from office_agent.imageprovider import normalize_generated_image
    image = normalize_generated_image(data.getvalue())
    calls = []
    def generate(request, config):
        calls.append((request.size, config.model))
        return image.raw
    service._provider_factory = lambda: SimpleNamespace(generate=generate)
    assert api.confirm_image_generation(plan_id, True)['ok']
    assert not api.confirm_image_generation(plan_id, True)['ok'] and len(calls) == 1
    target = tmp_path / 'chosen-images'; target.mkdir()
    app.window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(target)])
    plan = api.plan_image_save(plan_id)
    assert plan['ok'], plan
    saved = api.confirm_image_save(plan['result']['save_id'], True)
    assert saved['ok'], saved
    output = Path(saved['result']['path'])
    assert output.parent == target and output.read_bytes() == image.raw
    assert not list(service.root.glob('*.png'))
    rollback = api.plan_image_rollback(saved['result']['save_id'])
    assert api.confirm_image_rollback(rollback['result']['rollback_id'], True)['ok']
    assert not output.exists() and app.path is None


def test_context_budget_omits_complete_pairs_without_changing_visible_history(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    row = connection(); row['context_window'] = 1024
    save(api._connections, payload([row]))
    select(api._connections, 'official')
    calls = []
    def invoke(messages):
        calls.append(messages)
        return SimpleNamespace(content='结论保留20 ms与否定条件。' * 10)
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
    omitted = []
    for index in range(8):
        result = api.ask(('问题' + str(index) + '，保留20 ms，不改变范围。') * 10)
        assert result['ok'], result
        omitted.append(result['omitted_turns'])
        finish(app)
    assert max(omitted) > 0
    assert all([message['role'] for message in messages] == ['user','assistant'] * ((len(messages)-1)//2) + ['user'] for messages in calls)
    assert len(api.conversation_snapshot()['result']['sessions'][0]['turns']) == 8
    assert calls[-1][-1]['content'].startswith('问题7')


def test_history_conflict_reload_and_delete_when_saving_disabled(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=lambda messages: SimpleNamespace(content='ok')))
    assert api.ask('original')['ok']; finish(app)
    state = api.conversation_snapshot()['result']; identifier = state['current']
    path = Path(state['location'])
    body = json.loads(path.read_text(encoding='utf8'))
    body['sessions'][0]['title'] = 'external title'
    path.write_text(json.dumps(body, ensure_ascii=False), encoding='utf8')
    before = path.read_bytes()
    assert api.ask('new in memory')['ok']; finish(app)
    assert api.conversation_snapshot()['result']['warning'] == 'chat_history_conflict'
    assert path.read_bytes() == before
    restored = api.reload_conversations()['result']
    assert restored['sessions'][0]['title'] == 'external title'
    assert len(restored['sessions'][0]['turns']) == 1
    plan = api.plan_settings_update({'chat_history':False})
    assert api.confirm_settings_update(plan['plan']['plan_id'], True)['ok']
    assert api.delete_conversation(identifier, True)['ok']
    assert not json.loads(path.read_text(encoding='utf8'))['sessions']


def test_logout_fences_late_chat_and_does_not_restore_permissions(tmp_path, monkeypatch):
    app, api, active = setup(tmp_path, True)
    started, release, ended = threading.Event(), threading.Event(), threading.Event()
    def invoke(messages):
        started.set()
        assert release.wait(5)
        return SimpleNamespace(content='late private answer')
    class Client:
        def close(self):
            ended.set()
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke,http_client=Client()))
    assert api.ask('slow')['ok'] and started.wait(5)
    active[0] = False; api._connections.invalidate(); app._clear_private_state()
    release.set(); assert ended.wait(5)
    finish(app)
    assert api.conversation_snapshot()['reason'] == 'not_authenticated'
    assert not app.busy and api._connections.chat_running is None
    assert 'late private answer' not in json.dumps(app.poll())


def test_image_history_is_bounded_metadata_without_restart_authority(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    select(api._connections, 'official', 'image', 'image-model')
    assert api.ask('画一张图')['ok']
    session, service = app._image_session
    identifier = next(iter(service._plans))
    data = io.BytesIO(); Image.new('RGB', (20, 20), 'white').save(data, format='PNG')
    service._provider_factory = lambda: SimpleNamespace(generate=lambda *args: data.getvalue())
    assert api.confirm_image_generation(identifier, True)['ok']
    image_history = api.list_independent_image_history()
    assert image_history['ok'] and len(image_history['items']) == 1
    assert 'data:image' not in json.dumps(image_history) and 'fixture-private-' not in json.dumps(image_history)
    newapp = WebviewApp(); restored = Api(newapp, auth=api._auth, settings_root=tmp_path)
    assert restored.list_independent_image_history()['items'] == image_history['items']
    assert restored.confirm_image_generation(identifier, True)['reason'] == 'no_session'
    assert newapp.path is None and newapp.editor is None


@pytest.mark.parametrize('kind,status,expected', [('auth',401,'authentication_failed'),('missing',404,'model_not_found'),('limit',429,'rate_limited'),('server',503,'service_failed'),('timeout',None,'timeout'),('connect',None,'address_unreachable'),('format',None,'response_incompatible')])
def test_chat_errors_are_specific_sanitized_and_never_retried(tmp_path, monkeypatch, kind, status, expected):
    app, api, _ = setup(tmp_path)
    calls = []
    def invoke(messages):
        calls.append(messages)
        error = TimeoutError('sensitive provider detail') if kind == 'timeout' else ConnectionError('sensitive provider detail') if kind == 'connect' else ValueError('sensitive provider detail')
        if status:
            error.status_code = status
        raise error
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
    assert api.ask('keep this prompt')['ok']
    events = finish(app)
    error = next(event for event in events if event['type'] == 'chat_error')
    assert error['code'] == expected and error['prompt'] == 'keep this prompt'
    assert 'sensitive provider detail' not in json.dumps(events) and len(calls) == 1


def test_rejected_file_admission_keeps_previous_selection_and_sanitizes_error(tmp_path, monkeypatch):
    from office_agent.webui import bridge
    app = WebviewApp()
    old = tmp_path / 'previous.docx'; old.write_bytes(b'previous')
    selected = tmp_path / 'new.docx'; selected.write_bytes(b'candidate')
    marker = SimpleNamespace()
    app.path, app.editor = old, marker
    app.window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(selected)])
    def deny(*args, **kwargs):
        raise ValueError('private parser detail')
    monkeypatch.setattr(bridge, 'Editor', deny)
    result = app.select_file()
    assert not result['ok'] and 'private parser detail' not in json.dumps(result)
    assert app.path == old and app.editor is marker


def test_other_connection_credentials_never_enter_saved_chat(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    second = connection('custom', 'fixture-other-credential-two')
    save(api._connections, payload([connection(),second]))
    select(api._connections,'official')
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='echo '+second['api_key'])))
    assert api.ask('normal question')['ok']
    events=finish(app)
    state=api.conversation_snapshot()['result']
    assert second['api_key'] not in json.dumps(events)
    assert second['api_key'] not in Path(state['location']).read_text(encoding='utf8')


def test_configured_credentials_in_prompt_are_not_sent_or_saved(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    monkeypatch.setattr(api._connections,'build_model',lambda *args:pytest.fail('credential entered model request'))
    result=api.ask('my credential is '+connection()['api_key'])
    assert result['reason']=='chat_sensitive_input' and not app.busy
    assert not Path(api.conversation_snapshot()['result']['location']).exists()


@pytest.mark.parametrize('status,expected',[(401,'image_authentication_failed'),(403,'image_authentication_failed'),(429,'image_rate_limited'),(503,'image_service_failed'),('connect','image_address_unreachable')])
def test_image_provider_reason_is_additive_and_preserves_failure_contract(tmp_path,monkeypatch,status,expected):
    import httpx2
    from office_agent.imageprovider import OpenAIImageGenProvider
    app,api,_=setup(tmp_path)
    select(api._connections,'official','image','image-model')
    assert api.ask('one image')['ok']
    session,service=app._image_session
    identifier=next(iter(service._plans))
    calls=[]
    def response(request):
        calls.append(request)
        if status=='connect':
            raise httpx2.ConnectError('sensitive remote information')
        return httpx2.Response(status,content=b'sensitive remote information')
    service._provider_factory=lambda:OpenAIImageGenProvider(transport_factory=lambda:httpx2.MockTransport(response))
    result=api.confirm_image_generation(identifier,True)
    assert not result['ok'] and result['failure_detail']=='image_provider_failed'
    assert result['provider_error']==expected and len(calls)==1
    assert 'sensitive remote information' not in json.dumps(result)
    assert result['provider_request_state']=='attempt_committed'
