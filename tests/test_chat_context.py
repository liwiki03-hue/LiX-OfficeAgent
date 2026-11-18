import json
import threading
from pathlib import Path
from types import SimpleNamespace

import httpx2
import pytest

from tests.test_daily_flows import setup, finish
from tests.test_model_connections import connection, payload, save, select


def test_preview_matches_sent_pairs_and_does_not_save_or_authorize(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    configured = connection(); configured['context_window'] = 1024
    save(api._connections, payload([configured])); select(api._connections, 'official')
    calls = []
    def invoke(messages):
        calls.append(messages)
        return SimpleNamespace(content='结论保留20 ms与否定条件。' * 10, response_metadata={'finish_reason':'stop'})
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
    for number in range(6):
        assert api.ask(('问题' + str(number) + '，保留20 ms和限定条件。') * 10)['ok']; finish(app)
    state = api.conversation_snapshot()['result']
    path = Path(state['location']); before = path.read_bytes()
    question = '核对以上条件，不修改文件。' * 10
    preview = api.preview_chat_context(question)['result']
    assert len(calls) == 6 and path.read_bytes() == before
    assert preview['read_only'] and preview['omitted_turns'] > 0
    assert preview['estimated_input_tokens'] + preview['reply_token_reserve'] <= preview['context_window']
    assert [turn['id'] for turn in preview['included']] == [turn['id'] for turn in state['sessions'][0]['turns'][preview['omitted_turns']:]]
    assert not app.busy and app.pending_id is None and app.path is None and app._read_authorized_path is None
    response = api.ask(question, {'stream':False, 'context_revision':preview['context_revision']})
    assert response['context_turns'] == len(preview['included']) and response['omitted_turns'] == len(preview['omitted'])
    finish(app)
    assert [message['role'] for message in calls[-1]] == ['user','assistant'] * len(preview['included']) + ['user']
    assert calls[-1][-1]['content'] == question


@pytest.mark.parametrize('change', ['question','route','session','history'])
def test_stale_context_preview_never_starts_request(tmp_path, monkeypatch, change):
    app, api, _ = setup(tmp_path)
    calls = []
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=lambda messages: calls.append(messages) or SimpleNamespace(content='ok')))
    question = 'next'
    preview = api.preview_chat_context(question)['result']
    if change == 'question':
        question = 'changed'
    elif change == 'route':
        select(api._connections, 'custom')
    elif change == 'session':
        assert api.new_conversation()['ok']
    else:
        assert api.ask('intervening')['ok']; finish(app)
    before = len(calls)
    snapshot = api.conversation_snapshot()['result']
    result = api.ask(question, {'context_revision':preview['context_revision']})
    assert result == {'ok':False, 'reason':'chat_context_changed'}
    assert len(calls) == before and not app.busy
    assert api.conversation_snapshot()['result']['sessions'] == snapshot['sessions']


def test_preview_excludes_running_failed_turns_and_keeps_transfer_confirmation(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    started, release = threading.Event(), threading.Event()
    def invoke(messages):
        started.set()
        assert release.wait(5)
        return SimpleNamespace(content='answer', response_metadata={'finish_reason':'stop'})
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
    assert api.ask('running')['ok'] and started.wait(5)
    try:
        preview = api.preview_chat_context('next')['result']
        assert preview['context_turns'] == 0 and len(preview['excluded']) == 1
        assert app.busy
    finally:
        release.set(); finish(app)
    select(api._connections, 'custom')
    preview = api.preview_chat_context('next')['result']
    assert preview['transfer_required'] and len(preview['included']) == 1
    result = api.ask('next', {'context_revision':preview['context_revision']})
    assert result['reason'] == 'chat_transfer_required' and not app.busy


def test_preview_validation_auth_and_secret_redaction(tmp_path, monkeypatch):
    app, api, active = setup(tmp_path)
    monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('preview called model'))
    assert api.preview_chat_context('fixture-private-one')['reason'] == 'chat_sensitive_input'
    assert api.preview_chat_context('')['reason'] == 'empty'
    assert api.preview_chat_context('x' * 8001)['reason'] == 'empty'
    result = api.preview_chat_context('question')
    assert result['ok'] and 'fixture-private-' not in json.dumps(result)
    assert not app.busy
    active[0] = False
    assert api.preview_chat_context('question')['reason'] == 'not_authenticated'


def sse(content, reason):
    chunks = [{'id':'fixture', 'object':'chat.completion.chunk', 'created':1, 'model':'same-model',
               'choices':[{'index':0, 'delta':{'content':content}, 'finish_reason':None}]}]
    if reason is not None:
        chunks.append({'id':'fixture', 'object':'chat.completion.chunk', 'created':1, 'model':'same-model',
                       'choices':[{'index':0, 'delta':{}, 'finish_reason':reason}]})
    return ''.join('data: ' + json.dumps(chunk) + '\n\n' for chunk in chunks).encode() + (b'data: [DONE]\n\n' if reason else b'')


@pytest.mark.parametrize('reason,label', [('stop','问答已返回'), ('length','回复可能未完整'), (None,'回复完整性未确认')])
def test_actual_sdk_sse_completion_states_and_client_cleanup(tmp_path, monkeypatch, reason, label):
    app, api, _ = setup(tmp_path)
    requests, clients = [], []
    def handle(request):
        requests.append((str(request.url), request.headers['authorization'], json.loads(request.content)))
        return httpx2.Response(200, headers={'content-type':'text/event-stream'}, content=sse('真实SDK模拟SSE回复', reason))
    def client():
        value = httpx2.Client(transport=httpx2.MockTransport(handle), follow_redirects=False, trust_env=False)
        clients.append(value)
        return value
    monkeypatch.setattr('office_agent.webui.model_connections._http_client', client)
    assert api.ask('question')['ok']
    events = finish(app)
    turn = api.conversation_snapshot()['result']['sessions'][0]['turns'][0]
    assert turn['answer'] == '真实SDK模拟SSE回复' and turn['finish_reason'] == (reason or 'unknown')
    assert any(event.get('label') == label for event in events)
    assert len(requests) == 1 and clients[0].is_closed
    assert requests[0][:2] == ('https://official.example/v1/chat/completions','Bearer fixture-private-one')
    assert requests[0][2]['model'] == 'same-model' and requests[0][2]['stream'] is True
    assert 'fixture-private-' not in json.dumps(events)


@pytest.mark.parametrize('failure,expected', [(302,'redirect_refused'), (401,'authentication_failed'), (404,'model_not_found'),
    (429,'rate_limited'), (503,'service_failed'), ('timeout','timeout'), ('connect','address_unreachable'),
    ('malformed','response_incompatible')])
def test_actual_sdk_failures_do_not_retry_or_leak_response(tmp_path, monkeypatch, failure, expected):
    app, api, _ = setup(tmp_path)
    calls = []
    def handle(request):
        calls.append(request)
        if failure == 'timeout':
            raise httpx2.ReadTimeout('private remote body', request=request)
        if failure == 'connect':
            raise httpx2.ConnectError('private remote body', request=request)
        return httpx2.Response(200 if failure == 'malformed' else failure,
                              json={'error':{'message':'private remote body fixture-private-one'}})
    monkeypatch.setattr('office_agent.webui.model_connections._http_client', lambda: httpx2.Client(transport=httpx2.MockTransport(handle), follow_redirects=False, trust_env=False))
    assert api.ask('question', {'stream':False})['ok']
    events = finish(app)
    assert len(calls) == 1 and not app.busy
    assert next(event for event in events if event['type'] == 'chat_error')['code'] == expected
    assert 'private remote body' not in json.dumps(events) and 'fixture-private-' not in json.dumps(events)
    assert api.conversation_snapshot()['result']['sessions'][0]['turns'][0]['status'] == 'failed'


@pytest.mark.parametrize('stream', [True,False])
def test_ollama_done_reason_is_preserved(tmp_path, monkeypatch, stream):
    import importlib
    app, api, _ = setup(tmp_path)
    configured = connection(); configured.update(protocol='ollama', base_url='http://127.0.0.1:11434', models=[{'id':'same-model','capabilities':['chat']}])
    save(api._connections, payload([configured])); select(api._connections, 'official')
    module = importlib.import_module('ollama._client')
    original = module.httpx.Client
    requests, clients = [], []
    def handle(request):
        requests.append((str(request.url), request.headers.get('authorization'), json.loads(request.content)))
        body = {'model':'same-model', 'created_at':'2026-10-27T00:00:00Z', 'message':{'role':'assistant','content':'limited'}, 'done':True, 'done_reason':'length'}
        return module.httpx.Response(200, content=(json.dumps(body)+'\n').encode(), headers={'content-type':'application/x-ndjson'})
    def client(*args, **kwargs):
        value = original(*args, **kwargs, transport=module.httpx.MockTransport(handle))
        clients.append(value)
        return value
    monkeypatch.setattr(module.httpx, 'Client', client)
    assert api.ask('question', {'stream':stream})['ok']
    events = finish(app)
    assert api.conversation_snapshot()['result']['sessions'][0]['turns'][0]['finish_reason'] == 'length'
    assert any(event.get('label') == '回复可能未完整' for event in events)
    assert len(requests) == 1 and requests[0][:2] == ('http://127.0.0.1:11434/api/chat','Bearer fixture-private-one')
    assert requests[0][2]['stream'] is True and clients[0].is_closed
    assert any(event['type'] == 'answer_draft_start' for event in events) is stream


@pytest.mark.parametrize('stream', [True, False])
def test_leading_model_thought_is_absent_from_chat_events_history_and_followup(tmp_path, monkeypatch, stream):
    app, api, _ = setup(tmp_path, True)
    calls = []
    hidden = '不应保存的模型内部推理'
    answer = '平均延迟为120 ms[1]。\n```xml\n<think>字面示例</think>\n```'
    pieces = ['<th', 'ink>' + hidden, '</th', 'ink><think>另一段内部推理</think>', answer]
    def chunks(messages):
        calls.append(messages)
        for text in pieces:
            yield SimpleNamespace(content=text, response_metadata={})
        yield SimpleNamespace(content='', response_metadata={'finish_reason':'stop'})
    def invoke(messages):
        calls.append(messages)
        return SimpleNamespace(content=''.join(pieces), response_metadata={'finish_reason':'stop'})
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(stream=chunks, invoke=invoke))
    assert api.ask('平均延迟是多少？', {'stream':stream})['ok']
    events = finish(app)
    snapshot = api.conversation_snapshot()['result']
    turn = snapshot['sessions'][0]['turns'][0]
    assert turn['status'] == 'complete' and turn['answer'] == answer
    assert hidden not in json.dumps(events, ensure_ascii=False)
    assert hidden not in Path(snapshot['location']).read_text('utf-8')
    visible = ''.join(item.get('text', '') for item in events if item['type'] == 'answer_draft_delta')
    assert hidden not in visible and '另一段内部推理' not in visible
    assert api.ask('引用上一条回答', {'stream':False})['ok']
    followup_events = finish(app)
    assert len(calls) == 2 and hidden not in json.dumps(calls[-1], ensure_ascii=False)
    assert answer in [item['content'] for item in calls[-1] if item['role'] == 'assistant']
    assert hidden not in json.dumps(followup_events, ensure_ascii=False)


@pytest.mark.parametrize('stream', [True, False])
def test_unclosed_model_thought_fails_without_visible_or_persisted_reasoning(tmp_path, monkeypatch, stream):
    app, api, _ = setup(tmp_path, True)
    calls = []
    hidden = '未闭合的模型内部推理'
    def chunks(messages):
        calls.append(messages)
        yield SimpleNamespace(content='<thi', response_metadata={})
        yield SimpleNamespace(content='nk>' + hidden, response_metadata={'finish_reason':'stop'})
    def invoke(messages):
        calls.append(messages)
        return SimpleNamespace(content='<think>' + hidden, response_metadata={'finish_reason':'stop'})
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(stream=chunks, invoke=invoke))
    assert api.ask('只回答已知结论', {'stream':stream})['ok']
    events = finish(app)
    snapshot = api.conversation_snapshot()['result']
    turn = snapshot['sessions'][0]['turns'][0]
    assert turn['status'] == 'failed' and not turn['answer']
    assert len(calls) == 1 and not app.busy
    assert hidden not in json.dumps(events, ensure_ascii=False)
    assert hidden not in Path(snapshot['location']).read_text('utf-8')
    assert not any(item['type'] == 'answer_draft_delta' for item in events)
