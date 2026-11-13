import importlib
import json
from pathlib import Path
from types import SimpleNamespace

import httpx2
import pytest

from tests.test_daily_flows import setup, finish
from tests.test_model_connections import connection, payload, save, select


def local_connection(cid='official', window=1024):
    row = connection(cid)
    row.update(protocol='ollama',base_url='http://127.0.0.1:11434',context_window=window,
               models=[{'id':'same-model','capabilities':['chat','agent']}])
    return row


def test_overflow_preview_and_send_have_same_actionable_budget(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    row = connection(); row['context_window'] = 1024
    save(api._connections, payload([row])); select(api._connections, 'official')
    monkeypatch.setattr(api._connections, 'build_model', lambda *args:pytest.fail('overflow called model'))
    question = '保留严格限定条件、数字与单位。' * 300
    before = api.conversation_snapshot()['result']
    disk = Path(before['location']).read_bytes() if Path(before['location']).exists() else None
    preview, response = api.preview_chat_context(question), api.ask(question)
    assert preview['reason'] == response['reason'] == 'model_context_budget'
    assert preview['request_started'] is response['request_started'] is False
    assert preview['budget'] == response['budget']
    budget = response['budget']
    assert budget['context_window'] == 1024 and budget['reply_token_reserve'] == 512
    assert budget['estimated_excess_tokens'] == budget['estimated_input_tokens'] + 512 - 1024 > 0
    assert budget['ollama_num_ctx'] is None
    assert budget['route']['connection_id'] == 'official'
    assert 'fixture-private-' not in json.dumps(response)
    assert api.conversation_snapshot()['result']['sessions'] == before['sessions']
    assert (Path(before['location']).read_bytes() if Path(before['location']).exists() else None) == disk
    assert not app.busy and api.get_model_connections()['catalog']['current']['chat']['valid']


@pytest.mark.parametrize('window',[1024,2048,8192])
@pytest.mark.parametrize('stream',[False,True])
def test_actual_ollama_sdk_carries_configured_chat_window(tmp_path, monkeypatch, window, stream):
    app, api, _ = setup(tmp_path)
    save(api._connections, payload([local_connection(window=window)])); select(api._connections,'official')
    module = importlib.import_module('ollama._client')
    original, requests, clients = module.httpx.Client, [], []
    def handle(request):
        requests.append(json.loads(request.content))
        body = {'model':'same-model','created_at':'2026-10-27T00:00:00Z',
                'message':{'role':'assistant','content':'窗口一致'},'done':True,'done_reason':'stop'}
        return module.httpx.Response(200,content=(json.dumps(body)+'\n').encode())
    def client(*args,**kwargs):
        value = original(*args,**kwargs,transport=module.httpx.MockTransport(handle))
        clients.append(value)
        return value
    monkeypatch.setattr(module.httpx,'Client',client)
    preview = api.preview_chat_context('短问题')['result']
    result = api.ask('短问题',{'stream':stream,'context_revision':preview['context_revision']})
    assert result['ok']; finish(app)
    assert len(requests) == 1 and clients[0].is_closed
    assert requests[0]['options']['num_ctx'] == preview['ollama_num_ctx'] == window
    assert requests[0]['options']['num_predict'] == result['budget']['reply_token_reserve'] == 512
    assert result['budget']['context_window'] == window
    assert preview['runtime_verified'] is False


def test_agent_ollama_runtime_window_semantics_are_unchanged(tmp_path):
    _, api, _ = setup(tmp_path)
    save(api._connections,payload([local_connection(window=2048)]))
    route = select(api._connections,'official','agent')
    model = api._connections.build_model(route)
    try:
        assert model.num_ctx is None
    finally:
        model._client._client.close()


def test_captured_chat_budget_survives_next_selection(tmp_path):
    _, api, _ = setup(tmp_path)
    save(api._connections,payload([local_connection(window=1024),local_connection('custom',2048)]))
    route = select(api._connections,'official')
    select(api._connections,'custom')
    model = api._connections.build_model(route)
    try:
        assert model.num_ctx == 1024
        assert api.preview_chat_context('next')['result']['ollama_num_ctx'] == 2048
    finally:
        model._client._client.close()


def test_success_budget_does_not_migrate_v1_history(tmp_path,monkeypatch):
    app, api, _ = setup(tmp_path,True)
    monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='ok',response_metadata={'finish_reason':'stop'})))
    preview = api.preview_chat_context('short')['result']
    result = api.ask('short',{'stream':False})
    assert result['ok']; finish(app)
    assert result['budget']['estimated_input_tokens'] == preview['estimated_input_tokens']
    assert result['budget']['ollama_num_ctx'] is None and result['budget']['runtime_verified'] is False
    book = api.conversation_snapshot()['result']
    stored = json.loads(Path(book['location']).read_text(encoding='utf8'))
    assert stored['version'] == 1
    assert set(stored['sessions'][0]['turns'][0]) == {'id','question','answer','route','status','created','context_turns','omitted_turns','finish_reason'}


def test_openai_sdk_does_not_receive_ollama_runtime_options(tmp_path,monkeypatch):
    app,api,_ = setup(tmp_path)
    calls = []
    def handle(request):
        calls.append(json.loads(request.content))
        return httpx2.Response(200,json={'id':'fixture','object':'chat.completion','created':1,'model':'same-model',
            'choices':[{'index':0,'message':{'role':'assistant','content':'ok'},'finish_reason':'stop'}]})
    monkeypatch.setattr('office_agent.webui.model_connections._http_client',lambda:httpx2.Client(transport=httpx2.MockTransport(handle),follow_redirects=False,trust_env=False))
    assert api.ask('short',{'stream':False})['ok']; finish(app)
    assert len(calls)==1 and 'num_ctx' not in calls[0] and 'options' not in calls[0]
