import json
from types import SimpleNamespace

import httpx2
import pytest

from office_agent.chat_policy import output_parameter
from office_agent.settings_store import SettingsError
from tests.test_chat_context import sse
from tests.test_daily_flows import setup, finish
from tests.test_model_connections import connection, payload, save, select


@pytest.mark.parametrize('url,choice,expected', [
    ('https://api.deepseek.com/v1', 'auto', 'max_tokens'),
    ('https://api.deepseek.com/', 'auto', 'max_tokens'),
    ('https://API.DEEPSEEK.COM:443/v1/', 'auto', 'max_tokens'),
    ('https://api.deepseek.com/beta', 'auto', 'max_completion_tokens'),
    ('https://api.deepseek.com:444/v1', 'auto', 'max_completion_tokens'),
    ('https://api.deepseek.com.example/v1', 'auto', 'max_completion_tokens'),
    ('https://custom.example/v1', 'auto', 'max_completion_tokens'),
    ('https://api.deepseek.com/v1', 'max_completion_tokens', 'max_completion_tokens'),
    ('https://custom.example/v1', 'max_tokens', 'max_tokens'),
])
def test_output_parameter_uses_exact_endpoint_and_explicit_override(url, choice, expected):
    route = SimpleNamespace(protocol='openai', base_url=url, output_parameter=choice)
    assert output_parameter(route) == expected
    route.output_parameter = 'unsupported'
    with pytest.raises(SettingsError, match='chat_output_parameter_unsupported'):
        output_parameter(route)


def test_official_auto_preview_wire_and_linked_continuation(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path)
    row = connection()
    row['base_url'] = 'https://api.deepseek.com/v1'
    save(api._connections, payload([row])); select(api._connections, 'official')
    requests = []
    def handle(request):
        requests.append(json.loads(request.content))
        return httpx2.Response(200, headers={'content-type':'text/event-stream'},
                               content=sse('有界回答片段', 'length') + b'data: [DONE]\n\n')
    monkeypatch.setattr('office_agent.webui.model_connections._http_client',
                        lambda:httpx2.Client(transport=httpx2.MockTransport(handle), follow_redirects=False, trust_env=False))
    preview = api.preview_chat_context('合成原问题')['result']
    assert preview['reply_token_reserve'] == 512 and preview['output_parameter'] == 'max_tokens'
    first = api.ask('合成原问题', {'stream':True, 'context_revision':preview['context_revision']})
    assert first['ok'] and first['turn_id']
    finish(app)
    state = api.conversation_snapshot()['result']
    turn = state['sessions'][0]['turns'][0]
    assert turn['id'] == first['turn_id'] and turn['finish_reason'] == 'length'
    assert 'budget' not in turn
    binding = {'session_id':state['current'], **state['continuations'][0]}
    continuation = api.preview_chat_context('继续', {'continuation':binding})['result']
    assert continuation['reply_token_reserve'] == 512 and continuation['output_parameter'] == 'max_tokens'
    assert api.ask('继续', {'stream':True, 'continuation':binding, 'context_revision':continuation['context_revision']})['ok']
    finish(app)
    assert len(requests) == 2
    assert all(request['max_tokens'] == 512 and 'max_completion_tokens' not in request for request in requests)
    assert [message['content'] for message in requests[1]['messages']] == ['合成原问题','有界回答片段','继续']
    assert api.conversation_snapshot()['result']['sessions'][0]['turns'][1]['continuation_of'] == binding
    assert app.path is None and app.pending_id is None and not app.busy
