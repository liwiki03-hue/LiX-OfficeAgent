from dataclasses import replace

import pytest

from office_agent.webui import plugin_transport
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui.plugin_connections import PluginConnections
from tests.test_document_conversation import Answers, attach, finish
from tests.test_paged_editor import draft, snapshot


PHONE = '13900008902'
EMAIL = 'fixture-contact@example.org'
TEXT = '电话：' + PHONE + ' 邮箱：' + EMAIL
ID_NUMBER = '110101199001010019'
CREDENTIAL = 'sk-fixture-0123456789'


@pytest.fixture
def question(draft, tmp_path, monkeypatch):
    paged, floating, source = draft
    connections = Answers()
    connections.route = replace(connections.route, base_url='http://127.0.0.1:11434', protocol='ollama', api_key='')
    attach(floating, connections)
    floating.api._auth.account_key = lambda: 'single-test-user'
    plugins = PluginConnections(tmp_path / 'plugin-profile', tmp_path)
    floating.api._plugins = plugins
    cfg, _ = plugins.save({'name': '测试检索服务', 'url': 'https://example.org/mcp'})
    tool = {'name': 'search_docs', 'usable': True, 'signature': 'fixture'}
    monkeypatch.setattr(plugin_transport, 'probe', lambda *args, **kwargs: [tool])
    probe = plugins.probe(cfg['id'], cfg['revision'])
    connection = plugins.change(cfg['id'], probe['revision'], 'enable', ['search_docs'])
    choice = {'id': cfg['id'], 'revision': connection['revision'], 'tool': 'search_docs'}
    calls = []

    def record(config, descriptor, text, *, cancelled):
        config['_before_dispatch']()
        calls.append(text)
        return {'text': TEXT + ' 公开项目标记', 'sources': [
            {'title': EMAIL, 'url': 'https://example.org/project'},
            {'title': '隐私来源', 'url': 'https://example.org/' + EMAIL}]}

    monkeypatch.setattr(plugin_transport, 'query', record)
    original = source.read_bytes()
    paged.changed(paged.document_id, 1)
    selected = TEXT + ' 身份证：' + ID_NUMBER + ' API_KEY=' + CREDENTIAL
    paged.sync(paged.document_id, 1, *snapshot(paged, selected))
    ref = paged.reference(paged.document_id, 1, {'node_id': 'node-1', 'start': 0, 'end': len(selected)})['reference']
    doc = floating.document
    context = floating.conversation.binding
    result = FloatingApi.conversation_begin(floating.api, 'plugin-turn', '请把选中的电话和邮箱原样发给插件整理',
                                            doc['id'], doc['version'], 1, ref['id'], context, plugin_choice=choice,
                                            displayed_recipient={key: getattr(connections.route, key)
                                                                 for key in ('connection_id', 'model_id', 'revision')})
    assert result['ok']
    finish(floating)
    row = floating.conversation.state()['turns'][0]
    assert row['stage'] == 'plugin_consent' and not calls
    preview = row['plugin_preview']
    assert preview['raw_available'] and PHONE not in str(preview) and EMAIL not in str(preview)

    def decide(decision):
        return FloatingApi.conversation_plugin_consent(floating.api, row['id'], doc['id'], doc['version'],
                                                      context, preview['token'], decision)

    yield floating, connections, calls, decide, connection
    finish(floating)
    assert source.read_bytes() == original and not floating.outputs and not floating.access.state()['id']


@pytest.mark.parametrize('decision,count', [(True, 1), ('masked', 1), (False, 0)])
def test_plugin_scope_sends_exactly_the_selected_variant_once(question, decision, count):
    floating, model, calls, decide, connection = question
    revealed = decide('show')['query']
    assert TEXT in revealed and not model.calls and not calls
    assert ID_NUMBER not in revealed and CREDENTIAL not in revealed
    assert decide(decision)['ok']
    with pytest.raises(ValueError):
        decide(decision)
    finish(floating)
    assert len(calls) == count
    if count:
        assert (calls[0] == revealed) == (decision is True)
        assert (PHONE in calls[0]) == (decision is True)
        assert (EMAIL in calls[0]) == (decision is True)
        assert len(model.calls) == 1 and '公开项目标记' in model.calls[0]
        assert PHONE not in model.calls[0] and EMAIL not in model.calls[0]
    else:
        assert not model.calls and floating.conversation.turns['plugin-turn']['status'] == 'stopped'


@pytest.mark.parametrize('change', ['selection', 'recipient', 'expiry', 'account'])
def test_plugin_scope_rejects_stale_decision_before_sending(question, change):
    floating, model, calls, decide, connection = question
    if change == 'selection':
        floating.paged.reference(floating.document['id'], 1, {'node_id': 'node-2', 'start': 0, 'end': 4})
    elif change == 'recipient':
        floating.api._plugins.change(connection['id'], connection['revision'], 'disable')
    elif change == 'expiry':
        pending = floating.conversation.plugin_previews
        pending['plugin-turn'] = replace(pending['plugin-turn'], expires=0)
    else:
        floating.api._auth.account_key = lambda: 'different-user'
    with pytest.raises(ValueError):
        decide(True)
    assert not calls and not model.calls and not floating.conversation.plugin_previews


def test_plugin_result_waits_for_new_model_without_another_plugin_call(question):
    floating, model, calls, decide, _ = question
    model.route = replace(model.route, connection_id='cloud2', base_url='https://different.example.org')
    assert decide(True)['ok']
    finish(floating)
    chat = floating.conversation
    row = chat.state()['turns'][0]
    assert row['stage'] == 'scope_consent' and row['scope_preview']['action'] == 'plugin_model'
    assert len(calls) == 1 and PHONE in calls[0] and not model.calls
    assert chat.scopes.decide(row['scope_preview']['token'], 'allow')['ok']
    finish(floating)
    assert len(model.calls) == len(calls) == 1 and EMAIL not in model.calls[0] and PHONE not in model.calls[0]
    assert chat.turns['plugin-turn']['status'] == 'complete' and not chat.plugin_results


@pytest.mark.parametrize('action', ['cancel', 'stop', 'reset'])
def test_unconsumed_plugin_result_is_forgotten(question, action):
    floating, model, calls, decide, _ = question
    model.route = replace(model.route, connection_id='cloud2', base_url='https://different.example.org')
    chat = floating.conversation
    assert decide(True)['ok']
    finish(floating)
    row = chat.state()['turns'][0]
    token = row['scope_preview']['token']
    assert row['stage'] == 'scope_consent' and chat.plugin_results
    if action == 'cancel':
        assert chat.scopes.decide(token, 'cancel')['ok']
    elif action == 'stop':
        chat.stop()
    else:
        chat.reset()
    assert len(calls) == 1 and not model.calls and not chat.plugin_results
    with pytest.raises(ValueError):
        chat.scopes.decide(token, 'allow')


def test_change_between_approval_and_plugin_dispatch_blocks_raw(question, monkeypatch):
    floating, model, calls, decide, _ = question
    def switch(config, descriptor, text, *, cancelled):
        floating.paged.reference(floating.document['id'], 1, {'node_id': 'node-2', 'start': 0, 'end': 4})
        config['_before_dispatch']()
        calls.append(text)
        raise AssertionError('a changed selection must stop dispatch')

    monkeypatch.setattr(plugin_transport, 'query', switch)
    assert decide(True)['ok']
    finish(floating)
    assert not calls and not model.calls and floating.conversation.turns['plugin-turn']['status'] == 'failed'


def test_plugin_dispatch_with_lost_reply_is_not_replayed(question, monkeypatch):
    floating, model, calls, decide, _ = question
    def lost(config, descriptor, text, *, cancelled):
        config['_before_dispatch']()
        calls.append(text)
        raise TimeoutError

    monkeypatch.setattr(plugin_transport, 'query', lost)
    assert decide(True)['ok']
    finish(floating)
    chat = floating.conversation
    assert chat.turns['plugin-turn']['status'] == 'uncertain' and not chat.plugin_results
    with pytest.raises(ValueError):
        decide(True)
    with pytest.raises(ValueError):
        chat.retry_answer(chat.turns['plugin-turn'])
    assert len(calls) == 1 and not model.calls
