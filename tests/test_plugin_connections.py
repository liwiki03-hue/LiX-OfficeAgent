from threading import Event, Thread

import pytest

from office_agent.webui import plugin_catalog, plugin_transport
from office_agent.webui.plugin_connections import LEARN_URL, PluginConnections
from office_agent.webui.plugin_turns import public_keywords, safe_query


def test_manual_and_catalog_share_a_user_connection_and_revoke_old_grants(tmp_path, monkeypatch):
    store = PluginConnections(tmp_path / 'profile', tmp_path)
    official = dict(source='com.microsoft/microsoft-learn-mcp', url=LEARN_URL, publisher='com.microsoft')
    names = []

    def fetch(term, cursor):
        names.append(term)
        return {'items': [official] if term.lower() == 'microsoft-learn' else [], 'cursor': ''}

    monkeypatch.setattr(plugin_catalog, 'fetch_page', fetch)
    assert store.search('Microsoft Learn')['items'] == [official]
    assert store.search('Microsoft Learn')['items'] == [official]
    assert names == ['Microsoft Learn', 'microsoft-learn']
    config = {'name': 'Microsoft Learn', 'url': LEARN_URL}
    row, _ = store.save(dict(config, source=official['source']))
    manual, duplicate = store.save(config)
    assert duplicate and manual['id'] == row['id'] and not row['enabled']
    with pytest.raises(ValueError):
        store.save(dict(config, source='com.other/microsoft-learn-mcp'))
    tool = dict(name='microsoft_docs_search', usable=True, signature='sig1', query_key='query')
    monkeypatch.setattr(plugin_transport, 'probe', lambda *a, **kw: [tool.copy()])
    active = store.probe(row['id'], row['revision'])
    active = store.change(row['id'], active['revision'], 'enable', [tool['name']])
    question = '查微软官方文档，Excel 的 XLOOKUP 怎么写？'
    ticket = store.resolve(question, None)
    assert public_keywords(question, ticket) == 'Excel XLOOKUP'
    assert public_keywords(question + ' 文档内的专有名词', ticket) is None
    assert store.resolve('不要查微软官方文档，只看选区。', None) is None
    tool['signature'] = 'sig2'
    changed = store.probe(active['id'], active['revision'])
    assert not changed['enabled'] and not changed['allowed']
    with pytest.raises(ValueError):
        store.checked(ticket)
    assert PluginConnections(store.profile, tmp_path).list()[0]['revision'] == changed['revision']


def test_stop_and_disable_discard_inflight_result(tmp_path, monkeypatch):
    store = PluginConnections(tmp_path / 'profile', tmp_path)
    config, _ = store.save({'name':'test', 'url':'https://example.org/mcp'})
    tool = dict(name='search_docs', usable=True, signature='stable')
    monkeypatch.setattr(plugin_transport, 'probe', lambda *a, **kw: [tool])
    row = store.probe(config['id'], 1)
    row = store.change(row['id'], row['revision'], 'enable', ['search_docs'])
    ticket = store.resolve('', {'id':row['id'], 'tool':'search_docs', 'revision':row['revision']})
    started, release = Event(), Event()
    errors = []

    def deferred(config, descriptor, text, *, cancelled):
        started.set()
        assert release.wait(4) and cancelled()
        return {'text':'late', 'sources':[]}

    def ask():
        try:
            store.query(ticket, 'public words', lambda:False)
        except ValueError as error:
            errors.append(str(error))

    monkeypatch.setattr(plugin_transport, 'query', deferred)
    worker = Thread(target=ask)
    worker.start()
    try:
        assert started.wait(4)
        store.change(row['id'], row['revision'], 'disable')
    finally:
        release.set()
        worker.join(5)
    assert errors and not worker.is_alive() and not store.operations


def test_private_question_requires_preview_and_blocks_secrets():
    ticket = {'destination':'https://example.com/mcp', 'tool':'search_docs'}
    assert public_keywords('查微软官方文档，Excel 的 XLOOKUP 怎么写？', ticket) is None
    assert safe_query('查手册\n用户选中的一句话') == '查手册\n用户选中的一句话'
    for value in [r'C:\private\contract.docx', 'Bearer abcde0123456789', 'secret=private-fixture']:
        with pytest.raises(ValueError):
            safe_query(value)


def test_long_current_results_fit_the_selected_model_and_keep_only_seen_sources(tmp_path):
    from office_agent.agent import OfficeAgent
    from office_agent.webui.plugin_turns import evidence_state
    import json

    agent = OfficeAgent(object(), lambda *args: None, [], tmp_path, persist=False, decision_limit=1,
                        context_tokens=4096, reply_tokens=512)
    sources = [{'title': '资料 ' + str(index), 'url': f'https://example.org/source/{index}'} for index in range(8)]
    sections = [(source['title'] + ' 当前返回的用法。' * 230 + '\n' + source['url']) for source in sources]
    state, shown = evidence_state(agent, '怎么写？', 'Excel XLOOKUP', '\n\n'.join(sections), sources)
    sent = json.loads(state.required.goal)
    assert shown and shown == sent['sources'] and agent.estimate(state, []) < 4096 * .9
    assert all(source in sources and source['title'] in str(sent['external_evidence']) for source in shown)

