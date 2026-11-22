import json
import socket

import httpx
import pytest

from office_agent.webui import openalex_search as search
from office_agent.webui import plugin_transport as transport


SEARCH = {
    'name': 'search_works',
    'annotations': {'title': 'Search works', 'readOnlyHint': True, 'destructiveHint': False,
                    'idempotentHint': True, 'openWorldHint': True},
    'inputSchema': {'type': 'object', 'properties': {
        'query': {'type': 'string', 'maxLength': 2000},
        'limit': {'type': 'integer', 'minimum': 1, 'maximum': 50},
        'include_abstracts': {'type': 'boolean'}, 'oql': {'type': 'string', 'maxLength': 20000},
    }, 'additionalProperties': False},
}


def test_word_public_topic_previews_only_approved_fields():
    for question in ('查找有关 BM25 的研究，列两篇，附作者、年份和可打开的出处', '找两篇关于 BM25 的论文'):
        assert search.arguments(search.preview(question)) == {'query': 'BM25', 'limit': 2, 'include_abstracts': False}
    for query in ('找两篇关于选区的研究', '找12篇关于 BM25 的研究'):
        with pytest.raises(transport.PluginError):
            search.preview(query)
    with pytest.raises(transport.PluginError):
        search.arguments('{"query":"BM25","limit":2,"include_abstracts":false,"oql":"works"}')


def test_named_service_route_uses_an_enabled_connection(tmp_path, monkeypatch):
    from office_agent.webui.plugin_connections import PluginConnections

    store = PluginConnections(tmp_path / 'profile', tmp_path)
    row, _ = store.save({'name': 'OpenAlex', 'transport': 'http', 'url': search.URL, 'auth': 'oauth'})
    monkeypatch.setattr(transport, 'probe', lambda config, **kwargs: [transport._descriptor(SEARCH, config)])
    row = store.probe(row['id'], row['revision'])
    with pytest.raises(ValueError, match='启用'):
        store.resolve('用 OpenAlex 查两篇关于 BM25 的论文', None)
    enabled = store.change(row['id'], row['revision'], 'enable', [search.TOOL])
    ticket = store.resolve('用 OpenAlex 查两篇关于 BM25 的论文', None)
    assert ticket['id'] == enabled['id'] and ticket['tool'] == search.TOOL
    assert store.resolve('别用 OpenAlex 查，我只读当前段。', None) is None


def test_streamable_work_queries_validate_the_enabled_contract(monkeypatch):
    calls, searches = [], []
    listed = [SEARCH, dict(SEARCH, name='claim_author_profile')]
    records = [{'id': 'W123', 'title': 'Fixture BM25 A', 'year': 2009, 'authors': ['A. Author (Lab)'],
                'doi': '10.1234/synthetic-a', 'openalex_url': 'https://openalex.org/W123'},
               {'id': 'W124', 'title': 'Fixture BM25 B', 'year': 2023, 'authors': ['B. Researcher']},
               {'id': 'W125', 'title': 'Third result', 'year': 2024, 'authors': []}]
    status = 200

    async def service(request):
        assert request.headers['Host'] == 'mcp.openalex.org'
        if request.method != 'POST':
            return httpx.Response(405)
        assert request.headers.get('Authorization') == 'Bearer fixture-test'
        message = json.loads(request.content)
        calls.append(message['method'])
        if 'id' not in message:
            return httpx.Response(202)
        if message['method'] == 'initialize':
            result = {'protocolVersion': message['params']['protocolVersion'], 'capabilities': {'tools': {}},
                      'serverInfo': {'name': 'OpenAlex protocol fixture', 'version': 'controlled'}}
        elif message['method'] == 'tools/list':
            result = {'tools': listed}
        else:
            assert message['method'] == 'tools/call' and message['params']['name'] == search.TOOL
            searches.append(message['params']['arguments'])
            if status != 200:
                return httpx.Response(status)
            result = {'content': [{'type': 'text', 'text': json.dumps({'results': records})}]}
        return httpx.Response(200, json={'jsonrpc': '2.0', 'id': message['id'], 'result': result})

    real_dns = socket.getaddrinfo
    monkeypatch.setattr(socket, 'getaddrinfo', lambda host, *args, **kwargs:
                        real_dns('93.184.216.34' if host == 'mcp.openalex.org' else host, *args, **kwargs))
    monkeypatch.setattr(httpx, 'AsyncHTTPTransport', lambda **kwargs: httpx.MockTransport(service))
    monkeypatch.setattr(transport.plugin_oauth, 'token', lambda config: 'fixture-test')
    config = {'url': search.URL, 'transport': 'http', 'auth': 'oauth', 'id': 'openalex-fixture'}
    enabled, denied = transport.probe(config)
    assert enabled['usable'] and not denied['usable'] and searches == []
    assert enabled['server'] == {'name': 'OpenAlex protocol fixture', 'version': 'controlled'}
    assert not transport._descriptor(SEARCH, dict(config, url='https://example.org/mcp'))['usable']
    requested = search.preview('找两篇关于 BM25 的研究')
    result = transport.query(config, enabled, requested)
    assert searches == [{'query': 'BM25', 'limit': 2, 'include_abstracts': False}]
    assert [source['url'] for source in result['sources']] == ['https://openalex.org/W123', 'https://openalex.org/W124']
    assert all(item in result['text'] for item in ('Fixture BM25 A', 'A. Author', '2009', '10.1234/synthetic-a', 'B. Researcher', '2023'))
    assert 'Third result' not in result['text']
    records.clear()
    result = transport.query(config, enabled, requested)
    assert not result['sources'] and '未查到' in result['text']
    status = 429
    with pytest.raises(search.SearchRejected, match='请求受限'):
        transport.query(config, enabled, requested)
    listed[0] = dict(SEARCH, annotations={**SEARCH['annotations'], 'readOnlyHint': False})
    with pytest.raises(transport.PluginError, match='结构'):
        transport.query(config, enabled, requested)
    assert calls.count('tools/list') == 5 and len(searches) == 3
