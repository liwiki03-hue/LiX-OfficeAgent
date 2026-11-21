import base64
import hashlib
import json
import os
from pathlib import Path
import sys
from time import sleep
from urllib.parse import parse_qs, urlencode, urlsplit


TOOL = {
    'name': 'search_docs',
    'annotations': {'readOnlyHint': True, 'destructiveHint': False, 'openWorldHint': False},
    'inputSchema': {'type': 'object', 'properties': {'query': {'type': 'string', 'maxLength': 150}},
                    'required': ['query'], 'additionalProperties': False},
}


def rpc(request, tool=TOOL):
    method = request['method']
    if method == 'initialize':
        return {'protocolVersion': request['params']['protocolVersion'], 'capabilities': {'tools': {}},
                'serverInfo': {'name': 'fixture', 'version': '1'}}
    if method == 'tools/list':
        return {'tools': [tool, dict(tool, name='write_file')]}
    if method == 'tools/call':
        query = request['params']['arguments']['query']
        return {'content': [{'type': 'text', 'text': f'{query} https://example.com/docs/current'}]}
    raise AssertionError(method)


class Remote:
    host = 'https://127.0.0.1:24682'
    token = 'synthetic-transport-test-token'

    def __init__(self, monkeypatch):
        import httpx
        self.calls = []
        self.tool = TOOL
        self.require_auth = True
        self.challenge = self.code = ''
        self.deny = False
        self.revoked = False
        self.redirect = False
        monkeypatch.setattr(httpx, 'AsyncHTTPTransport', lambda **_: httpx.MockTransport(self.handle))

    def handle(self, request):
        import httpx
        path = request.url.path
        if self.redirect:
            return httpx.Response(302, headers={'location': 'https://example.com/elsewhere'})
        if path.startswith('/.well-known/oauth-protected-resource'):
            return httpx.Response(200, json={'resource': self.host + '/mcp', 'authorization_servers': [self.host]})
        if path == '/.well-known/oauth-authorization-server':
            return httpx.Response(200, json={'issuer': self.host, 'authorization_endpoint': self.host + '/authorize',
                'registration_endpoint': self.host + '/register', 'token_endpoint': self.host + '/token',
                'response_types_supported': ['code'], 'token_endpoint_auth_methods_supported': ['none'],
                'code_challenge_methods_supported': ['S256']})
        if path == '/register':
            return httpx.Response(201, json={**json.loads(request.content), 'client_id': 'lix-controlled'})
        if path == '/token':
            fields = parse_qs(request.content.decode())
            digest = hashlib.sha256(fields['code_verifier'][0].encode()).digest()
            assert base64.urlsafe_b64encode(digest).decode().rstrip('=') == self.challenge
            assert fields['code'] == [self.code]
            return httpx.Response(200, json={'access_token': self.token, 'token_type': 'bearer', 'expires_in': 240})
        assert path == '/mcp'
        if self.require_auth and (self.revoked or request.headers.get('Authorization') != 'Bearer ' + self.token):
            return httpx.Response(401, headers={'WWW-Authenticate': f'Bearer resource_metadata="{self.host}/.well-known/oauth-protected-resource"'})
        envelope = json.loads(request.content)
        self.calls.append(envelope)
        if 'id' not in envelope:
            return httpx.Response(202)
        result = rpc(envelope, self.tool)
        if envelope['method'] == 'tools/call':
            result['content'][0]['text'] += ' ' + self.token
        return httpx.Response(200, json={'jsonrpc': '2.0', 'id': envelope['id'], 'result': result})

    def open_browser(self, url):
        from urllib.error import HTTPError
        from urllib.request import urlopen
        import pytest
        fields = parse_qs(urlsplit(url).query)
        assert urlsplit(url).path == '/authorize'
        redirect = fields['redirect_uri'][0]
        assert urlsplit(redirect).hostname == '127.0.0.1'
        assert fields['code_challenge_method'] == ['S256']
        self.challenge = fields['code_challenge'][0]
        self.code = 'fixture-login-code'
        with pytest.raises(HTTPError) as denied:
            urlopen(redirect + '?' + urlencode({'code': self.code, 'state': 'wrong-state'}), timeout=2)
        assert denied.value.code == 400
        response = {'error': 'access_denied'} if self.deny else {'code': self.code}
        with urlopen(redirect + '?' + urlencode(dict(response, state=fields['state'][0])), timeout=2) as reply:
            assert reply.status == 200
        return True


def test_learn_contract_only_approves_real_query_tool():
    from mcp.types import CallToolResult, TextContent
    from office_agent.webui import plugin_transport as tr
    config = {'url': tr.LEARN, 'transport': 'http'}
    raw = {'name': 'microsoft_docs_search', 'annotations': {'readOnlyHint': True, 'destructiveHint': False},
           'inputSchema': {'type': 'object', 'properties': {'query': {'type': 'string', 'default': None}}}}
    descriptor = tr._descriptor(raw, config)
    assert descriptor['usable'] and descriptor['query_key'] == 'query' and descriptor['minimum'] == 1
    assert not tr._descriptor(raw, dict(config, url='https://example.com/api/mcp'))['usable']
    assert not tr._descriptor(dict(raw, name='microsoft_docs_fetch'), config)['usable']
    assert not tr._descriptor(dict(raw, annotations={'destructiveHint': True}), config)['usable']
    docs = [{'title': f'Excel {i}', 'content': 'XLOOKUP ' * 900, 'contentUrl': f'https://learn.microsoft.com/office/{i}'} for i in range(10)]
    response = CallToolResult(content=[TextContent(type='text', text=json.dumps(docs))])
    result = tr._output(response, set(), True)
    assert result['sources'][0]['title'] == 'Excel 0' and len(result['sources']) >= 2
    assert len(result['text'].encode()) < 20000 and all(source['url'] in result['text'] for source in result['sources'])


def test_http_uses_actual_schema_rechecks_permission_and_redacts(monkeypatch):
    import pytest
    from office_agent.webui import plugin_transport as tr
    remote = Remote(monkeypatch)
    config = dict(transport='http', auth='bearer', url=remote.host + '/mcp', private=True,
                  credential_ref='LIX_PLUGIN_SEARCH', _resolve_secret=lambda ref: remote.token)
    import asyncio
    async def browser_probe():
        return tr.probe(config)
    tool, unsafe = asyncio.run(browser_probe())
    assert tool['usable'] and not unsafe['usable']
    assert 'tools/call' not in [entry['method'] for entry in remote.calls]
    result = tr.query(config, tool, 'XLOOKUP')
    assert remote.token not in str(result) and 'XLOOKUP' in result['text']
    assert result['sources'][0]['url'] == 'https://example.com/docs/current'
    calls = [entry for entry in remote.calls if entry['method'] == 'tools/call']
    assert len(calls) == 1 and calls[0]['params']['arguments'] == {'query': 'XLOOKUP'}
    def revoked():
        raise ValueError('stale approval')
    with pytest.raises(tr.PluginError, match='许可已改变'):
        tr.query(dict(config, _before_dispatch=revoked), tool, 'XLOOKUP')
    remote.tool = dict(TOOL, annotations=dict(TOOL['annotations'], readOnlyHint=False))
    with pytest.raises(tr.PluginError, match='结构'):
        tr.query(config, tool, 'XLOOKUP')
    assert [entry for entry in remote.calls if entry['method'] == 'tools/call'] == calls
    remote.redirect = True
    with pytest.raises(tr.PluginError, match='重定向'):
        tr.probe(config)
    with pytest.raises(tr.PluginError, match='公网'):
        tr.probe(dict(config, source='example/catalog'))


def test_verified_reply_survives_slow_delete_and_remains_cancellable(monkeypatch):
    import asyncio
    import httpx
    import pytest
    from office_agent.webui import plugin_transport as tr
    requests = []
    stopping = False
    deletions = []
    async def service(request):
        if request.method == 'DELETE':
            assert request.headers['Mcp-Session-Id'] == 'fixture-session'
            assert request.headers['MCP-Protocol-Version']
            deletions.append('start')
            try:
                await asyncio.sleep(10)
            finally:
                deletions.append('closed')
        if request.method == 'GET':
            return httpx.Response(405)
        envelope = json.loads(request.content)
        requests.append(envelope['method'])
        if 'id' not in envelope:
            return httpx.Response(202)
        headers = {'Mcp-Session-Id': 'fixture-session'} if envelope['method'] == 'initialize' else {}
        return httpx.Response(200, headers=headers, json={'jsonrpc': '2.0', 'id': envelope['id'], 'result': rpc(envelope)})
    monkeypatch.setattr(httpx, 'AsyncHTTPTransport', lambda **_: httpx.MockTransport(service))
    monkeypatch.setattr(tr, '_REQUEST_SECONDS', .22)
    monkeypatch.setattr(tr, '_CLOSE_SECONDS', .8)
    config = dict(transport='http', auth='none', url='https://127.0.0.1:24701/mcp', private=True)
    tool = tr._descriptor(TOOL, config)
    def cancelled():
        return stopping and deletions.count('start') == 2
    answer = tr.query(config, tool, 'XLOOKUP', cancelled=cancelled)
    assert 'XLOOKUP' in answer['text'] and answer['sources'][0]['url'] == 'https://example.com/docs/current'
    assert requests.count('tools/call') == 1 and deletions == ['start', 'closed']
    stopping = True
    with pytest.raises(tr.PluginError, match='取消'):
        tr.query(config, tool, 'XLOOKUP', cancelled=cancelled)
    assert requests.count('tools/call') == 2 and deletions == ['start', 'closed'] * 2


def test_oauth_sdk_pkce_state_scope_expiry_cancel_disconnect(tmp_path, monkeypatch):
    import pytest
    if os.name != 'nt':
        pytest.skip('Windows user credential store')
    from office_agent.webui import plugin_oauth, plugin_transport as tr
    remote = Remote(monkeypatch)
    monkeypatch.setattr(plugin_oauth.webbrowser, 'open', remote.open_browser)
    config = dict(id='plugin_a', transport='http', auth='oauth', url=remote.host + '/mcp', private=True, _profile_dir=tmp_path)
    remote.deny = True
    with pytest.raises(tr.PluginError, match='取消'):
        tr.login(config)
    assert not tr.credential_available(config) and not list(tmp_path.rglob('*.bin'))
    remote.deny = False
    tools = tr.login(config)
    assert tools[0]['usable'] and tr.credential_available(config)
    assert remote.token.encode() not in next(tmp_path.rglob('*.bin')).read_bytes()
    assert not tr.credential_available(dict(config, id='plugin_b'))
    assert not tr.credential_available(dict(config, url=remote.host + '/other'))
    assert not tr.credential_available(dict(config, _profile_dir=tmp_path / 'different_user'))
    assert 'XLOOKUP' in tr.query(config, tools[0], 'XLOOKUP')['text']
    remote.revoked = True
    with pytest.raises(tr.PluginError, match='重新登录'):
        tr.query(config, tools[0], 'XLOOKUP')
    assert not tr.credential_available(config)
    remote.revoked = False
    assert tr.login(config)[0]['usable']
    with monkeypatch.context() as clock:
        from time import time
        clock.setattr(plugin_oauth.time, 'time', lambda: time() + 300)
        assert not tr.credential_available(config)
        with pytest.raises(tr.PluginError, match='登录'):
            tr.probe(config)
    assert tr.login(config)[0]['usable']
    tr.disconnect(config)
    assert not tr.credential_available(config) and not list(tmp_path.rglob('*.bin'))


def test_real_stdio_cancels_only_owned_process(tmp_path, monkeypatch):
    import threading
    import pytest
    if os.name != 'nt':
        pytest.skip('Windows job object')
    import win32api
    import win32event
    from pywintypes import error as WindowsError
    from office_agent.webui import plugin_transport as tr
    pid_file = tmp_path / 'owned.pid'
    def exited():
        try:
            process = win32api.OpenProcess(0x00100000, False, int(pid_file.read_text()))
        except WindowsError as error:
            if error.winerror != 87:
                raise
            return True
        try:
            return win32event.WaitForSingleObject(process, 1500) == win32event.WAIT_OBJECT_0
        finally:
            process.Close()
    monkeypatch.setenv('UNRELATED_SECRET_MARKER', 'must-not-propagate')
    config = dict(transport='stdio', auth='none', program=sys.executable,
                  args=[str(Path(__file__).resolve()), str(pid_file)], env_refs={'LOCAL_TOKEN': 'LIX_PLUGIN_LOCAL'},
                  _resolve_secret=lambda ref: 'synthetic-stdio-value')
    tool = tr.probe(config)[0]
    assert tool['usable']
    assert exited()
    result = tr.query(config, tool, 'XLOOKUP')
    assert 'XLOOKUP' in result['text'] and 'synthetic-stdio-value' not in str(result)
    assert exited()
    cancel = threading.Event()
    def stop():
        for _ in range(100):
            if (tmp_path / 'running').exists():
                cancel.set()
                return
            sleep(.05)
    watcher = threading.Thread(target=stop)
    watcher.start()
    try:
        with pytest.raises(tr.PluginError, match='取消'):
            tr.query(config, tool, 'slow', cancelled=cancel.is_set)
    finally:
        cancel.set()
        watcher.join()
    assert exited()


def test_learn_rechecks_tools_before_starting_its_separate_search_budget(monkeypatch):
    import asyncio
    from contextlib import asynccontextmanager
    from types import SimpleNamespace
    from mcp.types import CallToolResult, TextContent
    import pytest
    from office_agent.webui import plugin_transport as tr

    tool = dict(name='microsoft_docs_search', signature='fixed', usable=True, minimum=1, maximum=100, query_key='query')
    calls = []
    async def call(name, arguments):
        calls.append((name, arguments))
        await asyncio.sleep(.2)
        docs = [{'title':'Today', 'content':'XLOOKUP', 'contentUrl':'https://learn.microsoft.com/office/doc'}]
        return CallToolResult(content=[TextContent(type='text', text=json.dumps(docs))])

    @asynccontextmanager
    async def session(*args):
        yield SimpleNamespace(call_tool=call), set(), None, None

    async def discover(*args):
        await asyncio.sleep(.2)
        return [tool]

    config = {'transport':'http', 'url':tr.LEARN}
    monkeypatch.setattr(tr, '_session', session)
    monkeypatch.setattr(tr, '_discover', discover)
    monkeypatch.setattr(tr, '_LEARN_METADATA_SECONDS', .32)
    monkeypatch.setattr(tr, '_LEARN_QUERY_SECONDS', .32)
    assert tr.query(config, tool, 'Excel XLOOKUP')['sources'][0]['title'] == 'Today'
    monkeypatch.setattr(tr, '_LEARN_METADATA_SECONDS', .05)
    with pytest.raises(tr.PluginError, match='超时'):
        tr.query(config, tool, 'Excel XLOOKUP')
    assert calls == [('microsoft_docs_search', {'query': 'Excel XLOOKUP'})]


if __name__ == '__main__':
    target = Path(sys.argv[1])
    target.write_text(str(os.getpid()))
    assert os.environ['LOCAL_TOKEN'] == 'synthetic-stdio-value' and 'UNRELATED_SECRET_MARKER' not in os.environ
    for line in sys.stdin:
        message = json.loads(line)
        if 'id' not in message:
            continue
        value = rpc(message)
        if message['method'] == 'tools/call':
            if message['params']['arguments']['query'] == 'slow':
                target.with_name('running').touch()
                sleep(60)
            value['content'][0]['text'] += os.environ['LOCAL_TOKEN']
        print(json.dumps({'jsonrpc': '2.0', 'id': message['id'], 'result': value}), flush=True)
