import asyncio
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
import webbrowser
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from mcp.client.auth import OAuthClientProvider
from mcp.shared.auth import OAuthClientMetadata

from ..downloads import safe


class PluginError(ValueError):
    pass


_lock = threading.RLock()
_generations = {}


def _location(config):
    identity = json.dumps([config['id'], config['url'], config.get('oauth_issuer', '')], ensure_ascii=True)
    return safe(Path(config['_profile_dir']) / 'oauth' / (hashlib.sha256(identity.encode()).hexdigest() + '.bin'))


def _protect(data):
    import win32crypt
    return win32crypt.CryptProtectData(data, 'LiX MCP', None, None, None, 1)


def _unprotect(data):
    import win32crypt
    return win32crypt.CryptUnprotectData(data, None, None, None, 1)[1]


def token(config):
    path = _location(config)
    with _lock:
        if path.is_file() and path.stat().st_size <= 24000:
            try:
                value = json.loads(_unprotect(path.read_bytes()))
                if value['until'] is None or value['until'] > time.time():
                    return value['bearer']
            except Exception:
                pass
    raise PluginError('登录已失效，请在插件详情中重新登录。')


def disconnect(config):
    if config.get('auth') != 'oauth':
        return
    with _lock:
        path = _location(config)
        _generations[path] = _generations.get(path, 0) + 1
        path.unlink(missing_ok=True)


def invalidate(config, rejected):
    with _lock:
        try:
            current = token(config)
        except PluginError:
            return
        if current == rejected:
            disconnect(config)


class LoginStorage:
    def __init__(self, config):
        self.path = _location(config)
        with _lock:
            self.generation = _generations.get(self.path, 0)
        self.info = self.tokens = self.error = None

    def fail(self, reason):
        self.error = PluginError(reason)
        raise self.error

    async def get_tokens(self):
        return self.tokens

    async def set_tokens(self, tokens):
        self.tokens = tokens

    async def get_client_info(self):
        return self.info

    async def set_client_info(self, client_info):
        self.info = client_info

    def commit(self, cancelled):
        with _lock:
            if self.generation != _generations.get(self.path, 0) or cancelled():
                raise PluginError('登录已取消。')
            if not self.tokens or not self.tokens.access_token or len(self.tokens.access_token) > 8192:
                raise PluginError('服务未返回可用的登录凭据。')
            until = time.time() + self.tokens.expires_in if self.tokens.expires_in is not None else None
            data = _protect(json.dumps(dict(bearer=self.tokens.access_token, until=until)).encode())
            self.path.parent.mkdir(parents=True, exist_ok=True)
            temp = self.path.with_suffix('.' + secrets.token_hex(8) + '.tmp')
            try:
                with temp.open('xb') as stream:
                    stream.write(data)
                os.replace(temp, self.path)
                _generations[self.path] = self.generation + 1
            finally:
                temp.unlink(missing_ok=True)


@asynccontextmanager
async def authorization(config, origins, cancelled):
    if os.name != 'nt':
        raise PluginError('本机登录凭据存储需要 Windows。')
    from .plugin_transport import origin
    storage = LoginStorage(config)
    result = asyncio.get_running_loop().create_future()
    state = ''

    async def callback(reader, writer):
        ok = False
        try:
            data = await asyncio.wait_for(reader.readuntil(b'\r\n\r\n'), 3)
            lines = data.decode('ascii').split('\r\n')
            method, target, protocol = lines[0].split(' ')
            url = urlsplit(target)
            args = parse_qs(url.query, strict_parsing=True, max_num_fields=8)
            hosts = [line.partition(':')[2].strip() for line in lines[1:] if line.lower().startswith('host:')]
            received = args.get('state', [])
            if (method != 'GET' or protocol != 'HTTP/1.1' or url.path != '/callback' or url.netloc or url.scheme
                    or hosts != [address] or not state or len(received) != 1 or not hmac.compare_digest(received[0], state)
                    or result.done() or cancelled()):
                return
            if 'error' in args:
                result.set_result(None)
                ok = True
            elif len(args.get('code', [])) == 1 and 1 <= len(args['code'][0]) <= 4096:
                result.set_result((args['code'][0], state))
                ok = True
        except (ValueError, OSError, asyncio.IncompleteReadError, asyncio.LimitOverrunError, TimeoutError):
            pass
        finally:
            status = b'200 OK' if ok else b'400 Bad Request'
            writer.write(b'HTTP/1.1 ' + status + b'\r\nConnection: close\r\nCache-Control: no-store\r\nContent-Length: 0\r\n\r\n')
            writer.close()

    server = await asyncio.start_server(callback, '127.0.0.1', 0, limit=12000)
    address = '127.0.0.1:' + str(server.sockets[0].getsockname()[1])
    async with server:
        async def redirect(url):
            nonlocal state
            parsed = urlsplit(url)
            args = parse_qs(parsed.query, max_num_fields=32)
            if (origin(url) not in origins or len(args.get('state', [])) != 1 or cancelled()
                    or args.get('redirect_uri') != ['http://' + address + '/callback']):
                storage.fail('授权地址或登录状态与当前服务不符。')
            state = args['state'][0]
            if not await asyncio.to_thread(webbrowser.open, url):
                storage.fail('无法打开系统浏览器，请检查默认浏览器后重新登录。')

        async def receive():
            value = await result
            if value is None:
                storage.fail('用户取消了登录。')
            return value

        auth = OAuthClientProvider(config['url'], OAuthClientMetadata(client_name='LiX',
            redirect_uris=['http://' + address + '/callback'], token_endpoint_auth_method='none'),
            storage, redirect, receive, timeout=120)
        try:
            yield auth, storage
        finally:
            if not result.done():
                result.cancel()
