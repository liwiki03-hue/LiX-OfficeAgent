import asyncio
from concurrent.futures import Future, TimeoutError as FutureTimeout
from contextlib import AsyncExitStack, asynccontextmanager
from contextvars import ContextVar, copy_context
from datetime import timedelta
import hashlib
import ipaddress
import json
import logging
import os
from pathlib import Path
import re
import socket
import sys
import tempfile
import threading
from urllib.parse import parse_qsl, urlsplit
from urllib.request import proxy_bypass

import httpx
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.client.streamable_http import streamable_http_client

from ..mcp_reads import McpReadRegistry
from ..network import options as network_options
from . import openalex_search, plugin_oauth
from .plugin_oauth import PluginError, disconnect


LEARN = 'https://learn.microsoft.com/api/mcp'
SECRET_NAME = re.compile(r'LIX_PLUGIN_[A-Z][A-Z0-9_]{0,79}')
QUERY_KEYS = {'query', 'question', 'keywords', 'term', 'search_term'}
LINK = re.compile(r'https://[^\s<>"\'\[\]{}）)]+')
_private_operation = ContextVar('lix_private_plugin_operation', default=False)
_bridge_workers = threading.BoundedSemaphore(4)
_REQUEST_SECONDS = 26
_LEARN_METADATA_SECONDS = 42
_LEARN_QUERY_SECONDS = 20
_CLOSE_SECONDS = 2


def _allow_log(record):
    return not _private_operation.get()


for _logger_name in ('client', 'mcp.client.stdio', 'mcp.client.auth.oauth2', 'mcp.client.auth.utils',
                     'mcp.client.streamable_http', 'mcp.shared.tool_name_validation', 'httpx'):
    logging.getLogger(_logger_name).addFilter(_allow_log)


def origin(url):
    try:
        parsed = urlsplit(url)
        if (parsed.scheme != 'https' or not parsed.hostname or parsed.username is not None or parsed.password is not None
                or '\\' in url or any(ord(char) < 33 for char in url)):
            raise ValueError
        if parsed.port == 0:
            raise ValueError
        return (parsed.scheme, parsed.hostname.lower(), parsed.port or 443)
    except (TypeError, ValueError):
        raise PluginError('请填写没有账号、密码的 HTTPS 服务地址。') from None


def _secret(config, key):
    if not isinstance(key, str) or not SECRET_NAME.fullmatch(key):
        raise PluginError('凭据引用应使用 LIX_PLUGIN_ 开头的本机配置名。')
    try:
        value = config['_resolve_secret'](key)
    except Exception:
        value = None
    if not isinstance(value, str) or not value or len(value) > 8192 or re.search(r'[\x00-\x1f\x7f]', value):
        raise PluginError('找不到有效的本机凭据，请核对 .env 引用。')
    return value


def credential_available(config):
    try:
        if config.get('auth') == 'oauth':
            plugin_oauth.token(config)
        if config.get('auth') == 'bearer':
            _secret(config, config.get('credential_ref'))
        for ref in config.get('env_refs', {}).values():
            _secret(config, ref)
        return True
    except (PluginError, OSError, KeyError):
        return False


class _LimitedStream(httpx.AsyncByteStream):
    def __init__(self, stream):
        self.stream = stream

    async def __aiter__(self):
        size = 0
        async for chunk in self.stream:
            size += len(chunk)
            if size > 256 * 1024:
                raise PluginError('服务响应超过读取上限。')
            yield chunk

    async def aclose(self):
        await self.stream.aclose()


class _Transport(httpx.AsyncBaseTransport):
    def __init__(self, config, cancelled, origins):
        self.config = config
        self.endpoint = origin(config['url'])
        self.proxy = None
        if (config['url'].rstrip('/') == LEARN and config.get('auth', 'none') == 'none'
                and not config.get('private')):
            try:
                network = network_options(config.get('_settings_root'), default='auto')
                if network.mode != 'auto' or not proxy_bypass(self.endpoint[1]):
                    self.proxy = network.route(LEARN)
            except (ValueError, OSError):
                raise PluginError('请核对 Microsoft Learn 的网络代理设置。') from None
        self.pools = {target: httpx.AsyncHTTPTransport(proxy=self.proxy if target == self.endpoint else None,
                      trust_env=False, retries=1, limits=httpx.Limits(max_connections=4,
                      max_keepalive_connections=2)) for target in origins}
        self.origins = origins
        self.allow_private = config.get('private') is True and config.get('source', '') in ('', None, 'manual')
        self.cancelled = cancelled
        self.stage = '初始化'
        self.error = None
        self.login = False
        self.secrets = set()

    async def handle_async_request(self, request):
        try:
            target = origin(str(request.url))
            if target not in self.origins or self.cancelled() and request.method != 'DELETE':
                raise PluginError('连接已取消或目标来源改变，请核对地址。')
            host, port = target[1:]
            records = await asyncio.get_running_loop().getaddrinfo(host, port, type=socket.SOCK_STREAM)
            addresses = [ipaddress.ip_address(row[4][0]) for row in records]
            if not addresses or (not self.allow_private or target != self.endpoint) and any(not ip.is_global or ip.is_multicast for ip in addresses):
                raise PluginError('目录服务只能连接公网地址。私有端点请手动添加并确认。')
            pinned = next((ip for ip in addresses if ip.version == 4), addresses[0])
            through_proxy = self.proxy is not None and target == self.endpoint
            if through_proxy and (str(request.url).rstrip('/') != LEARN
                                  or any(key in request.headers for key in ('Authorization', 'X-Api-Key'))):
                raise PluginError('该代理连接仅用于匿名查询 Microsoft Learn。')
            url = request.url if through_proxy else request.url.copy_with(host=str(pinned))
            for key in ('Authorization', 'X-Api-Key'):
                if request.headers.get(key):
                    if target != self.endpoint:
                        raise PluginError('已阻止向其他来源发送插件凭据。')
                    self.secrets.add(request.headers[key].removeprefix('Bearer '))
            wire = httpx.Request(request.method, url, headers=request.headers, stream=request.stream,
                                 extensions={**request.extensions, 'sni_hostname': host})
            response = await self.pools[target].handle_async_request(wire)
            if response.is_redirect or response.headers.get('content-encoding', 'identity').lower() != 'identity':
                await response.aclose()
                raise PluginError('连接发生重定向或返回不支持的编码，请核对最终服务地址。')
            if response.status_code == 429 and self.config['url'].rstrip('/') == openalex_search.URL:
                await response.aclose()
                raise openalex_search.SearchRejected('OpenAlex 请求受限，请稍后重试或检查账户额度。')
            if response.status_code in {401, 403} and not self.login and request.method != 'DELETE':
                await response.aclose()
                if response.status_code == 401 and self.config.get('auth') == 'oauth':
                    plugin_oauth.invalidate(self.config, request.headers.get('Authorization', '').removeprefix('Bearer '))
                if self.config['url'].rstrip('/') == openalex_search.URL:
                    raise openalex_search.SearchRejected('OpenAlex 已拒绝授权，请在插件设置中重新登录。')
                raise PluginError('服务拒绝访问，请核对凭据或重新登录。')
            response.stream = _LimitedStream(response.stream)
            return response
        except (httpx.HTTPError, OSError, PluginError) as error:
            self.error = error if isinstance(error, PluginError) else PluginError(self.stage + '时连接失败或超时，请核对服务地址和网络。')
            raise self.error from None

    async def aclose(self):
        for pool in self.pools.values():
            await pool.aclose()


def _descriptor(raw, config):
    name = raw.get('name', '')
    schema = raw.get('inputSchema') or {}
    signature = hashlib.sha256(json.dumps({k: raw.get(k) for k in ('name', 'inputSchema', 'outputSchema', 'annotations')},
        sort_keys=True, separators=(',', ':'), ensure_ascii=False).encode()).hexdigest()
    result = dict(name=name, description=str(raw.get('description') or '')[:240], signature=signature,
                  usable=False, reason='只支持经审核的单字段只读检索工具。', query_key=None)
    if config.get('transport') == 'http' and config.get('url', '').rstrip('/') == openalex_search.URL:
        if openalex_search.supported(raw):
            return dict(result, usable=True, reason='', query_key='query', minimum=1, maximum=2000)
        return dict(result, reason='OpenAlex 仅启用已核对结构的 search_works。')
    properties = schema.get('properties') or {}
    official = config.get('transport') == 'http' and config.get('url', '').rstrip('/') == LEARN
    if type(properties) is not dict or len(properties) != 1:
        return result
    key, field = next(iter(properties.items()))
    if key not in QUERY_KEYS or type(field) is not dict or field.get('type') != 'string':
        return result
    if official:
        hints = raw.get('annotations') or {}
        if (name != 'microsoft_docs_search' or key != 'query' or schema.get('type') != 'object'
                or schema.get('required', []) not in ([], [key])
                or set(schema) - {'type', 'properties', 'required', 'additionalProperties', 'description', 'title'}
                or set(field) - {'type', 'description', 'title', 'default', 'minLength', 'maxLength'}
                or hints.get('readOnlyHint') is not True or hints.get('destructiveHint') is not False):
            return result
    else:
        try:
            McpReadRegistry([raw], [name], None)
        except (ValueError, TypeError):
            return result
    low, high = field.get('minLength', 1), field.get('maxLength', 2000)
    if type(low) is not int or type(high) is not int or not 0 <= low <= high or low > 2000:
        return result
    result.update(usable=True, reason='', query_key=key, minimum=max(1, low), maximum=min(2000, high))
    if 'enum' in field:
        result['choices'] = field['enum']
    return result


async def _discover(session, config):
    tools = []
    names = set()
    cursors = set()
    cursor = None
    for _ in range(8):
        page = await session.list_tools(cursor=cursor)
        for tool in page.tools:
            if len(tools) >= 128 or tool.name in names or not re.fullmatch(r'[\w.-]{1,128}', tool.name, re.ASCII):
                raise PluginError('工具列表重复或超出上限，请核对服务。')
            names.add(tool.name)
            tools.append(_descriptor(tool.model_dump(mode='json', by_alias=True, exclude_none=True), config))
        cursor = page.nextCursor
        if not cursor:
            return tools
        if cursor in cursors or len(cursor) > 1024:
            break
        cursors.add(cursor)
    raise PluginError('服务的工具分页超过上限。')


def _redact(text, secrets):
    for value in secrets:
        if value:
            text = text.replace(value, '[凭据已隐藏]')
    text = re.sub(r'(?i)\b(Bearer\s+)[a-z0-9._~+/-]+=*', r'\1[凭据已隐藏]', text)
    text = re.sub(r'(?i)\b((?:api[_-]?key|access[_-]?token|client_secret)\s*[=:]\s*)[^\s&,"\x27]+', r'\1[凭据已隐藏]', text)
    return text


def _source(url, official):
    try:
        _, host, _ = origin(url)
        parsed = urlsplit(url)
        params = parse_qsl(parsed.query, strict_parsing=True, max_num_fields=4)
    except (ValueError, TypeError):
        return False
    if (len(url) > 1500 or host == 'localhost' or host.endswith(('.local', '.localhost'))
            or official and host not in {'learn.microsoft.com', 'support.microsoft.com'}
            or any(key not in {'view', 'tabs', 'preserve-view'} or not re.fullmatch(r'[\w.,-]{1,80}', value, re.ASCII)
                   for key, value in params)):
        return False
    try:
        address = ipaddress.ip_address(host)
        return address.is_global and not address.is_multicast
    except ValueError:
        return True


def _records(parts, structured):
    rows = []
    if isinstance(structured, dict) and isinstance(structured.get('results'), list):
        values = [structured['results']]
    else:
        try:
            values = [json.loads(part.text) for part in parts]
        except (ValueError, RecursionError):
            return None
    for data in values:
        if isinstance(data, dict):
            data = data.get('results', [data])
        if not isinstance(data, list):
            return None
        for row in data:
            if not isinstance(row, dict) or not isinstance(row.get('contentUrl'), str) or not isinstance(row.get('content'), str):
                return None
            rows.append(row)
    return rows


def _output(response, secrets, official):
    if response.isError:
        raise PluginError('这次查询失败，请核对工具或稍后重试。')
    parts = response.content
    if not parts or len(parts) > 20 or any(item.type != 'text' for item in parts):
        raise PluginError('这次没有可用的文本结果。')
    records = _records(parts, response.structuredContent)
    if records is not None and official:
        sections, sources = [], []
        for row in records[:20]:
            url = row['contentUrl']
            if not _source(url, True) or _redact(url, secrets) != url or any(source['url'] == url for source in sources):
                continue
            title = _redact(str(row.get('title') or urlsplit(url).hostname)[:160], secrets)
            content = _redact(row['content'], secrets).encode('utf-8')[:2500].decode('utf-8', errors='ignore').strip()
            if content:
                section = title + '\n' + content + '\n' + url
                if sum(len(s.encode('utf-8')) + 2 for s in [*sections, section]) > 19500:
                    break
                sections.append(section)
                sources.append({'title': title, 'url': url})
            if len(sources) == 8:
                break
        if not sections:
            raise PluginError('这次没有查到带官方来源的资料，请更换关键词。')
        return {'text': '\n\n'.join(sections), 'sources': sources}
    text = _redact('\n'.join(item.text for item in parts).strip(), secrets)
    whole_links = None
    if official and len(text.encode('utf-8')) > 19500:
        matches = [match for match in LINK.finditer(text) if _source(match[0].rstrip('.,;:!，。；'), True)]
        first_link = next(iter(matches), None)
        if first_link is None:
            raise PluginError('本次结果缺少可核对的官方来源，请更换关键词。')
        first = max(0, first_link.start() - 2000) if len(text[:first_link.end()].encode('utf-8')) > 19000 else 0
        whole_links = {match[0].rstrip('.,;:!，。；') for match in matches}
        text = text[first:].encode('utf-8')[:19500].decode('utf-8', errors='ignore')
    if not text or len(text.encode('utf-8')) > 20000:
        raise PluginError('本次结果为空或超过文本上限，请缩小查询范围。')
    sources = []
    titles = {r['contentUrl']: _redact(str(r.get('title') or '')[:160], secrets) for r in (records or [])}
    for match in LINK.finditer(text):
        url = match[0].rstrip('.,;:!，。；')
        if not _source(url, official) or whole_links is not None and url not in whole_links:
            continue
        if url not in [source['url'] for source in sources]:
            sources.append({'title': titles.get(url) or urlsplit(url).hostname + urlsplit(url).path[:100], 'url': url})
        if len(sources) == 8:
            break
    if official and not sources:
        raise PluginError('本次结果缺少可核对的官方来源，请更换关键词。')
    return {'text': text, 'sources': sources}


def _local(config):
    program = Path(config.get('program', ''))
    args = config.get('args', [])
    refs = config.get('env_refs', {})
    if (os.name != 'nt' or not program.is_absolute() or not program.is_file() or program.suffix.lower() != '.exe'
            or str(program).startswith('\\\\') or program.name.lower() in {'cmd.exe', 'powershell.exe', 'pwsh.exe', 'wscript.exe', 'cscript.exe'}
            or not isinstance(args, list) or len(args) > 32 or any(not isinstance(x, str) or len(x) > 2048 or re.search(r'[\x00-\x1f]', x) for x in args)):
        raise PluginError('请选择已安装的本机 EXE，并逐项填写参数。')
    if (type(refs) is not dict or len(refs) > 12 or any(not re.fullmatch(r'[A-Z][A-Z0-9_]{0,63}', key)
            or key in {'PATH', 'COMSPEC', 'SYSTEMROOT', 'WINDIR'} for key in refs)):
        raise PluginError('本机环境变量名称无效。')
    return str(program), args, {key: _secret(config, ref) for key, ref in refs.items()}


@asynccontextmanager
async def _session(config, cancelled, login):
    secrets = set()
    transport = None
    try:
        async with AsyncExitStack() as stack:
            storage = None
            if config['transport'] == 'stdio':
                program, args, env = _local(config)
                secrets.update(env.values())
                cwd = stack.enter_context(tempfile.TemporaryDirectory(prefix='lix-plugin-'))
                stderr = stack.enter_context(open(os.devnull, 'w'))
                params = StdioServerParameters(command=sys.executable,
                    args=['-I', str(Path(__file__).with_name('plugin_stdio_worker.py')), json.dumps(list(env)), program, *args],
                    env=env, cwd=cwd)
                reader, writer = await stack.enter_async_context(stdio_client(params, errlog=stderr))
            elif config['transport'] == 'http':
                url = config['url']
                if urlsplit(url).query or urlsplit(url).fragment:
                    raise PluginError('请把凭据放入本机引用，服务地址不能含查询串或片段。')
                origins = {origin(url)}
                if config.get('oauth_issuer') and config.get('auth') == 'oauth':
                    origins.add(origin(config['oauth_issuer']))
                transport = _Transport(config, cancelled, origins)
                transport.login = login
                headers = {'Accept-Encoding': 'identity'}
                auth = None
                if login:
                    auth, storage = await stack.enter_async_context(plugin_oauth.authorization(config, origins, cancelled))
                elif config.get('auth') in {'bearer', 'oauth'}:
                    value = plugin_oauth.token(config) if config['auth'] == 'oauth' else _secret(config, config.get('credential_ref'))
                    key = config.get('api_header') or 'Authorization'
                    if key not in {'Authorization', 'X-Api-Key'}:
                        raise PluginError('不支持该凭据请求头。')
                    headers[key] = ('Bearer ' if key == 'Authorization' else '') + value
                    secrets.add(value)
                elif config.get('auth', 'none') != 'none':
                    raise PluginError('不支持该认证方式。')
                http = await stack.enter_async_context(httpx.AsyncClient(auth=auth, headers=headers, transport=transport,
                    timeout=httpx.Timeout(16, connect=6), trust_env=False, follow_redirects=False))
                get_session_id = None
                protocol = ''
                async def close_remote():
                    sid = get_session_id() if get_session_id else None
                    if not sid:
                        return
                    final_headers = {'Mcp-Session-Id': sid, 'Accept': 'application/json, text/event-stream'}
                    if protocol:
                        final_headers['MCP-Protocol-Version'] = protocol
                    if storage and storage.tokens:
                        final_headers['Authorization'] = 'Bearer ' + storage.tokens.access_token
                    try:
                        async with asyncio.timeout(_CLOSE_SECONDS / 2):
                            async with http.stream('DELETE', url, headers=final_headers, auth=None):
                                pass
                    except (httpx.HTTPError, OSError, ValueError):
                        pass
                stack.push_async_callback(close_remote)
                reader, writer, get_session_id = await stack.enter_async_context(
                    streamable_http_client(url, http_client=http, terminate_on_close=False))
            else:
                raise PluginError('不支持该传输方式。')
            session = await stack.enter_async_context(ClientSession(reader, writer, read_timeout_seconds=timedelta(seconds=125 if login else 18)))
            initialized = await session.initialize()
            protocol = initialized.protocolVersion
            if config.get('url', '').rstrip('/') == openalex_search.URL:
                config['_server'] = {'name': initialized.serverInfo.name[:160], 'version': initialized.serverInfo.version[:80]}
            yield session, secrets, storage, transport
    except Exception as error:
        if transport and transport.error:
            raise transport.error from None
        if storage and storage.error:
            raise storage.error from None
        errors = [error]
        while errors:
            nested = errors.pop()
            if isinstance(nested, PluginError):
                raise nested from None
            errors.extend(getattr(nested, 'exceptions', []))
        raise PluginError('无法完成 MCP 握手或调用，请核对服务地址、凭据与工具协议。') from None


def _run(config, tool, text, cancelled, login=False):
    owner_cancelled = cancelled
    abandoned = threading.Event()
    receipt = None
    close_deadline = None
    learn = config.get('transport') == 'http' and str(config.get('url', '')).rstrip('/') == LEARN and not login
    metadata_seconds = 125 if login else _LEARN_METADATA_SECONDS if learn else _REQUEST_SECONDS
    query_seconds = _LEARN_QUERY_SECONDS if learn and tool else 0
    request_deadline = None
    query_started = False
    def cancelled():
        return abandoned.is_set() or owner_cancelled()

    def verified(value):
        nonlocal receipt, close_deadline
        if config['transport'] == 'http':
            receipt = value
            close_deadline = asyncio.get_running_loop().time() + _CLOSE_SECONDS
        return value

    async def operation():
        nonlocal request_deadline, query_started
        async with _session(config, cancelled, login) as (session, secrets, storage, transport):
            if transport:
                transport.stage = '读取工具清单'
            rows = await _discover(session, config)
            if config.get('_server') and config.get('url', '').rstrip('/') == openalex_search.URL:
                for row in rows:
                    row['server'] = config['_server']
            if transport:
                secrets.update(transport.secrets)
            if cancelled():
                raise PluginError('本轮已取消。')
            if tool is None:
                for row in rows:
                    if _redact(json.dumps(row, ensure_ascii=False), secrets) != json.dumps(row, ensure_ascii=False):
                        raise PluginError('工具信息含有凭据，已拒绝保存。')
                if storage:
                    storage.commit(cancelled)
                return verified(rows)
            row = next((r for r in rows if r['name'] == tool.get('name')), None)
            if not row or not row['usable'] or row['signature'] != tool.get('signature'):
                raise PluginError('工具结构或授权已改变，请重新测试并选择工具。')
            if (not isinstance(text, str) or not row['minimum'] <= len(text) <= row['maximum']
                    or 'choices' in row and text not in row['choices'] or re.search(r'[\x00-\x08\x0b-\x1f]|[A-Za-z]:[\\/]|https?://|file:|\\\\|(?:^|\s)/[\w.]', text, re.I)
                    or any(value in text for value in secrets if value)):
                raise PluginError('查询范围包含路径、凭据或不受支持的内容，请仅提供查询词。')
            if cancelled():
                raise PluginError('本轮已取消。')
            openalex = config['transport'] == 'http' and config.get('url', '').rstrip('/') == openalex_search.URL
            arguments = openalex_search.arguments(text) if openalex else {row['query_key']: text}
            if transport:
                transport.stage = '查询资料'
            before_dispatch = config.get('_before_dispatch')
            if before_dispatch:
                try:
                    before_dispatch()
                except ValueError:
                    raise PluginError('派发前文档、插件或本轮许可已改变，请重新提问。') from None
            if query_seconds:
                now = asyncio.get_running_loop().time()
                if now >= request_deadline:
                    raise PluginError('核对 Microsoft Learn 工具清单超时，请重试。')
                request_deadline = now + query_seconds
                query_started = True
            result = await session.call_tool(row['name'], arguments)
            if cancelled():
                raise PluginError('本轮已取消。')
            if openalex:
                records = openalex_search.output(result, arguments['limit'])
                records['text'] = _redact(records['text'], secrets)
                for source in records['sources']:
                    source['title'] = _redact(source['title'], secrets)
                return verified(records)
            return verified(_output(result, secrets, str(config.get('url') or '').rstrip('/') == LEARN))

    async def run():
        nonlocal request_deadline
        if cancelled():
            raise PluginError('本轮已取消。')
        loop = asyncio.get_running_loop()
        request_deadline = loop.time() + metadata_seconds
        task = asyncio.create_task(operation())
        try:
            while not task.done():
                if cancelled():
                    raise PluginError('本轮已取消。')
                if close_deadline is not None and loop.time() >= close_deadline:
                    break
                if close_deadline is None and loop.time() >= request_deadline:
                    if learn:
                        phase = '本次检索' if query_started else '握手或工具清单'
                        raise PluginError('Microsoft Learn ' + phase + '超时，请稍后重试。')
                    raise PluginError('连接或调用超时，请核对服务后重试。')
                await asyncio.wait([task], timeout=.08)
            if receipt is None:
                return task.result()
        finally:
            if not task.done():
                task.cancel()
            await asyncio.gather(task, return_exceptions=True)
        if cancelled():
            raise PluginError('本轮已取消。')
        return receipt
    scope = _private_operation.set(True)
    try:
        try:
            asyncio.get_running_loop()
        except RuntimeError:
            return asyncio.run(run())
        if not _bridge_workers.acquire(blocking=False):
            raise PluginError('已有多个插件操作进行中，请稍后重试。')
        result = Future()
        def worker():
            try:
                result.set_result(asyncio.run(run()))
            except BaseException as error:
                result.set_exception(error)
            finally:
                _bridge_workers.release()
        context = copy_context()
        try:
            threading.Thread(target=context.run, args=(worker,), name='LiX-plugin-request', daemon=True).start()
        except BaseException:
            _bridge_workers.release()
            raise
        try:
            return result.result(timeout=metadata_seconds + query_seconds + _CLOSE_SECONDS + 8)
        except FutureTimeout:
            abandoned.set()
            raise PluginError('插件请求超时，请重试。') from None
    finally:
        _private_operation.reset(scope)


def probe(config, *, cancelled=lambda: False):
    return _run(config, None, '', cancelled)


def query(config, tool, query_text, *, cancelled=lambda: False):
    return _run(config, tool, query_text, cancelled)


def login(config, *, cancelled=lambda: False):
    if config.get('transport') != 'http' or config.get('auth') != 'oauth':
        raise PluginError('此连接未配置 OAuth。')
    return _run(config, None, '', cancelled, login=True)
