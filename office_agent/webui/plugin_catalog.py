import ipaddress
import re
import time
from urllib.parse import urlsplit

import httpx


REGISTRY = 'https://registry.modelcontextprotocol.io/v0.1/servers'
IDENTITY = re.compile(r'[A-Za-z0-9.-]+/[A-Za-z0-9._-]+\Z')


def clean_url(value):
    if not isinstance(value, str) or len(value) > 1500 or re.search(r'[\s\\\x00-\x1f{}]', value):
        raise ValueError('请填写完整的 HTTPS 服务地址。')
    try:
        parsed = urlsplit(value)
        port = parsed.port
        host = parsed.hostname
    except ValueError:
        raise ValueError('服务地址格式无效。') from None
    if parsed.scheme != 'https' or not host or parsed.username is not None or parsed.password is not None or parsed.query or parsed.fragment:
        raise ValueError('请输入 HTTPS 地址，凭据请使用本机引用。')
    host = host.encode('idna').decode('ascii').lower()
    authority = '[' + host + ']' if ':' in host else host
    return 'https://' + authority + (':' + str(port) if port not in (None, 443) else '') + (parsed.path or '/')


def public_url(raw):
    try:
        value = clean_url(raw)
        name = urlsplit(value).hostname.rstrip('.')
        if name == 'localhost' or name.endswith(('.local', '.localhost', '.internal')) or '.' not in name:
            return ''
        try:
            if not ipaddress.ip_address(name).is_global:
                return ''
        except ValueError:
            pass
        return value
    except (TypeError, ValueError, UnicodeError):
        return ''


def fetch_page(term, cursor):
    params = {'search': term, 'version': 'latest', 'limit': '15'}
    if cursor:
        params['cursor'] = cursor
    deadline = time.monotonic() + 15
    with httpx.stream('GET', REGISTRY, params=params, follow_redirects=False, timeout=8) as response:
        response.raise_for_status()
        data = bytearray()
        for chunk in response.iter_bytes(32768):
            data.extend(chunk)
            if len(data) > 1024 * 1024 or time.monotonic() > deadline:
                raise ValueError('目录响应过大或超时。')
    import json
    page = json.loads(data)
    rows = page['servers']
    if not isinstance(rows, list) or len(rows) > 50:
        raise ValueError('目录返回格式不兼容。')
    items = {}
    for row in rows:
        server = row.get('server', {})
        identity = server.get('name', '')
        if not IDENTITY.fullmatch(identity) or len(identity) > 200:
            continue
        remotes = [public_url(remote.get('url')) for remote in server.get('remotes', [])
                   if remote.get('type') == 'streamable-http']
        packages = [{'name': str(pkg.get('identifier', ''))[:160], 'version': str(pkg.get('version', ''))[:80],
                     'registry': str(pkg.get('registryType', ''))[:40], 'runtime': str(pkg.get('runtimeHint', ''))[:80]}
                    for pkg in server.get('packages', [])[:5]]
        items[identity] = dict(source=identity, publisher=identity.split('/')[0], name=str(server.get('title') or identity)[:100],
                               description=str(server.get('description', ''))[:260], url=next(filter(None, remotes), ''),
                               website=public_url(server.get('websiteUrl') or (server.get('repository') or {}).get('url', '')),
                               version=str(server.get('version', ''))[:80], packages=packages)
    next_page = page.get('metadata', {}).get('nextCursor') or ''
    if not isinstance(next_page, str) or len(next_page) > 1500:
        raise ValueError('目录翻页标识无效。')
    return {'items': list(items.values()), 'cursor': next_page}

