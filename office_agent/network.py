from dataclasses import dataclass
import ipaddress
import os
from pathlib import Path
import ssl
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, HTTPSHandler, ProxyHandler, build_opener, getproxies


@dataclass(frozen=True, repr=False)
class NetworkOptions:
    mode: str = 'direct'
    proxy: str = ''
    bypass: str = ''

    def __post_init__(self):
        if self.mode not in {'auto', 'direct', 'manual'}:
            raise ValueError('network_mode_invalid')
        if self.mode == 'manual' or self.proxy:
            value = urlsplit(self.proxy)
            if (value.scheme not in {'http', 'https'} or not value.hostname or value.path not in {'', '/'}
                    or value.query or value.fragment or any(c.isspace() for c in self.proxy) or value.port == 0):
                raise ValueError('network_proxy_invalid')

    def route(self, url):
        parsed = urlsplit(url)
        host = (parsed.hostname or '').lower()
        if not host or parsed.scheme not in {'http', 'https'}:
            raise ValueError('network_target_invalid')
        try:
            local = ipaddress.ip_address(host).is_loopback
        except ValueError:
            local = host == 'localhost' or host.endswith('.localhost')
        if local or self.mode == 'direct':
            return None
        for item in self.bypass.split(','):
            item = item.strip().lower()
            if item == '*':
                return None
            if '/' in item:
                try:
                    if ipaddress.ip_address(host) in ipaddress.ip_network(item, strict=False):
                        return None
                except ValueError:
                    pass
            name = item.lstrip('.')
            if name and (host == name or host.endswith('.' + name) or f'{host}:{parsed.port}' == name):
                return None
        if self.mode == 'manual':
            return self.proxy
        proxies = getproxies()
        value = proxies.get(parsed.scheme) or proxies.get('all')
        if not value:
            return None
        NetworkOptions('manual', value)
        return value


def options(root=None, *, default='direct'):
    from dotenv import dotenv_values
    root = Path(root or Path.cwd()).absolute()
    path = root / '.env'
    from .downloads import safe
    safe(path)
    values = {}
    if path.exists():
        if path.is_symlink() or path.stat().st_size > 128 * 1024:
            raise ValueError('network_settings_invalid')
        values = dotenv_values(path, interpolate=False)
    mode = values.get('OFFICE_NETWORK_MODE') or os.environ.get('OFFICE_NETWORK_MODE', default)
    proxy = values.get('OFFICE_NETWORK_PROXY') or os.environ.get('OFFICE_NETWORK_PROXY', '')
    bypass = values.get('OFFICE_NO_PROXY') or os.environ.get('NO_PROXY', os.environ.get('no_proxy', ''))
    return NetworkOptions(mode, proxy, bypass)


def endpoint_identity(url):
    from urllib.parse import urlunsplit
    parsed = urlsplit(url)
    host = (parsed.hostname or '').lower()
    if ':' in host:
        host = '[' + host + ']'
    port = parsed.port
    authority = host + (':' + str(port) if port is not None and port != (443 if parsed.scheme == 'https' else 80) else '')
    return urlunsplit((parsed.scheme.lower(), authority, parsed.path.rstrip('/'), parsed.query, parsed.fragment))


def transport(url, *, root=None):
    import httpx2
    return httpx2.HTTPTransport(proxy=options(root).route(url), retries=0, trust_env=False)


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, newurl):
        return None


def urlopen(request, timeout=30):
    import certifi
    url = request.full_url if hasattr(request, 'full_url') else request
    proxy = options().route(url)
    opener = build_opener(ProxyHandler({'http': proxy, 'https': proxy} if proxy else {}),
                          HTTPSHandler(context=ssl.create_default_context(cafile=certifi.where())), _NoRedirect())
    return opener.open(request, timeout=timeout)
