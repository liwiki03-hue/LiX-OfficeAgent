import json
import time
from urllib.parse import urlsplit

import httpx2

from .settings_definition import PROVIDER_DEFAULTS
from .settings_store import SettingsError, _validate

MAX_RESPONSE = 262144
TIMEOUT = 5


def prepare_probe(store, draft):
    if type(draft) is not dict or not draft.keys() <= {'provider', 'model', 'base_url', 'api_key', 'clear_key'}:
        raise SettingsError()
    values = dict(draft)
    clear = values.pop('clear_key', False)
    if type(clear) is not bool:
        raise SettingsError()
    _validate(values)
    provider = values.get('provider')
    if provider not in PROVIDER_DEFAULTS:
        raise SettingsError()
    defaults = PROVIDER_DEFAULTS[provider]
    url = values.get('base_url') or defaults['base_url']
    parsed = urlsplit(url)
    if provider != 'ollama' and parsed.scheme != 'https':
        raise SettingsError()
    if provider == 'ollama' and parsed.scheme == 'http' and parsed.hostname not in ('127.0.0.1', 'localhost', '::1'):
        raise SettingsError()
    source = store._values(store._read(store.root / '.env', True))
    saved_url = source.get(provider.upper() + '_BASE_URL') or defaults['base_url']
    key = '' if clear or provider == 'ollama' else values.get('api_key') or (source.get(provider.upper() + '_API_KEY') if url.rstrip('/') == saved_url.rstrip('/') else '') or ''
    if key:
        _validate({'api_key': key})
    return provider, url, values.get('model') or defaults['model'], key


def probe_connection(prepared, transport=None, *, include_models=False):
    provider, url, model, key = prepared
    headers = {'Accept': 'application/json', 'Accept-Encoding': 'identity'}
    if key:
        headers['Authorization'] = 'Bearer ' + key
    endpoint = url.rstrip('/') + ('/api/tags' if provider == 'ollama' else '/models')
    started = time.monotonic()
    from .network import transport as network_transport
    try:
        with httpx2.Client(timeout=TIMEOUT, follow_redirects=False, trust_env=False,
                           transport=transport or network_transport(endpoint)) as client:
            with client.stream('GET', endpoint, headers=headers) as response:
                code = response.status_code
                if code in (401, 403):
                    return {'status': 'authentication_failed'}
                if code in (404, 405, 501):
                    return {'status': 'models_unsupported'}
                if 300 <= code < 400:
                    return {'status': 'redirect_refused'}
                if code != 200:
                    return {'status': 'service_failed'}
                if response.headers.get('Content-Encoding', 'identity') != 'identity':
                    return {'status': 'models_unverified'}
                body = bytearray()
                for block in response.iter_raw():
                    if time.monotonic() - started > TIMEOUT:
                        return {'status': 'timeout'}
                    body.extend(block)
                    if len(body) > MAX_RESPONSE:
                        return {'status': 'response_too_large'}
                payload = json.loads(body)
                if type(payload) is not dict:
                    return {'status': 'models_unverified'}
                rows = payload.get('models' if provider == 'ollama' else 'data')
                if type(rows) is not list:
                    return {'status': 'models_unverified'}
                names = [row.get('name' if provider == 'ollama' else 'id') for row in rows if type(row) is dict]
                result = {'status': 'reachable', 'model_listed': model in names}
                if include_models:
                    from .settings_definition import MODEL_PATTERN
                    import re
                    result['models'] = list(dict.fromkeys(name for name in names if type(name) is str and re.fullmatch(MODEL_PATTERN, name) and (not key or key not in name)))[:64]
                    result['truncated'] = len(names) > 64
                return result
    except httpx2.TimeoutException:
        return {'status': 'timeout'}
    except (ValueError, UnicodeError):
        return {'status': 'models_unverified'}
    except Exception:
        return {'status': 'connection_failed'}
