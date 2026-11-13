from dataclasses import dataclass, field
import hashlib
import hmac
import json
import re
import secrets
import time
from urllib.parse import urlsplit

from .settings_definition import MODEL_PATTERN, PROVIDER_DEFAULTS
from .settings_store import SettingsError, _File, _Plan, MAX_ENV_BYTES, TTL_SECONDS, _validate


ENV_NAME = 'OFFICE_MODEL_CONNECTIONS'
CAPABILITIES = ('chat', 'agent', 'image')
ID = r'[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}'


def _string(value, maximum):
    if type(value) is not str or not value.strip() or len(value) > maximum or any(ord(c) < 32 or ord(c) == 127 or 0xD800 <= ord(c) <= 0xDFFF for c in value):
        raise SettingsError('settings_invalid')
    return value


def validate_catalog(value):
    if type(value) is not dict or set(value) != {'connections', 'defaults'} or type(value['connections']) is not list or len(value['connections']) > 16:
        raise SettingsError()
    ids, names, selections, credentials = set(), set(), set(), []
    for row in value['connections']:
        if type(row) is not dict or set(row) != {'id', 'name', 'protocol', 'base_url', 'api_key', 'enabled', 'models', 'context_window'}:
            raise SettingsError()
        if not re.fullmatch(ID, _string(row['id'], 64)) or row['id'] in ids or type(row['enabled']) is not bool:
            raise SettingsError()
        ids.add(row['id'])
        _string(row['name'], 80)
        if row['name'].strip().casefold() in names:
            raise SettingsError()
        names.add(row['name'].strip().casefold())
        if row['protocol'] not in ('openai', 'ollama'):
            raise SettingsError()
        _validate({k: row[k] for k in ('base_url', 'api_key', 'context_window')})
        parsed = urlsplit(row['base_url'])
        if not (parsed.scheme == 'https' or row['protocol'] == 'ollama' and parsed.scheme == 'http' and parsed.hostname in ('localhost', '127.0.0.1', '::1')):
            raise SettingsError()
        if type(row['models']) is not list or len(row['models']) > 64:
            raise SettingsError()
        models = set()
        for model in row['models']:
            if type(model) is not dict or not {'id', 'capabilities'} <= set(model) or not set(model) <= {'id', 'capabilities', 'output_parameter'} or not re.fullmatch(MODEL_PATTERN, _string(model['id'], 200)) or model['id'] in models:
                raise SettingsError()
            if 'output_parameter' in model and model['output_parameter'] not in ('auto', 'max_tokens', 'max_completion_tokens', 'num_predict', 'unsupported'):
                raise SettingsError()
            if model.get('output_parameter', 'auto') not in (('auto', 'num_predict', 'unsupported') if row['protocol'] == 'ollama' else ('auto', 'max_tokens', 'max_completion_tokens', 'unsupported')):
                raise SettingsError()
            models.add(model['id'])
            caps = model['capabilities']
            if type(caps) is not list or not caps or any(type(c) is not str or c not in CAPABILITIES for c in caps) or len(set(caps)) != len(caps):
                raise SettingsError()
            if row['protocol'] == 'ollama' and 'image' in caps:
                raise SettingsError()
            if 'image' in caps and not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}', model['id']):
                raise SettingsError()
            for cap in caps:
                selections.add((row['id'], model['id'], cap))
        if row['api_key']:
            credentials.append(row['api_key'])
    if type(value['defaults']) is not dict or not value['defaults'].keys() <= set(CAPABILITIES):
        raise SettingsError()
    for cap, selected in value['defaults'].items():
        if type(selected) is not dict or set(selected) != {'connection_id', 'model_id', 'capability'} or selected.get('capability') != cap:
            raise SettingsError()
        if any(type(selected.get(k)) is not str for k in selected) or tuple(selected[k] for k in ('connection_id', 'model_id', 'capability')) not in selections:
            raise SettingsError()
        if not next(r for r in value['connections'] if r['id'] == selected['connection_id'])['enabled']:
            raise SettingsError()
    public = json.dumps([{k: v for k, v in row.items() if k != 'api_key'} for row in value['connections']], ensure_ascii=False)
    if any(key in public for key in credentials):
        raise SettingsError()
    return value


def _from_values(values):
    raw = values.get(ENV_NAME)
    if raw is not None:
        try:
            return validate_catalog(json.loads(raw))
        except Exception:
            raise SettingsError('settings_file_invalid') from None
    rows, defaults = [], {}
    for provider, suggested in PROVIDER_DEFAULTS.items():
        prefix = provider.upper()
        if not any(values.get(prefix + suffix) for suffix in ('_MODEL', '_BASE_URL', '_API_KEY')) and values.get('MODEL_PROVIDER') != provider:
            continue
        row = {'id': 'legacy-' + provider, 'name': {'deepseek': 'DeepSeek', 'dashscope': '百炼', 'ollama': 'Ollama'}[provider],
               'protocol': 'ollama' if provider == 'ollama' else 'openai', 'base_url': values.get(prefix + '_BASE_URL') or suggested['base_url'],
               'api_key': values.get(prefix + '_API_KEY') or '', 'enabled': True,
               'context_window': int(values.get('OFFICE_CONTEXT_WINDOW') or 8192),
               'models': [{'id': values.get(prefix + '_MODEL') or suggested['model'], 'capabilities': ['chat', 'agent']}]}
        rows.append(row)
        if provider == values.get('MODEL_PROVIDER'):
            for cap in ('chat', 'agent'):
                defaults[cap] = {'connection_id': row['id'], 'model_id': row['models'][0]['id'], 'capability': cap}
    if all(values.get(k) for k in ('IMAGE_MODEL', 'IMAGE_BASE_URL', 'IMAGE_API_KEY')):
        rows.append({'id': 'legacy-image', 'name': '原生图接口', 'protocol': 'openai', 'base_url': values['IMAGE_BASE_URL'],
                     'api_key': values['IMAGE_API_KEY'], 'enabled': True, 'context_window': 8192,
                     'models': [{'id': values['IMAGE_MODEL'], 'capabilities': ['image']}]})
        defaults['image'] = {'connection_id': 'legacy-image', 'model_id': values['IMAGE_MODEL'], 'capability': 'image'}
    try:
        return validate_catalog({'connections': rows, 'defaults': defaults})
    except Exception:
        raise SettingsError('settings_file_invalid') from None


def from_values(values):
    try:
        catalog = _from_values(values)
        public = json.dumps(redact(catalog), ensure_ascii=False)
        if any(v and v in public for k, v in values.items() if any(marker in k.upper() for marker in ('KEY', 'TOKEN', 'SECRET', 'PASSWORD'))):
            raise SettingsError()
        return catalog
    except Exception:
        raise SettingsError('settings_file_invalid') from None


def redact(catalog):
    return json.loads(json.dumps({'connections': [{**{k: v for k, v in row.items() if k != 'api_key'}, 'api_key_configured': bool(row['api_key'])} for row in catalog['connections']],
            'defaults': catalog['defaults']}))


def plan_catalog(store, payload, session, revision):
    store._session(session)
    if store._fused:
        raise SettingsError('settings_uncertain')
    expected = store._connection_snapshots.get(revision) if type(revision) is str else None
    if expected is None:
        raise SettingsError('settings_source_changed')
    source = store._same(store.root / '.env', expected, True)
    old = from_values(store._values(source))
    if type(payload) is not dict or set(payload) != {'connections', 'defaults'} or type(payload['connections']) is not list or len(payload['connections']) > 16:
        raise SettingsError()
    rows, differences = [], []
    old_rows = {r['id']: r for r in old['connections']}
    for supplied in payload['connections']:
        if type(supplied) is not dict:
            raise SettingsError()
        row = dict(supplied)
        action = row.pop('key_action', None)
        key = row.pop('api_key', '')
        previous = old_rows.get(row.get('id'), {})
        if action not in ('keep', 'replace', 'clear') or (action != 'replace' and key) or (action == 'replace' and not key):
            raise SettingsError()
        row['api_key'] = key if action == 'replace' else previous.get('api_key', '') if action == 'keep' else ''
        rows.append(row)
        if row != previous:
            differences.append({'connection': row.get('name'), 'before': {k: v for k, v in previous.items() if k != 'api_key'},
                                'after': {k: v for k, v in row.items() if k != 'api_key'}, 'key_action': action})
    desired_catalog = validate_catalog({'connections': rows, 'defaults': payload['defaults']})
    for cid, row in old_rows.items():
        if cid not in {r['id'] for r in rows}:
            differences.append({'connection': row['name'], 'operation': 'delete'})
    if old['defaults'] != payload['defaults']:
        differences.append({'field': 'defaults', 'before': old['defaults'], 'after': payload['defaults']})
    public_text = json.dumps(differences, ensure_ascii=False)
    if any(row['api_key'] and row['api_key'] in public_text for row in [*old['connections'], *rows]):
        raise SettingsError()
    if old == desired_catalog:
        raise SettingsError('settings_no_changes')
    encoded = json.dumps(desired_catalog, ensure_ascii=True, separators=(',', ':')).replace("'", '\\u0027')
    lines = source.data.decode('utf-8').splitlines(keepends=True)
    from .settings_store import ASSIGNMENT, MAX_PENDING
    if len(store._saved) >= MAX_PENDING:
        raise SettingsError('settings_capacity')
    newline = '\r\n' if b'\r\n' in source.data else '\n'
    output, found = [], False
    for line in lines:
        match = ASSIGNMENT.match(line)
        if match and match[1] == ENV_NAME:
            output.append(ENV_NAME + "='" + encoded + "'" + newline)
            found = True
        else:
            output.append(line)
    if not found:
        if output and not output[-1].endswith(('\n', '\r')):
            output.append(newline)
        output.append(ENV_NAME + "='" + encoded + "'" + newline)
    desired = ''.join(output).encode('utf-8')
    if len(desired) > MAX_ENV_BYTES:
        raise SettingsError('settings_capacity')
    from_values(store._values(_File(desired, ())))
    plan_id = store._id()
    store._pending[plan_id] = _Plan(session, time.monotonic(), source, desired)
    return {'plan_id': plan_id, 'differences': differences, 'confirmation_required': True, 'wrote_nothing': True,
            'expires_in_seconds': TTL_SECONDS, 'restart_required': True, 'web_routes_immediate': True}


@dataclass(frozen=True, repr=False)
class ModelRoute:
    connection_id: str
    connection_name: str
    model_id: str
    capability: str
    revision: str
    protocol: str
    base_url: str
    context_window: int
    api_key: str = field(repr=False)
    epoch: int = field(repr=False)
    output_parameter: str = field(default='auto', repr=False)

    def public(self):
        return {k: getattr(self, k) for k in ('connection_id', 'connection_name', 'model_id', 'capability', 'revision')}


class ConnectionState:
    def __init__(self):
        self._salt = secrets.token_bytes(32)
        self.epoch = 0
        self.selected = {}
        self.mode = 'agent'
        self.outcomes = {}
        self.observed = {}
        self.lists = {}

    def invalidate(self):
        self.epoch += 1
        self.selected.clear()
        self.outcomes.clear()
        self.observed.clear()
        self.lists.clear()
        self.mode = 'agent'

    def identity(self, row):
        return hmac.new(self._salt, json.dumps(row, sort_keys=True, ensure_ascii=True).encode(), hashlib.sha256).hexdigest()[:32]

    def resolve(self, catalog, capability):
        selection = self.selected.get(capability) or catalog['defaults'].get(capability)
        if not selection:
            raise SettingsError('model_selection_required')
        row = next((r for r in catalog['connections'] if r['id'] == selection['connection_id']), None)
        if row is None or not row['enabled'] or row['protocol'] != 'ollama' and not row['api_key']:
            raise SettingsError('model_selection_invalid')
        if not any(m['id'] == selection['model_id'] and capability in m['capabilities'] for m in row['models']):
            raise SettingsError('model_selection_invalid')
        revision = self.identity(row)
        if selection.get('revision', revision) != revision:
            raise SettingsError('model_configuration_changed')
        self.selected[capability] = {**selection, 'revision': revision}
        model = next(m for m in row['models'] if m['id'] == selection['model_id'])
        return ModelRoute(row['id'], row['name'], selection['model_id'], capability, revision, row['protocol'], row['base_url'], row['context_window'], row['api_key'], self.epoch, model.get('output_parameter', 'auto'))

    def public(self, catalog):
        public = redact(catalog)
        for row, source in zip(public['connections'], catalog['connections']):
            revision = row['revision'] = self.identity(source)
            row['model_list_status'] = self.lists.get((row['id'], revision), 'not_requested')
            changed = row['id'] in self.observed and self.observed[row['id']] != revision
            for model in row['models']:
                if changed:
                    for cap in model['capabilities']:
                        self.outcomes.setdefault((row['id'], model['id'], cap, revision), 'changed')
                model['states'] = {cap: self.outcomes.get((row['id'], model['id'], cap, revision), 'unverified') for cap in model['capabilities']}
            self.observed[row['id']] = revision
        current = {}
        for cap in CAPABILITIES:
            try:
                current[cap] = {'valid': True, **self.resolve(catalog, cap).public()}
            except SettingsError as exc:
                current[cap] = {'valid': False, 'reason': exc.code, **(self.selected.get(cap) or catalog['defaults'].get(cap) or {})}
        return {**public, 'current': current, 'mode': self.mode}
