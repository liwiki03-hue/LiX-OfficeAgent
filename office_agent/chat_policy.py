import hashlib
import json
from urllib.parse import urlsplit

from .settings_store import SettingsError


DEFAULT_REPLY_TOKENS = 512
MAX_ANSWER_CHARACTERS = 32000
MAX_STREAM_CHUNKS = 2048
STREAM_SECONDS = 90


def output_parameter(route):
    choice = getattr(route, 'output_parameter', 'auto')
    if choice == 'unsupported':
        raise SettingsError('chat_output_parameter_unsupported')
    if choice != 'auto':
        return choice
    if route.protocol == 'ollama':
        return 'num_predict'
    endpoint = urlsplit(route.base_url)
    if (endpoint.scheme == 'https' and endpoint.hostname == 'api.deepseek.com' and
            endpoint.port in (None, 443) and endpoint.path.rstrip('/') in ('', '/v1') and
            not endpoint.username and not endpoint.password and not endpoint.query and not endpoint.fragment):
        return 'max_tokens'
    return 'max_completion_tokens'


def reply_policy(values, route=None):
    raw = values.get('OFFICE_CHAT_REPLY_TOKENS', '512')
    if type(raw) is not str or not raw.isascii() or not raw.isdecimal() or not 64 <= int(raw) <= 16384:
        raise SettingsError('settings_file_invalid')
    parameter = output_parameter(route) if route else None
    return {'reply_token_reserve': int(raw), 'output_parameter': parameter,
            'configured': 'OFFICE_CHAT_REPLY_TOKENS' in values,
            'max_answer_characters': MAX_ANSWER_CHARACTERS, 'max_stream_chunks': MAX_STREAM_CHUNKS,
            'stream_seconds': STREAM_SECONDS}


def turn_digest(turn):
    return hashlib.sha256(json.dumps({key: turn[key] for key in ('id', 'question', 'answer', 'route', 'status', 'finish_reason')},
                                    ensure_ascii=False, sort_keys=True).encode('utf8')).hexdigest()


def continuation_chain(row, binding):
    if type(binding) is not dict or set(binding) != {'session_id', 'turn_id', 'source_digest'}:
        raise SettingsError('chat_continuation_invalid')
    if not row or binding['session_id'] != row['id']:
        raise SettingsError('chat_continuation_changed')
    parent = next((turn for turn in row['turns'] if turn['id'] == binding['turn_id']), None)
    if parent is None or turn_digest(parent) != binding['source_digest']:
        raise SettingsError('chat_continuation_changed')
    if parent['status'] != 'complete' or not parent['answer'] or parent['finish_reason'] not in ('length', 'unknown', 'local_limit'):
        raise SettingsError('chat_continuation_unavailable')
    chain, seen = [], set()
    current = parent
    while current:
        if current['id'] in seen or len(chain) >= 16:
            raise SettingsError('chat_continuation_limit')
        seen.add(current['id'])
        chain.insert(0, current)
        link = current.get('continuation_of')
        if not link:
            break
        ancestor = next((turn for turn in row['turns'] if turn['id'] == link['turn_id']), None)
        if ancestor is None or turn_digest(ancestor) != link['source_digest']:
            raise SettingsError('chat_continuation_changed')
        current = ancestor
    return parent, chain


def public_continuations(row):
    return [{'turn_id': turn['id'], 'source_digest': turn_digest(turn)} for turn in row['turns']
            if turn['status'] == 'complete' and turn['answer'] and turn['finish_reason'] in ('length', 'unknown', 'local_limit')]


def safe_usage(reply):
    usage = getattr(reply, 'usage_metadata', None)
    metadata = getattr(reply, 'response_metadata', {}) or {}
    if type(usage) is not dict:
        usage = metadata.get('token_usage') or {}
    result = {}
    for name, alias in [('input_tokens', 'prompt_tokens'), ('output_tokens', 'completion_tokens'), ('total_tokens', 'total_tokens')]:
        value = usage.get(name, usage.get(alias))
        if type(value) is int and 0 <= value <= 1000000000:
            result[name] = value
    if not result:
        for name, key in [('input_tokens', 'prompt_eval_count'), ('output_tokens', 'eval_count')]:
            value = metadata.get(key)
            if type(value) is int and 0 <= value <= 1000000000:
                result[name] = value
    return result
