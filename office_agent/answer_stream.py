from __future__ import annotations

import uuid
import json

from .turn_metrics import response_usage
from .visible_answer import LeadingThoughtFilter, visible_answer


class StreamError(ValueError):
    pass


def _get(value, key, default=None):
    return value.get(key, default) if isinstance(value, dict) else getattr(value,key,default)


def _text(content):
    if isinstance(content,str):
        return content
    if isinstance(content,list):
        if any(not isinstance(item,dict) or item.get('type') not in {'text','text_delta'} or not isinstance(item.get('text'),str) for item in content):
            raise StreamError('unsupported_stream_content')
        return ''.join(item['text'] for item in content)
    raise StreamError('unsupported_stream_content')


def _result(response, content):
    result = {'content': content}
    if _get(response, 'tool_calls') is not None:
        result['tool_calls'] = _get(response, 'tool_calls')
    metadata = _get(response, 'response_metadata', {}) or {}
    terminal = {key: metadata[key] for key in ('finish_reason', 'stop_reason')
                if key in metadata and metadata[key] in {'stop', 'end_turn', 'tool_calls', 'tool_use', 'length', 'max_tokens', 'content_filter', 'error', 'unknown'}}
    if terminal:
        result['response_metadata'] = terminal
    usage = response_usage(response)
    if usage:
        result['usage_metadata'] = usage
    return result


def _terminal(response, *, called=False):
    if _get(response, 'invalid_tool_calls', []):
        raise StreamError('incomplete_stream')
    metadata = _get(response, 'response_metadata', {}) or {}
    finish = metadata.get('finish_reason', metadata.get('stop_reason'))
    if finish is not None and finish not in {'stop', 'end_turn', 'tool_calls', 'tool_use'}:
        raise StreamError('incomplete_stream')
    if (finish in {'tool_calls', 'tool_use'} or called) and not _get(response, 'tool_calls', []):
        raise StreamError('incomplete_tool_stream')


def streamed_response(bound, messages, callback, stop):
    if callback is None or not callable(getattr(bound,'stream',None)):
        response = bound.invoke(messages)
        _terminal(response)
        return _result(response, visible_answer(_text(_get(response, 'content', ''))))
    identifier = uuid.uuid4().hex
    active = True
    def emit(kind, **values):
        if callback({'type':kind,'stream_id':identifier,**values}) is False:
            raise StreamError('stream_scope_expired')
    iterator = None
    visibility = LeadingThoughtFilter()
    visible = ''
    try:
        emit('answer_draft_start')
        iterator = iter(bound.stream(messages))
        response, count, length, called = None, 0, 0, False
        for chunk in iterator:
            if stop():
                raise StreamError('stream_cancelled')
            count += 1
            if count > 2048:
                raise StreamError('stream_chunk_limit')
            content = _text(_get(chunk,'content',''))
            length += len(content)
            if length > 20000:
                raise StreamError('stream_answer_limit')
            fragment = visibility.feed(content)
            visible += fragment
            additional = _get(chunk,'additional_kwargs',{}) or {}
            tool_chunk = bool(_get(chunk,'tool_calls',[]) or _get(chunk,'tool_call_chunks',[]) or _get(chunk,'invalid_tool_calls',[]) or additional.get('function_call') or additional.get('tool_calls'))
            if tool_chunk:
                called = True
                if active:
                    emit('answer_draft_cancel')
                    active = False
            if active and fragment:
                emit('answer_draft_delta',text=fragment)
            if isinstance(chunk,dict):
                if response is None:
                    response = {'content':'','tool_calls':[], 'response_metadata':{}, 'usage_metadata':{}, 'invalid_tool_calls':[]}
                if not isinstance(response,dict):
                    raise StreamError('mixed_stream_types')
                response['content'] += content
                response['tool_calls'].extend(chunk.get('tool_calls',[]))
                response['response_metadata'].update(chunk.get('response_metadata',{}))
                response['invalid_tool_calls'].extend(chunk.get('invalid_tool_calls',[]))
                for key, value in response_usage(chunk).items():
                    response['usage_metadata'][key] = response['usage_metadata'].get(key, 0) + value
            else:
                if isinstance(response,dict):
                    raise StreamError('mixed_stream_types')
                response = chunk if response is None else response+chunk
        if stop():
            raise StreamError('stream_cancelled')
        tail = visibility.finish()
        visible += tail
        if active and tail:
            emit('answer_draft_delta', text=tail)
        if response is None:
            raise StreamError('incomplete_stream')
        _terminal(response, called=called)
        raw_calls = _get(response,'tool_call_chunks',[]) or []
        def unique(pairs):
            result = {}
            for key,value in pairs:
                if key in result:
                    raise StreamError('ambiguous_tool_stream')
                result[key] = value
            return result
        for raw in raw_calls:
            try:
                arguments = json.loads(raw['args'],object_pairs_hook=unique)
            except (ValueError,KeyError,TypeError):
                raise StreamError('incomplete_tool_stream') from None
            if not isinstance(arguments,dict) or not any(call.get('name') == raw.get('name') and call.get('id') == raw.get('id') and call.get('args') == arguments for call in _get(response,'tool_calls',[])):
                raise StreamError('incomplete_tool_stream')
        return _result(response, visible)
    except Exception as error:
        if active:
            try:
                emit('answer_draft_cancel')
            except Exception:
                pass
        if isinstance(error,StreamError):
            raise
        raise StreamError('stream_failed') from None
    finally:
        close = getattr(iterator,'close',None)
        if callable(close):
            close()
