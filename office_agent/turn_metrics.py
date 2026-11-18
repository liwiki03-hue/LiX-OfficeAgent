from __future__ import annotations

from contextlib import contextmanager
from math import isfinite
from threading import RLock
from time import perf_counter
import re
import uuid


STAGES = frozenset({'queue', 'context', 'model', 'authorization', 'tool', 'validation', 'apply', 'reopen', 'render', 'editor', 'publication'})
COUNTERS = frozenset({'decisions', 'model_attempts', 'tool_requests', 'tool_executions', 'tool_reuse', 'tool_failures', 'model_failures', 'roundtrips_avoided'})
OUTCOMES = frozenset({'ok', 'error', 'cancelled', 'timeout', 'rejected', 'stale'})
MARKS = frozenset({'proposal_ready', 'first_effective_answer', 'completed'})
TOKEN_FIELDS = ('input_tokens', 'output_tokens', 'total_tokens', 'reasoning_tokens')


def _get(value, key, default=None):
    return value.get(key, default) if isinstance(value, dict) else getattr(value, key, default)


def _tokens(value):
    return type(value) is int and 0 <= value <= 1000000000


def response_usage(reply):
    metadata = _get(reply, 'response_metadata', {})
    metadata = metadata if type(metadata) is dict else {}
    sources = [_get(reply, 'usage_metadata'), metadata.get('token_usage'), metadata.get('usage'), _get(reply, 'usage')]
    if type(reply) is dict:
        sources.append(reply)
    sources.append(metadata)
    result = {}
    aliases = {
        'input_tokens': ('input_tokens', 'prompt_tokens', 'prompt_eval_count'),
        'output_tokens': ('output_tokens', 'completion_tokens', 'eval_count'),
        'total_tokens': ('total_tokens',),
        'reasoning_tokens': ('reasoning_tokens',),
    }
    for source in sources:
        if type(source) is not dict:
            continue
        for field, names in aliases.items():
            for name in names:
                if field not in result and _tokens(source.get(name)):
                    result[field] = source[name]
        for name, field in (('output_token_details', 'reasoning'), ('completion_tokens_details', 'reasoning_tokens')):
            details = source.get(name)
            if 'reasoning_tokens' not in result and type(details) is dict and _tokens(details.get(field)):
                result['reasoning_tokens'] = details[field]
    return result


def _usage_summary(samples):
    reported = {field: sum(field in sample for sample in samples) for field in TOKEN_FIELDS}
    result = {field: sum(sample[field] for sample in samples if field in sample) if reported[field] else None for field in TOKEN_FIELDS}
    available = sum(bool(sample) for sample in samples)
    result.update({
        'requests': len(samples),
        'reported_requests': available,
        'unavailable_requests': len(samples) - available,
        'field_reported_requests': reported,
        'complete': bool(samples) and all('input_tokens' in sample and 'output_tokens' in sample for sample in samples),
    })
    return result


def _union(intervals):
    end, total = None, 0.0
    for start, finish in sorted(intervals):
        if end is None or start >= end:
            total += finish - start
        elif finish > end:
            total += finish - end
        end = finish if end is None else max(end, finish)
    return total


class TurnMetrics:
    def __init__(self, *, clock=perf_counter, tool_names=()):
        self._clock = clock
        self._lock = RLock()
        self._started = self._now()
        self.turn_id = uuid.uuid4().hex
        self._tools = frozenset()
        self.register_tools(tool_names)
        self._spans = {}
        self._marks = {}
        self._counts = {key: 0 for key in sorted(COUNTERS)}
        self._usage = []
        self._estimates = []

    def register_tools(self, names):
        if isinstance(names, (str, bytes)):
            raise ValueError('metrics_tool_name_invalid')
        try:
            values = tuple(names)
        except TypeError:
            raise ValueError('metrics_tool_name_invalid') from None
        if any(type(name) is not str or re.fullmatch(r'[A-Za-z][A-Za-z0-9_.:-]{0,95}', name) is None for name in values):
            raise ValueError('metrics_tool_name_invalid')
        with self._lock:
            self._tools = self._tools.union(values)

    def _now(self):
        value = self._clock()
        if type(value) not in (int, float) or not isfinite(value):
            raise ValueError('metrics_clock_invalid')
        return float(value)

    def now(self):
        with self._lock:
            return self._now()

    def is_complete(self):
        with self._lock:
            return 'completed' in self._marks

    def _at(self, value, now):
        value = now if value is None else value
        if type(value) not in (int, float) or not isfinite(value) or value < self._started or value > now:
            raise ValueError('metrics_timestamp_invalid')
        return float(value)

    def start(self, stage, *, tool=None, at=None):
        with self._lock:
            if stage not in STAGES or (tool is not None and (stage != 'tool' or tool not in self._tools)):
                raise ValueError('metrics_span_invalid')
            identifier = uuid.uuid4().hex
            self._spans[identifier] = {'id': identifier, 'stage': stage, 'tool': tool, 'start': self._at(at, self._now()), 'end': None, 'outcome': None}
            return identifier

    def stop(self, identifier, *, outcome='ok', at=None):
        with self._lock:
            if identifier not in self._spans or outcome not in OUTCOMES:
                raise ValueError('metrics_span_invalid')
            span = self._spans[identifier]
            if span['end'] is not None:
                raise ValueError('metrics_span_finished')
            end = self._at(at, self._now())
            if end < span['start']:
                raise ValueError('metrics_timestamp_invalid')
            span.update(end=end, outcome=outcome)

    @contextmanager
    def span(self, stage, *, tool=None, at=None):
        identifier = self.start(stage, tool=tool, at=at)
        try:
            yield identifier
        except BaseException:
            self.stop(identifier, outcome='error')
            raise
        else:
            self.stop(identifier)

    def count(self, name, amount=1):
        with self._lock:
            if name not in COUNTERS or type(amount) is not int or amount < 0:
                raise ValueError('metrics_count_invalid')
            self._counts[name] += amount

    def mark(self, name, *, at=None):
        with self._lock:
            if name not in MARKS:
                raise ValueError('metrics_mark_invalid')
            value = self._at(at, self._now())
            self._marks.setdefault(name, value)

    def record_usage(self, reply):
        with self._lock:
            self._usage.append(response_usage(reply))

    def record_estimate(self, *, input_tokens=None, output_tokens=None, total_tokens=None):
        values = {'input_tokens': input_tokens, 'output_tokens': output_tokens, 'total_tokens': total_tokens}
        if any(value is not None and not _tokens(value) for value in values.values()):
            raise ValueError('metrics_estimate_invalid')
        with self._lock:
            self._estimates.append({key: value for key, value in values.items() if value is not None})

    def snapshot(self):
        with self._lock:
            now = self._now()
            end = self._marks.get('completed', now)
            spans, intervals = [], {stage: [] for stage in sorted(STAGES)}
            outcomes = {outcome: 0 for outcome in sorted(OUTCOMES)}
            for span in self._spans.values():
                finish = min(span['end'] if span['end'] is not None else end, end)
                finish = max(span['start'], finish)
                intervals[span['stage']].append((span['start'], finish))
                row = {'id': span['id'], 'stage': span['stage'], 'start_seconds': span['start'] - self._started,
                       'duration_seconds': finish - span['start'], 'active': span['end'] is None, 'outcome': span['outcome']}
                if span['tool'] is not None:
                    row['tool'] = span['tool']
                if span['outcome'] is not None:
                    outcomes[span['outcome']] += 1
                spans.append(row)
            stage_wall = {stage: _union(values) for stage, values in intervals.items()}
            observed = _union([interval for values in intervals.values() for interval in values])
            elapsed = max(0.0, end - self._started)
            return {
                'schema': 1,
                'turn_id': self.turn_id,
                'elapsed_seconds': elapsed,
                'marks_seconds': {name: self._marks[name] - self._started if name in self._marks else None
                                  for name in sorted(MARKS) if name != 'proposal_ready' or name in self._marks},
                'stage_wall_seconds': stage_wall,
                'observed_wall_seconds': observed,
                'unattributed_wall_seconds': max(0.0, elapsed - observed),
                'cross_stage_overlap_seconds': max(0.0, sum(stage_wall.values()) - observed),
                'counts': dict(self._counts),
                'outcomes': outcomes,
                'spans': spans,
                'usage': {'server': _usage_summary(self._usage), 'local_estimate': _usage_summary(self._estimates)},
            }
