from collections import deque
from contextlib import contextmanager
from contextvars import ContextVar
from functools import wraps
import threading
import time
import uuid


_current = ContextVar('office_review_diagnostic', default=None)


@contextmanager
def review_phase(name):
    trace = _current.get()
    began = time.monotonic()
    record = {'stage': name, 'status': 'running'}
    if trace is not None and len(trace['stages']) < 64:
        trace['stages'].append(record)
    try:
        yield
        record['status'] = 'returned'
    except BaseException as error:
        record.update(status='failed', error_type=type(error).__name__)
        raise
    finally:
        record['milliseconds'] = round((time.monotonic() - began) * 1000, 3)


def review_stage(name):
    def decorate(action):
        @wraps(action)
        def measured(*args, **kwargs):
            with review_phase(name):
                return action(*args, **kwargs)
        return measured
    return decorate


def review_request(action):
    @wraps(action)
    def measured(self, *args, **kwargs):
        trace = {'id': uuid.uuid4().hex, 'operation': action.__name__, 'status': 'running', 'stages': []}
        with self.diagnostic_lock:
            self.diagnostics.append(trace)
        token = _current.set(trace)
        began = time.monotonic()
        try:
            result = action(self, *args, **kwargs)
            trace['status'] = 'returned' if isinstance(result, dict) and result.get('ok', True) else 'blocked'
            return result
        except BaseException as error:
            trace.update(status='failed', error_type=type(error).__name__)
            raise
        finally:
            trace['milliseconds'] = round((time.monotonic() - began) * 1000, 3)
            _current.reset(token)
    return measured


def initialize_review_diagnostics(review):
    review.diagnostic_lock = threading.Lock()
    review.diagnostics = deque(maxlen=16)
