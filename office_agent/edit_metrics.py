from contextlib import contextmanager
from functools import wraps
from threading import RLock


PHASES = {'preparing': 'context', 'serializing': 'context', 'backup': 'publication',
          'working_file': 'publication', 'reopening': 'reopen', 'rendering': 'render',
          'publishing': 'publication', 'validation': 'validation'}


class EditMetrics:
    def __init__(self, collector, conversation_id, current):
        self.collector = collector
        self.conversation_id = conversation_id
        self.id = collector.turn_id
        self._current = current
        self._lock = RLock()
        self._spans = set()
        self._phase = None

    def current(self):
        try:
            return bool(self._current()) and not self.collector.is_complete()
        except Exception:
            return False

    def start(self, stage):
        with self._lock:
            if not self.current():
                return None
            try:
                token = self.collector.start(stage)
            except Exception:
                return None
            self._spans.add(token)
            return token

    def stop(self, token, *, outcome='ok'):
        with self._lock:
            if token not in self._spans:
                return
            self._spans.remove(token)
            try:
                self.collector.stop(token, outcome=outcome)
            except Exception:
                pass

    @contextmanager
    def measure(self, stage):
        token = self.start(stage)
        try:
            yield token
        except BaseException:
            self.stop(token, outcome='error')
            raise
        else:
            self.stop(token)

    def phase(self, name):
        stage = PHASES.get(name)
        if stage is None:
            return
        with self._lock:
            self.finish_phase()
            self._phase = self.start(stage)

    def finish_phase(self, *, outcome='ok'):
        with self._lock:
            token, self._phase = self._phase, None
            self.stop(token, outcome=outcome)

    def finish(self, *, outcome='ok'):
        with self._lock:
            self._phase = None
            for token in list(self._spans):
                self.stop(token, outcome=outcome)


def for_edit(floating, conversation_id):
    conversation = getattr(floating, 'conversation', None)
    method = getattr(conversation, 'edit_trace', None)
    if not conversation_id or not callable(method):
        return None
    try:
        return method(conversation_id)
    except Exception:
        return None


@contextmanager
def measure(trace, stage):
    if trace is None:
        yield None
    else:
        with trace.measure(stage) as token:
            yield token


def publication_metrics(method):
    @wraps(method)
    def measured(self, *args, **kwargs):
        trace = getattr(self, 'metrics_trace', None)
        if trace is not None:
            trace.phase('validation')
        outcome = 'ok'
        try:
            return method(self, *args, **kwargs)
        except BaseException:
            outcome = 'error'
            raise
        finally:
            if trace is not None:
                trace.finish_phase(outcome=outcome)
    return measured
