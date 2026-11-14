import atexit
from datetime import datetime, timezone
import json
from logging.handlers import RotatingFileHandler
import logging
from pathlib import Path
import queue
import re
import threading
import time


FIELDS = frozenset({'module', 'action', 'correlation', 'elapsed_ms', 'result', 'code', 'build', 'host', 'expected_sha256', 'actual_sha256', 'received', 'attempts'})
_instances = {}
_lock = threading.RLock()


class _Handler(RotatingFileHandler):
    def handleError(self, record):
        raise OSError('diagnostic_io_failed')


class Diagnostics:
    def __init__(self, root, *, limit=5 * 1024 * 1024, backups=2):
        self.path = Path(root).absolute() / '.runtime/diagnostics/operations.jsonl'
        self.limit, self.backups = limit, backups
        self.queue = queue.Queue(128)
        self.stop = threading.Event()
        self.dropped = 0
        self.error = ''
        self.thread = None

    def _safe(self):
        from .downloads import safe
        for name in [self.path, *[self.path.with_name(self.path.name + '.' + str(n)) for n in range(1, self.backups + 1)]]:
            safe(name)

    def record(self, **fields):
        if self.stop.is_set():
            return False
        row = {'time': datetime.now(timezone.utc).isoformat()}
        for key, value in fields.items():
            if key not in FIELDS:
                continue
            if key in {'elapsed_ms', 'received', 'attempts'}:
                if type(value) is int and 0 <= value <= (86400000 if key == 'elapsed_ms' else 3 if key == 'attempts' else 1024 * 1024 * 1024):
                    row[key] = value
            elif key in {'expected_sha256', 'actual_sha256'}:
                if isinstance(value, str) and re.fullmatch('[a-f0-9]{64}', value):
                    row[key] = value
            elif key == 'host':
                if isinstance(value, str) and re.fullmatch(r'[a-zA-Z0-9.-]{1,253}', value):
                    row[key] = value
            elif isinstance(value, str) and re.fullmatch(r'[A-Za-z0-9_.:-]{1,96}', value):
                row[key] = value
        if not {'module', 'action', 'result'} <= row.keys():
            return False
        try:
            self.queue.put_nowait(row)
        except queue.Full:
            self.dropped += 1
            return False
        with _lock:
            if self.thread is None:
                self.thread = threading.Thread(target=self._write, name='lix-diagnostics', daemon=True)
                self.thread.start()
        return True

    def _write(self):
        handler = None
        try:
            self._safe()
            self.path.parent.mkdir(parents=True, exist_ok=True)
            handler = _Handler(self.path, maxBytes=self.limit, backupCount=self.backups, encoding='utf-8')
            while not self.stop.is_set() or not self.queue.empty():
                try:
                    row = self.queue.get(timeout=.1)
                except queue.Empty:
                    continue
                try:
                    self._safe()
                    handler.emit(logging.LogRecord('lix.operations', logging.INFO, '', 0, json.dumps(row, ensure_ascii=True), (), None))
                    handler.flush()
                except OSError:
                    self.error = 'diagnostic_io_failed'
                finally:
                    self.queue.task_done()
        except (OSError, ValueError):
            self.error = 'diagnostic_unavailable'
        finally:
            if handler is not None:
                handler.close()

    def preview(self):
        self._safe()
        rows = []
        if self.path.is_file():
            with self.path.open('rb') as stream:
                stream.seek(max(0, self.path.stat().st_size - 65536))
                if stream.tell():
                    stream.readline()
                for line in stream:
                    if len(line) <= 2048:
                        try:
                            row = json.loads(line)
                            if isinstance(row, dict) and row.keys() <= FIELDS | {'time'}:
                                rows.append(row)
                        except (ValueError, UnicodeError):
                            pass
        return {'events': rows[-100:], 'dropped': self.dropped, 'code': self.error,
                'path': str(self.path), 'includes_documents': False, 'includes_configuration': False}

    def close(self):
        self.stop.set()
        if self.thread is not None:
            self.thread.join(2)
        return self.thread is None or not self.thread.is_alive()


def get(root=None):
    root = str(Path(root or Path.cwd()).absolute())
    with _lock:
        if root not in _instances:
            _instances[root] = Diagnostics(root)
        return _instances[root]


def record(module, action, result, **fields):
    return get().record(module=module, action=action, result=result, **fields)


@atexit.register
def close_all():
    for instance in list(_instances.values()):
        instance.close()
