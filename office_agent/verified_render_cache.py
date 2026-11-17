from collections import OrderedDict
import hashlib
import json
import threading
import time


_entries = OrderedDict()
_lock = threading.RLock()
_limit = 32 * 1024 * 1024


def key(root, input_sha256, projection_sha256, runtime_sha256, environment, implementation):
    if not environment.get('complete'):
        return None
    data = [str(root.resolve()).casefold(), input_sha256, projection_sha256, runtime_sha256,
            environment, implementation, 'static_docx_pdf_v1']
    return hashlib.sha256(json.dumps(data, sort_keys=True).encode('utf-8')).hexdigest()


def _remove(identifier):
    entry = _entries.pop(identifier, None)
    if entry:
        entry['timer'].cancel()


def clear():
    with _lock:
        for identifier in list(_entries):
            _remove(identifier)


def _expire(identifier, deadline):
    with _lock:
        entry = _entries.get(identifier)
        if entry and entry['deadline'] == deadline:
            _remove(identifier)


def get(identifier):
    if identifier is None:
        return None
    with _lock:
        entry = _entries.get(identifier)
        if not entry:
            return None
        if entry['deadline'] <= time.monotonic() or hashlib.sha256(entry['data']).hexdigest() != entry['sha256']:
            _remove(identifier)
            return None
        _entries.move_to_end(identifier)
        return entry['data']


def offer(identifier, data):
    if identifier is None or not isinstance(data, bytes) or not data.startswith(b'%PDF-') or len(data) > _limit // 2:
        return False
    with _lock:
        _remove(identifier)
        while _entries and (len(_entries) >= 4 or sum(len(entry['data']) for entry in _entries.values()) + len(data) > _limit):
            _remove(next(iter(_entries)))
        deadline = time.monotonic() + 300
        timer = threading.Timer(300, _expire, args=(identifier, deadline))
        timer.daemon = True
        _entries[identifier] = {'data': data, 'sha256': hashlib.sha256(data).hexdigest(), 'deadline': deadline, 'timer': timer}
        timer.start()
        return True
