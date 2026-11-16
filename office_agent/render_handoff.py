import atexit
import hashlib
import json
from pathlib import Path
import sys
import threading
import time


MAX_BYTES = 32 * 1024 * 1024
SECONDS = 300
_lock = threading.RLock()
_entry = None
_previous = None


def implementation_identity():
    from .writes import digest
    folder = Path(__file__).parent
    names = ('fidelity.py', 'fidelity_worker.py', 'range_edits.py', 'range_alignment.py',
             'render_identity.py', 'render_handoff.py', 'layout_checks.py','pdf_reading_order.py')
    return hashlib.sha256(json.dumps([digest(Path(sys.executable)),
        {name: digest(folder / name) for name in names}], sort_keys=True).encode()).hexdigest()


def _key(root, report):
    if (report.get('inputs_unchanged') is not True or not report.get('render_environment', {}).get('complete')
            or report.get('render', {}).get('status') not in {'equivalent', 'different'}
            or len(report.get('static_projection_sha256', [])) != 2
            or report.get('conversion_identity') != implementation_identity()):
        return None
    fields = ('source_sha256', 'output_sha256', 'static_projection_sha256',
              'runtime_sha256', 'render_environment', 'conversion_identity')
    return json.dumps([str(Path(root).resolve()), {name: report[name] for name in fields}], sort_keys=True)


def clear():
    with _lock:
        for entry in _entries():
            _remove(entry)


def _entries():
    return tuple(entry for entry in (_entry, _previous) if entry is not None)


def _remove(entry):
    global _entry, _previous
    entry['timer'].cancel()
    if _entry is entry:
        _entry = None
    if _previous is entry:
        _previous = None


def _expire(key, deadline):
    with _lock:
        for entry in _entries():
            if entry['key'] == key and entry['deadline'] == deadline:
                _remove(entry)


def offer(root, report, pdfs):
    global _entry, _previous
    from .fidelity import bounded_bytes
    try:
        key = _key(root, report)
        if key is None:
            return False
        paths = tuple(pdfs)
        if len(paths) != 2 or any(path.is_symlink() or not path.is_file() for path in paths) or sum(path.stat().st_size for path in paths) > MAX_BYTES:
            return False
        payloads = tuple(bounded_bytes(path) for path in paths)
        if len(payloads) != 2 or sum(map(len, payloads)) > MAX_BYTES or any(not data.startswith(b'%PDF-') for data in payloads):
            return False
        with _lock:
            now = time.monotonic()
            for entry in _entries():
                if entry['deadline'] <= now or entry['key'] == key:
                    _remove(entry)
            retained = next((entry for entry in _entries() if sum(map(len, entry['payloads'])) + sum(map(len, payloads)) <= MAX_BYTES), None)
            for entry in _entries():
                if entry is not retained:
                    _remove(entry)
            _previous = retained
            deadline = now + SECONDS
            timer = threading.Timer(SECONDS, _expire, args=(key, deadline))
            timer.daemon = True
            _entry = {'key': key, 'payloads': payloads, 'hashes': tuple(hashlib.sha256(data).hexdigest() for data in payloads),
                      'deadline': deadline, 'timer': timer}
            try:
                timer.start()
            except RuntimeError:
                _remove(_entry)
                return False
        return True
    except (OSError, ValueError):
        return False


def take(root, report):
    with _lock:
        if not _entries():
            return None
        try:
            now = time.monotonic()
            for entry in _entries():
                if now >= entry['deadline']:
                    _remove(entry)
            key = _key(root, report)
            entry = next((item for item in _entries() if item['key'] == key), None)
            if entry is None:
                clear()
                return None
            try:
                if tuple(hashlib.sha256(data).hexdigest() for data in entry['payloads']) != entry['hashes']:
                    return None
                return entry['payloads']
            finally:
                _remove(entry)
        except (OSError, ValueError):
            clear()
            return None


atexit.register(clear)
