"""Host-only local renderer, not yet mounted in production UI/model routes.

No arbitrary commands or disk payloads. SVG is untrusted until host validation.
Job containment is NOT a network/filesystem security sandbox.
"""
from dataclasses import dataclass
import hashlib
import hmac
from pathlib import Path
import secrets
import socket
import sys
import threading
import time
from typing import Callable

from .diagram_channel import MAGIC, read_exact, receive_message, send_message
from .diagram_process import OwnedRun, _run_owned
from .diagrams import DiagramError, DiagramSource, validate_source, validate_svg

_slots = threading.BoundedSemaphore(2)
_disabled = threading.Event()  # Restart required after unconfirmed receiver cleanup.


@dataclass(frozen=True)
class RenderedDiagram:
    source: DiagramSource
    svg: str
    execution: OwnedRun

    def evidence(self):
        return {**self.source.evidence(), 'rendered': True, 'svg_bytes': len(self.svg.encode('utf-8')),
                'svg_sha256': hashlib.sha256(self.svg.encode('utf-8')).hexdigest(),
                'empty_confirmed': self.execution.empty_confirmed,
                'elapsed_seconds': self.execution.elapsed_seconds}


def _worker_command(port, key):
    return (sys.executable, str(Path(__file__).with_name('diagram_worker.py').resolve()),
            str(port), key.hex())


def _validated_response(payload):
    if type(payload) is not dict or type(payload.get('ok')) is not bool:
        raise DiagramError('renderer_failed')
    if payload['ok'] is False:
        if set(payload) != {'ok', 'code'}:
            raise DiagramError('renderer_failed')
        code = payload['code'] if payload['code'] in ('unsafe_svg', 'unsupported_source', 'render_timeout') else 'renderer_failed'
        raise DiagramError(code)
    if set(payload) != {'ok', 'svg', 'bridge_absent'} or payload['bridge_absent'] is not True:
        raise DiagramError('renderer_failed')
    return validate_svg(payload['svg'])


def render_diagram(source: str, *, expected_kind: str | None = None,
                   timeout_ms: int = 8000, cancelled: Callable[[], bool] | None = None) -> RenderedDiagram:
    """Read-only host preview. Never exports files or confers write authority.

    Caller must bind cancellation/results to its session/epoch before UI mounting.
    Uses the current interpreter; startup capability is not inferred from versions.
    """
    if (type(timeout_ms) is not int or not 100 <= timeout_ms <= 8000
            or cancelled is not None and not callable(cancelled)):
        raise DiagramError('unsupported_source')
    spec = validate_source(source, expected_kind)
    if cancelled is not None and cancelled():
        raise DiagramError('render_cancelled')
    if _disabled.is_set():
        raise DiagramError('renderer_cleanup_failed')
    if not _slots.acquire(blocking=False):
        raise DiagramError('renderer_failed')
    listener = thread = None
    stop, done, fault, user_cancel, expired = (threading.Event() for _ in range(5))
    state = {'payload': None, 'connection': None}
    deadline = time.monotonic() + timeout_ms / 1000
    key = secrets.token_bytes(32)
    try:
        if _disabled.is_set():
            raise DiagramError('renderer_cleanup_failed')
        listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        if hasattr(socket, 'SO_EXCLUSIVEADDRUSE'):
            listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        listener.bind(('127.0.0.1', 0))
        listener.listen(2)
        listener.settimeout(0.1)
        def receive():
            try:
                # Invalid authentication never receives diagram text. Attempts bounded.
                attempts = 0
                while attempts < 8:
                    if stop.is_set() or time.monotonic() >= deadline:
                        raise TimeoutError('renderer connection deadline')
                    try:
                        connection, address = listener.accept()
                    except socket.timeout:
                        continue
                    attempts += 1
                    with connection:
                        state['connection'] = connection
                        connection.settimeout(0.1)
                        if address[0] != '127.0.0.1':
                            continue
                        header = read_exact(connection, len(MAGIC) + len(key), deadline, stop)
                        if not hmac.compare_digest(header, MAGIC + key):
                            continue
                        send_message(connection, {'source': spec.source, 'kind': spec.kind})
                        state['payload'] = receive_message(connection, deadline, stop)
                        return
                raise ValueError('renderer authentication failed')
            except TimeoutError:
                if not stop.is_set():
                    expired.set()
            except Exception:
                if not stop.is_set():
                    fault.set()
            finally:
                done.set()
        thread = threading.Thread(target=receive, name='owned-diagram-channel', daemon=False)
        thread.start()
        def abort():
            if cancelled is not None and cancelled():
                user_cancel.set()
            return user_cancel.is_set() or expired.is_set() or fault.is_set()
        try:
            execution = _run_owned(_worker_command(listener.getsockname()[1], key),
                                   timeout_ms=max(100, int(max(0, deadline - time.monotonic()) * 1000)),
                                   cancelled=abort)
        except DiagramError as exc:
            if exc.code == 'renderer_cleanup_failed':
                _disabled.set()
            raise
        if execution.empty_confirmed is not True:
            _disabled.set()
            raise DiagramError('renderer_cleanup_failed')
        if execution.code in ('process_completed', 'process_failed'):
            done.wait(min(1.0, max(0.0, deadline - time.monotonic())))
        if user_cancel.is_set():
            raise DiagramError('render_cancelled')
        if expired.is_set() or execution.code == 'render_timeout':
            raise DiagramError('render_timeout')
        if fault.is_set() or not done.is_set():
            raise DiagramError('renderer_failed')
        safe = _validated_response(state['payload'])
        if execution.code != 'process_completed' or not execution.empty_confirmed:
            raise DiagramError('renderer_failed')
        return RenderedDiagram(spec, safe, execution)
    finally:
        stop.set()
        try:
            if listener is not None:
                listener.close()
            connection = state['connection']
            if connection is not None:
                try:
                    connection.shutdown(socket.SHUT_RDWR)
                except OSError:
                    pass
                connection.close()
            if thread is not None:
                thread.join(1.0)
                if thread.is_alive():
                    _disabled.set()
                    raise DiagramError('renderer_cleanup_failed')
        finally:
            _slots.release()
