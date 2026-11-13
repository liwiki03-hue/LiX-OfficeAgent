"""Host-only raster receipt from a current native SVG preview, never a browser image.

The restricted SVG is sent only to an owned fixed-command QtSvg child. A job
limits lifetime/descendants, NOT filesystem/network access. PNG is decoded and
canonicalized on the host before it can become save authority.
"""
import base64
from dataclasses import dataclass
import hashlib
import hmac
from io import BytesIO
from pathlib import Path
import re
import secrets
import socket
import sys
import threading
import time
from xml.etree import ElementTree as ET

from PIL import Image, UnidentifiedImageError

from .diagram_channel import MAGIC, read_exact, receive_message, send_message
from .diagram_process import _run_owned
from .diagrams import DiagramError, validate_svg

MAX_PNG_BYTES = 1_000_000
_disabled = threading.Event()
_slots = threading.BoundedSemaphore(2)


@dataclass(frozen=True)
class RasterReceipt:
    preview: object  # Exact host-held RenderedDiagram, not client-reconstructible.
    png: bytes
    sha256: str
    width: int
    height: int
    renderer: str = 'host-qt-svg-v1'


def normalize_png(data, *, expected_size=None):
    if type(data) is not bytes or not 0 < len(data) <= MAX_PNG_BYTES or not data.startswith(b'\x89PNG\r\n\x1a\n'):
        raise DiagramError('raster_invalid')
    try:
        with Image.open(BytesIO(data)) as image:
            if (image.format != 'PNG' or getattr(image, 'n_frames', 1) != 1
                    or image.mode not in {'RGBA', 'RGB', 'P', 'L', 'LA'}):
                raise ValueError()
            w, h = image.size
            if not 1 <= w <= 2048 or not 1 <= h <= 2048 or w * h > 4_194_304:
                raise ValueError()
            if expected_size is not None and (w, h) != expected_size:
                raise ValueError()
            image.load()
            # Re-encode pixels only: no worker-supplied PNG metadata survives.
            clean = image.convert('RGB')
            sink = BytesIO()
            clean.save(sink, format='PNG')
            encoded = sink.getvalue()
            if len(encoded) > MAX_PNG_BYTES:
                raise ValueError()
            return encoded, w, h
    except (OSError, ValueError, UnidentifiedImageError, Image.DecompressionBombError):
        raise DiagramError('raster_invalid') from None


def verify_png(data, width, height):
    """Disk bytes must be exactly the canonical host PNG, not just decodable."""
    canonical, w, h = normalize_png(data, expected_size=(width, height))
    if canonical != data or (w, h) != (width, height):
        raise DiagramError('raster_invalid')
    return True


def rasterize(preview, *, timeout_ms=8000):
    from .diagram_service import RenderedDiagram
    from .diagram_process import OwnedRun
    if (type(preview) is not RenderedDiagram or type(preview.execution) is not OwnedRun
            or preview.execution.code != 'process_completed' or preview.execution.exit_code != 0
            or preview.execution.empty_confirmed is not True or type(timeout_ms) is not int
            or not 100 <= timeout_ms <= 8000):
        raise DiagramError('raster_invalid')
    svg = validate_svg(preview.svg)
    if svg != preview.svg:
        raise DiagramError('raster_invalid')
    view = ET.fromstring(svg).attrib['viewBox'].split()
    width, height = (int(round(float(n))) for n in view[2:])
    if not 1 <= width <= 2048 or not 1 <= height <= 2048 or width * height > 4_194_304:
        raise DiagramError('raster_invalid')
    if _disabled.is_set():
        raise DiagramError('renderer_cleanup_failed')
    if not _slots.acquire(blocking=False):
        raise DiagramError('raster_busy')
    listener = thread = None
    stop, done, fault = threading.Event(), threading.Event(), threading.Event()
    state = {'payload': None, 'connection': None}
    deadline = time.monotonic() + timeout_ms / 1000
    key = secrets.token_bytes(32)
    try:
        listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        if hasattr(socket, 'SO_EXCLUSIVEADDRUSE'):
            listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        listener.bind(('127.0.0.1', 0))
        listener.listen(2)
        listener.settimeout(0.1)
        def receive():
            try:
                for _ in range(8):
                    if stop.is_set() or time.monotonic() >= deadline:
                        break
                    try:
                        connection, addr = listener.accept()
                    except socket.timeout:
                        continue
                    with connection:
                        state['connection'] = connection
                        connection.settimeout(0.1)
                        if addr[0] != '127.0.0.1' or not hmac.compare_digest(read_exact(connection, len(MAGIC) + len(key), deadline, stop), MAGIC + key):
                            continue
                        send_message(connection, {'svg': svg})
                        state['payload'] = receive_message(connection, deadline, stop)
                        return
                raise ValueError('raster connection failed')
            except Exception:
                if not stop.is_set():
                    fault.set()
            finally:
                done.set()
        thread = threading.Thread(target=receive, name='owned-raster-channel', daemon=False)
        thread.start()
        command = (sys.executable, str(Path(__file__).with_name('diagram_raster_worker.py').resolve()),
                   str(listener.getsockname()[1]), key.hex())
        try:
            execution = _run_owned(command, timeout_ms=max(100, int(max(0, deadline - time.monotonic()) * 1000)),
                                   cancelled=lambda: fault.is_set())
        except DiagramError as exc:
            if exc.code == 'renderer_cleanup_failed':
                _disabled.set()
            raise
        if execution.empty_confirmed is not True:
            _disabled.set()
            raise DiagramError('renderer_cleanup_failed')
        if execution.code in {'process_completed', 'process_failed'}:
            done.wait(min(1.0, max(0.0, deadline - time.monotonic())))
        if execution.code == 'render_timeout':
            raise DiagramError('raster_timeout')
        if fault.is_set() or not done.is_set() or execution.code != 'process_completed':
            raise DiagramError('raster_failed')
        response = state['payload']
        if type(response) is not dict or set(response) != {'ok', 'png'} or response['ok'] is not True or type(response['png']) is not str or len(response['png']) > 1_400_000:
            raise DiagramError('raster_failed')
        try:
            raw = base64.b64decode(response['png'], validate=True)
        except ValueError:
            raise DiagramError('raster_invalid') from None
        png, w, h = normalize_png(raw, expected_size=(width, height))
        return RasterReceipt(preview, png, hashlib.sha256(png).hexdigest(), w, h)
    finally:
        stop.set()
        try:
            if listener is not None:
                listener.close()
            if state['connection'] is not None:
                try:
                    state['connection'].shutdown(socket.SHUT_RDWR)
                except OSError:
                    pass
                state['connection'].close()
            if thread is not None:
                thread.join(1)
                if thread.is_alive():
                    _disabled.set()
                    raise DiagramError('renderer_cleanup_failed')
        finally:
            _slots.release()
