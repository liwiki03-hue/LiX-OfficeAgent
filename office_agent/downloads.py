import hashlib
from http.client import HTTPException
import json
import os
from pathlib import Path
import socket
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
import uuid


class DownloadError(ValueError):
    def __init__(self, code, **details):
        super().__init__(code)
        self.details = details


def safe(path):
    path = Path(path).absolute()
    for item in [*reversed(path.parents), path]:
        if item.is_symlink() or item.exists() and getattr(item.lstat(), 'st_file_attributes', 0) & 0x400:
            raise DownloadError('download_path')
    return path


def digest(path):
    result = hashlib.sha256()
    with safe(path).open('rb') as stream:
        while chunk := stream.read(1024 * 1024):
            result.update(chunk)
    return result.hexdigest()


def _fetch(url, target, expected, *, size=None, limit=64 * 1024 * 1024, seconds=120,
          chunk_bytes=4 * 1024 * 1024, opener=None, progress=None, cancelled=None):
    from urllib.request import Request
    from .network import urlopen
    opener, progress, cancelled = opener or urlopen, progress or (lambda **kw: None), cancelled or (lambda: False)
    parsed = urlsplit(url)
    if parsed.scheme != 'https' or parsed.username or parsed.password or parsed.fragment:
        raise DownloadError('download_url')
    if len(expected) != 64 or any(c not in '0123456789abcdef' for c in expected):
        raise DownloadError('download_checksum_spec')
    target = safe(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        if not target.is_file() or target.stat().st_size > limit or size is not None and target.stat().st_size != size or digest(target) != expected:
            raise DownloadError('download_cache_changed')
        return target
    partial, metadata = safe(target.with_name(target.name + '.partial')), safe(target.with_name(target.name + '.partial.json'))
    identity = {'url': url, 'sha256': expected, 'bytes': size}
    saved = {}
    if metadata.exists():
        if metadata.stat().st_size > 4096:
            raise DownloadError('download_identity')
        saved = json.loads(metadata.read_text('utf-8'))
        if any(saved.get(key) != value for key, value in identity.items()):
            raise DownloadError('download_identity')
    if partial.exists() and (not partial.is_file() or partial.stat().st_size > (size if size is not None else limit)):
        raise DownloadError('partial_invalid')
    deadline, attempts = time.monotonic() + seconds, 0
    while True:
        if cancelled():
            raise DownloadError('download_cancelled')
        if time.monotonic() >= deadline:
            raise DownloadError('download_timeout')
        attempts += 1
        count = partial.stat().st_size if partial.exists() else 0
        try:
            with partial.open('ab') as stream:
                while size is None or count < size:
                    if cancelled():
                        raise DownloadError('download_cancelled')
                    remaining = deadline - time.monotonic()
                    if remaining <= 0:
                        raise DownloadError('download_timeout')
                    last = min(count + chunk_bytes, size) - 1 if size is not None else None
                    headers = {'User-Agent': 'LiX-pinned-setup', 'Accept-Encoding': 'identity', 'Range': f'bytes={count}-{last if last is not None else ""}'}
                    if saved.get('etag'):
                        headers['If-Range'] = saved['etag']
                    with opener(Request(url, headers=headers), timeout=min(20, max(.1, remaining))) as response:
                        if response.geturl() != url:
                            raise DownloadError('download_identity')
                        received_headers = getattr(response, 'headers', {})
                        etag = received_headers.get('ETag')
                        if saved.get('etag') and etag != saved['etag']:
                            raise DownloadError('download_identity')
                        response_end = None
                        if response.status == 206:
                            import re
                            match = re.fullmatch(r'bytes (\d+)-(\d+)/(\d+)', received_headers.get('Content-Range', ''))
                            if not match or int(match[1]) != count or size is not None and (int(match[2]) != last or int(match[3]) != size):
                                raise DownloadError('download_range')
                            response_end = int(match[2]) + 1
                            if int(match[3]) > limit or response_end <= count or response_end > int(match[3]):
                                raise DownloadError('download_size')
                            if size is None:
                                size = int(match[3])
                        elif response.status == 200 and count == 0:
                            response_end = size
                        else:
                            raise DownloadError('download_range')
                        length = received_headers.get('Content-Length')
                        if length is not None:
                            if not length.isascii() or not length.isdigit() or int(length) > limit or response_end is not None and int(length) != response_end - count:
                                raise DownloadError('download_size')
                            response_end = count + int(length)
                        saved = {**identity, 'etag': etag if isinstance(etag, str) and len(etag) <= 256 else None}
                        temporary = safe(metadata.with_name(metadata.name + '.tmp'))
                        temporary.write_text(json.dumps(saved), encoding='utf-8')
                        os.replace(temporary, metadata)
                        read = getattr(response, 'read1', response.read)
                        while True:
                            if cancelled():
                                raise DownloadError('download_cancelled')
                            if time.monotonic() >= deadline:
                                raise DownloadError('download_timeout')
                            block = read(min(65536, (response_end - count + 1) if response_end is not None else 65536))
                            if not block:
                                break
                            if count + len(block) > limit or response_end is not None and count + len(block) > response_end:
                                raise DownloadError('download_size')
                            stream.write(block)
                            count += len(block)
                            if count % (1024 * 1024) < len(block):
                                stream.flush()
                                progress(state='download', received=count, total=size or response_end or 0, attempt=attempts)
                        stream.flush()
                        os.fsync(stream.fileno())
                        if response_end is not None and count != response_end:
                            raise EOFError('download_incomplete')
                    progress(state='download', received=count, total=size or response_end or 0, attempt=attempts)
                    if size is None:
                        break
            break
        except (HTTPError, URLError, TimeoutError, socket.timeout, HTTPException, ConnectionError, EOFError) as error:
            retryable = not isinstance(error, HTTPError) or error.code == 429 or 500 <= error.code <= 599
            if attempts >= 3 or not retryable:
                raise
            delay = min(2 ** (attempts - 1), 4)
            if isinstance(error, HTTPError):
                retry_after = error.headers.get('Retry-After', '') if error.headers else ''
                if retry_after.isascii() and retry_after.isdigit():
                    delay = min(int(retry_after), 8)
                elif retry_after:
                    from datetime import datetime, timezone
                    from email.utils import parsedate_to_datetime
                    try:
                        delay = min(8, max(0, (parsedate_to_datetime(retry_after) - datetime.now(timezone.utc)).total_seconds()))
                    except (TypeError, ValueError, OverflowError):
                        pass
                error.close()
            progress(state='retry', received=partial.stat().st_size, total=size or 0, attempt=attempts, code='download_retry')
            until = min(deadline, time.monotonic() + delay)
            while time.monotonic() < until:
                if cancelled():
                    raise DownloadError('download_cancelled')
                time.sleep(min(.1, max(0, until - time.monotonic())))
    progress(state='verify', received=count, total=size or count, attempt=attempts)
    actual = digest(partial)
    if size is not None and count != size or actual != expected:
        raise DownloadError('download_checksum', expected_sha256=expected, actual_sha256=actual, received=count, attempts=attempts)
    os.link(partial, target)
    partial.unlink()
    metadata.unlink(missing_ok=True)
    return target


def quarantine(target, expected, confirmed=False):
    target = safe(target)
    if confirmed is not True or not target.is_file() or digest(target) == expected:
        raise DownloadError('download_quarantine_denied')
    destination = safe(target.with_name(target.name + '.invalid-' + uuid.uuid4().hex))
    target.rename(destination)
    return destination


def fetch(url, target, expected, **kwargs):
    from .diagnostics import record
    started = time.monotonic()
    record('network', 'download', 'started')
    try:
        result = _fetch(url, target, expected, **kwargs)
        record('network', 'download', 'verified', elapsed_ms=int((time.monotonic() - started) * 1000))
        return result
    except Exception as error:
        record('network', 'download', 'failed', code=str(error) if isinstance(error, DownloadError) else 'http_' + str(error.code) if isinstance(error, HTTPError) else type(error).__name__,
               host=urlsplit(url).hostname or '', elapsed_ms=int((time.monotonic() - started) * 1000), **(error.details if isinstance(error, DownloadError) else {}))
        raise
