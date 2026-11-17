import hashlib
from http.client import HTTPException
import json
import os
from pathlib import Path
import socket
import ssl
import time
import sys
import uuid
from urllib.error import HTTPError, URLError

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from office_agent.network import urlopen

from tools.install_frontend_assets import checked


ROOT = Path(__file__).absolute().parents[1]
DOWNLOAD_SECONDS = 600
CHUNK_BYTES = 4 * 1024 * 1024


class PreparationError(ValueError):
    def __init__(self, code, system_code=None):
        self.code = code
        self.system_code = system_code
        super().__init__(code)


def report(state, *, code='', received=0, total=0, attempt=1, system_code=None):
    target = os.environ.get('LIX_RENDERER_REPORT')
    if not target:
        return
    path = checked(Path(target), missing=True)
    temporary = checked(path.with_name(path.name + '.tmp'), missing=True)
    value = {'schema': 1, 'state': state, 'code': code, 'received': received, 'total': total, 'attempt': attempt,
             'system_code': system_code, 'updated_at': time.time()}
    with temporary.open('w', encoding='utf-8') as stream:
        json.dump(value, stream)
    os.replace(temporary, path)


def digest(path):
    value = hashlib.sha256()
    with checked(path).open('rb') as stream:
        while chunk := stream.read(1024 * 1024):
            value.update(chunk)
    return value.hexdigest()


def download(folder, spec):
    from office_agent.downloads import fetch, DownloadError
    folder = checked(folder, missing=True)
    folder.mkdir(parents=True, exist_ok=True)
    archive = checked(folder / ('LibreOffice_' + spec['version'] + '_Win_x86-64.msi'), missing=True)
    partial = checked(folder / (archive.name + '.partial'), missing=True)
    if not partial.exists():
        candidates = []
        with os.scandir(folder) as entries:
            for index, entry in enumerate(entries):
                if index >= 128:
                    raise PreparationError('download_directory_limit')
                suffix = entry.name.removeprefix(archive.name + '.').removesuffix('.partial')
                if entry.name.startswith(archive.name + '.') and entry.name.endswith('.partial') and len(suffix) == 32 and all(c in '0123456789abcdef' for c in suffix):
                    candidate = checked(Path(entry.path))
                    if candidate.is_file() and candidate.stat().st_size <= spec['bytes']:
                        candidates.append(candidate)
        if candidates:
            previous = max(candidates, key=lambda value: (value.stat().st_size, value.name))
            os.link(previous, partial)
            previous.unlink()
    try:
        return fetch(spec['url'], archive, spec['sha256'], size=spec['bytes'], limit=spec['bytes'],
                     seconds=DOWNLOAD_SECONDS, chunk_bytes=CHUNK_BYTES, opener=urlopen, progress=report)
    except DownloadError as error:
        raise PreparationError(str(error)) from None
    except EOFError:
        raise PreparationError('download_incomplete') from None

def extract(archive, destination):
    from office_agent.diagram_process import _WindowsJob
    command = Path(os.environ['SystemRoot']) / 'System32/msiexec.exe'
    process, job = None, _WindowsJob()
    try:
        process = job.spawn((str(command), '/a', str(archive), '/qn', 'TARGETDIR=' + str(destination)))
        end = time.monotonic() + 180
        while process.poll() is None and time.monotonic() < end:
            time.sleep(.1)
        if process.poll() is None:
            raise PreparationError('extract_timeout')
        if process.poll() != 0:
            raise PreparationError('installer_service_unavailable' if process.poll() == 1601 else 'extract_failed', process.poll())
    finally:
        empty = job.close()
        if process:
            process.close()
        if not empty:
            raise ValueError('Renderer extraction process cleanup could not be confirmed')


def read_receipt(path, binary, spec):
    path = checked(path)
    if not path.is_file() or path.stat().st_size > 65536:
        raise PreparationError('runtime_receipt_invalid')
    with path.open('rb') as stream:
        data = stream.read(65537)
    if len(data) > 65536:
        raise PreparationError('runtime_receipt_invalid')
    try:
        saved = json.loads(data)
    except (ValueError, UnicodeError):
        raise PreparationError('runtime_receipt_invalid') from None
    if (not isinstance(saved, dict) or saved.get('archive_sha256') != spec['sha256']
            or not checked(binary).is_file() or saved.get('executable_sha256') != digest(binary)):
        raise PreparationError('runtime_integrity_changed')
    return saved


def publish_receipt(path, value):
    path = checked(path, missing=True)
    temporary = checked(path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp'), missing=True)
    with temporary.open('x', encoding='utf-8') as stream:
        json.dump(value, stream, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    if os.name == 'nt':
        os.rename(temporary, path)
    else:
        os.link(temporary, path)
        temporary.unlink()


def prepare(root):
    root = checked(root.absolute())
    metadata = checked(root / 'release/WINDOWS_INSTALLER.json')
    if metadata.stat().st_size > 64 * 1024:
        raise ValueError('Installer metadata exceeds the limit')
    spec = json.loads(metadata.read_text('utf-8'))['renderer']
    if (spec['version'] != '26.2.6.2' or spec['bytes'] != 373293056
            or spec['sha256'] != '788ce7d4b56460357f57552cb8cd848e8f2254f28f6e51fa61e0c65a18096373'
            or spec['url'] != 'https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.6.2/win/x86_64/LibreOffice_26.2.6.2_Win_x86-64.msi'):
        raise ValueError('Renderer is not the pinned official release')
    runtime = checked(root / '.runtime', missing=True)
    runtime.mkdir(exist_ok=True)
    checked(runtime)
    destination = checked(runtime / 'libreoffice', missing=True)
    receipt = checked(runtime / 'installed-renderer.json', missing=True)
    binary = destination / 'program/soffice.com'
    if destination.exists():
        if not receipt.exists():
            embedded = checked(destination / '.lix-renderer.json', missing=True)
            if not embedded.exists():
                raise ValueError('Renderer directory already exists without a valid receipt; retained')
            saved = read_receipt(embedded, binary, spec)
            publish_receipt(receipt, saved)
            return {'ok': True, 'status': 'receipt_recovered', 'rendering_verified': False}
        read_receipt(receipt, binary, spec)
        return {'ok': True, 'status': 'already_verified', 'rendering_verified': False}
    if receipt.exists():
        raise PreparationError('runtime_receipt_without_binary')
    bundled = checked(root.parent / 'runtime' / ('LibreOffice_' + spec['version'] + '_Win_x86-64.msi'), missing=True)
    if bundled.exists():
        report('verify')
        if not bundled.is_file() or bundled.stat().st_size != spec['bytes'] or digest(bundled) != spec['sha256']:
            raise PreparationError('download_checksum')
        archive = bundled
    else:
        archive = download(runtime / 'downloads', spec)
    staging = checked(runtime / 'renderer-preparation', missing=True)
    staging.mkdir(exist_ok=True)
    with os.scandir(staging) as entries:
        if sum(1 for _, entry in zip(range(4), entries)) >= 4:
            raise PreparationError('runtime_staging_limit')
    attempt = checked(staging / uuid.uuid4().hex, missing=True)
    attempt.mkdir()
    staged = checked(attempt / 'payload', missing=True)
    report('extract')
    extract(archive, staged)
    staged_binary = checked(staged / 'program/soffice.com')
    if not staged_binary.is_file():
        raise ValueError('Expected renderer executable is missing')
    result = {'ok': True, 'archive_url': spec['url'], 'archive_sha256': spec['sha256'],
              'executable_sha256': digest(staged_binary), 'rendering_verified': False,
              'deployment': 'user-local administrative extraction; no Office registration'}
    publish_receipt(staged / '.lix-renderer.json', result)
    if (not staged.resolve(strict=True).is_relative_to(runtime.resolve(strict=True))
            or not destination.resolve().is_relative_to(runtime.resolve(strict=True)) or destination.exists()):
        raise PreparationError('runtime_publish_conflict')
    os.rename(staged, destination)
    publish_receipt(receipt, result)
    try:
        attempt.rmdir()
    except OSError:
        pass
    return result


def main(root=ROOT):
    try:
        report('starting')
        result = prepare(root)
        report('ready')
        print(json.dumps(result))
        return 0
    except (OSError, ValueError, KeyError, HTTPException) as error:
        if isinstance(error, PreparationError):
            code = error.code
        elif isinstance(error, HTTPError):
            code = 'download_http'
        elif isinstance(error, (TimeoutError, socket.timeout)):
            code = 'download_timeout'
        elif isinstance(error, URLError):
            if isinstance(error.reason, (TimeoutError, socket.timeout)):
                code = 'download_timeout'
            else:
                code = 'download_tls' if isinstance(error.reason, ssl.SSLError) else 'download_network'
        elif isinstance(error, HTTPException):
            code = 'download_incomplete'
        elif isinstance(error, PermissionError):
            code = 'disk_permission'
        else:
            code = 'runtime_integrity_or_extraction'
        try:
            report('failed', code=code, system_code=getattr(error, 'system_code', None))
        except (OSError, ValueError):
            pass
        print('Renderer preparation failed: ' + code)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
