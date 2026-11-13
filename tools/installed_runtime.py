import hashlib
import json
import os
from pathlib import Path
import time
from urllib.request import Request, urlopen
import uuid

from tools.install_frontend_assets import checked


ROOT = Path(__file__).absolute().parents[1]


def digest(path):
    value = hashlib.sha256()
    with checked(path).open('rb') as stream:
        while chunk := stream.read(1024 * 1024):
            value.update(chunk)
    return value.hexdigest()


def download(folder, spec):
    folder = checked(folder, missing=True)
    folder.mkdir(parents=True, exist_ok=True)
    checked(folder)
    archive = checked(folder / ('LibreOffice_' + spec['version'] + '_Win_x86-64.msi'), missing=True)
    if archive.exists():
        if archive.stat().st_size != spec['bytes'] or digest(archive) != spec['sha256']:
            raise ValueError('Existing renderer download changed; retained, not overwritten')
        return archive
    partial = checked(folder / (archive.name + '.' + uuid.uuid4().hex + '.partial'), missing=True)
    end, count = time.monotonic() + 600, 0
    request = Request(spec['url'], headers={'User-Agent': 'LiX-installed-runtime-setup'})
    with urlopen(request, timeout=30) as response, partial.open('xb') as stream:
        if response.status != 200 or not response.geturl().startswith('https://'):
            raise ValueError('Renderer archive request failed')
        while chunk := response.read(1024 * 1024):
            count += len(chunk)
            if count > spec['bytes'] or time.monotonic() > end:
                raise ValueError('Renderer download budget exceeded; partial retained')
            stream.write(chunk)
        stream.flush()
        os.fsync(stream.fileno())
    if count != spec['bytes'] or digest(partial) != spec['sha256']:
        raise ValueError('Renderer archive checksum mismatch; nothing extracted')
    os.link(partial, archive)
    partial.unlink()
    return archive


def extract(archive, destination):
    from office_agent.diagram_process import _WindowsJob
    command = Path(os.environ['SystemRoot']) / 'System32/msiexec.exe'
    process, job = None, _WindowsJob()
    try:
        process = job.spawn((str(command), '/a', str(archive), '/qn', 'TARGETDIR=' + str(destination)))
        end = time.monotonic() + 180
        while process.poll() is None and time.monotonic() < end:
            time.sleep(.1)
        if process.poll() != 0:
            raise ValueError('Renderer extraction failed or timed out; incomplete files retained')
    finally:
        empty = job.close()
        if process:
            process.close()
        if not empty:
            raise ValueError('Renderer extraction process cleanup could not be confirmed')


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
        if not receipt.is_file() or receipt.stat().st_size > 64 * 1024:
            raise ValueError('Renderer directory already exists without a valid receipt; retained')
        saved = json.loads(receipt.read_text('utf-8'))
        if (saved.get('archive_sha256') != spec['sha256'] or not checked(binary).is_file()
                or saved.get('executable_sha256') != digest(binary)):
            raise ValueError('Installed renderer changed; repair requires user action')
        return {'ok': True, 'status': 'already_verified', 'rendering_verified': False}
    archive = download(runtime / 'downloads', spec)
    extract(archive, destination)
    if not checked(binary).is_file():
        raise ValueError('Expected renderer executable is missing')
    result = {'ok': True, 'archive_url': spec['url'], 'archive_sha256': spec['sha256'],
              'executable_sha256': digest(binary), 'rendering_verified': False,
              'deployment': 'user-local administrative extraction; no Office registration'}
    with receipt.open('x', encoding='utf-8') as stream:
        json.dump(result, stream, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    return result


if __name__ == '__main__':
    try:
        print(json.dumps(prepare(ROOT)))
    except (OSError, ValueError, KeyError) as error:
        raise SystemExit('Renderer preparation failed: ' + type(error).__name__) from None
