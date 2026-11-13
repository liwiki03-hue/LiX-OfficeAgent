import hashlib
import io
import json
import os
from pathlib import Path, PurePosixPath
import tarfile
import time
from urllib.request import Request, urlopen
import uuid


MAX_ARCHIVE = 64 * 1024 * 1024
MAX_EXPANDED = 256 * 1024 * 1024
URL = 'https://registry.npmjs.org/superdoc/-/superdoc-1.47.2.tgz'


def checked(path, *, missing=False):
    if not path.is_absolute() or '..' in path.parts:
        raise ValueError('Expected an absolute project path')
    for part in [*reversed(path.parents), path]:
        try:
            info = part.lstat()
        except FileNotFoundError:
            if missing:
                continue
            raise
        if part.is_symlink() or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise ValueError('Refusing a linked or redirected frontend path')
    return path


def sha(data):
    return hashlib.sha256(data).hexdigest()


def spec(root):
    path = checked(root / 'release/SUPERDOC_SOURCE.json')
    if path.stat().st_size > 128 * 1024:
        raise ValueError('Frontend metadata exceeds the limit')
    row = json.loads(path.read_text('utf-8'))
    names = {'package/dist/superdoc.min.js': 'superdoc.min.js', 'package/dist/style.css': 'style.css', 'package/LICENSE': 'LICENSE'}
    if row.get('package') != 'superdoc' or row.get('version') != '1.47.2' or row.get('archive_url') != URL:
        raise ValueError('Frontend version or official source is not the pinned version')
    assets = row.get('assets', [])
    if len(assets) != 3 or {item.get('path'): item.get('local') for item in assets} != names:
        raise ValueError('Frontend asset list does not match the pinned package')
    for value in [row.get('archive_sha256'), *[item.get('sha256') for item in assets]]:
        if not isinstance(value, str) or len(value) != 64 or any(c not in '0123456789abcdef' for c in value):
            raise ValueError('Frontend SHA256 is invalid')
    return row


def download(root, expected):
    directory = checked(root / '.runtime/downloads', missing=True)
    directory.mkdir(parents=True, exist_ok=True)
    checked(directory)
    archive = checked(directory / 'superdoc-1.47.2.tgz', missing=True)
    if archive.exists():
        if not archive.is_file() or archive.stat().st_size > MAX_ARCHIVE or sha(archive.read_bytes()) != expected:
            raise ValueError('Existing frontend archive differs from the pinned hash; retained, not overwritten')
        return archive.read_bytes(), False
    partial = checked(directory / ('superdoc-' + uuid.uuid4().hex + '.partial'), missing=True)
    deadline, total = time.monotonic() + 120, 0
    request = Request(URL, headers={'User-Agent': 'LiX-pinned-frontend-setup', 'Accept': 'application/octet-stream'})
    with urlopen(request, timeout=30) as response, partial.open('xb') as stream:
        if response.status != 200 or response.geturl() != URL:
            raise ValueError('Official frontend download returned an unexpected address or status')
        while chunk := response.read(1024 * 1024):
            total += len(chunk)
            if total > MAX_ARCHIVE or time.monotonic() > deadline:
                raise ValueError('Frontend download exceeded its size or time budget; partial retained')
            stream.write(chunk)
        stream.flush()
        os.fsync(stream.fileno())
    data = partial.read_bytes()
    if sha(data) != expected:
        raise ValueError('Official frontend archive hash mismatch; partial retained, nothing installed')
    checked(archive, missing=True)
    os.link(partial, archive)
    partial.unlink()
    return data, True


def unpack(data, assets):
    wanted = {item['path']: item for item in assets}
    result, seen, total, count = {}, set(), 0, 0
    deadline = time.monotonic() + 30
    with tarfile.open(fileobj=io.BytesIO(data), mode='r:gz') as archive:
        for member in archive:
            count += 1
            path = PurePosixPath(member.name)
            total += member.size
            if (count > 5000 or total > MAX_EXPANDED or member.size > 64 * 1024 * 1024 or member.size < 0
                    or time.monotonic() > deadline or member.name in seen or path.is_absolute()
                    or '..' in path.parts or '\\' in member.name or ':' in member.name
                    or not (member.isdir() or member.isfile())):
                raise ValueError('Frontend archive contains unsafe or oversized entries')
            seen.add(member.name)
            if member.name not in wanted:
                continue
            if not member.isfile() or member.size > 8 * 1024 * 1024:
                raise ValueError('Frontend asset is not a bounded regular file')
            with archive.extractfile(member) as stream:
                value = stream.read(8 * 1024 * 1024 + 1)
            if len(value) != member.size or sha(value) != wanted[member.name]['sha256']:
                raise ValueError('Frontend asset hash mismatch; nothing installed')
            result[wanted[member.name]['local']] = value
    if set(result) != {item['local'] for item in assets}:
        raise ValueError('Pinned frontend archive is missing required files')
    return result


def prepare(root):
    root = checked(root.absolute())
    metadata = spec(root)
    destination = checked(root / 'office_agent/webui/static/vendor/superdoc', missing=True)
    missing = []
    for item in metadata['assets']:
        path = checked(destination / item['local'], missing=True)
        if path.exists():
            if not path.is_file() or path.stat().st_size > 8 * 1024 * 1024 or sha(path.read_bytes()) != item['sha256']:
                raise ValueError('Existing frontend asset changed; refusing to overwrite ' + item['local'])
        else:
            missing.append(item['local'])
    if not missing:
        return {'status': 'already_verified', 'package': 'superdoc', 'version': '1.47.2', 'network': False, 'created': []}
    data, network = download(root, metadata['archive_sha256'])
    if sha(data) != metadata['archive_sha256']:
        raise ValueError('Frontend archive no longer matches the pinned hash')
    values = unpack(data, metadata['assets'])
    destination.mkdir(parents=True, exist_ok=True)
    checked(destination)
    for name in missing:
        target = checked(destination / name, missing=True)
        with target.open('xb') as stream:
            stream.write(values[name])
            stream.flush()
            os.fsync(stream.fileno())
        if sha(checked(target).read_bytes()) != sha(values[name]):
            raise ValueError('Frontend write verification failed; do not start the editor')
    for item in metadata['assets']:
        if sha(checked(destination / item['local']).read_bytes()) != item['sha256']:
            raise ValueError('Frontend files changed during preparation; verification did not complete')
    return {'status': 'verified', 'package': 'superdoc', 'version': '1.47.2', 'network': network, 'created': missing,
            'archive_sha256': metadata['archive_sha256']}


if __name__ == '__main__':
    try:
        print(json.dumps(prepare(Path(__file__).absolute().parents[1])))
    except (ValueError, OSError, tarfile.TarError) as error:
        raise SystemExit('Frontend preparation failed: ' + str(error)) from None
