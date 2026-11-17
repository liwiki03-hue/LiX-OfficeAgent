import hashlib
import json
import os
from pathlib import Path
import sys

from tools.install_frontend_assets import checked


def seal(folder):
    folder = checked(Path(folder).absolute())
    if Path(sys.prefix).resolve() != folder.resolve() or not (folder / '.lix-owner').is_file():
        raise ValueError('offline_environment_identity')
    if os.name == 'nt':
        folder = Path('\\\\?\\' + str(folder))
    rows, total = [], 0
    for current, directories, files in os.walk(folder, followlinks=False):
        for name in directories:
            if getattr((Path(current) / name).lstat(), 'st_file_attributes', 0) & 0x400:
                raise ValueError('offline_environment_link')
        directories[:] = [name for name in directories if name != '__pycache__']
        for name in files:
            if name in {'.lix-inventory.json', '.lix-inventory.tmp'} or name.endswith('.pyc'):
                continue
            path = Path(current) / name
            if path.is_symlink() or getattr(path.lstat(), 'st_file_attributes', 0) & 0x400 or path.stat().st_nlink != 1:
                raise ValueError('offline_environment_link')
            size = path.stat().st_size
            total += size
            if len(rows) >= 25000 or total > 2 * 1024**3 or size > 512 * 1024**2:
                raise ValueError('offline_environment_budget')
            with path.open('rb') as stream:
                digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            rows.append({'path': path.relative_to(folder).as_posix(), 'bytes': size, 'sha256': digest})
    data = json.dumps({'schema': 1, 'id': 'environment', 'files': sorted(rows, key=lambda row: row['path'])}, ensure_ascii=False).encode()
    if len(data) > 4 * 1024**2:
        raise ValueError('offline_environment_inventory_budget')
    target = folder / '.lix-inventory.json'
    temporary = folder / '.lix-inventory.tmp'
    for path in (target, temporary):
        if path.is_symlink() or path.exists() and getattr(path.lstat(), 'st_file_attributes', 0) & 0x400:
            raise ValueError('offline_environment_link')
    with temporary.open('xb') as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, target)
    print(json.dumps({'sealed_files': len(rows), 'bytes': total, 'network_requests': 0}))


if __name__ == '__main__':
    seal(Path(sys.prefix))
