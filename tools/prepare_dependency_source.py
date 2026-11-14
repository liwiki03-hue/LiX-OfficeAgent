import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tomllib

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from office_agent.downloads import safe


SOURCES = {'aliyun': 'https://mirrors.aliyun.com/pypi/simple/', 'pypi': 'https://pypi.org/simple/'}


def identity(data):
    lock = tomllib.loads(data)
    return sorted((row['name'], row['version'], tuple(sorted(asset['hash'] for asset in [*row.get('wheels', []), *([row['sdist']] if 'sdist' in row else [])]))) for row in lock['package'])


def prepare(root, target, source):
    root, target = safe(root), safe(target)
    if source not in SOURCES or target.exists() or not target.is_relative_to(root):
        raise ValueError('A new directory within the project is required')
    project = (root / 'pyproject.toml').read_text('utf8')
    lock = (root / 'uv.lock').read_text('utf8')
    before = (hashlib.sha256(project.encode('utf8')).hexdigest(), hashlib.sha256(lock.encode('utf8')).hexdigest())
    original = identity(lock)
    if source == 'pypi':
        project = project.replace('https://mirrors.aliyun.com/pypi/simple/', SOURCES[source])
        lock = lock.replace('https://mirrors.aliyun.com/pypi/simple/', SOURCES[source]).replace('https://mirrors.aliyun.com/pypi/packages/', 'https://files.pythonhosted.org/packages/')
    target.mkdir(parents=True)
    (target / 'pyproject.toml').write_text(project, 'utf8')
    (target / 'uv.lock').write_text(lock, 'utf8')
    uv = shutil.which('uv')
    if not uv:
        raise ValueError('Existing uv is required')
    result = subprocess.run([uv, 'lock', '--project', str(target), '--python', sys.executable, '--offline'], capture_output=True, timeout=60)
    final = (target / 'uv.lock').read_text('utf8')
    unchanged = identity(final) == original
    report = {'source': source, 'index': SOURCES[source], 'uv_exit': result.returncode, 'versions_and_artifact_hashes_preserved': unchanged,
              'source_files_unchanged': before == (hashlib.sha256((root / 'pyproject.toml').read_text('utf8').encode('utf8')).hexdigest(), hashlib.sha256((root / 'uv.lock').read_text('utf8').encode('utf8')).hexdigest()),
              'candidate_lock_sha256': hashlib.sha256(final.encode('utf8')).hexdigest(), 'network_requests': 0,
              'downloads_verified': False, 'installed': False, 'offline_lock_check': result.returncode == 0 and unchanged}
    if result.returncode:
        report['uv_code'] = 'cache_access_denied' if b'Access is denied' in result.stderr or b'os error 5' in result.stderr else 'offline_lock_not_verified'
    (target / 'SOURCE_SELECTION.json').write_text(json.dumps(report, indent=2), 'utf8')
    if not unchanged:
        raise ValueError('Candidate changed package versions or artifact hashes; not suitable for installation')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--target', type=Path, required=True)
    parser.add_argument('--source', choices=list(SOURCES), required=True)
    args = parser.parse_args()
    print(json.dumps(prepare(args.root.absolute(), args.target.absolute(), args.source)))
