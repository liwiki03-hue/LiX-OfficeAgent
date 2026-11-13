import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import zipfile


MAX_FILE = 64 * 1024 * 1024
MAX_TOTAL = 512 * 1024 * 1024
FORBIDDEN = {'.git', '.venv', '.env', '.runtime', '.office_agent', '__pycache__',
             '.pytest_cache', '.startup-logs', 'node_modules', 'ebwebview'}
PATTERNS = {
    'private-key': re.compile(rb'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----'),
    'provider-token': re.compile(rb'\bsk-[A-Za-z0-9_-]{24,}\b'),
    'github-token': re.compile(rb'\b(?:ghp_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{40,})\b'),
    'aws-access-id': re.compile(rb'\bAKIA[0-9A-Z]{16}\b'),
}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def safe_name(name):
    if not isinstance(name, str) or '\\' in name or ':' in name or '\x00' in name:
        raise ValueError('invalid relative file name')
    parts = PurePosixPath(name).parts
    if not parts or name != '/'.join(parts) or PurePosixPath(name).is_absolute():
        raise ValueError('invalid relative file name')
    if any(p in {'.', '..'} or p.endswith((' ', '.')) for p in parts):
        raise ValueError('invalid path component')
    if any(p.lower() in FORBIDDEN for p in parts):
        raise ValueError('private directory or configuration excluded')
    if any(p.lower().startswith('.env') and p != '.env.example' for p in parts):
        raise ValueError('private configuration excluded')
    if Path(name).suffix.lower() in {'.log', '.db', '.sqlite', '.pyc', '.pyo', '.msi', '.exe'}:
        raise ValueError('runtime or private output excluded')
    return parts


def read_file(root, name):
    parts = safe_name(name)
    current = root
    for part in parts:
        current = current / part
        info = current.lstat()
        if current.is_symlink() or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise ValueError('links and reparse points excluded')
    if not stat.S_ISREG(info.st_mode) or info.st_size > MAX_FILE or info.st_nlink > 1:
        raise ValueError('not a bounded independent regular file')
    with current.open('rb') as stream:
        data = stream.read(MAX_FILE + 1)
    after = current.lstat()
    if len(data) > MAX_FILE or (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns) != (
            after.st_dev, after.st_ino, after.st_size, after.st_mtime_ns):
        raise ValueError('file changed while reading')
    return data


def read_json(path):
    if path.stat().st_size > 4 * 1024 * 1024:
        raise ValueError('manifest too large')
    return json.loads(path.read_text(encoding='utf-8'))


def scan(name, data, exceptions):
    hits = []
    for rule, pattern in PATTERNS.items():
        for match in pattern.finditer(data):
            allowed = any(item.get('path') == name and item.get('file_sha256') == sha(data)
                          and item.get('rule') == rule and item.get('match_sha256') == sha(match.group())
                          and item.get('reason') for item in exceptions)
            if not allowed:
                hits.append({'path': name, 'rule': rule, 'line': data.count(b'\n', 0, match.start()) + 1})
    return hits


def collect(root, policy):
    if policy.get('schema') != 1 or not isinstance(policy.get('files'), list) or not policy['files']:
        raise ValueError('invalid file policy')
    rows, issues, seen, total = [], [], set(), 0
    for name in policy['files']:
        safe_name(name)
        if name.casefold() in seen:
            raise ValueError('duplicate or case-colliding path')
        seen.add(name.casefold())
        data = read_file(root, name)
        total += len(data)
        if total > MAX_TOTAL or len(seen) > 10000:
            raise ValueError('source budget exceeded')
        issues.extend(scan(name, data, policy.get('scan_exceptions', [])))
        rows.append({'path': name, 'bytes': len(data), 'sha256': sha(data)})
    if issues:
        raise ValueError(json.dumps({'code': 'secret-pattern-review-required', 'issues': issues}, ensure_ascii=False))
    return {'schema': 1, 'publication_authorized': False, 'files': rows, 'total_bytes': total,
            'scope': 'explicit current files only; not Git history or proof of license rights',
            'policy_sha256': sha(json.dumps(policy, sort_keys=True, ensure_ascii=False).encode())}


def verify(root, manifest, policy):
    current = collect(root, policy)
    if current != manifest:
        raise ValueError('source or policy changed; regenerate and review manifest')
    return current


def check_git_index(root, policy):
    if not (root / '.git').exists():
        return {'verified': False, 'reason': 'no_git_repository'}
    result = subprocess.run(['git', '-C', str(root), 'ls-files', '-z'], capture_output=True, timeout=15, check=False)
    if result.returncode or len(result.stdout) > 4 * 1024 * 1024:
        raise ValueError('Git index inspection failed')
    names = result.stdout.decode('utf-8', errors='strict').split('\x00')
    allowed = set(policy['files'])
    for name in filter(None, names):
        safe_name(name)
        if name not in allowed:
            raise ValueError('Git index contains a file outside the reviewed public policy')
    missing = sorted(allowed - set(names))
    if missing:
        raise ValueError('Git index is missing reviewed public files')
    return {'verified': True, 'files': len(allowed), 'history_scanned': False}


def export(root, target, manifest, policy):
    verify(root, manifest, policy)
    target = target.absolute()
    if target.exists():
        raise ValueError('target must be a new directory')
    for parent in (target.parent, *target.parents):
        if parent.exists() and (parent.is_symlink() or getattr(parent.lstat(), 'st_file_attributes', 0) & 0x400):
            raise ValueError('target parent must not be a link')
    target.mkdir(parents=False)
    for item in manifest['files']:
        data = read_file(root, item['path'])
        if sha(data) != item['sha256']:
            raise ValueError('source changed during export; candidate incomplete')
        output = target.joinpath(*safe_name(item['path']))
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open('xb') as stream:
            stream.write(data)
    verify(target, manifest, policy)
    with (target / 'SOURCE_MANIFEST.json').open('x', encoding='utf-8') as stream:
        json.dump(manifest, stream, ensure_ascii=False, indent=2)
    return {'files': len(manifest['files']), 'bytes': manifest['total_bytes'],
            'source_verified': True, 'publication_authorized': False}


def archive(root, target, manifest, policy):
    verify(root, manifest, policy)
    with target.open('xb') as stream, zipfile.ZipFile(stream, 'w', compression=zipfile.ZIP_DEFLATED) as output:
        for item in manifest['files']:
            data = read_file(root, item['path'])
            if sha(data) != item['sha256']:
                raise ValueError('source changed during archive; archive incomplete')
            info = zipfile.ZipInfo(item['path'], date_time=(2026, 11, 12, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            output.writestr(info, data)
        output.writestr('SOURCE_MANIFEST.json', json.dumps(manifest, ensure_ascii=False, indent=2))
    with zipfile.ZipFile(target) as package:
        if package.testzip() is not None:
            raise ValueError('archive reopen failed')
        for item in manifest['files']:
            if sha(package.read(item['path'])) != item['sha256']:
                raise ValueError('archive member verification failed')
    return {'sha256': sha(target.read_bytes()), 'bytes': target.stat().st_size,
            'files': len(manifest['files']), 'publication_authorized': False}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['manifest', 'check', 'export', 'archive'])
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--policy', type=Path, default=Path('release/PUBLIC_FILES.json'))
    parser.add_argument('--manifest', type=Path, default=Path('SOURCE_MANIFEST.json'))
    parser.add_argument('--target', type=Path)
    parser.add_argument('--require-git', action='store_true')
    args = parser.parse_args()
    root = args.root.resolve(strict=True)
    policy = read_json(args.policy)
    if args.action == 'manifest':
        result = collect(root, policy)
        with args.manifest.open('x', encoding='utf-8') as stream:
            json.dump(result, stream, ensure_ascii=False, indent=2)
        print(json.dumps({'files': len(result['files']), 'bytes': result['total_bytes']}))
    elif args.action == 'check':
        result = verify(root, read_json(args.manifest), policy)
        index = check_git_index(root, policy)
        if args.require_git and not index['verified']:
            raise ValueError('Git index required but unavailable')
        print(json.dumps({'verified': True, 'files': len(result['files']), 'git_index': index}))
    else:
        if args.target is None:
            parser.error('--target is required')
        action = archive if args.action == 'archive' else export
        print(json.dumps(action(root, args.target, read_json(args.manifest), policy)))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError) as exc:
        raise SystemExit(str(exc)) from None
