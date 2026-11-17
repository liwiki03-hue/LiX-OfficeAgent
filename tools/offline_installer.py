import argparse
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import subprocess
import tarfile
import time
import tomllib
import zipfile
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

from tools.build_windows_installer import checked, runtime_files, prepare_compiler_inputs
from tools import release_source


ROOT = Path(__file__).absolute().parents[1]
MIB = 1024 * 1024
BUDGETS = {
    'application': (5000, 256 * MIB, 64 * MIB),
    'python': (15000, 768 * MIB, 128 * MIB),
    'wheels': (512, 2048 * MIB, 512 * MIB),
    'renderer': (20000, 2048 * MIB, 256 * MIB),
    'prerequisites': (16, 768 * MIB, 384 * MIB),
    'licenses': (2000, 128 * MIB, 32 * MIB),
}
PYTHON = {'version': '3.12.14', 'build': '20260929', 'bytes': 46425719,
          'sha256': '28728baf30b65e263f0b25c5a85be8226e7ab0d212fbadd6a8f0f796139fa804',
          'url': 'https://github.com/astral-sh/python-build-standalone/releases/download/20260929/cpython-3.12.14%2B20260929-x86_64-pc-windows-msvc-install_only.tar.gz'}


def sha(path):
    with checked(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def artifact_url(value):
    if value.startswith('https://mirrors.aliyun.com/pypi/packages/'):
        return 'https://files.pythonhosted.org/packages/' + value.split('/pypi/packages/', 1)[1]
    return value


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8') as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)
        stream.write('\n')


def relative(name):
    if not isinstance(name, str) or len(name) > 240 or '\\' in name:
        raise ValueError('offline_path')
    parts = PurePosixPath(name).parts
    if not parts or '/'.join(parts) != name:
        raise ValueError('offline_path')
    for part in parts:
        if part in ('.', '..') or part.endswith((' ', '.')) or any(c in part for c in ':*?"<>|') or any(ord(c) < 32 for c in part):
            raise ValueError('offline_path')
        if part.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1, 10)), *(f'LPT{i}' for i in range(1, 10))}:
            raise ValueError('offline_path')
    return name


def copy_file(source, target):
    source, target = checked(source), checked(target, missing=True)
    target.parent.mkdir(parents=True, exist_ok=True)
    with source.open('rb') as incoming, target.open('xb') as output:
        shutil.copyfileobj(incoming, output, MIB)
    if sha(source) != sha(target):
        raise ValueError('offline_copy_changed')


def tree(folder):
    folder = checked(folder)
    seen, result = set(), []
    for current, directories, files in os.walk(folder, followlinks=False):
        if len(result) + len(directories) + len(files) > 30000:
            raise ValueError('offline_tree_limit')
        for name in directories:
            checked(Path(current) / name)
        for name in sorted(files):
            path = checked(Path(current) / name)
            name = relative(path.relative_to(folder).as_posix())
            if name.casefold() in seen or not path.is_file() or path.stat().st_nlink != 1:
                raise ValueError('offline_tree_identity')
            seen.add(name.casefold())
            result.append((name, path))
    return sorted(result)


def fetch(url, target, expected=None, size=None, maximum=512 * MIB):
    parsed = urlsplit(url)
    allowed = ('github.com', 'release-assets.githubusercontent.com', 'objects.githubusercontent.com',
               'files.pythonhosted.org', 'go.microsoft.com', 'download.microsoft.com',
               'msedge.sf.dl.delivery.mp.microsoft.com', 'msedge.b.tlu.dl.delivery.mp.microsoft.com')
    if parsed.scheme != 'https' or parsed.hostname not in allowed or parsed.username or parsed.password:
        raise ValueError('offline_download_origin')
    target = checked(target, missing=True)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        if not expected or sha(target) != expected or size is not None and target.stat().st_size != size:
            raise ValueError('offline_input_changed')
        return target
    partial = target.with_name(target.name + '.partial')
    if partial.exists():
        raise ValueError('offline_partial_retained')
    deadline = time.monotonic() + 600
    with urlopen(Request(url, headers={'User-Agent': 'LiX-offline-inputs'}), timeout=30) as response, partial.open('xb') as stream:
        final = urlsplit(response.geturl())
        if final.scheme != 'https' or final.hostname not in allowed:
            raise ValueError('offline_download_redirect')
        count = 0
        while chunk := response.read(MIB):
            count += len(chunk)
            if count > maximum or time.monotonic() > deadline:
                raise ValueError('offline_download_budget')
            stream.write(chunk)
    if size is not None and count != size or expected is not None and sha(partial) != expected:
        raise ValueError('offline_download_integrity')
    partial.rename(target)
    return target


def extract_python(archive, destination):
    if sha(archive) != PYTHON['sha256']:
        raise ValueError('offline_python_identity')
    destination = checked(destination, missing=True)
    if destination.exists():
        raise ValueError('offline_python_destination_exists')
    seen, total = set(), 0
    with tarfile.open(archive, 'r:gz') as package:
        for index, member in enumerate(package):
            name = member.name.rstrip('/')
            if index > 15000 or not name.startswith('python/') and name != 'python':
                raise ValueError('offline_python_archive')
            if name == 'python':
                continue
            name = relative(name[7:])
            if name.casefold() in seen or not (member.isfile() or member.isdir()):
                raise ValueError('offline_python_archive')
            seen.add(name.casefold())
            total += member.size
            if total > 768 * MIB or member.size > 128 * MIB:
                raise ValueError('offline_python_budget')
            target = checked(destination / name, missing=True)
            if member.isdir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with package.extractfile(member) as source, target.open('xb') as output:
                    shutil.copyfileobj(source, output, MIB)


def wheel_plan(root, exported, built=None):
    from packaging.markers import default_environment
    from packaging.requirements import Requirement
    from packaging.tags import sys_tags
    from packaging.utils import canonicalize_name, parse_wheel_filename
    environment = default_environment()
    if environment['python_full_version'] != PYTHON['version'] or environment['sys_platform'] != 'win32':
        raise ValueError('offline_build_requires_windows_python_3_12_14')
    packages = tomllib.loads((root / 'uv.lock').read_text('utf-8'))['package']
    lock = {(canonicalize_name(row['name']), row['version']): row for row in packages}
    tags = {value: index for index, value in enumerate(sys_tags())}
    joined = exported.read_text('utf-8').replace('\\\n', ' ')
    rows = []
    for line in joined.splitlines():
        if not line.strip() or line.lstrip().startswith('#'):
            continue
        requirement = Requirement(line.split(' --hash=')[0].strip())
        if requirement.url or requirement.extras or str(requirement.specifier).count('==') != 1:
            raise ValueError('offline_requirement_unsupported')
        if requirement.marker and not requirement.marker.evaluate(environment):
            continue
        version = str(requirement.specifier)[2:]
        name = canonicalize_name(requirement.name)
        candidates = []
        for item in lock[(name, version)].get('wheels', []):
            filename = urlsplit(item['url']).path.rsplit('/', 1)[-1]
            wheel_name, wheel_version, _, wheel_tags = parse_wheel_filename(filename)
            ranks = [tags[tag] for tag in wheel_tags if tag in tags]
            if ranks and wheel_name == name and str(wheel_version) == version:
                candidates.append((min(ranks), filename, item))
        if not candidates:
            item = lock[(name, version)].get('sdist', {})
            built_row = (built or {}).get(name)
            if not built_row or built_row.get('source_sha256') != item.get('hash', '').removeprefix('sha256:'):
                raise ValueError('offline_wheel_missing:' + name)
            rows.append(built_row['wheel'])
            continue
        _, filename, item = sorted(candidates, key=lambda row: (row[0], row[1]))[0]
        value = item['hash'].removeprefix('sha256:')
        if '--hash=sha256:' + value not in line:
            raise ValueError('offline_export_lock_mismatch')
        rows.append({'name': name, 'version': version, 'file': filename, 'url': artifact_url(item['url']), 'sha256': value, 'bytes': item.get('size')})
    if not 1 <= len(rows) <= 256 or len({row['name'] for row in rows}) != len(rows):
        raise ValueError('offline_requirements_count')
    return rows


def collect(root, destination):
    destination = checked(destination, missing=True)
    destination.mkdir(parents=True, exist_ok=True)
    export = destination / 'requirements-universal.txt'
    result = subprocess.run(['uv', 'export', '--locked', '--no-dev', '--no-default-groups', '--no-emit-project',
                             '--format', 'requirements-txt', '--output-file', str(export)], cwd=root, capture_output=True, timeout=90)
    if result.returncode:
        raise ValueError('offline_export_failed')
    built_path = destination / 'BUILT_WHEELS.json'
    built = json.loads(built_path.read_text('utf-8')) if built_path.exists() else {}
    while True:
        try:
            rows = wheel_plan(root, export, built)
            break
        except ValueError as error:
            if not str(error).startswith('offline_wheel_missing:'):
                raise
            name = str(error).split(':', 1)[1]
            if name != 'proxy-tools' or name in built:
                raise
            row = next(row for row in tomllib.loads((root / 'uv.lock').read_text('utf-8'))['package'] if row['name'] == name)
            source = row['sdist']
            archive = fetch(artifact_url(source['url']), destination / 'sources' / urlsplit(source['url']).path.rsplit('/', 1)[-1], source['hash'].removeprefix('sha256:'), source.get('size'))
            constraints = destination / 'build-constraints.txt'
            constraints.write_text('setuptools==80.9.0\nwheel==0.45.1\n', 'utf-8')
            output = destination / 'built-wheels'
            result = subprocess.run(['uv', 'build', '--wheel', '--no-sources', '--build-constraint', str(constraints), '--out-dir', str(output), str(archive)], capture_output=True, timeout=180)
            (destination / 'source-build.log').write_bytes(result.stdout + result.stderr)
            if result.returncode:
                raise ValueError('offline_source_build_failed:' + name)
            files = list(output.glob('*.whl'))
            if len(files) != 1:
                raise ValueError('offline_source_build_ambiguous')
            wheel = files[0]
            built[name] = {'source_sha256': sha(archive), 'source_url': source['url'], 'constraints_sha256': sha(constraints),
                           'wheel': {'name': name, 'version': row['version'], 'file': wheel.name, 'url': source['url'], 'sha256': sha(wheel), 'bytes': wheel.stat().st_size, 'built_locally': True}}
            with built_path.open('w', encoding='utf-8') as stream:
                json.dump(built, stream, indent=2)
            copy_file(wheel, destination / 'wheels' / wheel.name)
    write_json(destination / 'WHEELS.json', {'schema': 1, 'lock_sha256': sha(root / 'uv.lock'), 'python': PYTHON, 'wheels': rows})
    folder = destination / 'wheels'
    with ThreadPoolExecutor(max_workers=4) as pool:
        list(pool.map(lambda row: fetch(row['url'], folder / row['file'], row['sha256'], row['bytes']), [row for row in rows if not row.get('built_locally')]))
    archive = fetch(PYTHON['url'], destination / 'python.tar.gz', PYTHON['sha256'], PYTHON['bytes'])
    extract_python(archive, destination / 'python')
    write_json(destination / 'PYTHON_FILES.json', {'archive': PYTHON, 'files': [
        {'path': name, 'bytes': path.stat().st_size, 'sha256': sha(path)} for name, path in tree(destination / 'python')]})
    (destination / 'requirements.txt').write_text(''.join(f"{row['name']}=={row['version']} --hash=sha256:{row['sha256']}\n" for row in rows), 'utf-8')
    return {'wheels': len(rows), 'wheel_bytes': sum((folder / row['file']).stat().st_size for row in rows), 'python': PYTHON}


def component(target, identity, files):
    count, maximum, single = BUDGETS[identity]
    if not 0 < len(files) <= count or sum(row['bytes'] for row in files) > maximum or any(row['bytes'] > single for row in files):
        raise ValueError('offline_component_budget:' + identity)
    data = {'schema': 1, 'id': identity, 'files': files}
    path = target / 'manifests' / (identity + '.json')
    write_json(path, data)
    if path.stat().st_size > 4 * MIB:
        raise ValueError('offline_manifest_budget')
    return {'id': identity, 'path': 'manifests/' + path.name, 'bytes': path.stat().st_size, 'sha256': sha(path),
            'file_count': len(files), 'total_bytes': sum(row['bytes'] for row in files)}


def stage(root, target, uv_archive, license_root, inputs, renderer, prerequisites, build_id, internal=False):
    root, target, inputs, renderer = checked(root), checked(target, missing=True), checked(inputs), checked(renderer)
    if target.exists() or not re.fullmatch(r'[a-z0-9][a-z0-9-]{0,63}', build_id):
        raise ValueError('offline_new_target_and_build_required')
    source = release_source.collect(root, release_source.read_json(root / 'release/PUBLIC_FILES.json'))
    metadata = release_source.read_json(root / 'release/WINDOWS_INSTALLER.json')
    frontend = release_source.read_json(root / 'release/SUPERDOC_SOURCE.json')
    if not internal and (not frontend.get('rebuild_verified') or not frontend.get('fixed_corresponding_commit')):
        raise ValueError('offline_superdoc_distribution_blocked')
    wheels = release_source.read_json(inputs / 'WHEELS.json')
    if wheels['lock_sha256'] != sha(root / 'uv.lock') or wheels['python'] != PYTHON:
        raise ValueError('offline_lock_changed')
    built = release_source.read_json(inputs / 'BUILT_WHEELS.json') if (inputs / 'BUILT_WHEELS.json').exists() else {}
    if wheel_plan(root, inputs / 'requirements-universal.txt', built) != wheels['wheels']:
        raise ValueError('offline_wheel_plan_changed')
    requirements = ''.join(f"{row['name']}=={row['version']} --hash=sha256:{row['sha256']}\n" for row in wheels['wheels'])
    if (inputs / 'requirements.txt').read_text('utf-8') != requirements:
        raise ValueError('offline_requirements_changed')
    runtime = runtime_files(uv_archive, metadata, license_root)
    python_files = release_source.read_json(inputs / 'PYTHON_FILES.json')
    if python_files.get('archive') != PYTHON:
        raise ValueError('offline_python_source')
    python_rows = {row['path']: row for row in python_files['files']}
    renderer_files = release_source.read_json(inputs / 'RENDERER.json')
    if renderer_files.get('archive') != metadata['renderer']:
        raise ValueError('offline_renderer_source')
    renderer_rows = {row['path']: row for row in renderer_files['files']}
    prerequisites = release_source.read_json(checked(prerequisites))
    if prerequisites.get('schema') != 1 or {row['id'] for row in prerequisites.get('files', [])} != {'webview2', 'net48'}:
        raise ValueError('offline_prerequisites_missing')
    target.mkdir()
    collected = {name: [] for name in BUDGETS}
    seen = set()
    def add(identity, name, path=None, data=None, expected=None):
        relative(name)
        if name.casefold() in seen:
            raise ValueError('offline_duplicate_path')
        seen.add(name.casefold())
        destination = target / name
        if path is not None:
            copy_file(path, destination)
        else:
            destination.parent.mkdir(parents=True, exist_ok=True)
            with destination.open('xb') as stream:
                stream.write(data)
        row = {'path': name, 'bytes': destination.stat().st_size, 'sha256': sha(destination)}
        if expected and row['sha256'] != expected:
            raise ValueError('offline_resource_changed')
        collected[identity].append(row)
    for row in source['files']:
        add('application', 'app/' + row['path'], root / row['path'], expected=row['sha256'])
    for name, data in runtime.items():
        add('licenses' if name.startswith('licenses/') else 'application', name, data=data)
    for row in frontend['assets']:
        name = 'office_agent/webui/static/vendor/superdoc/' + row['local']
        if ('app/' + name).casefold() not in seen:
            add('application', 'app/' + name, root / name, expected=row['sha256'])
    for name, path in tree(inputs / 'python'):
        row = python_rows.pop(name, None)
        if not row or row['bytes'] != path.stat().st_size:
            raise ValueError('offline_python_files')
        add('python', 'runtime/python/' + name, path, expected=row['sha256'])
    if python_rows:
        raise ValueError('offline_python_files_missing')
    for row in wheels['wheels']:
        path = inputs / 'wheels' / row['file']
        if row['bytes'] is not None and path.stat().st_size != row['bytes']:
            raise ValueError('offline_wheel_size')
        add('wheels', 'runtime/wheels/' + row['file'], path, expected=row['sha256'])
        with zipfile.ZipFile(path) as archive:
            if len(archive.infolist()) > 20000:
                raise ValueError('offline_wheel_archive_limit')
            for item in archive.infolist():
                name = item.filename
                base = name.rsplit('/', 1)[-1].lower()
                if not item.is_dir() and ('.dist-info/licenses/' in name.lower() or base in {'license', 'license.txt', 'license.md', 'copying', 'notice', 'metadata'}):
                    if item.file_size > 8 * MIB:
                        raise ValueError('offline_license_limit')
                    relative(name)
                    add('licenses', 'licenses/wheels/' + row['name'] + '/' + name, data=archive.read(item))
    add('wheels', 'runtime/requirements.txt', inputs / 'requirements.txt')
    add('licenses', 'licenses/WHEELS.json', inputs / 'WHEELS.json')
    for name, path in tree(renderer):
        row = renderer_rows.pop(name, None)
        if not row or row['bytes'] != path.stat().st_size:
            raise ValueError('offline_renderer_files')
        add('renderer', 'app/.runtime/libreoffice/' + name, path, expected=row['sha256'])
    if renderer_rows:
        raise ValueError('offline_renderer_files_missing')
    if not (target / 'app/.runtime/libreoffice/program/soffice.com').is_file():
        raise ValueError('offline_renderer_missing')
    for row in prerequisites['files']:
        path = checked(Path(row['path']))
        pinned = next(item for item in release_source.read_json(root / 'release/OFFLINE_INSTALLER.json')['prerequisites'] if item['id'] == row['id'])
        if any(pinned[key] != row[key] for key in ('bytes', 'sha256', 'url', 'signer')):
            raise ValueError('offline_prerequisite_identity')
        if row.get('signature_status') != 'Valid' or 'Microsoft' not in row.get('signer', '') or not row['url'].startswith('https://'):
            raise ValueError('offline_prerequisite_signature')
        add('prerequisites', 'runtime/prerequisites/' + row['id'] + '.exe', path, expected=row['sha256'])
    public_prerequisites = {'schema': 1, 'files': [{key: value for key, value in row.items() if key != 'path'} for row in prerequisites['files']]}
    add('licenses', 'licenses/PREREQUISITES.json', data=(json.dumps(public_prerequisites, ensure_ascii=False, indent=2) + '\n').encode())
    add('licenses', 'licenses/PYTHON_FILES.json', inputs / 'PYTHON_FILES.json')
    add('licenses', 'licenses/RENDERER.json', inputs / 'RENDERER.json')
    components = [component(target, name, rows) for name, rows in collected.items()]
    manifest = {'schema': 2, 'channel': 'offline-preview', 'deployment_mode': 'offline', 'build_id': build_id,
                'version': re.search(r'^version = "([0-9.]+)"', (root / 'pyproject.toml').read_text('utf-8'), re.M)[1],
                'platform': 'windows-x64', 'python': PYTHON['version'], 'python_build': PYTHON,
                'lock_sha256': sha(root / 'uv.lock'), 'components': components, 'environment_mode': 'assemble-local',
                'bundles_python_runtime': True, 'bundles_wheelhouse': True, 'bundles_python_environment': False,
                'bundles_superdoc_compiled_assets': True, 'bundles_renderer': True, 'requires_first_run_network': False,
                'internal_only': bool(internal), 'distribution_blocked': not frontend.get('rebuild_verified'),
                'total_bytes': sum(row['total_bytes'] for row in components), 'files': []}
    write_json(target / 'LiX.install.json', manifest)
    write_json(target / 'bootstrap.json', {'manifest_sha256': sha(target / 'LiX.install.json'), 'version': manifest['version'], 'python': manifest['python']})
    (target / 'INSTALL-NOTES.txt').write_text('LiX 离线内部预览\n\n从本地组件装配运行环境，不调用模型。云端AI与在线目录仍需网络和配置。\n组件完整性、原生安装与干净机验收分别记录；当前未签名。\nSuperDoc对应源码与重建未收口，本候选仅供内部核验，禁止公开分发。\n用户数据默认保留；暂不支持覆盖已有不同构建，升级前保留原安装。\n', 'utf-8-sig')
    files = [row for values in collected.values() for row in values] + [{key: row[key] for key in ('path', 'bytes', 'sha256')} for row in components]
    prepare_compiler_inputs(root, target, dict(manifest, files=files))
    return manifest, source


def main():
    parser = argparse.ArgumentParser()
    commands = parser.add_subparsers(dest='command', required=True)
    gather = commands.add_parser('collect')
    gather.add_argument('--destination', required=True, type=Path)
    build = commands.add_parser('stage')
    for name in ('target', 'uv-archive', 'license-root', 'inputs', 'renderer', 'prerequisites'):
        build.add_argument('--' + name, required=True, type=Path)
    build.add_argument('--build-id', required=True)
    build.add_argument('--internal-only', action='store_true')
    args = parser.parse_args()
    if args.command == 'collect':
        print(json.dumps(collect(ROOT, args.destination)))
    else:
        manifest, source = stage(ROOT, args.target, args.uv_archive, args.license_root, args.inputs, args.renderer, args.prerequisites, args.build_id, args.internal_only)
        write_json(args.target / 'STAGE.json', {'version': manifest['version'], 'bundles_renderer': True, 'channel': manifest['channel'], 'public_source_files': len(source['files']), 'payload_files': sum(row['file_count'] for row in manifest['components']), 'source_manifest_sha256': hashlib.sha256(json.dumps(source, sort_keys=True).encode()).hexdigest()})
        print(json.dumps({'channel': manifest['channel'], 'bytes': manifest['total_bytes'], 'components': len(manifest['components']), 'internal_only': manifest['internal_only']}))


if __name__ == '__main__':
    main()
