import base64
import hashlib
import io
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import sys
import time
import uuid
import zipfile


MAX_PACKAGE = 10 * 1024 * 1024
MAX_EXPANDED = 30 * 1024 * 1024


class UpdateError(ValueError):
    pass


def version(value):
    if type(value) is not str or not re.fullmatch(r'(0|[1-9]\d{0,4})\.(0|[1-9]\d{0,4})\.(0|[1-9]\d{0,4})', value):
        raise UpdateError('update_version_invalid')
    return tuple(map(int, value.split('.')))


def safe_path(root, name, *, managed=False, internal=False):
    if type(name) is not str or '\\' in name or any(char in name for char in ':<>"|?*') or any(ord(char) < 32 for char in name) or len(name) > 240:
        raise UpdateError('update_path_refused')
    parts = PurePosixPath(name).parts
    if not parts or '/'.join(parts) != name or any((part.startswith('.') and not (internal and index == 0 and part == '.office-agent-updates')) or part.endswith(('.', ' ')) or part.upper().split('.')[0] in {'CON', 'PRN', 'AUX', 'NUL', *('COM' + str(i) for i in range(10)), *('LPT' + str(i) for i in range(10))} for index, part in enumerate(parts)):
        raise UpdateError('update_path_refused')
    if managed and name not in {'pyproject.toml', 'uv.lock'} and (len(parts) < 2 or parts[0] != 'office_agent' or '__pycache__' in parts or
                    parts[-1] in {'update_channel.json', 'update_inventory.json'} or Path(name).suffix not in {'.py', '.js', '.css', '.html', '.json'}):
        raise UpdateError('update_path_refused')
    target = root.joinpath(*parts)
    if root not in target.parents:
        raise UpdateError('update_path_refused')
    for path in (target, *target.parents):
        if path.exists() or path.is_symlink():
            value = path.lstat()
            if value.st_nlink > 1 and stat.S_ISREG(value.st_mode) or getattr(value, 'st_file_attributes', 0) & 0x400 or stat.S_ISLNK(value.st_mode):
                raise UpdateError('update_path_refused')
    return target


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_json(path, limit=1024 * 1024):
    if not path.is_file() or path.stat().st_size > limit:
        raise UpdateError('update_metadata_invalid')
    return json.loads(path.read_bytes())


def local_info(root):
    value = read_json(safe_path(root, 'office_agent/release_info.json'))
    version(value['version'])
    return value


def channel(root):
    value = read_json(safe_path(root, 'office_agent/update_channel.json'), 4096)
    if set(value) != {'repository', 'public_key'}:
        raise UpdateError('update_channel_invalid')
    if not value['repository'] and not value['public_key']:
        return value
    if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9-]{0,38}/[A-Za-z0-9_.-]{1,100}', value['repository']):
        raise UpdateError('update_channel_invalid')
    if len(base64.b64decode(value['public_key'], validate=True)) != 32:
        raise UpdateError('update_channel_invalid')
    return value


def inventory(root):
    result = {}
    for path in sorted((root / 'office_agent').rglob('*')):
        if path.is_file() and path.suffix in {'.py', '.js', '.css', '.html', '.json'} and '__pycache__' not in path.parts and path.name not in {'update_inventory.json', 'update_channel.json'}:
            name = path.relative_to(root).as_posix()
            result[name] = digest(safe_path(root, name, managed=True))
    for name in ('pyproject.toml', 'uv.lock'):
        result[name] = digest(safe_path(root, name))
    return {'version': local_info(root)['version'], 'files': result,
            'dependencies': dependency_identity({name: safe_path(root, name).read_bytes() for name in ('pyproject.toml', 'uv.lock')})}


def dependency_identity(contents):
    result = {}
    for name in ('pyproject.toml', 'uv.lock'):
        text = contents[name].decode('utf8').replace('\r\n', '\n')
        if name == 'pyproject.toml':
            text = re.sub(r'(?m)^version = "[^"\n]+"$', 'version = "APP_VERSION"', text, count=1)
        else:
            text = re.sub(r'(\[\[package\]\]\nname = "agent-quickstart"\n)version = "[^"\n]+"', r'\g<1>version = "APP_VERSION"', text, count=1)
        result[name] = hashlib.sha256(text.encode()).hexdigest()
    return result


def check_installed(root):
    expected = read_json(safe_path(root, 'office_agent/update_inventory.json'))
    current = inventory(root)
    if expected != current:
        raise UpdateError('update_local_changes')
    return current


def unpack(body, public_key, target_version):
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
    if len(body) > MAX_PACKAGE:
        raise UpdateError('update_capacity')
    with zipfile.ZipFile(io.BytesIO(body)) as archive:
        names = archive.namelist()
        if len(names) > 1024 or len({name.casefold() for name in names}) != len(names):
            raise UpdateError('update_package_invalid')
        if sum(item.file_size for item in archive.infolist()) > MAX_EXPANDED or any(item.file_size > 4 * 1024 * 1024 or stat.S_ISLNK(item.external_attr >> 16) for item in archive.infolist()):
            raise UpdateError('update_capacity')
        metadata = archive.read('manifest.json')
        if len(metadata) > 200000 or len(archive.read('manifest.sig')) != 64:
            raise UpdateError('update_package_invalid')
        try:
            Ed25519PublicKey.from_public_bytes(base64.b64decode(public_key, validate=True)).verify(archive.read('manifest.sig'), metadata)
        except Exception:
            raise UpdateError('update_signature_invalid') from None
        manifest = json.loads(metadata)
        if set(manifest) != {'format', 'version', 'files', 'dependencies'} or manifest['format'] != 'office-agent-update-v1' or manifest['version'] != target_version:
            raise UpdateError('update_package_invalid')
        version(manifest['version'])
        files = manifest['files']
        if type(files) is not dict or not files or set(names) != {'manifest.json', 'manifest.sig', *('payload/' + name for name in files)}:
            raise UpdateError('update_package_invalid')
        contents = {}
        for name, fingerprint in files.items():
            safe_path(Path.cwd().resolve(), name, managed=True)
            value = archive.read('payload/' + name)
            if type(fingerprint) is not str or hashlib.sha256(value).hexdigest() != fingerprint:
                raise UpdateError('update_package_invalid')
            contents[name] = value
        release = json.loads(contents['office_agent/release_info.json'])
        if release['version'] != target_version:
            raise UpdateError('update_version_invalid')
        if dependency_identity(contents) != manifest['dependencies']:
            raise UpdateError('update_package_invalid')
        return manifest, contents


def write_json(path, value):
    temporary = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    with temporary.open('xb') as stream:
        stream.write(json.dumps(value, ensure_ascii=False, indent=2).encode('utf8'))
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def install(root, ticket, *, replace=os.replace):
    state = safe_path(root, '.office-agent-updates', internal=True)
    lock = state / 'install.lock'
    try:
        with lock.open('xb') as stream:
            stream.write(str(os.getpid()).encode())
    except FileExistsError:
        raise UpdateError('update_install_locked') from None
    try:
        return _install(root, ticket, replace=replace)
    finally:
        lock.unlink()


def _install(root, ticket, *, replace=os.replace):
    state = safe_path(root, '.office-agent-updates', internal=True)
    if ticket.parent != state or not re.fullmatch('[a-f0-9]{32}-ticket.json', ticket.name):
        raise UpdateError('update_path_refused')
    safe_path(root, 'office_agent/release_info.json')
    for path in (state, ticket):
        if path.is_symlink() or getattr(path.lstat(), 'st_file_attributes', 0) & 0x400 or path.is_file() and path.stat().st_nlink != 1:
            raise UpdateError('update_path_refused')
    request = read_json(ticket, 16000)
    if set(request) != {'id', 'version', 'sha256', 'created'} or not re.fullmatch('[a-f0-9]{32}', request['id']) or ticket.name != request['id'] + '-ticket.json' or not 0 <= time.time() - request['created'] <= 300:
        raise UpdateError('update_plan_expired')
    token = request['id']
    package = state / (token + '.zip')
    if package.is_symlink() or getattr(package.lstat(), 'st_file_attributes', 0) & 0x400 or package.stat().st_nlink != 1 or package.stat().st_size > MAX_PACKAGE or digest(package) != request['sha256']:
        raise UpdateError('update_package_invalid')
    before = check_installed(root)
    if version(request['version']) <= version(before['version']):
        raise UpdateError('update_version_invalid')
    manifest, contents = unpack(package.read_bytes(), channel(root)['public_key'], request['version'])
    if manifest['dependencies'] != before['dependencies']:
        raise UpdateError('update_dependencies_changed')
    paths = sorted(set(before['files']) | set(contents))
    for name in paths:
        safe_path(root, name, managed=True)
    backup = state / (token + '-backup')
    backup.mkdir()
    for name in before['files']:
        destination = backup / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(safe_path(root, name, managed=True), destination)
        if digest(destination) != before['files'][name]:
            raise UpdateError('update_local_changes')
    previous_inventory = (root / 'office_agent/update_inventory.json').read_bytes()
    (backup / 'office_agent/update_inventory.json').write_bytes(previous_inventory)
    audit = state / (token + '-audit.json')
    write_json(audit, {'status': 'prepared', 'from': before['version'], 'to': request['version']})
    changed = []
    try:
        if check_installed(root) != before:
            raise UpdateError('update_local_changes')
        write_json(audit, {'status': 'installing', 'from': before['version'], 'to': request['version']})
        for name in paths:
            target = safe_path(root, name, managed=True)
            if name in before['files'] and (not target.is_file() or digest(target) != before['files'][name]) or name not in before['files'] and target.exists():
                raise UpdateError('update_local_changes')
            changed.append(name)
            if name not in contents:
                target.unlink()
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            temporary = target.with_name(target.name + '.' + token + '.new')
            with temporary.open('xb') as stream:
                stream.write(contents[name])
                stream.flush()
                os.fsync(stream.fileno())
            replace(temporary, target)
        after = inventory(root)
        if after != {'version': request['version'], 'files': manifest['files'], 'dependencies': before['dependencies']}:
            raise UpdateError('update_verification_failed')
        write_json(root / 'office_agent/update_inventory.json', after)
        write_json(audit, {'status': 'installed', 'from': before['version'], 'to': request['version']})
        ticket.unlink()
        return 'installed'
    except Exception:
        try:
            for name in reversed(changed):
                target = safe_path(root, name, managed=True)
                if name in before['files']:
                    shutil.copyfile(backup / name, target)
                elif target.exists():
                    target.unlink()
            (root / 'office_agent/update_inventory.json').write_bytes(previous_inventory)
            if inventory(root) != before:
                raise UpdateError('update_uncertain')
            write_json(audit, {'status': 'rolled_back', 'from': before['version'], 'to': request['version']})
        except Exception:
            write_json(audit, {'status': 'uncertain', 'from': before['version'], 'to': request['version']})
            raise UpdateError('update_uncertain') from None
        raise UpdateError('update_install_failed_rolled_back') from None


def restart_application(root):
    import subprocess
    launcher = safe_path(root, 'Start-OfficeAgent.cmd')
    environment = dict(os.environ, OFFICE_AGENT_RESTART_ENTRY=str(launcher))
    command = Path(os.environ.get('SystemRoot', 'C:/Windows')) / 'System32/cmd.exe'
    return subprocess.Popen(f'"{command}" /d /v:off /s /c ""%OFFICE_AGENT_RESTART_ENTRY%""',
                            cwd=root, env=environment, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))


def worker(root, ticket, pid, restart=True):
    import ctypes
    if sys.platform != 'win32':
        raise UpdateError('update_platform_unsupported')
    if ticket.parent != safe_path(root, '.office-agent-updates', internal=True) or not re.fullmatch('[a-f0-9]{32}-ticket.json', ticket.name):
        raise UpdateError('update_path_refused')
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.OpenProcess.argtypes = [ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32]
    kernel.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel.OpenProcess(0x100000, False, pid)
    if not handle:
        raise UpdateError('update_parent_unavailable')
    try:
        write_json(ticket.with_suffix('.ready'), {'pid': os.getpid(), 'parent': pid})
        if kernel.WaitForSingleObject(handle, 120000) != 0:
            raise UpdateError('update_exit_timeout')
    finally:
        kernel.CloseHandle(handle)
    result = install(root, ticket)
    if result == 'installed' and restart:
        restart_application(root)


if __name__ == '__main__':
    try:
        worker(Path(sys.argv[1]).absolute(), Path(sys.argv[2]), int(sys.argv[3]), restart='--no-restart' not in sys.argv[4:])
    except Exception as error:
        try:
            root, ticket = Path(sys.argv[1]).absolute(), Path(sys.argv[2])
            state = safe_path(root, '.office-agent-updates', internal=True)
            if ticket.parent == state and re.fullmatch('[a-f0-9]{32}-ticket.json', ticket.name):
                audit = state / (ticket.name.removesuffix('-ticket.json') + '-audit.json')
                report = read_json(audit, 16000)
                if report['status'] == 'queued':
                    code = str(error) if isinstance(error, UpdateError) and re.fullmatch('update_[a-z_]{1,60}', str(error)) else 'update_failed'
                    write_json(audit, {**report, 'status': 'refused', 'code': code})
        except Exception:
            pass
        sys.exit(1)
