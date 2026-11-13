import hashlib
import json
import os
from pathlib import Path
import platform


def _stamp(path):
    info = path.stat()
    if path.is_symlink() or not path.is_file() or getattr(info, 'st_file_attributes', 0) & 0x400:
        raise ValueError('resource_link')
    return (str(path), info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns)


def _content(stamp):
    path = Path(stamp[0])
    if stamp[3] > 256 * 1024 * 1024:
        raise ValueError('resource_capacity')
    result = hashlib.sha256()
    received = 0
    with path.open('rb') as stream:
        while data := stream.read(1024 * 1024):
            received += len(data)
            if received > stamp[3]:
                raise ValueError('resource_changed')
            result.update(data)
    if _stamp(path) != stamp:
        raise ValueError('resource_changed')
    return result.hexdigest()


def environment_identity(runtime):
    try:
        if os.name != 'nt':
            raise ValueError('unsupported_environment')
        import win32api
        import winreg
        from .fidelity import RENDER_DPI, _profile
        import inspect
        font_roots = [Path(win32api.GetWindowsDirectory()) / 'Fonts',
                      Path(os.environ['LOCALAPPDATA']) / 'Microsoft/Windows/Fonts',
                      runtime.parent.parent / 'share/fonts']
        resources = set()
        for folder in font_roots:
            if folder.is_symlink() or (folder.exists() and getattr(folder.stat(), 'st_file_attributes', 0) & 0x400):
                raise ValueError('resource_link')
            if not folder.exists():
                continue
            for base, directories, files in os.walk(folder, followlinks=False):
                if any((Path(base) / name).is_symlink() or getattr((Path(base) / name).stat(), 'st_file_attributes', 0) & 0x400 for name in directories):
                    raise ValueError('resource_link')
                for name in files:
                    path = Path(base) / name
                    if path.suffix.lower() in {'.ttf', '.otf', '.ttc', '.fon', '.fnt'}:
                        _stamp(path)
                        resources.add(path.resolve())
                        if len(resources) > 4096:
                            raise ValueError('resource_capacity')
        if not resources:
            raise ValueError('fonts_unavailable')
        registry = []
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for suffix in ('Fonts', 'FontSubstitutes', 'FontLink\\SystemLink'):
                try:
                    with winreg.OpenKey(hive, 'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\' + suffix) as key:
                        count = winreg.QueryInfoKey(key)[1]
                        if count > 4096:
                            raise ValueError('resource_capacity')
                        values = [winreg.EnumValue(key, index) for index in range(count)]
                except FileNotFoundError:
                    values = []
                registry.append([str(hive), suffix, values])
                if suffix == 'Fonts':
                    for name, value, kind in values:
                        if type(value) is not str:
                            continue
                        registered = Path(value)
                        if not registered.is_absolute():
                            candidates = [folder / registered for folder in font_roots if (folder / registered).is_file()]
                            if not candidates and registered.name == str(registered):
                                continue
                            if len(candidates) != 1:
                                raise ValueError('font_registration_unknown')
                            registered = candidates[0]
                        if registered.resolve() not in resources:
                            if not registered.exists() and any(folder in registered.parents for folder in font_roots):
                                continue
                            from win32com.shell import shell, shellcon
                            programs = [Path(shell.SHGetFolderPath(0, value, 0, 0)) for value in (shellcon.CSIDL_PROGRAM_FILES, shellcon.CSIDL_PROGRAM_FILESX86)]
                            if registered.suffix.lower() not in {'.ttf', '.otf', '.ttc', '.fon', '.fnt'} or not any(folder in registered.parents for folder in programs):
                                raise ValueError('font_registration_unknown')
                            _stamp(registered)
                            resources.add(registered.resolve())
        font_count = len(resources)
        for path in runtime.parent.iterdir():
            if path.is_file() and path.suffix.lower() in {'.com', '.exe', '.dll', '.ini', '.rc', '.rdb'}:
                _stamp(path)
                resources.add(path.resolve())
        registry_root = runtime.parent.parent / 'share/registry'
        if registry_root.is_dir():
            for base, directories, files in os.walk(registry_root, followlinks=False):
                if any((Path(base) / name).is_symlink() for name in directories):
                    raise ValueError('resource_link')
                for name in files:
                    if Path(name).suffix.lower() in {'.xcd', '.xcu', '.xcs'}:
                        _stamp(Path(base) / name)
                        resources.add((Path(base) / name).resolve())
        stamps = sorted((_stamp(path) for path in resources))
        if len(stamps) > 4096 or sum(stamp[3] for stamp in stamps) > 2 * 1024 * 1024 * 1024:
            raise ValueError('resource_capacity')
        manifest = {'files': [(stamp, _content(stamp)) for stamp in stamps], 'registry': registry,
                    'profile': hashlib.sha256(inspect.getsource(_profile).encode()).hexdigest(), 'dpi': RENDER_DPI, 'os': platform.version()}
        if stamps != sorted(_stamp(path) for path in resources):
            raise ValueError('resource_changed')
        return {'complete': True, 'sha256': hashlib.sha256(json.dumps(manifest, sort_keys=True).encode()).hexdigest(),
                'fonts': font_count, 'resources': len(stamps), 'method': 'windows-registered-system-user-runtime-fonts-and-renderer-content-v1',
                'scope': 'registered Windows/system/user/runtime fonts, font substitution/link registry, renderer binaries/configuration and isolated profile; external ephemeral font providers not proven'}
    except Exception as error:
        return {'complete': False, 'sha256': None, 'method': 'unavailable-or-outside-bounded-resource-roots',
                'reason': str(error) if type(error) is ValueError and str(error) in {'resource_link', 'resource_capacity', 'resource_changed', 'fonts_unavailable', 'font_registration_unknown', 'unsupported_environment'} else type(error).__name__,
                'scope': 'fresh conversion permitted; reuse refused because necessary font/configuration identity is not established'}
