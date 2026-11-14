import os
from pathlib import Path
import struct
import sys


def _executable(path):
    path = Path(path).expanduser().absolute()
    if path.name.casefold() != 'msedge.exe' or not path.is_file() or path.is_symlink():
        return None
    try:
        with path.open('rb') as stream:
            header = stream.read(64)
            if len(header) != 64 or header[:2] != b'MZ':
                return None
            offset = struct.unpack_from('<I', header, 60)[0]
            if not 64 <= offset <= 2 * 1024 * 1024:
                return None
            stream.seek(offset)
            if stream.read(4) != b'PE\0\0':
                return None
    except OSError:
        return None
    return path


def find_edge(required=None):
    configured = os.getenv('LIX_EDGE_PATH') or os.getenv('OFFICE_EDGE_PATH')
    if configured:
        found = _executable(configured)
        if found is None:
            raise OSError('Configured Edge executable is invalid')
        return found
    candidates = []
    if sys.platform == 'win32':
        import winreg
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for flags in (winreg.KEY_READ | winreg.KEY_WOW64_64KEY, winreg.KEY_READ | winreg.KEY_WOW64_32KEY):
                try:
                    with winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe', 0, flags) as key:
                        candidates.append(winreg.QueryValueEx(key, None)[0].strip('"'))
                except OSError:
                    pass
    for name in ('ProgramFiles', 'ProgramFiles(x86)', 'LOCALAPPDATA'):
        if os.getenv(name):
            candidates.append(str(Path(os.environ[name]) / 'Microsoft/Edge/Application/msedge.exe'))
    for path in dict.fromkeys(candidates):
        found = _executable(path)
        if found:
            return found
    if required is True or required is None and os.getenv('CI', '').casefold() in {'1', 'true'}:
        raise OSError('Edge is required; configure LIX_EDGE_PATH before browser tests')
    return None
