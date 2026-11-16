import importlib.metadata
import json
import os
from pathlib import Path
import platform
import sys

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def inspect(root):
    import playwright
    node = Path(playwright.__file__).resolve().parent / 'driver/node.exe'
    from office_agent.browser_runtime import find_edge
    edge = find_edge(required=False)
    runtime = root / '.runtime/libreoffice/program/soffice.com'
    checks = {'windows': sys.platform == 'win32', 'python_312': sys.version_info[:2] == (3, 12),
              'node': node.is_file(), 'edge': edge is not None, 'renderer': runtime.is_file(),
              'pythonnet': importlib.metadata.version('pythonnet') == '3.0.5'}
    environment = {'python_utf8_mode': sys.flags.utf8_mode,
                   'stdout_encoding': sys.stdout.encoding, 'filesystem_encoding': sys.getfilesystemencoding()}
    if sys.platform == 'win32':
        import ctypes
        from ctypes import wintypes
        user = ctypes.windll.user32
        environment['desktop_pixels'] = [user.GetSystemMetrics(0), user.GetSystemMetrics(1)]
        environment['oem_code_page'] = ctypes.windll.kernel32.GetOEMCP()
        area = wintypes.RECT()
        if user.SystemParametersInfoW(0x30, 0, ctypes.byref(area), 0):
            environment['working_area'] = [area.left, area.top, area.right - area.left, area.bottom - area.top]
    return {'ok': all(checks.values()), 'checks': checks, 'python': platform.python_version(),
            'node_directory': str(node.parent), 'environment': environment,
            'real_office': False, 'real_service': False}


if __name__ == '__main__':
    result = inspect(Path(__file__).resolve().parents[1])
    print(json.dumps(result))
    if result['ok'] and os.environ.get('GITHUB_PATH'):
        with open(os.environ['GITHUB_PATH'], 'a', encoding='utf-8') as stream:
            stream.write(result['node_directory'] + '\n')
    raise SystemExit(0 if result['ok'] else 1)
