import importlib.metadata
import json
import os
from pathlib import Path
import sys

from tools.install_frontend_assets import checked, prepare


ROOT = Path(__file__).absolute().parents[1]


def webview_version():
    if sys.platform != 'win32':
        return None
    import winreg
    locations = [
        (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'),
        (winreg.HKEY_CURRENT_USER, r'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'),
    ]
    for hive, key in locations:
        try:
            with winreg.OpenKey(hive, key, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as opened:
                value, kind = winreg.QueryValueEx(opened, 'pv')
            if kind == winreg.REG_SZ and value and value != '0.0.0.0':
                parts = value.split('.')
                if len(parts) == 4 and all(part.isascii() and part.isdigit() for part in parts):
                    return value
        except OSError:
            continue
    return None


def inspect(root):
    import clr
    import webview.platforms.winforms
    import docx
    import openpyxl
    import pptx
    import sqlite_vec
    from office_agent import fidelity
    root = checked(root.absolute())
    assets = prepare(root)
    checks = {'windows_x64': sys.platform == 'win32' and sys.maxsize > 2**32,
              'python': sys.version_info[:3] == (3, 12, 14),
              'pythonnet': importlib.metadata.version('pythonnet') == '3.0.5',
              'webview2': bool(webview_version()), 'frontend': assets['status'] in ('already_verified', 'verified'),
              'renderer_executable': fidelity.runtime_path() == checked(root / '.runtime/libreoffice/program/soffice.com').resolve()}
    return {'ok': all(checks.values()), 'checks': checks, 'gui_visible_verified': False,
            'rendering_verified': False, 'model_requests': 0, 'new_windows_verified': False}


if __name__ == '__main__':
    import threading
    watchdog = threading.Timer(45, lambda: os._exit(124))
    watchdog.daemon = True
    watchdog.start()
    try:
        result = inspect(ROOT)
        print(json.dumps(result))
    except Exception as error:
        print(json.dumps({'ok': False, 'error_type': type(error).__name__, 'model_requests': 0}))
        raise SystemExit(1) from None
    finally:
        watchdog.cancel()
    raise SystemExit(0 if result['ok'] else 1)
