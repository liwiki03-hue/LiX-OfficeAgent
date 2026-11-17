import importlib.metadata
import json
import os
from pathlib import Path
import sys

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tools.install_frontend_assets import checked, spec, sha


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
    import importlib
    root = checked(root.absolute())
    details = {}
    def run(name, operation, next_step):
        try:
            passed = bool(operation())
            details[name] = {'status': 'passed' if passed else 'failed', 'next_step': '' if passed else next_step}
        except Exception as error:
            passed = False
            details[name] = {'status': 'failed', 'code': type(error).__name__, 'next_step': next_step}
        return passed
    def frontend():
        metadata = spec(root)
        return all(sha(checked(root / 'office_agent/webui/static/vendor/superdoc' / item['local']).read_bytes()) == item['sha256'] for item in metadata['assets'])
    def renderer():
        from office_agent import fidelity
        from tools.installed_runtime import digest
        binary = checked(root / '.runtime/libreoffice/program/soffice.com')
        offline = checked(root.parent / 'LiX.install.json', missing=True)
        if os.environ.get('LIX_DEPLOYMENT_MODE') == 'offline':
            if not offline.is_file() or offline.stat().st_size > 2 * 1024 * 1024:
                return False
            manifest = json.loads(offline.read_text('utf-8'))
            if manifest.get('schema') != 2 or manifest.get('deployment_mode') != 'offline':
                return False
            component = next(item for item in manifest['components'] if item['id'] == 'renderer')
            listing = checked(root.parent / component['path'])
            if listing.stat().st_size != component['bytes'] or digest(listing) != component['sha256']:
                return False
            rows = json.loads(listing.read_text('utf-8'))['files']
            expected = next(row for row in rows if row['path'] == 'app/.runtime/libreoffice/program/soffice.com')
            return fidelity.runtime_path() == binary.resolve() and digest(binary) == expected['sha256']
        receipt = checked(root / '.runtime/installed-renderer.json')
        if receipt.stat().st_size > 65536:
            return False
        saved = json.loads(receipt.read_text('utf-8'))
        return (fidelity.runtime_path() == binary.resolve() and saved.get('executable_sha256') == digest(binary)
                and saved.get('archive_sha256') == '788ce7d4b56460357f57552cb8cd848e8f2254f28f6e51fa61e0c65a18096373')
    checks = {'windows_x64': sys.platform == 'win32' and sys.maxsize > 2**32,
              'python': sys.version_info[:3] == (3, 12, 14),
              'pythonnet': run('pythonnet', lambda: importlib.metadata.version('pythonnet') == '3.0.5' and importlib.import_module('clr') and importlib.import_module('webview.platforms.winforms'), '重新准备第 1 阶段'),
              'webview2': run('webview2', webview_version, '安装或修复 Microsoft WebView2 Runtime'),
              'frontend': run('frontend', frontend, '重新准备第 2 阶段'),
              'renderer_executable': run('renderer_executable', renderer, '重新准备第 3 阶段')}
    for name in ('docx', 'openpyxl', 'pptx', 'sqlite_vec'):
        checks[name] = run(name, lambda name=name: importlib.import_module(name), '重新准备第 1 阶段')
    for name, passed in checks.items():
        details.setdefault(name, {'status': 'passed' if passed else 'failed', 'next_step': '' if passed else '使用安装器固定的 Python 3.12.14 x64 重新准备'})
    if os.environ.get('LIX_DEPLOYMENT_MODE') == 'offline':
        def packages():
            import re
            from packaging.requirements import Requirement
            path = checked(root.parent / 'runtime/requirements.txt')
            if path.stat().st_size > 128 * 1024:
                return False
            lines = path.read_text('utf-8').splitlines()
            if not 1 <= len(lines) <= 256:
                return False
            for line in lines:
                if not re.fullmatch(r'[a-z0-9-]+==[A-Za-z0-9.+_-]+ --hash=sha256:[a-f0-9]{64}', line):
                    return False
                requirement = Requirement(line.split(' --hash=')[0])
                if importlib.metadata.version(requirement.name) not in requirement.specifier:
                    return False
            return True
        checks['locked_packages'] = run('locked_packages', packages, '使用同一离线安装介质重新装配运行依赖')
    return {'ok': all(checks.values()), 'checks': checks, 'gui_visible_verified': False,
            'deployment_mode': os.environ.get('LIX_DEPLOYMENT_MODE', 'network-preview'),
            'cloud_ai_requires_network': True,
            'details': details, 'network_requests': 0, 'preparation_performed': False,
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
