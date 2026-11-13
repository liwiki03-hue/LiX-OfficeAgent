import importlib.metadata
import json
import os
from pathlib import Path
import platform
import sys


def inspect(root):
    import playwright
    node = Path(playwright.__file__).resolve().parent / 'driver/node.exe'
    edge = Path('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe')
    runtime = root / '.runtime/libreoffice/program/soffice.com'
    checks = {'windows': sys.platform == 'win32', 'python_312': sys.version_info[:2] == (3, 12),
              'node': node.is_file(), 'edge': edge.is_file(), 'renderer': runtime.is_file(),
              'pythonnet': importlib.metadata.version('pythonnet') == '3.0.5'}
    return {'ok': all(checks.values()), 'checks': checks, 'python': platform.python_version(),
            'node_directory': str(node.parent), 'real_office': False, 'real_service': False}


if __name__ == '__main__':
    result = inspect(Path(__file__).resolve().parents[1])
    print(json.dumps(result))
    if result['ok'] and os.environ.get('GITHUB_PATH'):
        with open(os.environ['GITHUB_PATH'], 'a', encoding='utf-8') as stream:
            stream.write(result['node_directory'] + '\n')
    raise SystemExit(0 if result['ok'] else 1)
