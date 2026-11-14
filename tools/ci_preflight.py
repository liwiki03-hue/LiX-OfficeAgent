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
    return {'ok': all(checks.values()), 'checks': checks, 'python': platform.python_version(),
            'node_directory': str(node.parent), 'real_office': False, 'real_service': False}


if __name__ == '__main__':
    result = inspect(Path(__file__).resolve().parents[1])
    print(json.dumps(result))
    if result['ok'] and os.environ.get('GITHUB_PATH'):
        with open(os.environ['GITHUB_PATH'], 'a', encoding='utf-8') as stream:
            stream.write(result['node_directory'] + '\n')
    raise SystemExit(0 if result['ok'] else 1)
