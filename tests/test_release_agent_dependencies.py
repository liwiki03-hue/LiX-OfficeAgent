import ast
import json
from pathlib import Path, PurePosixPath


ROOT = Path(__file__).resolve().parents[1]


def test_public_python_sources_include_their_local_agent_dependencies():
    names = set(json.loads((ROOT / 'release/PUBLIC_FILES.json').read_text('utf-8'))['files'])
    missing = []
    for name in sorted(names):
        if not name.endswith('.py'):
            continue
        parent = PurePosixPath(name).parent.parts
        for node in ast.walk(ast.parse((ROOT / name).read_text('utf-8-sig'), filename=name)):
            modules = []
            if isinstance(node, ast.Import):
                modules = [alias.name.split('.') for alias in node.names]
            elif isinstance(node, ast.ImportFrom):
                base = list(parent[:len(parent) - node.level + 1]) if node.level else []
                modules = [base + node.module.split('.')] if node.module else [base + [alias.name] for alias in node.names]
            for parts in modules:
                if not parts or parts[0] != 'office_agent':
                    continue
                path = '/'.join(parts)
                candidates = (path + '.py', path + '/__init__.py')
                for target in candidates:
                    if (ROOT / target).is_file() and target not in names:
                        missing.append((name, node.lineno, target))
    assert not missing, missing
