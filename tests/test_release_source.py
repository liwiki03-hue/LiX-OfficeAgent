import importlib.util
import json
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('release_source', ROOT / 'tools/release_source.py')
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)


def setup_source(tmp_path):
    root = tmp_path / 'source'
    root.mkdir()
    (root / 'README.md').write_text('synthetic public source', encoding='utf-8')
    policy = {'schema': 1, 'files': ['README.md'], 'scan_exceptions': []}
    return root, policy


def test_explicit_files_only_export_and_source_conflict(tmp_path):
    root, policy = setup_source(tmp_path)
    (root / '.env').write_text('PRIVATE=not-read', encoding='utf-8')
    manifest = release.collect(root, policy)
    target = tmp_path / 'candidate'
    result = release.export(root, target, manifest, policy)
    assert result['source_verified'] and not result['publication_authorized']
    assert not (target / '.env').exists()
    assert json.loads((target / 'SOURCE_MANIFEST.json').read_text()) == manifest
    with pytest.raises(ValueError, match='new directory'):
        release.export(root, target, manifest, policy)
    (root / 'README.md').write_text('changed', encoding='utf-8')
    with pytest.raises(ValueError, match='source or policy changed'):
        release.export(root, tmp_path / 'rejected', manifest, policy)
    assert not (tmp_path / 'rejected').exists()


@pytest.mark.parametrize('name', ['../secret', '/absolute', 'C:/secret', 'a\\b', '.env',
                                 '.env.private', '.runtime/secret', 'a/../b', 'a./b', 'x.db'])
def test_private_and_escaping_paths_rejected(name):
    with pytest.raises(ValueError):
        release.safe_name(name)


def test_case_collision_and_reparse_refused(tmp_path, monkeypatch):
    root, policy = setup_source(tmp_path)
    with pytest.raises(ValueError, match='duplicate'):
        release.collect(root, {**policy, 'files': ['README.md', 'readme.md']})
    original = Path.is_symlink
    monkeypatch.setattr(Path, 'is_symlink', lambda path: path == root / 'README.md' or original(path))
    with pytest.raises(ValueError, match='links'):
        release.collect(root, policy)


def test_secret_report_redacts_match_and_exact_exception_is_stale_after_change(tmp_path):
    root, policy = setup_source(tmp_path)
    token = 'sk-' + 'testcredential' * 4
    data = token.encode()
    (root / 'README.md').write_bytes(data)
    with pytest.raises(ValueError) as error:
        release.collect(root, policy)
    assert token not in str(error.value) and 'provider-token' in str(error.value)
    exception = {'path': 'README.md', 'rule': 'provider-token', 'file_sha256': release.sha(data),
                 'match_sha256': release.sha(data), 'reason': 'synthetic test only'}
    release.collect(root, {**policy, 'scan_exceptions': [exception]})
    (root / 'README.md').write_bytes(data + b' changed')
    with pytest.raises(ValueError, match='review-required'):
        release.collect(root, {**policy, 'scan_exceptions': [exception]})


def test_manifest_tamper_and_target_symlink_parent_refused(tmp_path, monkeypatch):
    root, policy = setup_source(tmp_path)
    manifest = release.collect(root, policy)
    manifest['files'][0]['sha256'] = '0' * 64
    with pytest.raises(ValueError, match='changed'):
        release.verify(root, manifest, policy)
    original = Path.is_symlink
    monkeypatch.setattr(Path, 'is_symlink', lambda path: path == tmp_path or original(path))
    with pytest.raises(ValueError, match='target parent'):
        release.export(root, tmp_path / 'candidate', release.collect(root, policy), policy)


def test_runtime_dependency_and_pinned_workflow():
    import tomllib
    project = tomllib.loads((ROOT / 'pyproject.toml').read_text(encoding='utf-8'))
    assert 'python-pptx>=1,<2' in project['project']['dependencies']
    assert project['project']['license'] == 'AGPL-3.0-only'
    assert 'GNU AFFERO GENERAL PUBLIC LICENSE' in (ROOT / 'LICENSE').read_text()
    assert 'python-pptx>=1,<2' not in project['dependency-groups']['dev']
    installer = (ROOT / 'tools/Install-DocumentRuntime.ps1').read_text(encoding='utf-8')
    assert 'downloadarchive.documentfoundation.org/libreoffice/old/' in installer
    assert '788CE7D4B56460357F57552CB8CD848E8F2254F28F6E51FA61E0C65A18096373' in installer
    workflow = (ROOT / '.github/workflows/verify.yml').read_text(encoding='utf-8')
    assert 'contents: read' in workflow and 'uv run pytest -q' in workflow
    assert 'pull_request_target' not in workflow and '${{ secrets.' not in workflow
    import re
    actions = re.findall(r'uses: (.+)', workflow)
    assert actions and all(re.fullmatch(r'[\w-]+/[\w-]+@[a-f0-9]{40}', item) for item in actions)


def test_index_missing_and_unlisted_private_file_fail_closed(tmp_path, monkeypatch):
    from types import SimpleNamespace
    root, policy = setup_source(tmp_path)
    assert release.check_git_index(root, policy)['reason'] == 'no_git_repository'
    (root / '.git').mkdir()
    def result(names):
        monkeypatch.setattr(release.subprocess, 'run', lambda *a, **kw: SimpleNamespace(returncode=0, stdout=names))
    result(b'README.md\x00')
    assert release.check_git_index(root, policy)['verified']
    result(b'README.md\x00.env\x00')
    with pytest.raises(ValueError, match='private'):
        release.check_git_index(root, policy)
    result(b'README.md\x00internal-plan.md\x00')
    with pytest.raises(ValueError, match='outside'):
        release.check_git_index(root, policy)
    result(b'')
    with pytest.raises(ValueError, match='missing'):
        release.check_git_index(root, policy)


def test_source_archive_reopens_with_exact_manifest_and_no_private_files(tmp_path):
    import zipfile
    root, policy = setup_source(tmp_path)
    (root / '.env').write_text('private configuration excluded', encoding='utf-8')
    manifest = release.collect(root, policy)
    target = tmp_path / 'source.zip'
    result = release.archive(root, target, manifest, policy)
    assert not result['publication_authorized'] and result['sha256'] == release.sha(target.read_bytes())
    with zipfile.ZipFile(target) as package:
        assert set(package.namelist()) == {'README.md', 'SOURCE_MANIFEST.json'}
        assert json.loads(package.read('SOURCE_MANIFEST.json')) == manifest
    with pytest.raises(FileExistsError):
        release.archive(root, target, manifest, policy)


def test_public_python_imports_have_their_local_tool_dependencies():
    import ast
    policy = json.loads((ROOT / 'release/PUBLIC_FILES.json').read_text('utf-8'))
    names = set(policy['files'])
    missing = []
    for name in sorted(names):
        if not name.endswith('.py'):
            continue
        tree = ast.parse((ROOT / name).read_text('utf-8-sig'), filename=name)
        for node in ast.walk(tree):
            modules = []
            if isinstance(node, ast.Import):
                modules = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and not node.level and node.module:
                modules = [node.module]
                if node.module in {'tools', 'examples'}:
                    modules = [node.module + '.' + alias.name for alias in node.names]
            for module in modules:
                if module.split('.')[0] not in {'tools', 'examples', 'main'}:
                    continue
                path = module.replace('.', '/')
                if path + '.py' not in names and not any(item.startswith(path + '/') for item in names):
                    missing.append((name, node.lineno, module))
    assert not missing, missing
