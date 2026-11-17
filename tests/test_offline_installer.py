import hashlib
import json
from pathlib import Path
import shutil
import sys
import uuid

import pytest

from office_agent.diagram_process import _run_owned
from tools import build_windows_installer as installer
from tools import offline_installer as offline


ROOT = Path(__file__).absolute().parents[1]


@pytest.mark.parametrize('name', ['../escape', '/absolute', 'C:/file', 'a\\b', 'a//b', 'a/./b', 'CON.txt', 'a/LPT1', 'a/file.', 'a/file ', 'a\x00b'])
def test_offline_paths_reject_ambiguous_windows_names(name):
    with pytest.raises(ValueError, match='offline_path'):
        offline.relative(name)


def test_offline_component_has_fixed_budget_and_detects_duplicate_input(tmp_path):
    with pytest.raises(ValueError, match='component_budget'):
        offline.component(tmp_path, 'wheels', [{'path': 'runtime/wheels/file', 'bytes': 513 * offline.MIB, 'sha256': '0' * 64}])
    source = tmp_path / 'source'
    source.mkdir()
    (source / 'value').write_bytes(b'fixed')
    offline.copy_file(source / 'value', tmp_path / 'target/value')
    assert (tmp_path / 'target/value').read_bytes() == b'fixed'
    with pytest.raises(FileExistsError):
        offline.copy_file(source / 'value', tmp_path / 'target/value')


def test_offline_wheels_cannot_use_foreign_or_unhashed_requirements(tmp_path):
    (tmp_path / 'uv.lock').write_text('version = 1\npackage = []\n', 'utf-8')
    path = tmp_path / 'requirements.txt'
    path.write_text('package @ https://example.invalid/content\n', 'utf-8')
    with pytest.raises(ValueError, match='requirement_unsupported'):
        offline.wheel_plan(tmp_path, path)


@pytest.fixture(scope='module')
def offline_compiled(tmp_path_factory):
    if sys.platform != 'win32':
        pytest.skip('Windows compiled offline launcher')
    folder = tmp_path_factory.mktemp('offline-launcher')
    payload = folder / 'payload'
    payload.mkdir()
    compiler_root = folder / 'compiler-source'
    for name in ('tools/windows/installed_launcher.cs', 'tools/windows/process_job.cs',
                 'tools/windows/bootstrap_network.cs', 'tools/windows/offline_manifest.cs',
                 'office_agent/webui/static/brand/lix.ico'):
        target = compiler_root / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / name, target)
    launcher = compiler_root / 'tools/windows/installed_launcher.cs'
    source = launcher.read_text('utf-8')
    assert source.count('"LiXInstalledRuntime"') == 1
    mutex_name = 'LiXOfflineFixture-' + uuid.uuid4().hex
    launcher.write_text(source.replace('"LiXInstalledRuntime"', '"' + mutex_name + '"'), 'utf-8')
    (folder / 'fixture-mutex.txt').write_text(mutex_name, 'utf-8')
    content = {
        'application': ('app/uv.lock', b'locked fixture'),
        'python': ('runtime/python/python.exe', b'MZfixture'),
        'wheels': ('runtime/requirements.txt', b'fixed==1 --hash=sha256:' + b'0' * 64),
        'renderer': ('app/.runtime/libreoffice/program/soffice.com', b'MZrenderer'),
        'prerequisites': ('runtime/prerequisites/net48.exe', b'MZprerequisite'),
        'licenses': ('licenses/NOTICE', b'fixture license'),
    }
    components = []
    for identity, (name, data) in content.items():
        path = payload / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        components.append(offline.component(payload, identity, [{'path': name, 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}]))
    manifest = {'schema': 2, 'version': '0.7.11', 'channel': 'offline-preview', 'deployment_mode': 'offline', 'platform': 'windows-x64',
                'python': '3.12.14', 'build_id': 'test-offline', 'environment_mode': 'assemble-local', 'requires_first_run_network': False,
                'bundles_python_runtime': True, 'bundles_wheelhouse': True, 'lock_sha256': offline.sha(payload / 'app/uv.lock'),
                'components': components, 'total_bytes': sum(item['total_bytes'] for item in components)}
    offline.write_json(payload / 'LiX.install.json', manifest)
    offline.write_json(payload / 'bootstrap.json', {'version': '0.7.11', 'python': '3.12.14', 'manifest_sha256': offline.sha(payload / 'LiX.install.json')})
    result = _run_owned((sys.executable, '-I', '-c', 'import sys;from pathlib import Path;sys.path.insert(0,sys.argv[1]);from tools.build_windows_installer import compile_launcher;compile_launcher(Path(sys.argv[3]),Path(sys.argv[2]))', str(ROOT), str(payload), str(compiler_root)), timeout_ms=30000)
    assert result.empty_confirmed and result.exit_code == 0
    return payload


def test_offline_fixture_retains_busy_refusal_without_touching_application_mutex(offline_compiled):
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
    kernel.CreateMutexW.restype = wintypes.HANDLE
    kernel.ReleaseMutex.argtypes = (wintypes.HANDLE,)
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    name = (offline_compiled.parent / 'fixture-mutex.txt').read_text('utf-8')
    handle = kernel.CreateMutexW(None, True, name)
    assert handle and ctypes.get_last_error() != 183
    try:
        result = _run_owned((str(offline_compiled / 'LiX.exe'), '--check'), timeout_ms=20000)
        assert result.empty_confirmed and result.exit_code == 32
    finally:
        assert kernel.ReleaseMutex(handle)
        assert kernel.CloseHandle(handle)


@pytest.mark.parametrize('damage', [None, 'python', 'component', 'manifest'])
def test_actual_offline_check_never_creates_storage_or_repairs_missing_files(offline_compiled, tmp_path, monkeypatch, damage):
    target = tmp_path / 'LiX 中文 离线'
    shutil.copytree(offline_compiled, target)
    user = tmp_path / 'user'
    user.mkdir()
    (user / 'LiX').write_bytes(b'not a directory; diagnostics cannot be created')
    monkeypatch.setenv('LOCALAPPDATA', str(user))
    if damage == 'python':
        (target / 'runtime/python/python.exe').write_bytes(b'changed')
    elif damage == 'component':
        (target / 'manifests/python.json').unlink()
    elif damage == 'manifest':
        (target / 'LiX.install.json').write_text('{}', 'utf-8')
    before = sorted(str(path.relative_to(target)) for path in target.rglob('*'))
    result = _run_owned((str(target / 'LiX.exe'), '--check'), timeout_ms=20000)
    assert result.empty_confirmed and result.exit_code == (42 if damage is None else 1)
    assert sorted(str(path.relative_to(target)) for path in target.rglob('*')) == before
    assert (user / 'LiX').is_file()


def test_offline_public_build_refuses_unresolved_superdoc_before_creating_payload(tmp_path):
    with pytest.raises(ValueError, match='superdoc_distribution_blocked'):
        offline.stage(ROOT, tmp_path / 'never-created', tmp_path, tmp_path, tmp_path, tmp_path, tmp_path, 'test-build')
    assert not (tmp_path / 'never-created').exists()


def test_offline_environment_inventory_verifies_long_local_paths_and_rejects_changed_content(offline_compiled, tmp_path, monkeypatch):
    target = tmp_path / 'local-long-path'
    shutil.copytree(offline_compiled, target)
    user = tmp_path / 'profile'
    monkeypatch.setenv('LOCALAPPDATA', str(user))
    root = str(target) + '\\'
    manifest_hash = offline.sha(target / 'LiX.install.json')
    environment = user / 'LiX/Environments' / (manifest_hash[:16] + '-' + hashlib.sha256(root.encode('utf-8')).hexdigest()[:16])
    scripts = environment / 'Scripts'
    scripts.mkdir(parents=True)
    (scripts / 'python.exe').write_bytes(b'MZenvironment')
    name = 'Lib/site-packages/' + '/'.join(['long-package-segment-1234567890'] * 5) + '/fixed.bin'
    path = Path('\\\\?\\' + str(environment / name))
    path.parent.mkdir(parents=True)
    path.write_bytes(b'kept')
    assert len(str(environment / name)) > 260
    rows = [{'path': 'Scripts/python.exe', 'bytes': 13, 'sha256': hashlib.sha256(b'MZenvironment').hexdigest()},
            {'path': name, 'bytes': 4, 'sha256': hashlib.sha256(b'kept').hexdigest()}]
    inventory = environment / '.lix-inventory.json'
    inventory.write_text(json.dumps({'schema': 1, 'id': 'environment', 'files': rows}), 'utf-8')
    ready = target / 'app/.runtime/installed-ready.txt'
    ready.write_text(json.dumps({'schema': 2, 'manifest_sha256': manifest_hash, 'lock_sha256': offline.sha(target / 'app/uv.lock'),
        'installation': root, 'environment': str(environment), 'python_sha256': offline.sha(target / 'runtime/python/python.exe'),
        'environment_python_sha256': offline.sha(scripts / 'python.exe'), 'inventory_sha256': offline.sha(inventory), 'preflight_passed': True}), 'utf-8')
    before = _run_owned((str(target / 'LiX.exe'), '--check'), timeout_ms=20000)
    assert before.empty_confirmed and before.exit_code == 0
    path.write_bytes(b'changed')
    after = _run_owned((str(target / 'LiX.exe'), '--check'), timeout_ms=20000)
    assert after.empty_confirmed and after.exit_code == 42
