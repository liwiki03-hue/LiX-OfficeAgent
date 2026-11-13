import hashlib
import io
import json
from pathlib import Path
import sys
import time
import zipfile

import pytest

from tools import build_windows_installer as installer
from tools import installed_runtime


ROOT = Path(__file__).absolute().parents[1]


def fixture_source(tmp_path):
    root = tmp_path / 'public-source'
    (root / 'release').mkdir(parents=True)
    (root / 'README.md').write_text('Public application source', encoding='utf-8')
    (root / 'pyproject.toml').write_text('[project]\nversion = "0.7.11"\n', encoding='utf-8')
    (root / '.env').write_text('PRIVATE_VALUE=do-not-copy', encoding='utf-8')
    archive = tmp_path / 'uv.zip'
    with zipfile.ZipFile(archive, 'w') as out:
        out.writestr('uv.exe', b'MZsynthetic-no-execution')
        out.writestr('uvw.exe', b'synthetic')
        out.writestr('uvx.exe', b'synthetic')
    spec = json.loads((ROOT / 'release/WINDOWS_INSTALLER.json').read_text('utf-8'))
    spec['uv'].update(bytes=archive.stat().st_size, sha256=hashlib.sha256(archive.read_bytes()).hexdigest())
    spec['uv']['licenses'] = []
    (root / 'release/WINDOWS_INSTALLER.json').write_text(json.dumps(spec), encoding='utf-8')
    policy = {'schema': 1, 'files': ['README.md', 'pyproject.toml', 'release/WINDOWS_INSTALLER.json'], 'scan_exceptions': []}
    (root / 'release/PUBLIC_FILES.json').write_text(json.dumps(policy), encoding='utf-8')
    return root, archive, spec


def test_installer_stages_only_reviewed_source_and_keeps_private_configuration(tmp_path):
    root, archive, _ = fixture_source(tmp_path)
    payload = tmp_path / 'payload'
    manifest, source = installer.stage(root, payload, archive, tmp_path)
    assert len(source['files']) == 3
    assert not (payload / 'app/.env').exists()
    assert (root / '.env').read_text() == 'PRIVATE_VALUE=do-not-copy'
    assert (payload / 'app/README.md').read_bytes() == (root / 'README.md').read_bytes()
    assert manifest['channel'] == 'network-preview'
    with pytest.raises(ValueError, match='new directory'):
        installer.stage(root, payload, archive, tmp_path)


def test_installer_rejects_tampered_runtime_before_creating_payload(tmp_path):
    root, archive, _ = fixture_source(tmp_path)
    value = bytearray(archive.read_bytes())
    value[10] ^= 1
    archive.write_bytes(value)
    payload = tmp_path / 'payload'
    with pytest.raises(ValueError, match='hash'):
        installer.stage(root, payload, archive, tmp_path)
    assert not payload.exists()


def test_build_inputs_retain_install_identity_for_reinstall_without_exporting_private_config(tmp_path):
    root, archive, _ = fixture_source(tmp_path)
    payload = tmp_path / 'payload'
    manifest, _ = installer.stage(root, payload, archive, tmp_path)
    installer.prepare_compiler_inputs(root, payload, manifest)
    lines = (payload / 'FILES.iss').read_text('utf-8-sig').splitlines()
    retained = [line for line in lines if 'uninsneveruninstall' in line]
    assert len(retained) == 1
    assert 'LiX.install.json' in retained[0]
    assert all('.env"' not in line for line in lines)
    assert not (payload / 'LiX.exe').exists()


def test_installer_preserves_upstream_license_and_rejects_tampering(tmp_path):
    _, archive, spec = fixture_source(tmp_path)
    license_file = tmp_path / 'LICENSE-UV'
    license_file.write_bytes(b'copyright and upstream license\n')
    spec['uv']['licenses'] = [{'file': license_file.name, 'sha256': hashlib.sha256(license_file.read_bytes()).hexdigest()}]
    assert installer.runtime_files(archive, spec, tmp_path)['licenses/LICENSE-UV'] == license_file.read_bytes()
    license_file.write_bytes(b'changed')
    with pytest.raises(ValueError, match='license differs'):
        installer.runtime_files(archive, spec, tmp_path)


def test_installer_rejects_redirected_staging_parent(tmp_path, monkeypatch):
    root, archive, _ = fixture_source(tmp_path)
    parent = tmp_path / 'redirected'
    parent.mkdir()
    original = Path.is_symlink
    monkeypatch.setattr(Path, 'is_symlink', lambda value: value == parent or original(value))
    with pytest.raises(ValueError, match='link or reparse'):
        installer.stage(root, parent / 'payload', archive, tmp_path)


def test_renderer_hash_failure_does_not_extract_or_replace(tmp_path, monkeypatch):
    class Response(io.BytesIO):
        status = 200

        def geturl(self):
            return 'https://official.example/archive'

    monkeypatch.setattr(installed_runtime, 'urlopen', lambda *args, **kwargs: Response(b'changed'))
    spec = {'version': 'fixture', 'bytes': 7, 'sha256': '0' * 64, 'url': 'https://official.example/archive'}
    with pytest.raises(ValueError, match='checksum'):
        installed_runtime.download(tmp_path, spec)
    assert not list(tmp_path.glob('*.msi'))
    assert len(list(tmp_path.glob('*.partial'))) == 1


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows installed launcher')
def test_actual_compiled_launcher_checks_integrity_without_network_or_launch(tmp_path, monkeypatch):
    from office_agent.diagram_process import _WindowsJob
    root, archive, _ = fixture_source(tmp_path)
    payload = tmp_path / 'LiX 中文 & (install)'
    installer.stage(root, payload, archive, tmp_path)
    installer.compile_launcher(ROOT, payload)
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path / 'isolated-user'))

    def run():
        process, job = None, _WindowsJob()
        try:
            process = job.spawn((str(payload / 'LiX.exe'), '--check'))
            end = time.monotonic() + 15
            while process.poll() is None and time.monotonic() < end:
                time.sleep(.05)
            return process.poll()
        finally:
            empty = job.close()
            if process:
                process.close()
            assert empty

    assert run() == 42
    assert not (payload / 'app/.runtime').exists()
    assert not (tmp_path / 'isolated-user/LiX/Environments').exists()
    (payload / 'app/.env').write_text('LOCAL_ONLY=preserved', encoding='utf-8')
    assert run() == 42
    (payload / 'app/README.md').write_text('changed installation', encoding='utf-8')
    assert run() == 1
    assert (payload / 'app/.env').read_text() == 'LOCAL_ONLY=preserved'
