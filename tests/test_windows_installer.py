import hashlib
import io
import json
from pathlib import Path
import shutil
import sys
import uuid
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


class ArchiveResponse(io.BytesIO):
    def __init__(self, data, status=206, headers=None):
        super().__init__(data)
        self.status = status
        self.headers = headers or {}

    def geturl(self):
        return 'https://official.example/archive'


def archive_spec(data):
    return {'version': 'fixture', 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest(),
            'url': 'https://official.example/archive'}


def test_renderer_chunk_download_resumes_legacy_partial_without_repeating_bytes(tmp_path, monkeypatch):
    data = b'0123456789abcdef'
    spec = archive_spec(data)
    previous = tmp_path / ('LibreOffice_fixture_Win_x86-64.msi.' + 'a' * 32 + '.partial')
    previous.write_bytes(data[:5])
    ranges = []

    def receive(request, timeout):
        ranges.append(request.get_header('Range'))
        start, last = map(int, ranges[-1].split('=')[1].split('-'))
        return ArchiveResponse(data[start:last + 1], headers={
            'Content-Range': f'bytes {start}-{last}/{len(data)}', 'Content-Length': str(last - start + 1)})

    monkeypatch.setattr(installed_runtime, 'urlopen', receive)
    monkeypatch.setattr(installed_runtime, 'CHUNK_BYTES', 4)
    result = installed_runtime.download(tmp_path, spec)
    assert result.read_bytes() == data
    assert ranges == ['bytes=5-8', 'bytes=9-12', 'bytes=13-15']
    assert not previous.exists()
    assert not list(tmp_path.glob('*.partial'))
    assert installed_runtime.download(tmp_path, spec) == result
    assert len(ranges) == 3


@pytest.mark.parametrize('status,headers,expected', [
    (200, {}, 'download_range'),
    (206, {'Content-Range': 'bytes 0-5/10'}, 'download_range'),
    (206, {'Content-Range': 'bytes 4-9/10', 'Content-Length': '99'}, 'download_size'),
])
def test_renderer_bad_resume_response_never_appends(tmp_path, monkeypatch, status, headers, expected):
    partial = tmp_path / 'LibreOffice_fixture_Win_x86-64.msi.partial'
    partial.write_bytes(b'0123')
    monkeypatch.setattr(installed_runtime, 'urlopen', lambda *a, **k: ArchiveResponse(b'unsafe', status, headers))
    with pytest.raises(installed_runtime.PreparationError, match=expected):
        installed_runtime.download(tmp_path, archive_spec(b'0123456789'))
    assert partial.read_bytes() == b'0123'
    assert not list(tmp_path.glob('*.msi'))


def test_renderer_timeout_resumes_partial_in_a_bounded_retry(tmp_path, monkeypatch):
    calls = []

    class Interrupted(ArchiveResponse):
        def read1(self, size):
            if self.tell():
                raise TimeoutError('private network detail')
            return super().read1(4)

    def receive(request, timeout):
        calls.append(request.get_header('Range'))
        if len(calls) == 1:
            return Interrupted(b'0123456789', 200)
        return ArchiveResponse(b'456789', headers={'Content-Range': 'bytes 4-9/10'})

    monkeypatch.setattr(installed_runtime, 'urlopen', receive)
    assert installed_runtime.download(tmp_path, archive_spec(b'0123456789')).read_bytes() == b'0123456789'
    assert calls == ['bytes=0-9', 'bytes=4-9']


def test_renderer_error_receipt_is_safe_and_does_not_report_ready(tmp_path, monkeypatch, capsys):
    report = tmp_path / 'renderer.json'
    monkeypatch.setenv('LIX_RENDERER_REPORT', str(report))

    def fail(root):
        raise TimeoutError('secret-token-and-local-path')

    monkeypatch.setattr(installed_runtime, 'prepare', fail)
    assert installed_runtime.main(tmp_path) == 1
    saved = json.loads(report.read_text('utf-8'))
    assert saved['state'] == 'failed'
    assert saved['code'] == 'download_timeout'
    assert 'secret-token' not in report.read_text('utf-8') + capsys.readouterr().out


def test_renderer_global_budget_preserves_data_without_more_requests(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr(installed_runtime, 'DOWNLOAD_SECONDS', 0)
    monkeypatch.setattr(installed_runtime, 'urlopen', lambda *a, **k: calls.append(a))
    with pytest.raises(installed_runtime.PreparationError, match='download_timeout'):
        installed_runtime.download(tmp_path, archive_spec(b'0123456789'))
    assert calls == []
    assert not list(tmp_path.glob('*.msi'))


def test_bundle_stages_checked_archive_and_declares_network_still_needed(tmp_path):
    root, archive, spec = fixture_source(tmp_path)
    renderer = tmp_path / 'renderer.msi'
    renderer.write_bytes(b'synthetic-offline-renderer')
    spec['renderer'].update(bytes=renderer.stat().st_size, sha256=hashlib.sha256(renderer.read_bytes()).hexdigest())
    (root / 'release/WINDOWS_INSTALLER.json').write_text(json.dumps(spec), encoding='utf-8')
    manifest, _ = installer.stage(root, tmp_path / 'payload', archive, tmp_path, renderer)
    row = next(item for item in manifest['files'] if item['path'].endswith('.msi'))
    assert (tmp_path / 'payload' / row['path']).read_bytes() == renderer.read_bytes()
    assert row['sha256'] == spec['renderer']['sha256']
    assert manifest['bundles_renderer'] is True
    assert '不是全离线安装包' in (tmp_path / 'payload/INSTALL-NOTES.txt').read_text('utf-8-sig')
    renderer.write_bytes(b'tampered')
    with pytest.raises(ValueError, match='pinned official archive'):
        installer.stage(root, tmp_path / 'rejected', archive, tmp_path, renderer)
    assert not (tmp_path / 'rejected').exists()


def test_bundled_renderer_is_verified_before_extract_without_network(tmp_path, monkeypatch):
    root = tmp_path / 'app'
    (root / 'release').mkdir(parents=True)
    spec = json.loads((ROOT / 'release/WINDOWS_INSTALLER.json').read_text('utf-8'))
    (root / 'release/WINDOWS_INSTALLER.json').write_text(json.dumps(spec), encoding='utf-8')
    bundle = tmp_path / 'runtime/LibreOffice_26.2.6.2_Win_x86-64.msi'
    bundle.parent.mkdir()
    with bundle.open('wb') as stream:
        stream.truncate(spec['renderer']['bytes'])
    original_digest = installed_runtime.digest
    monkeypatch.setattr(installed_runtime, 'digest', lambda path: spec['renderer']['sha256'] if path == bundle else original_digest(path))
    monkeypatch.setattr(installed_runtime, 'urlopen', lambda *a, **k: pytest.fail('unexpected network request'))
    extracted = []

    def extract(archive, destination):
        extracted.append(archive)
        (destination / 'program').mkdir(parents=True)
        (destination / 'program/soffice.com').write_bytes(b'synthetic-binary')

    monkeypatch.setattr(installed_runtime, 'extract', extract)
    assert installed_runtime.prepare(root)['ok'] is True
    assert extracted == [bundle]
    assert installed_runtime.prepare(root)['status'] == 'already_verified'
    assert extracted == [bundle]


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows installed launcher')
def test_actual_compiled_launcher_checks_integrity_without_network_or_launch(tmp_path, monkeypatch):
    from office_agent.diagram_process import _run_owned
    root, archive, _ = fixture_source(tmp_path)
    payload = tmp_path / 'LiX 中文 & (install)'
    installer.stage(root, payload, archive, tmp_path)
    compiler_root = tmp_path / 'compiler-source'
    for name in ('tools/windows/installed_launcher.cs', 'tools/windows/process_job.cs',
                 'tools/windows/bootstrap_network.cs', 'tools/windows/offline_manifest.cs',
                 'office_agent/webui/static/brand/lix.ico'):
        target = compiler_root / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / name, target)
    launcher = compiler_root / 'tools/windows/installed_launcher.cs'
    source = launcher.read_text('utf-8')
    assert source.count('"LiXInstalledRuntime"') == 1
    launcher.write_text(source.replace('"LiXInstalledRuntime"', '"LiXNetworkFixture-' + uuid.uuid4().hex + '"'), 'utf-8')
    compiled = _run_owned((sys.executable, '-I', '-c', 'import sys,faulthandler;from pathlib import Path;trace=open(sys.argv[3],"w");sys.stdout=sys.stderr=trace;faulthandler.dump_traceback_later(8,repeat=True,file=trace);sys.path.insert(0,sys.argv[1]);from tools.build_windows_installer import compile_launcher;compile_launcher(Path(sys.argv[4]),Path(sys.argv[2]));faulthandler.cancel_dump_traceback_later()', str(ROOT), str(payload), str(tmp_path / 'compiler-stack.log'), str(compiler_root)), timeout_ms=30000)
    assert compiled.empty_confirmed
    assert compiled.exit_code == 0, 'Launcher compilation did not complete within its isolated 30-second boundary'
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path / 'isolated-user'))

    def run():
        result = _run_owned((sys.executable, '-I', '-c', 'import subprocess,sys;sys.exit(subprocess.run([sys.argv[1],"--check"],timeout=15).returncode)', str(payload / 'LiX.exe')), timeout_ms=20000)
        assert result.empty_confirmed
        assert result.code != 'render_timeout', 'Compiled launcher check exceeded independent 20-second boundary'
        return result.exit_code

    assert run() == 42
    assert not (payload / 'app/.runtime').exists()
    assert not (tmp_path / 'isolated-user/LiX/Environments').exists()
    (payload / 'app/.env').write_text('LOCAL_ONLY=preserved', encoding='utf-8')
    assert run() == 42
    (payload / 'app/README.md').write_text('changed installation', encoding='utf-8')
    assert run() == 1
    assert (payload / 'app/.env').read_text() == 'LOCAL_ONLY=preserved'
