from pathlib import Path
import shutil
import subprocess

import pytest


ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize('state,message', [('installed', 'Runtime already exists'),
                                          ('invalid_archive', 'Archive checksum mismatch'),
                                          ('partial_directory', 'Partial download must be a regular file')])
def test_installer_refuses_unsafe_existing_state_without_network(tmp_path, state, message):
    shell = shutil.which('pwsh')
    if shell is None:
        pytest.skip('PowerShell 7 required for offline installer regression')
    downloads = tmp_path / '.runtime/downloads'
    downloads.mkdir(parents=True)
    archive = downloads / 'LibreOffice_26.2.6.2_Win_x86-64.msi'
    if state == 'installed':
        (tmp_path / '.runtime/libreoffice').mkdir()
    elif state == 'invalid_archive':
        archive.write_bytes(b'synthetic invalid MSI; no executable content')
    else:
        Path(str(archive) + '.partial').mkdir()
    result = subprocess.run([shell, '-NoProfile', '-File', str(ROOT / 'tools/Install-DocumentRuntime.ps1'),
                             '-ProjectRoot', str(tmp_path)], capture_output=True, timeout=15, check=False)
    output = (result.stdout + result.stderr).decode('utf-8', errors='replace')
    assert result.returncode != 0 and message in output
    assert not (tmp_path / '.runtime/libreoffice-manifest.json').exists()
    if state == 'invalid_archive':
        assert archive.read_bytes() == b'synthetic invalid MSI; no executable content'
