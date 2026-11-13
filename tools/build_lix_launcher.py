import os
from pathlib import Path
import re
import subprocess
import tempfile


ROOT = Path(__file__).resolve().parents[1]


def build(output=None):
    if os.name != 'nt':
        raise RuntimeError('LiX launcher requires Windows .NET Framework')
    compiler = Path(os.environ['SystemRoot']) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if not compiler.is_file():
        compiler = Path(os.environ['SystemRoot']) / 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
    if not compiler.is_file():
        raise RuntimeError('Windows .NET Framework C# compiler is unavailable')
    destination = Path(output).resolve() if output else ROOT / 'LiX.exe'
    version = re.search(r'^version = "(\d+\.\d+\.\d+)"$', (ROOT / 'pyproject.toml').read_text('utf-8'), re.M)
    if not version:
        raise ValueError('Project version is missing')
    with tempfile.TemporaryDirectory(prefix='lix-build-', dir=destination.parent) as temporary:
        folder = Path(temporary)
        metadata = folder / 'version.cs'
        metadata.write_text(
            'using System.Reflection;\n'
            '[assembly: AssemblyTitle("LiX")]\n'
            '[assembly: AssemblyProduct("LiX")]\n'
            '[assembly: AssemblyDescription("LiX desktop launcher")]\n'
            '[assembly: AssemblyVersion("' + version[1] + '.0")]\n', encoding='utf-8')
        executable = folder / 'LiX.exe'
        result = subprocess.run([
            str(compiler), '/nologo', '/target:winexe', '/platform:anycpu', '/optimize+',
            '/reference:System.Windows.Forms.dll', '/out:' + str(executable),
            '/win32icon:' + str(ROOT / 'office_agent/webui/static/brand/lix.ico'),
            str(ROOT / 'tools/lix_launcher.cs'), str(metadata),
        ], cwd=ROOT, env=dict(os.environ, TEMP=str(folder), TMP=str(folder)), capture_output=True, timeout=60)
        if result.returncode:
            raise RuntimeError('C# compiler failed: ' + result.stdout.decode('mbcs', errors='replace')[:4000])
        executable.replace(destination)
    return destination


if __name__ == '__main__':
    print(build())
