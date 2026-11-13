import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import re
import subprocess
import zipfile

from tools import release_source


ROOT = Path(__file__).absolute().parents[1]


def checked(path, missing=False):
    path = Path(os.path.abspath(path))
    for entry in [*reversed(path.parents), path]:
        try:
            attributes = entry.lstat()
        except FileNotFoundError:
            if missing:
                continue
            raise
        if entry.is_symlink() or getattr(attributes, 'st_file_attributes', 0) & 0x400:
            raise ValueError('Installer path contains a link or reparse point')
    return path


def digest(data):
    return hashlib.sha256(data).hexdigest()


def runtime_files(archive, metadata, license_root):
    archive = checked(archive)
    expected = metadata['uv']
    if archive.stat().st_size != expected['bytes']:
        raise ValueError('UV archive size differs from the pinned release')
    data = archive.read_bytes()
    if digest(data) != expected['sha256']:
        raise ValueError('UV archive hash differs from the pinned release')
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        items = package.infolist()
        if len(items) != 3 or {item.filename for item in items} != {'uv.exe', 'uvw.exe', 'uvx.exe'}:
            raise ValueError('Unexpected UV release archive entries')
        item = package.getinfo('uv.exe')
        if item.file_size > 64 * 1024 * 1024 or item.external_attr >> 16 & 0o170000 == 0o120000:
            raise ValueError('UV executable is oversized or linked')
        executable = package.read(item)
        if not executable.startswith(b'MZ'):
            raise ValueError('UV release is not a Windows executable')
    files = {'runtime/uv.exe': executable}
    for row in expected['licenses']:
        path = checked(Path(license_root) / row['file'])
        if path.stat().st_size > 128 * 1024:
            raise ValueError('UV license exceeds the bound')
        value = path.read_bytes()
        if digest(value) != row['sha256']:
            raise ValueError('UV license differs from the pinned upstream text')
        files['licenses/' + row['file']] = value
    return files


def stage(root, target, uv_archive, license_root):
    root, target = checked(root), checked(target, missing=True)
    if target.exists():
        raise ValueError('Installer staging requires a new directory')
    policy = release_source.read_json(root / 'release/PUBLIC_FILES.json')
    source = release_source.collect(root, policy)
    metadata = release_source.read_json(root / 'release/WINDOWS_INSTALLER.json')
    version = re.search(r'^version = "(\d+\.\d+\.\d+)"\r?$', (root / 'pyproject.toml').read_text('utf-8'), re.M)
    if not version or metadata.get('schema') != 1 or metadata.get('channel') != 'network-preview':
        raise ValueError('Installer metadata is incompatible')
    files = runtime_files(uv_archive, metadata, license_root)
    for row in source['files']:
        data = release_source.read_file(root, row['path'])
        if digest(data) != row['sha256']:
            raise ValueError('Public source changed while staging the installer')
        files['app/' + row['path']] = data
    manifest = {'schema': 1, 'version': version[1], 'channel': 'network-preview',
                'python': metadata['python'], 'uv_version': metadata['uv']['version'],
                'files': [{'path': name, 'bytes': len(data), 'sha256': digest(data)}
                          for name, data in sorted(files.items())]}
    manifest_data = (json.dumps(manifest, ensure_ascii=False, indent=2) + '\n').encode('utf-8')
    files['LiX.install.json'] = manifest_data
    target.mkdir()
    for name, data in files.items():
        path = target / name
        path.parent.mkdir(exist_ok=True, parents=True)
        with path.open('xb') as stream:
            stream.write(data)
    (target / 'INSTALL-NOTES.txt').write_text(
        'LiX 联网预览安装版\n\n'
        '仅安装到当前用户，无需管理员权限。首次启动须联网准备固定 Python、锁定依赖、文档编辑资源和渲染器，可能需要数分钟及约 3 GB 可用空间。\n\n'
        '不捆绑开发者的密钥、配置、历史、文档或虚拟环境。请在软件设置中配置自己的模型；安装与首次准备不调用模型。\n\n'
        '本候选未签名，尚未完成全新 Windows 机器与正式发行验证。安装不代表功能全量验收。\n\n'
        '卸载移除安装清单中的程序与快捷方式；保留本机配置、历史和下载缓存。用户文档及修改副本不属于卸载范围。\n\n'
        '项目自有代码：AGPL-3.0-only；随包附公开源码及原有第三方通知。文档编辑资源在用户主动准备时按固定版本获取。\n',
        encoding='utf-8-sig')
    configuration = {'manifest_sha256': digest(manifest_data), 'version': version[1], 'python': metadata['python']}
    (target / 'bootstrap.json').write_text(json.dumps(configuration), encoding='utf-8')
    return manifest, source


def compile_launcher(root, payload):
    if os.name != 'nt':
        raise RuntimeError('Windows .NET compiler is required')
    compiler = Path(os.environ['SystemRoot']) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    checked(compiler)
    version = json.loads((payload / 'bootstrap.json').read_text('utf-8'))['version']
    metadata = payload / 'launcher_version.cs'
    metadata.write_text('using System.Reflection;[assembly: AssemblyTitle("LiX")]'
                        '[assembly: AssemblyProduct("LiX")][assembly: AssemblyVersion("' + version + '.0")]', encoding='utf-8')
    args = [str(compiler), '/nologo', '/target:winexe', '/platform:x64', '/optimize+',
            '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll',
            '/reference:System.Web.Extensions.dll', '/out:' + str(payload / 'LiX.exe'),
            '/win32icon:' + str(root / 'office_agent/webui/static/brand/lix.ico'),
            '/resource:' + str(payload / 'bootstrap.json') + ',LiX.Bootstrap',
            str(root / 'tools/windows/installed_launcher.cs'), str(root / 'tools/windows/process_job.cs'), str(metadata)]
    result = subprocess.run(args, capture_output=True, timeout=60)
    if result.returncode:
        raise RuntimeError('Launcher compilation failed: ' + result.stdout.decode('mbcs', 'replace')[:5000])


def prepare_compiler_inputs(root, payload, manifest):
    version = manifest['version']
    (payload / 'launcher_version.cs').write_text(
        'using System.Reflection;[assembly: AssemblyTitle("LiX")]'
        '[assembly: AssemblyProduct("LiX")][assembly: AssemblyVersion("' + version + '.0")]', encoding='utf-8')
    names = [item['path'] for item in manifest['files']] + ['LiX.install.json', 'LiX.exe', 'INSTALL-NOTES.txt']
    rows = []
    for name in names:
        relative = Path(name)
        directory = str(relative.parent).replace('.', '', 1) if relative.parent == Path('.') else str(relative.parent)
        destination = '{app}' + ('\\' + directory if directory else '')
        flags = 'ignoreversion uninsneveruninstall' if name == 'LiX.install.json' else 'ignoreversion'
        rows.append('Source: "' + str(payload / relative) + '"; DestDir: "' + destination + '"; Flags: ' + flags)
    (payload / 'FILES.iss').write_text('\n'.join(rows) + '\n', encoding='utf-8-sig')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', type=Path, default=ROOT)
    parser.add_argument('--payload', type=Path, required=True)
    parser.add_argument('--uv-archive', type=Path, required=True)
    parser.add_argument('--license-root', type=Path, required=True)
    parser.add_argument('--stage-only', action='store_true')
    args = parser.parse_args()
    if not args.stage_only:
        parser.error('Build installers through tools/build_windows_installer.ps1; this entry only stages source')
    root, payload = checked(args.root), checked(args.payload, missing=True)
    manifest, source = stage(root, payload, args.uv_archive, args.license_root)
    prepare_compiler_inputs(root, payload, manifest)
    receipt = dict(version=manifest['version'], public_source_files=len(source['files']),
                   payload_files=len(manifest['files']),
                   source_manifest_sha256=digest(json.dumps(source, sort_keys=True).encode()))
    (payload / 'STAGE.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(receipt))


if __name__ == '__main__':
    main()
