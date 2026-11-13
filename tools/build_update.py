import argparse
import base64
import io
import json
from pathlib import Path
import re
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from office_agent.update_core import MAX_PACKAGE, channel, check_installed, inventory, local_info, version, write_json


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['keygen', 'configure', 'bump', 'inventory', 'build'])
    parser.add_argument('--private-key', type=Path)
    parser.add_argument('--public-key')
    parser.add_argument('--repository')
    parser.add_argument('--version')
    parser.add_argument('--title')
    parser.add_argument('--date')
    parser.add_argument('--item', action='append')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    if args.action == 'keygen':
        if args.private_key is None or ROOT.resolve() == args.private_key.resolve() or ROOT.resolve() in args.private_key.resolve().parents:
            parser.error('Private key must be explicitly stored outside this project')
        key = Ed25519PrivateKey.generate()
        with args.private_key.open('xb') as stream:
            stream.write(key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
        print(base64.b64encode(key.public_key().public_bytes_raw()).decode())
    elif args.action == 'configure':
        value = {'repository': args.repository or '', 'public_key': args.public_key or ''}
        if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9-]{0,38}/[A-Za-z0-9_.-]{1,100}', value['repository']) or len(base64.b64decode(value['public_key'], validate=True)) != 32:
            parser.error('Valid owner/repo and base64 Ed25519 public key required')
        write_json(ROOT / 'office_agent/update_channel.json', value)
        channel(ROOT)
        print('Configured publisher repository and public key; no network request made')
    elif args.action == 'bump':
        info = local_info(ROOT)
        if not args.version or version(args.version) <= version(info['version']) or not args.title or not args.item or not args.date:
            parser.error('Increasing --version, --date, --title and at least one --item required')
        info['version'] = args.version
        info['history'].insert(0, {'version': args.version, 'date': args.date, 'title': args.title, 'items': args.item})
        project = ROOT / 'pyproject.toml'
        text = project.read_text(encoding='utf8')
        project.write_text(re.sub(r'(?m)^version = "[^"\n]+"$', f'version = "{args.version}"', text, count=1), encoding='utf8')
        write_json(ROOT / 'office_agent/release_info.json', info)
        print('Version and changelog updated; run uv lock, then inventory and build')
    else:
        info = local_info(ROOT)
        if f'version = "{info["version"]}"' not in (ROOT / 'pyproject.toml').read_text(encoding='utf8'):
            parser.error('Project version differs from release metadata')
        if args.action == 'inventory':
            write_json(ROOT / 'office_agent/update_inventory.json', inventory(ROOT))
            print('Created managed-code inventory; do this only for a reviewed release, never to bypass local-change protection')
            return
        current = check_installed(ROOT)
        source = channel(ROOT)
        if args.private_key is None or args.output is None or not source['repository']:
            parser.error('Configured channel, --private-key and --output required')
        if ROOT.resolve() in args.private_key.resolve().parents:
            parser.error('Signing key must be stored outside this project')
        key = serialization.load_pem_private_key(args.private_key.read_bytes(), password=None)
        if not isinstance(key, Ed25519PrivateKey) or base64.b64encode(key.public_key().public_bytes_raw()).decode() != source['public_key']:
            parser.error('Signing key does not match the configured publisher key')
        manifest = {'format': 'office-agent-update-v1', **current}
        metadata = json.dumps(manifest, sort_keys=True, separators=(',', ':')).encode()
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('manifest.json', metadata)
            archive.writestr('manifest.sig', key.sign(metadata))
            for name in current['files']:
                archive.writestr('payload/' + name, (ROOT / name).read_bytes())
        if buffer.tell() > MAX_PACKAGE:
            parser.error('Update exceeds 10 MiB; use a reviewed full distribution')
        target = args.output / f'office-agent-{info["version"]}.zip'
        with target.open('xb') as stream:
            stream.write(buffer.getvalue())
        print(str(target))


if __name__ == '__main__':
    main()
