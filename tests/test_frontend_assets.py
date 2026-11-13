import importlib.util
import io
import json
from pathlib import Path
import tarfile

import pytest


MODULE = importlib.util.spec_from_file_location('frontend_assets', Path(__file__).resolve().parents[1] / 'tools/install_frontend_assets.py')
assets = importlib.util.module_from_spec(MODULE)
MODULE.loader.exec_module(assets)


def fixture(root):
    target = root / 'release/SUPERDOC_SOURCE.json'
    target.parent.mkdir()
    entries = {'package/dist/superdoc.min.js': ('superdoc.min.js', b'synthetic javascript'),
               'package/dist/style.css': ('style.css', b'synthetic stylesheet'),
               'package/LICENSE': ('LICENSE', b'synthetic license fixture')}
    stream = io.BytesIO()
    with tarfile.open(fileobj=stream, mode='w:gz') as archive:
        for name, (_, data) in entries.items():
            entry = tarfile.TarInfo(name)
            entry.size = len(data)
            archive.addfile(entry, io.BytesIO(data))
    data = stream.getvalue()
    row = {'package': 'superdoc', 'version': '1.47.2', 'archive_url': assets.URL,
           'archive_sha256': assets.sha(data), 'assets': [{'path': name, 'local': local, 'sha256': assets.sha(value)}
                                                        for name, (local, value) in entries.items()]}
    target.write_text(json.dumps(row), encoding='utf-8')
    return data, row


def test_pinned_frontend_install_is_exact_idempotent_and_never_overwrites(tmp_path, monkeypatch):
    data, row = fixture(tmp_path)
    calls = []
    monkeypatch.setattr(assets, 'download', lambda root, expected: (calls.append(expected) or data, True))
    first = assets.prepare(tmp_path)
    assert first['status'] == 'verified' and len(first['created']) == 3 and len(calls) == 1
    assert assets.prepare(tmp_path)['status'] == 'already_verified' and len(calls) == 1
    target = tmp_path / 'office_agent/webui/static/vendor/superdoc/style.css'
    target.write_bytes(b'local changes')
    with pytest.raises(ValueError, match='refusing to overwrite'):
        assets.prepare(tmp_path)
    assert target.read_bytes() == b'local changes' and len(calls) == 1


def test_frontend_bad_hash_and_redirected_target_are_rejected_before_writes(tmp_path, monkeypatch):
    data, row = fixture(tmp_path)
    row['assets'][0]['sha256'] = '0' * 64
    (tmp_path / 'release/SUPERDOC_SOURCE.json').write_text(json.dumps(row), encoding='utf-8')
    monkeypatch.setattr(assets, 'download', lambda *a: (data, False))
    with pytest.raises(ValueError, match='asset hash'):
        assets.prepare(tmp_path)
    assert not (tmp_path / 'office_agent').exists()
    original = Path.is_symlink
    monkeypatch.setattr(Path, 'is_symlink', lambda path: path == tmp_path / 'office_agent' or original(path))
    (tmp_path / 'office_agent').mkdir()
    with pytest.raises(ValueError, match='redirected'):
        assets.prepare(tmp_path)
