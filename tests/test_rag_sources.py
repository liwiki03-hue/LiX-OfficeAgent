"""Tiny synthetic OOXML only; no Office/provider/environment inputs."""
import hashlib
import io
import os
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

from office_agent import rag_sources as sources
from office_agent.rag import structure_chunks
from office_agent.rag_sources import SourceError, SourceExtractor
from office_agent.writes import Workspace

W = sources.W[1:-1]
S = sources.S[1:-1]
R = sources.R[1:-1]
P = sources.P[1:-1]
C = sources.C[1:-1]


def package(tmp_path, kind='docx', body=None, extras=None, compression=zipfile.ZIP_STORED):
    if kind == 'docx':
        core = 'word/document.xml'
        content = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
        xml = body or f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body></w:document>'
        parts = {core: xml}
    else:
        core = 'xl/workbook.xml'
        content = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
        parts = {core: f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>',
                 'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{P}"><Relationship Id="r1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
                 'xl/worksheets/sheet1.xml': body or f'<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row></sheetData></worksheet>'}
    parts['[Content_Types].xml'] = f'<Types xmlns="{C}"><Override PartName="/{core}" ContentType="{content}"/></Types>'
    parts.update(extras or {})
    path = tmp_path / ('source.' + kind)
    with zipfile.ZipFile(path, 'w', compression=compression) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return path


def extractor(tmp_path):
    return SourceExtractor(Workspace(tmp_path))


def fails(ex, name, code):
    with pytest.raises(SourceError) as caught:
        ex.extract(name)
    assert str(caught.value) == caught.value.code == code


def test_word_order_positions_and_no_writes(tmp_path, monkeypatch):
    body = f'''<w:document xmlns:w="{W}"><w:body>
    <w:p><w:r><w:t>before</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>left</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>right</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>'''
    path = package(tmp_path, body=body)
    before = {p.name: p.read_bytes() for p in tmp_path.iterdir()}
    ex = extractor(tmp_path)
    original_open = Path.open
    def read_only(path, mode='r', *args, **kwargs):
        assert mode == 'rb'
        return original_open(path, mode, *args, **kwargs)
    with monkeypatch.context() as patch:
        patch.setattr(Path, 'open', read_only)
        result = ex.extract(path.name)
        assert ex.fingerprint(result['doc_id']) == result['fingerprint']
    assert [b['type'] for b in result['blocks']] == ['paragraph', 'heading', 'table_row', 'paragraph']
    assert [b['text'] for b in result['blocks']] == ['before', 'Title', 'left | right', 'after']
    assert result['blocks'][2]['location'] == 'body:3/table/row:1'
    assert all(set(b) == {'source', 'location', 'type', 'text', 'group'} for b in result['blocks'])
    assert structure_chunks(result['blocks'])
    assert result['fingerprint'] == hashlib.sha256(path.read_bytes()).hexdigest()
    assert result['doc_id'] == hashlib.sha256(path.name.casefold().encode()).hexdigest()
    assert before == {p.name: p.read_bytes() for p in tmp_path.iterdir()}


def test_excel_actual_sparse_rows_formula_no_eval_shared_strings(tmp_path):
    xml = f'''<worksheet xmlns="{S}"><dimension ref="A1:XFD1048576"/><sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>SUM(1,2)</f><v>999</v></c></row>
    <row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>tail</t></is></c></row></sheetData></worksheet>'''
    path = package(tmp_path, 'xlsx', xml, {'xl/sharedStrings.xml': f'<sst xmlns="{S}"><si><t>shared</t></si></sst>'})
    blocks = extractor(tmp_path).extract(path.name)['blocks']
    assert [b['text'] for b in blocks] == ['A1: shared | B1: =SUM(1,2)', 'XFD1048576: tail']
    assert [b['location'] for b in blocks] == ['sheet:1:Data/row:1', 'sheet:1:Data/row:1048576']


def test_text_and_unknown_identifiers_stale_deleted_replaced(tmp_path):
    path = tmp_path / 'notes.txt'
    path.write_bytes('one\n\ntwo\n'.encode())
    ex = extractor(tmp_path)
    result = ex.extract('notes.txt')
    assert [b['location'] for b in result['blocks']] == ['line:1', 'line:3']
    with pytest.raises(SourceError, match='^invalid$'):
        ex.fingerprint('../notes.txt')
    path.write_text('updated', encoding='utf-8')
    assert ex.fingerprint(result['doc_id']) != result['fingerprint']
    path.unlink()
    with pytest.raises(SourceError, match='^source_changed$'):
        ex.fingerprint(result['doc_id'])
    path.write_text('one\n\ntwo\n', encoding='utf-8')
    with pytest.raises(SourceError, match='^source_changed$'):
        ex.fingerprint(result['doc_id'])


@pytest.mark.parametrize('name', ['../outside.txt', '/outside.txt', 'C:outside.txt', 'a\\b.txt', 'a/../b.txt', 'a//b.txt', 'a./b.txt', 'x' * 129])
def test_invalid_names(tmp_path, name):
    fails(extractor(tmp_path), name, 'invalid')


@pytest.mark.parametrize('text,code', [(b'\xff', 'invalid'), (b'', 'invalid'), (b'password=fixture-secret', 'sensitive'), (b'\x00', 'invalid'), (b'x' * 8001, 'capacity')])
def test_text_rejections(tmp_path, text, code):
    (tmp_path / 'notes.txt').write_bytes(text)
    fails(extractor(tmp_path), 'notes.txt', code)


def test_sensitive_source_unsupported_and_trusted_workspace(tmp_path):
    fails(extractor(tmp_path), '.env.txt', 'sensitive')
    fails(extractor(tmp_path), 'script.py', 'unsupported')
    with pytest.raises(SourceError, match='^invalid$'):
        SourceExtractor(SimpleNamespace(root=tmp_path))


def test_hardlink_and_symlink_simulation(tmp_path, monkeypatch):
    path = tmp_path / 'notes.txt'
    path.write_text('hello', encoding='utf-8')
    os.link(path, tmp_path / 'hard.txt')
    fails(extractor(tmp_path), 'notes.txt', 'invalid')
    (tmp_path / 'hard.txt').unlink()
    original = Path.lstat
    def reparse(path, *args, **kwargs):
        info = original(path, *args, **kwargs)
        if path.name == 'notes.txt':
            return SimpleNamespace(st_mode=info.st_mode, st_nlink=1, st_file_attributes=0x400)
        return info
    ex = extractor(tmp_path)
    monkeypatch.setattr(Path, 'lstat', reparse)
    fails(ex, 'notes.txt', 'invalid')


def test_ancestor_reparse_and_workspace_guard(tmp_path, monkeypatch):
    sub = tmp_path / 'sub'
    sub.mkdir()
    (sub / 'notes.txt').write_text('hello', encoding='utf-8')
    ex = extractor(tmp_path)
    original = Path.lstat
    def reparse(path, *args, **kwargs):
        info = original(path, *args, **kwargs)
        if path == sub:
            return SimpleNamespace(st_mode=info.st_mode, st_nlink=info.st_nlink, st_file_attributes=0x400)
        return info
    with monkeypatch.context() as patch:
        patch.setattr(Path, 'lstat', reparse)
        fails(ex, 'sub/notes.txt', 'invalid')
    monkeypatch.setattr(ex._workspace, 'file', lambda *a, **kw: (_ for _ in ()).throw(ValueError('private path')))
    fails(ex, 'sub/notes.txt', 'invalid')


def test_changed_during_parse_and_root_mutation(tmp_path, monkeypatch):
    path = tmp_path / 'notes.txt'
    path.write_text('first', encoding='utf-8')
    ex = extractor(tmp_path)
    parse = ex._parse
    def changed(*args):
        result = parse(*args)
        path.write_text('second', encoding='utf-8')
        return result
    monkeypatch.setattr(ex, '_parse', changed)
    fails(ex, 'notes.txt', 'source_changed')
    assert not ex._bindings
    ex._workspace.root = tmp_path.parent
    fails(ex, 'notes.txt', 'source_changed')


@pytest.mark.parametrize('extras,code', [
    ({'../evil.xml': '<a/>'}, 'invalid'),
    ({'word/vbaProject.bin': b'macro'}, 'unsupported'),
    ({'junk.xml': '<broken'}, 'invalid'),
    ({'junk.xml': '<!DOCTYPE a [<!ENTITY secret "secret">]><a>&secret;</a>'}, 'unsupported'),
    ({'junk.xml': '<!DOCTYPE a SYSTEM "file:///private"><a/>'}, 'unsupported'),
    ({'word/_rels/document.xml.rels': f'<Relationships xmlns="{P}"><Relationship Id="r1" Type="link" Target="https://example.test" TargetMode="External"/></Relationships>'}, 'unsupported'),
    ({'xl/externalLinks/externalLink1.xml': '<a/>'}, 'unsupported'),
])
def test_hostile_ooxml(tmp_path, extras, code):
    path = package(tmp_path, extras=extras)
    fails(extractor(tmp_path), path.name, code)


def test_zipbomb_and_caps(tmp_path, monkeypatch):
    path = package(tmp_path, extras={'bomb.bin': b'0' * 100_000}, compression=zipfile.ZIP_DEFLATED)
    fails(extractor(tmp_path), path.name, 'capacity')
    path = package(tmp_path)
    monkeypatch.setattr(sources, 'MAX_ZIP_ENTRIES', 1)
    fails(extractor(tmp_path), path.name, 'capacity')


@pytest.mark.parametrize('formula', ['WEBSERVICE("https://example.test")', '[other.xlsx]Sheet1!A1', 'HYPERLINK("file:///private")'])
def test_external_formulas_rejected(tmp_path, formula):
    path = package(tmp_path, 'xlsx', f'<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f>{formula}</f></c></row></sheetData></worksheet>')
    fails(extractor(tmp_path), path.name, 'unsupported')


def test_capacity_blocks_total_bytes_and_id_bindings(tmp_path, monkeypatch):
    path = tmp_path / 'notes.txt'
    path.write_text('a\nb\nc', encoding='utf-8')
    ex = extractor(tmp_path)
    with monkeypatch.context() as patch:
        patch.setattr(sources, 'MAX_BLOCKS', 2)
        fails(ex, path.name, 'capacity')
    with monkeypatch.context() as patch:
        patch.setattr(sources, 'MAX_TOTAL_TEXT', 2)
        fails(ex, path.name, 'capacity')
    with monkeypatch.context() as patch:
        patch.setattr(sources, 'MAX_SOURCE_BYTES', 2)
        fails(ex, path.name, 'capacity')
    ex.extract(path.name)
    (tmp_path / 'other.txt').write_text('other', encoding='utf-8')
    monkeypatch.setattr(sources, 'MAX_IDENTITIES', 1)
    fails(ex, 'other.txt', 'capacity')


def test_encrypted_zip_flags_rejected(tmp_path):
    path = package(tmp_path)
    data = bytearray(path.read_bytes())
    # Synthetic encrypted-member declarations; never decrypt or prompt for secrets.
    for marker, offset in [(b'PK\x03\x04', 6), (b'PK\x01\x02', 8)]:
        position = data.find(marker)
        data[position + offset] |= 1
    path.write_bytes(data)
    fails(extractor(tmp_path), path.name, 'unsupported')


def test_symlink_mode_simulation(tmp_path, monkeypatch):
    import stat
    path = tmp_path / 'notes.txt'
    path.write_text('hello', encoding='utf-8')
    ex = extractor(tmp_path)
    original = Path.lstat
    def linked(path, *args, **kwargs):
        if path.name == 'notes.txt':
            return SimpleNamespace(st_mode=stat.S_IFLNK, st_nlink=1, st_file_attributes=0)
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'lstat', linked)
    fails(ex, path.name, 'invalid')


def test_bad_core_and_empty_table(tmp_path):
    path = package(tmp_path, body=f'<w:wrong xmlns:w="{W}"/>')
    fails(extractor(tmp_path), path.name, 'invalid')
    path = package(tmp_path, body=f'<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>')
    fails(extractor(tmp_path), path.name, 'invalid')


def test_word_external_field_and_excel_secret(tmp_path):
    path = package(tmp_path, body=f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:instrText>DDEAUTO server topic</w:instrText><w:t>cached</w:t></w:r></w:p></w:body></w:document>')
    fails(extractor(tmp_path), path.name, 'unsupported')
    path = package(tmp_path, 'xlsx', f'<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>api_key=synthetic-fixture</t></is></c></row></sheetData></worksheet>')
    fails(extractor(tmp_path), path.name, 'sensitive')


def test_zip_perentry_total_and_xml_nodes(tmp_path, monkeypatch):
    path = package(tmp_path)
    ex = extractor(tmp_path)
    for limit in ('MAX_ENTRY_BYTES', 'MAX_ZIP_BYTES', 'MAX_XML_NODES'):
        with monkeypatch.context() as patch:
            patch.setattr(sources, limit, 1)
            fails(ex, path.name, 'capacity')


def test_workspace_root_identity_simulation(tmp_path, monkeypatch):
    (tmp_path / 'notes.txt').write_text('hello', encoding='utf-8')
    ex = extractor(tmp_path)
    original = Path.lstat
    def replaced(path, *args, **kwargs):
        info = original(path, *args, **kwargs)
        if path == tmp_path:
            return SimpleNamespace(st_dev=info.st_dev, st_ino=info.st_ino + 1,
                                   st_mode=info.st_mode, st_nlink=info.st_nlink, st_file_attributes=0)
        return info
    monkeypatch.setattr(Path, 'lstat', replaced)
    fails(ex, 'notes.txt', 'source_changed')
