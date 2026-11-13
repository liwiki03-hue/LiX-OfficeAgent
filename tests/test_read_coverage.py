import base64
from copy import deepcopy
import json
import io
from pathlib import Path
import struct
import zipfile

from docx import Document
from openpyxl import Workbook
from pydantic import ValidationError
import pytest

from office_agent.agent import ToolSpec, StrictModel, capability_pack
from office_agent.app import (DirectorySearchArgs, PdfReadArgs, PptxSlideArgs, build_dispatcher,
                              build_tools, select_tools, split_read_write)
from office_agent.directory_search import DirectoryIndex
from office_agent.documents import classify_document
from office_agent.legacy_reader import read_legacy_document, validate_legacy
from office_agent.pdf_reader import read_pdf, read_pdf_payload
from office_agent.presentations import read_pptx_slide, read_pptx_structure
from office_agent.rag_sources import SourceError
from office_agent.writes import Editor, GuardError, WordReplace, Workspace, digest
from test_fidelity import pdf_bytes


P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'


def make_pptx(path, *, long_text=False, external=False):
    def relationships(rows):
        return f'<Relationships xmlns="{REL}">' + ''.join(
            f'<Relationship Id="{identifier}" Type="{R}/{kind}" Target="{target}"/>' for identifier, kind, target in rows) + '</Relationships>'

    def shape(value, identifier=1, title=False):
        placeholder = '<p:ph type="title"/>' if title else ''
        return f'<p:sp><p:nvSpPr><p:cNvPr id="{identifier}" name="Box"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{value}</a:t></a:r></a:p></p:txBody></p:sp>'

    wrappers = f'xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" xmlns:c="{C}"'
    title = shape('First title', title=True)
    text = shape('x' * 3000 + 'END' if long_text else 'Budget 7319', 2)
    picture = '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Image"/></p:nvPicPr><p:blipFill><a:blip r:embed="image"/></p:blipFill></p:pic>'
    table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc gridSpan="2"><a:txBody><a:p><a:r><a:t>Merged cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    chart = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>'
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>',
        '_rels/.rels': relationships([('office', 'officeDocument', 'ppt/presentation.xml')]),
        'ppt/presentation.xml': f'<p:presentation {wrappers}><p:sldIdLst><p:sldId id="256" r:id="first"/><p:sldId id="257" r:id="second"/></p:sldIdLst></p:presentation>',
        'ppt/_rels/presentation.xml.rels': relationships([('first', 'slide', 'slides/slide2.xml'), ('second', 'slide', 'slides/slide1.xml')]),
        'ppt/slides/slide2.xml': f'<p:sld {wrappers}><p:cSld><p:spTree>{title}{text}{picture}{table}{chart}</p:spTree></p:cSld></p:sld>',
        'ppt/slides/slide1.xml': f'<p:sld {wrappers}><p:cSld><p:spTree>{shape("Second title", title=True)}</p:spTree></p:cSld></p:sld>',
        'ppt/slides/_rels/slide2.xml.rels': relationships([('image', 'image', '../media/image1.png'), ('chart', 'chart', '../charts/chart1.xml'), ('notes', 'notesSlide', '../notesSlides/notesSlide1.xml')]),
        'ppt/charts/chart1.xml': f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly amount</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>',
        'ppt/notesSlides/notesSlide1.xml': f'<p:notes {wrappers}><p:cSld><p:spTree>{shape("Speaker notes")}</p:spTree></p:cSld></p:notes>',
        'ppt/media/image1.png': base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lZkAAAAASUVORK5CYII='),
    }
    if external:
        parts['ppt/slides/_rels/slide2.xml.rels'] = f'<Relationships xmlns="{REL}"><Relationship Id="image" Type="{R}/image" Target="https://example.invalid/x.png" TargetMode="External"/></Relationships>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return path


def test_pptx_order_notes_tables_chart_and_images(tmp_path):
    path = make_pptx(tmp_path / 'slides.pptx')
    original = digest(path)
    report = read_pptx_structure(path, limit=1)
    assert report['slide_count'] == 2 and report['slides'][0]['title'] == 'First title'
    assert report['truncated'] is True and report['next_offset'] == 1
    assert read_pptx_structure(path, offset=1)['slides'][0]['title'] == 'Second title'
    slide = read_pptx_slide(path)
    by_type = {item['type']: item for item in slide['items']}
    assert by_type['notes']['text'] == 'Speaker notes'
    assert by_type['table_cell']['text'] == 'Merged cell'
    assert by_type['table_cell']['merge'] == {'gridSpan': '2'}
    assert by_type['chart']['text'] == 'Quarterly amount'
    assert by_type['image']['part'] == 'ppt/media/image1.png'
    assert slide['items'][0]['anchor']['coordinate_space'] == 'parent'
    assert digest(path) == original and slide['source_sha256'] == original


def test_pptx_partial_text_continuation_has_no_loss(tmp_path):
    path = make_pptx(tmp_path / 'long.pptx', long_text=True)
    cursor, text_offset, collected = 1, 0, ''
    for _ in range(30):
        result = read_pptx_slide(path, offset=cursor, text_offset=text_offset, limit=1, max_chars=200)
        item = result['items'][0]
        assert item['item_index'] == 1 and len(item['text']) <= 200
        collected += item['text']
        if not item['text_truncated']:
            break
        assert result['next_offset'] == 1 and result['next_text_offset'] > text_offset
        cursor, text_offset = result['next_offset'], result['next_text_offset']
    assert collected == 'x' * 3000 + 'END'


@pytest.mark.parametrize('options', [{'limit': 51}, {'slide_index': -1}, {'slide_index': 2}, {'offset': 20001}, {'max_chars': True}, {'text_offset': 30000}])
def test_pptx_invalid_bounds_refused(tmp_path, options):
    path = make_pptx(tmp_path / 'slides.pptx')
    original = digest(path)
    with pytest.raises(SourceError):
        read_pptx_slide(path, **options)
    assert digest(path) == original


def test_pptx_external_relationship_refused_and_source_unchanged(tmp_path):
    path = make_pptx(tmp_path / 'unsafe.pptx', external=True)
    original = digest(path)
    with pytest.raises(SourceError):
        read_pptx_structure(path)
    assert digest(path) == original


def test_pdf_page_and_char_continuations_and_coordinates(tmp_path):
    path = tmp_path / 'test.pdf'
    path.write_bytes(pdf_bytes(text='Hello ' * 100, pages=2))
    original = digest(path)
    first = read_pdf_payload(path, max_chars=200)
    assert len(first['pages'][0]['text']) == 200
    assert first['truncated'] and first['next_page'] == 0 and first['next_char_offset'] == 200
    assert len(first['pages'][0]['runs'][0]['bbox']) == 4
    second = read_pdf_payload(path, page=0, char_offset=200, max_chars=20000)
    assert first['pages'][0]['text'] + second['pages'][0]['text'] == 'Hello ' * 100
    assert second['pages'][-1]['page'] == 1
    assert second['next_page'] is None and second['ocr_performed'] is False
    assert digest(path) == original


def test_pdf_empty_layer_does_not_claim_page_is_blank(tmp_path):
    path = tmp_path / 'empty.pdf'
    path.write_bytes(pdf_bytes(text=''))
    report = read_pdf_payload(path)
    assert report['pages'][0]['text_layer_empty'] is True
    assert report['ocr_performed'] is False
    assert report['pages'][0]['text'] == ''


@pytest.mark.parametrize('options', [{'limit': 6}, {'page': -1}, {'char_offset': 2000001}, {'layout': 1}, {'max_chars': 199}])
def test_pdf_invalid_bounds_fail_before_worker(tmp_path, options, monkeypatch):
    path = tmp_path / 'test.pdf'
    path.write_bytes(pdf_bytes())
    monkeypatch.setattr('office_agent.pdf_reader._run', lambda *a, **k: pytest.fail('worker must not start'))
    with pytest.raises(SourceError):
        read_pdf(path, **options)


def word(path, text):
    document = Document()
    document.add_paragraph(text)
    document.save(path)
    return path


def test_directory_index_returns_sources_locations_and_skips_legacy(tmp_path):
    first = word(tmp_path / 'first.docx', '核准预算7319元')
    second = word(tmp_path / 'second.docx', '核准预算8000元')
    (tmp_path / 'old.doc').write_bytes(b'unsupported legacy')
    (tmp_path / 'nested').mkdir()
    word(tmp_path / 'nested' / 'third.docx', '核准预算9000元')
    before = {path: digest(path) for path in (first, second)}
    result = DirectoryIndex(Workspace(tmp_path), first).build().query('核准预算', top_k=1)
    assert result['document_count'] == 2 and len(result['hits']) == 1
    assert result['truncated'] is True and result['candidate_count'] == 2
    assert result['hits'][0]['source'] in {'first.docx', 'second.docx'}
    assert 'body:1/paragraph' in result['hits'][0]['location']
    assert any(item['source'] == 'old.doc' and item['reason'] == 'legacy_format' for item in result['inventory'])
    assert any(item['source'] == 'nested' and item['status'] == 'skipped' for item in result['inventory'])
    assert result['semantic_available'] is False and result['index_persisted'] is False
    assert {path: digest(path) for path in before} == before


class Embedder:
    identity = {'channel': 'local', 'model': 'synthetic-test-embedding', 'dimension': 2}

    def embed(self, texts):
        return [[1, 0] if 'vehicle' in text or 'car' in text else [0, 1] for text in texts]


def test_hybrid_cross_document_query_and_model_or_source_change_requires_rebuild(tmp_path):
    first = word(tmp_path / 'vehicle.docx', 'vehicle registration')
    other = word(tmp_path / 'fruit.docx', 'apple orange')
    embedder = Embedder()
    index = DirectoryIndex(Workspace(tmp_path), first, embedder=embedder).build()
    result = index.query('car')
    assert result['hits'][0]['source'] == 'vehicle.docx'
    assert result['retrieval'] == 'inverted_index_and_vector_rrf'
    embedder.identity = {**embedder.identity, 'model': 'different-test-model'}
    with pytest.raises(SourceError, match='source_changed'):
        index.query('car')
    embedder.identity = dict(Embedder.identity)
    word(other, 'changed source')
    with pytest.raises(SourceError, match='source_changed'):
        index.query('car')


def test_directory_candidates_each_use_workspace_guard_and_cap_twenty(tmp_path, monkeypatch):
    for index in range(21):
        word(tmp_path / f'{index:02}.docx', 'budget 17')
    selected = tmp_path / '00.docx'
    workspace = Workspace(tmp_path)
    original = workspace.file
    seen = []

    def guarded(path, **kwargs):
        seen.append(Path(path).name)
        if Path(path).name == '01.docx':
            raise GuardError('simulated link risk')
        return original(path, **kwargs)

    monkeypatch.setattr(workspace, 'file', guarded)
    index = DirectoryIndex(workspace, selected).build()
    report = index.query('budget')
    assert set(seen) == {f'{index:02}.docx' for index in range(21)}
    assert report['document_count'] == 20
    rejected = next(item for item in report['inventory'] if item['source'] == '01.docx')
    assert rejected['reason'] == 'path_or_parse_rejected'


def test_directory_twenty_first_safe_document_is_reported(tmp_path):
    for index in range(21):
        word(tmp_path / f'{index:02}.docx', 'budget 17')
    report = DirectoryIndex(Workspace(tmp_path), tmp_path / '00.docx').build().query('budget')
    assert report['document_count'] == 20
    assert next(item for item in report['inventory'] if item['source'] == '20.docx')['reason'] == 'document_limit'


@pytest.mark.parametrize('data,suffix', [(b'invalid', '.doc'), (b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'0' * 64, '.xls')])
def test_invalid_legacy_rejected_without_native_process(data, suffix, tmp_path, monkeypatch):
    path = tmp_path / ('legacy' + suffix)
    path.write_bytes(data)
    monkeypatch.setattr('office_agent.legacy_reader._run', lambda *a, **k: pytest.fail('must not start'))
    with pytest.raises(SourceError):
        read_legacy_document(path)
    assert path.read_bytes() == data


def test_new_tool_packs_are_read_only_and_old_write_tools_stay_hidden(tmp_path):
    tools = build_tools((tmp_path / 'selected.pptx',))
    for kind, expected in [('presentation', {'read_pptx_structure', 'read_pptx_slide'}), ('pdf', {'read_pdf'}), ('directory', {'search_directory_documents'})]:
        selected = select_tools(tools, kind)
        read, write = split_read_write(selected)
        assert not write and all(tool.access == 'read' for tool in read)
        assert expected <= {tool.name for tool in read}
        assert expected <= {tool.name for tool in capability_pack(tools, kind)}
    assert 'read_pptx_slide' not in {tool.name for tool in capability_pack(tools, 'word')}


def test_dispatcher_checks_selected_path_version_and_does_not_create_write_plans(tmp_path):
    source = make_pptx(tmp_path / 'source.pptx')
    other = make_pptx(tmp_path / 'other.pptx')
    editor = Editor(tmp_path)
    dispatcher = build_dispatcher(editor, source)
    assert dispatcher('read_pptx_slide', {'path': str(source)})['read_only'] is True
    assert classify_document(source).kind == 'presentation'
    with pytest.raises(ValueError, match='explicitly selected'):
        dispatcher('read_pptx_slide', {'path': str(other)})
    with pytest.raises(ValidationError):
        dispatcher('read_pptx_slide', {'path': str(source), 'confirmed': True})
    source.write_bytes(source.read_bytes() + b'changed')
    with pytest.raises(ValueError, match='changed'):
        dispatcher('read_pptx_slide', {'path': str(source)})
    assert not editor.plans and not editor.audit.exists()


def test_schema_context_compaction_keeps_field_rules_and_descriptions():
    from pydantic import Field

    class Args(StrictModel):
        value: int = Field(ge=1, le=5, description='bounded value')

    Args.__doc__ = 'Repeated tool-level explanation'
    schema = ToolSpec(name='test', description='Host tool explanation', schema_model=Args).model_schema()
    assert schema['function']['description'] == 'Host tool explanation'
    value = schema['function']['parameters']['properties']['value']
    assert value['minimum'] == 1 and value['maximum'] == 5 and value['description'] == 'bounded value'
    assert schema['function']['parameters']['additionalProperties'] is False


@pytest.mark.parametrize('schema,args', [(PdfReadArgs, {'page': 2000}), (PptxSlideArgs, {'limit': 51}), (DirectorySearchArgs, {'query': 'x', 'top_k': 11})])
def test_model_schemas_enforce_limits(schema, args):
    with pytest.raises(ValidationError):
        schema.model_validate({'path': 'file', **args})


@pytest.mark.parametrize('risk', ['passive', 'vba', 'encrypted', 'external_field', 'formula', 'external_book', 'defined_name', 'macro_sheet'])
def test_legacy_static_admission_rejects_active_streams_and_records(monkeypatch, risk):
    import olefile

    word_data = bytearray(32)
    word_data[:2] = b'\xec\xa5'
    streams = {('WordDocument',): bytes(word_data), ('0Table',): b'passive text'}
    suffix = '.doc'
    if risk == 'vba':
        streams[('Macros', 'VBA')] = b'code'
    elif risk == 'encrypted':
        word_data[10:12] = struct.pack('<H', 0x100)
        streams[('WordDocument',)] = bytes(word_data)
    elif risk == 'external_field':
        streams[('0Table',)] = 'INCLUDETEXT https://example.invalid'.encode('utf-16-le')
    elif risk in {'formula', 'external_book', 'defined_name', 'macro_sheet'}:
        suffix = '.xls'
        bof = struct.pack('<HHHH', 0x0809, 4, 0x0600, 5)
        record = {'formula': 0x0006, 'external_book': 0x01AE, 'defined_name': 0x0018, 'macro_sheet': 0x0085}[risk]
        payload = b'\x00' * 5 + b'\x01' if risk == 'macro_sheet' else b'unsafe'
        streams = {('Workbook',): bof + struct.pack('<HH', record, len(payload)) + payload}

    class SyntheticOle:
        def __enter__(self):
            return self

        def __exit__(self, *args):
            return None

        def listdir(self):
            return [list(name) for name in streams]

        def get_size(self, name):
            return len(streams[tuple(name)])

        def openstream(self, name):
            return io.BytesIO(streams[tuple(name)])

    monkeypatch.setattr(olefile, 'OleFileIO', lambda *a, **k: SyntheticOle())
    payload = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'fixture'
    if risk == 'passive':
        validate_legacy(payload, suffix)
    else:
        with pytest.raises(SourceError, match='unsupported'):
            validate_legacy(payload, suffix)


@pytest.mark.parametrize('available', [False, True])
def test_legacy_fallback_for_detection_or_activation_failure(tmp_path, monkeypatch, available):
    import office_agent.com
    import office_agent.legacy_reader

    source = tmp_path / 'legacy.doc'
    source.write_bytes(b'fixture; mocked readers only')
    original = digest(source)
    monkeypatch.setattr(office_agent.com, 'capabilities', lambda: {'word': {'available': available}})

    def unavailable(*args, **kwargs):
        raise RuntimeError('synthetic COM activation refusal')

    monkeypatch.setattr(office_agent.com, 'extract_office', unavailable)
    calls = []

    def fallback(path, **kwargs):
        calls.append((path, kwargs))
        return {'backend': 'synthetic_read_only', 'read_only': True}

    monkeypatch.setattr(office_agent.legacy_reader, 'read_legacy_document', fallback)
    editor = Editor(tmp_path)
    result = build_dispatcher(editor, source)('com_read', {'path': str(source), 'include_tables': True, 'region': None})
    assert result['read_only'] is True
    assert calls == [(source, {'include_tables': True, 'region': None})]
    assert digest(source) == original and not editor.plans and not editor.audit.exists()


def test_capacity_skips_remain_visible_in_model_directory_coverage(tmp_path, monkeypatch):
    from office_agent.rag_sources import SourceExtractor
    from office_agent.agent import _read_context

    selected = word(tmp_path / 'good.docx', 'budget 17')
    word(tmp_path / 'large.docx', 'fixture simulates capacity refusal')
    original = SourceExtractor.extract

    def extract(self, name):
        if name == 'large.docx':
            raise SourceError('capacity')
        return original(self, name)

    monkeypatch.setattr(SourceExtractor, 'extract', extract)
    report = DirectoryIndex(Workspace(tmp_path), selected).build().query('budget')
    assert report['corpus_truncated'] is True and report['coverage_complete'] is False
    assert report['skipped_document_count'] == 1
    view = json.loads(_read_context('search_directory_documents', report, {}))
    assert view['coverage_complete'] is False and view['skipped_document_count'] == 1
