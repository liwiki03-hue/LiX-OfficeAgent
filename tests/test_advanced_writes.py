import json
from pathlib import Path
import zipfile

from docx import Document
from openpyxl import Workbook, load_workbook
from PIL import Image
from pptx import Presentation
import pytest

from office_agent import advanced_writes as advanced
from office_agent.advanced_writes import (WordNumberedList, WordLocalImage, WordToc,
    PptxReplace, PptxSlide, ExcelRecalculate, DocumentExport, formula_results)
from office_agent.writes import Editor, GuardError, WordReplace, digest


@pytest.fixture
def samples(tmp_path):
    word = tmp_path/'source.docx'
    doc = Document()
    doc.add_paragraph('Original paragraph')
    doc.save(word)
    image = tmp_path/'picture.png'
    Image.new('RGB', (32, 20), 'blue').save(image)
    excel = tmp_path/'source.xlsx'
    book = Workbook()
    sheet = book.active
    sheet['A1'], sheet['A2'], sheet['B1'] = 2, 3, '=SUM(A1:A2)*2'
    book.save(excel)
    literal = tmp_path/'literal.xlsx'
    sheet['B1'] = 'ordinary'
    book.save(literal)
    pptx = tmp_path/'source.pptx'
    deck = Presentation()
    for text in ('First', 'Second'):
        slide = deck.slides.add_slide(deck.slide_layouts[5])
        slide.shapes.title.text = text
    deck.save(pptx)
    return {'word': word, 'excel': excel, 'literal': literal, 'pptx': pptx, 'image': image}


def requests(samples):
    return [WordNumberedList(path=str(samples['word']), items=['One', 'Two']),
            WordLocalImage(path=str(samples['word']), image_name='picture.png'),
            WordToc(path=str(samples['word'])),
            PptxReplace(path=str(samples['pptx']), slide_index=0, shape_id=2, old='First', new='Changed'),
            PptxSlide(path=str(samples['pptx']), action='delete', slide_index=0),
            PptxSlide(path=str(samples['pptx']), action='add_blank', slide_index=0),
            ExcelRecalculate(path=str(samples['excel'])),
            DocumentExport(path=str(samples['word']), format='pdf'),
            DocumentExport(path=str(samples['literal']), format='csv', sheet='Sheet')]


@pytest.mark.parametrize('index', range(9))
@pytest.mark.parametrize('confirmation', [False, 1, 'true'])
def test_new_writes_require_exact_human_confirmation(samples, tmp_path, index, confirmation):
    editor = Editor(tmp_path)
    request = requests(samples)[index]
    source = Path(request.path)
    before = source.read_bytes()
    plan = editor.plan(request)
    assert plan.confirmation_required
    with pytest.raises(GuardError, match='confirmation'):
        editor.execute(plan.id, confirmed=confirmation)
    assert source.read_bytes() == before
    assert plan.id in editor.plans
    assert not (tmp_path/'.office_agent').exists()


@pytest.mark.parametrize('index', [0, 1, 2, 3, 4, 5, 8])
def test_new_writes_preserve_source_audit_and_rollback(samples, tmp_path, monkeypatch, index):
    monkeypatch.setattr(advanced, 'render_gate', lambda *args: {'render': {'status': 'different'}, 'test_double': True})
    editor = Editor(tmp_path)
    request = requests(samples)[index]
    source = Path(request.path)
    before = source.read_bytes()
    plan = editor.plan(request)
    log = editor.execute(plan.id, confirmed=True)
    assert log['status'] == 'verified' and source.read_bytes() == before
    output = Path(log['output'])
    assert output.is_file() and Path(log['backup']).read_bytes() == before
    record = json.loads((tmp_path/'.office_agent'/plan.id/'audit.json').read_text(encoding='utf-8'))
    assert record['output_sha256'] == digest(output)
    if index == 0:
        assert [p.text for p in Document(output).paragraphs] == ['Original paragraph', 'One', 'Two']
    elif index == 1:
        assert len(Document(output).inline_shapes) == 1
    elif index == 2:
        assert 'page numbers' in json.loads(plan.after)['notice']
    elif index in {3, 4, 5}:
        deck = Presentation(output)
        assert len(deck.slides) == {3: 2, 4: 1, 5: 3}[index]
        if index == 3:
            assert deck.slides[0].shapes.title.text == 'Changed'
    else:
        assert output.read_text(encoding='utf-8-sig') == '2,ordinary\n3,\n'
    with pytest.raises(GuardError):
        editor.rollback(plan.id)
    assert editor.rollback(plan.id, confirmed=True)['status'] == 'rolled_back'
    assert not output.exists() and source.read_bytes() == before


def test_render_error_refuses_publication_and_records_failure(samples, tmp_path, monkeypatch):
    def fail(*args):
        error = GuardError('render validation failed')
        error.render_diff = {'render': {'status': 'error', 'code': 'worker_timeout'}}
        raise error
    monkeypatch.setattr(advanced, 'render_gate', fail)
    editor = Editor(tmp_path)
    plan = editor.plan(requests(samples)[0])
    with pytest.raises(GuardError):
        editor.execute(plan.id, confirmed=True)
    log = json.loads((tmp_path/'.office_agent'/plan.id/'audit.json').read_text(encoding='utf-8'))
    assert log['status'] == 'failed' and log['render_diff']['render']['code'] == 'worker_timeout'
    assert not Path(log['output']).exists() and digest(samples['word']) == plan.source_sha256


def test_existing_write_uses_opted_in_render_gate(samples, tmp_path, monkeypatch):
    called = []
    monkeypatch.setattr(advanced, 'render_gate', lambda *args: called.append(args) or {'render': {'status': 'different'}})
    editor = Editor(tmp_path, render_validation=True, require_confirmation=True)
    plan = editor.plan(WordReplace(path=str(samples['word']), paragraph_index=0, old='Original', new='Updated'))
    result = editor.execute(plan.id, confirmed=True)
    assert called and 'render_diff' in result


@pytest.mark.parametrize('formula', ['=NOW()', '=A1/0', '=B1', '=SUM(Sheet2!A1)', '=WEBSERVICE("x")', '=2**999', '=INDIRECT("A1")'])
def test_recalculation_rejects_unverifiable_formulas(samples, tmp_path, formula):
    book = load_workbook(samples['excel'])
    book.active['B1'] = formula
    book.save(samples['excel'])
    before = digest(samples['excel'])
    with pytest.raises(GuardError):
        Editor(tmp_path).plan(ExcelRecalculate(path=str(samples['excel'])))
    assert digest(samples['excel']) == before


def test_formula_expected_value_is_independently_calculated(samples):
    assert formula_results(samples['excel']) == {'Sheet': {'B1': 10}}


def test_cache_transplant_only_after_actual_matches(samples, tmp_path, monkeypatch):
    source = samples['excel']
    original = advanced._parts(source)
    def native(path, mode, extension, runs):
        parts = advanced._parts(path)
        root = advanced.ET.fromstring(parts['xl/worksheets/sheet1.xml'])
        cell = next(n for n in root.iter(advanced.S+'c') if n.get('r') == 'B1')
        cell.find(advanced.S+'v').text = '10'
        parts['xl/worksheets/sheet1.xml'] = advanced._xml(root)
        target = tmp_path/'native.xlsx'
        advanced._zip(parts, target)
        return target.read_bytes()
    monkeypatch.setattr(advanced, '_native', native)
    monkeypatch.setattr(advanced, 'render_gate', lambda *args: {'test_double': True})
    editor = Editor(tmp_path)
    plan = editor.plan(ExcelRecalculate(path=str(source)))
    result = editor.execute(plan.id, confirmed=True)
    assert result['calculation']['status'] == 'independently_verified'
    output = Path(result['output'])
    assert load_workbook(output, data_only=True).active['B1'].value == 10
    assert load_workbook(output, data_only=False).active['B1'].value == '=SUM(A1:A2)*2'
    advanced.verify_change(original, advanced._parts(output), plan.request)


@pytest.mark.parametrize('text', ['=CMD()', '+CMD()', '@SUM(A1)', '  =CMD()', '-bad'])
def test_csv_rejects_injection(samples, tmp_path, text):
    book = load_workbook(samples['literal'])
    book.active['A1'] = text
    book.save(samples['literal'])
    with pytest.raises(GuardError):
        Editor(tmp_path).plan(DocumentExport(path=str(samples['literal']), format='csv', sheet='Sheet'))


def test_secondary_image_changed_since_preview_refuses(samples, tmp_path):
    editor = Editor(tmp_path)
    plan = editor.plan(requests(samples)[1])
    Image.new('RGB', (32, 20), 'red').save(samples['image'])
    with pytest.raises(GuardError, match='inputs changed'):
        editor.execute(plan.id, confirmed=True)
    assert not list(tmp_path.glob('*.agent-*'))


def test_non_target_package_mutation_is_rejected(samples):
    before = advanced._parts(samples['word'])
    after = dict(before)
    after['word/styles.xml'] += b' '
    with pytest.raises(GuardError, match='non-target'):
        advanced.verify_change(before, after, requests(samples)[0])


def test_advanced_pack_requires_separate_write_unlock(samples):
    from office_agent.app import build_tools, select_tools, split_read_write
    from office_agent.agent import capability_pack
    tools = build_tools((samples['word'],))
    selected = select_tools(tools, 'advanced')
    read, write = split_read_write(selected)
    assert set(advanced.TOOL_MODELS) == {spec.name for spec in write}
    assert not any(spec.name.startswith('plan_') for spec in read)
    assert set(advanced.TOOL_MODELS) <= {spec.name for spec in capability_pack(tools, 'advanced')}

@pytest.mark.parametrize('kind', ['pptx_replace', 'pptx_slide', 'excel_recalculate'])
def test_unexpected_render_change_rejected(samples, monkeypatch, kind):
    from office_agent import fidelity
    report = {'inputs_unchanged': True, 'render': {'status': 'different', 'equivalent': False,
        'page_count_equal': True, 'source_pages': 2, 'output_pages': 2,
        'pages': [{'page': 2, 'text_equal': False, 'size_equal': True, 'pixels_equal': False}]}}
    monkeypatch.setattr(fidelity, 'compare_documents', lambda *args: report)
    request = next(r for r in requests(samples) if r.kind == kind)
    with pytest.raises(GuardError, match='render validation'):
        advanced.render_gate(samples['word'].parent, samples['word'], samples['word'], request)


def test_stale_primary_source_refused(samples, tmp_path):
    editor = Editor(tmp_path)
    plan = editor.plan(requests(samples)[0])
    doc = Document(samples['word'])
    doc.add_paragraph('Changed later')
    doc.save(samples['word'])
    before = digest(samples['word'])
    with pytest.raises(GuardError, match='source changed'):
        editor.execute(plan.id, confirmed=True)
    assert digest(samples['word']) == before
    assert not list(tmp_path.glob('*.agent-*'))


def test_changed_output_cannot_be_rolled_back(samples, tmp_path, monkeypatch):
    monkeypatch.setattr(advanced, 'render_gate', lambda *args: {'test_double': True})
    editor = Editor(tmp_path)
    plan = editor.plan(requests(samples)[0])
    result = editor.execute(plan.id, confirmed=True)
    output = Path(result['output'])
    output.write_bytes(b'new owner content')
    with pytest.raises(GuardError, match='output changed'):
        editor.rollback(plan.id, confirmed=True)
    assert output.read_bytes() == b'new owner content'


def test_presentation_host_can_offer_write_gate_without_changing_read_pack(samples):
    from office_agent.app import build_tools, select_tools, split_read_write
    read, write = split_read_write(select_tools(build_tools((samples['pptx'],)), 'presentation_edit'))
    assert {t.name for t in write} == {'plan_pptx_replace', 'plan_pptx_slide'}
    assert 'request_write_tools' in {t.name for t in read}


def test_complex_presentation_refused(samples, tmp_path):
    deck = Presentation(samples['pptx'])
    deck.slides[0].shapes.add_picture(str(samples['image']), 0, 0)
    deck.save(samples['pptx'])
    with pytest.raises(GuardError, match='complex'):
        Editor(tmp_path).plan(requests(samples)[3])


def test_internal_relationship_cannot_escape_before_render(samples):
    path = samples['word']
    parts = advanced._parts(path)
    root = advanced.ET.fromstring(parts['word/_rels/document.xml.rels'])
    root[0].set('Target', '../../outside.xml')
    parts['word/_rels/document.xml.rels'] = advanced._xml(root)
    altered = path.with_name('unsafe.docx')
    advanced._zip(parts, altered)
    with pytest.raises(GuardError, match='unsafe_relationship'):
        advanced._parts(altered)


def test_package_uppercase_suffix_supported(samples, tmp_path, monkeypatch):
    path = tmp_path/'UPPER.DOCX'
    path.write_bytes(samples['word'].read_bytes())
    monkeypatch.setattr(advanced, 'render_gate', lambda *args: {'test_double': True})
    editor = Editor(tmp_path)
    plan = editor.plan(WordNumberedList(path=str(path), items=['One']))
    assert editor.execute(plan.id, confirmed=True)['status'] == 'verified'


def test_unexpected_render_exception_is_auditable(samples, monkeypatch):
    from office_agent import fidelity
    def fail(*args):
        raise OSError('private error detail')
    monkeypatch.setattr(fidelity, 'compare_documents', fail)
    with pytest.raises(GuardError) as captured:
        advanced.render_gate(samples['word'].parent, samples['word'], samples['word'])
    assert captured.value.render_diff['render'] == {'status': 'error', 'equivalent': None, 'code': 'render_validation_failed'}
    assert 'private' not in str(captured.value)


def test_mismatched_native_formula_result_is_not_published(samples, tmp_path, monkeypatch):
    def native(path, mode, extension, runs):
        parts = advanced._parts(path)
        root = advanced.ET.fromstring(parts['xl/worksheets/sheet1.xml'])
        cell = next(n for n in root.iter(advanced.S+'c') if n.get('r') == 'B1')
        cell.find(advanced.S+'v').text = '999'
        parts['xl/worksheets/sheet1.xml'] = advanced._xml(root)
        result = tmp_path/'incorrect-native.xlsx'
        advanced._zip(parts, result)
        return result.read_bytes()
    monkeypatch.setattr(advanced, '_native', native)
    editor = Editor(tmp_path)
    plan = editor.plan(ExcelRecalculate(path=str(samples['excel'])))
    with pytest.raises(GuardError, match='independent expected'):
        editor.execute(plan.id, confirmed=True)
    log = json.loads((tmp_path/'.office_agent'/plan.id/'audit.json').read_text(encoding='utf-8'))
    assert log['status'] == 'failed' and not Path(log['output']).exists()
    assert digest(samples['excel']) == plan.source_sha256
