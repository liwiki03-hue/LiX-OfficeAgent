import base64
import io
import time
import zipfile
from types import SimpleNamespace

from docx import Document
from openpyxl import load_workbook
import pytest

from office_agent.canvas_ranges import parse_tsv, read_range
from office_agent import office_canvas as engine
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from office_agent.writes import GuardError
from tests.test_office_canvas import fixtures, operation
from tests.test_paged_editor import draft


def sync_book(service, book, ids, revision, bootstrap=False):
    stream = io.BytesIO(); book.save(stream)
    manifest = [{'id':identifier,'text':paragraph.text} for identifier, paragraph in zip(ids, book.paragraphs)]
    if revision != service.revision:
        service.changed(service.document_id, revision)
    return service.sync(service.document_id, revision, base64.b64encode(stream.getvalue()).decode(), manifest, bootstrap)


def test_new_and_rewritten_paragraphs_get_current_draft_locator_without_rebinding_source(draft):
    service, floating, source = draft
    original = source.read_bytes(); source_locators = dict(floating.locators)
    book = Document(source)
    sync_book(service, book, ['a','b','c'], 0, True)
    initial = {row['editor_id']:row for row in service.id_map(service.document_id)['rows']}
    old = initial['b']['draft_locator_id']
    book.paragraphs[1].runs[0].text = '实验于2026年完成，温度120 ℃。'
    added = book.add_paragraph('实验于2026年完成，温度120 ℃。')
    book.paragraphs[1]._p.addprevious(added._p)
    service.changed(service.document_id, 1)
    with pytest.raises(GuardError, match='同步'):
        service.id_map(service.document_id)
    sync_book(service, book, ['a','new','rewritten','c'], 1)
    current = {row['editor_id']:row for row in service.id_map(service.document_id)['rows']}
    assert current['c']['host_locator_id'] == initial['c']['host_locator_id']
    assert current['c']['paragraph_index'] == 3 and current['c']['draft_revision'] == 1
    assert current['new']['host_locator_id'] is None and current['rewritten']['host_locator_id'] is None
    assert len({row['draft_locator_id'] for row in current.values()}) == 4
    reference = service.reference_locator(service.document_id, 1, current['new']['draft_locator_id'], 0, 2)['reference']
    assert reference['node_id'] == 'new' and reference['old'] == '实验'
    assert reference['source_host_locator_id'] is None and reference['locator_scope'] == 'current_draft'
    with pytest.raises(GuardError, match='过期'):
        service.reference_locator(service.document_id, 1, old, 0, 2)
    with pytest.raises(GuardError):
        service.reference_locator(service.document_id, 0, old, 0, 2)
    assert source.read_bytes() == original and floating.locators == source_locators


def test_empty_source_map_never_reinitializes_by_current_index_and_same_snapshot_keeps_ids(draft):
    service, floating, source = draft
    book = Document(source)
    sync_book(service, book, ['a','b','c'], 0, True)
    sync_book(service, book, ['x','y','z'], 1)
    assert all(row['host_locator_id'] is None for row in service.id_map(service.document_id)['rows'])
    current = service.id_map(service.document_id)
    sync_book(service, book, ['x','y','z'], 1)
    assert service.id_map(service.document_id) == current
    sync_book(service, book, ['u','v','w'], 2)
    rows = service.id_map(service.document_id)['rows']
    assert all(row['host_locator_id'] is None for row in rows)
    locator = rows[0]['draft_locator_id']
    floating.epoch += 1
    with pytest.raises(GuardError, match='过期'):
        service.reference_locator(service.document_id, 2, locator, 0, 1)


def test_first_nonbootstrap_sync_does_not_invent_original_locator_and_invalid_manifest_is_atomic(draft):
    service, _, source = draft
    book = Document(source)
    sync_book(service, book, ['a','b','c'], 1)
    before = service.id_map(service.document_id)
    assert all(row['host_locator_id'] is None for row in before['rows'])
    with pytest.raises(GuardError, match='清单'):
        sync_book(service, book, ['a','a','c'], 1)
    assert service.id_map(service.document_id) == before


def test_locator_ignores_zip_container_time_but_rejects_unversioned_content(draft, monkeypatch):
    service, _, source = draft
    book = Document(source)
    sync_book(service, book, ['a','b','c'], 0, True)
    before = service.id_map(service.document_id)
    from office_agent.webui import paged_editor
    normalize = paged_editor.normalize

    def repacked(data, original=None):
        result = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(normalize(data, original))) as archive:
            with zipfile.ZipFile(result, 'w') as target:
                for item in archive.infolist():
                    target.writestr(zipfile.ZipInfo(item.filename, (1981,1,1,0,0,0)), archive.read(item.filename))
        return result.getvalue()

    monkeypatch.setattr(paged_editor, 'normalize', repacked)
    sync_book(service, book, ['a','b','c'], 0)
    assert service.id_map(service.document_id) == before
    book.paragraphs[1].runs[0].text = '未增加版本的替换'
    with pytest.raises(GuardError, match='同一草稿版本'):
        sync_book(service, book, ['a','b','c'], 0)
    assert service.id_map(service.document_id) == before


def test_changed_identity_invalidates_old_reference_without_restoring_approval(draft):
    service, _, source = draft
    book = Document(source)
    sync_book(service, book, ['a','b','c'], 0, True)
    old = service.reference(service.document_id, 0, {'node_id':'b','start':0,'end':2})['reference']
    service.pending_reference_id = old['id']
    sync_book(service, book, ['a','new-id','c'], 0)
    assert not service.references and service.pending_reference_id is None
    with pytest.raises(GuardError, match='过期'):
        service.reference_locator(service.document_id, 0, old['draft_locator_id'], 0, 2)


@pytest.fixture
def sheet_draft(tmp_path):
    source = fixtures(tmp_path)[0]
    app = WebviewApp()
    api = SimpleNamespace(_app=app, _auth=SimpleNamespace(status=lambda:{'authenticated':True}),
                          _memory_key=lambda:'self-created-range-test', _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating = FloatingSession(api)
    assert floating.admit(str(source))['ok']
    assert floating.read(floating.read_token, True)['ok']
    service = floating.canvas
    service.open(floating.document['id'], floating.document['version'])
    return service, floating, source


def prepare(service, text='125\n6', mode='numbers', start='B2'):
    return service.prepare_range(service.document_id, service.revision, 'sheet-0', start, text, mode)


def commit(service, result, confirmed=True):
    return service.commit_range(service.document_id, service.revision, result['plan']['id'], confirmed)


def test_range_preview_cancel_commit_group_undo_redo_and_formula_recalculation(sheet_draft):
    service, _, source = sheet_draft
    original = source.read_bytes(); before = service.data
    preview = prepare(service)
    assert service.data == before and service.revision == 0 and service.operations == []
    assert preview['changed_cells'] == 2 and preview['draft_only']
    assert commit(service, preview, False)['cancelled']
    assert service.data == before and not service.undo_groups
    result = commit(service, prepare(service))
    assert result['draft_revision'] == 1 and result['operation_count'] == 2 and result['undo_count'] == 1
    book = load_workbook(io.BytesIO(service.data))
    assert book['实验']['B2'].value == 125 and book['实验']['B3'].value == 6
    assert book['实验']['B2'].font.bold and book['实验']['B2'].border.bottom.style == 'thin'
    cells = {cell['id']:cell for cell in service.view(service.document_id, 1, 'sheet-0')['view']['cells']}
    assert cells['C2']['formula'] == 'B2*2' and cells['C2']['calculation']['value'] == '250'
    assert service.undo(service.document_id, 1)['redo_count'] == 1
    assert service.data == original and not service.operations
    result = service.redo(service.document_id, 2)
    assert result['draft_revision'] == 3 and result['operation_count'] == 2 and result['redo_count'] == 0
    assert load_workbook(io.BytesIO(service.data))['实验']['B3'].value == 6
    assert engine.verify(original, service.data, '.xlsx', service.operations)['non_target_parts_exact']
    assert source.read_bytes() == original


@pytest.mark.parametrize('start,text,mode', [('B2','125\t1','numbers'),('A4','x\ty','text'),('IV10000','x\ty','text'),('B2','1\t2\n3','text')])
def test_invalid_late_cell_or_shape_never_partially_commits(sheet_draft, start, text, mode):
    service, _, source = sheet_draft
    before = service.data
    with pytest.raises(GuardError):
        prepare(service, text, mode, start)
    assert service.data == before == source.read_bytes() and service.revision == 0
    assert not service.operations and not service.undo_groups and not service.redo_groups


def test_range_plan_identity_expiry_protection_and_repeated_commit(sheet_draft):
    service, floating, _ = sheet_draft
    first = prepare(service)
    second = prepare(service, '126\n7')
    with pytest.raises(GuardError, match='过期'):
        commit(service, first, False)
    assert service.range_plan['id'] == second['plan']['id']
    floating.protected_cells.add(('实验','B3'))
    with pytest.raises(GuardError, match='过期'):
        commit(service, second)
    assert service.revision == 0 and not service.operations
    floating.protected_cells.clear()
    expired = prepare(service)
    service.range_plan['deadline'] = time.monotonic()-1
    with pytest.raises(GuardError, match='过期'):
        commit(service, expired)
    valid = prepare(service)
    commit(service, valid)
    with pytest.raises(GuardError, match='过期'):
        commit(service, valid)
    assert service.revision == 1 and len(service.operations) == 2


def test_batch_protection_and_capacity_are_checked_before_commit(sheet_draft):
    service, floating, source = sheet_draft
    floating.protected_cells.add(('实验','B3'))
    with pytest.raises(GuardError, match='保护'):
        prepare(service)
    assert service.data == source.read_bytes() and service.revision == 0
    floating.protected_cells.clear()
    operations = [operation('D'+str(index), str(index), '') for index in range(1, 102)]
    with pytest.raises(GuardError, match='100'):
        service.apply_many(service.document_id, 0, operations)
    with pytest.raises(GuardError, match='重复'):
        service.apply_many(service.document_id, 0, [operation('D1','1',''),operation('D1','2','')])
    assert not service.operations and service.data == source.read_bytes()


def test_literal_paste_and_copy_do_not_evaluate_formula_or_infer_identifiers(sheet_draft):
    service, _, source = sheet_draft
    text = '=WEBSERVICE("https://invalid.example")\t0012\n@SUM(A1)\t-5'
    result = commit(service, prepare(service, text, 'text', 'D2'))
    assert result['group_operations'] == 4
    book = load_workbook(io.BytesIO(service.data))['实验']
    assert book['D2'].value.startswith('=WEBSERVICE') and book['D2'].data_type == 's'
    assert book['E2'].value == '0012' and book['E3'].value == '-5'
    copied = service.copy_range(service.document_id, 1, 'sheet-0', 'D2', 'E3')
    assert parse_tsv(copied['tsv']) == parse_tsv(text)
    formula = service.copy_range(service.document_id, 1, 'sheet-0', 'C2', 'C2')
    assert formula['tsv'] == '=B2*2\n'
    assert source.read_bytes() == service.original


def test_new_edit_clears_redo_and_protected_redo_is_refused(sheet_draft):
    service, floating, _ = sheet_draft
    commit(service, prepare(service))
    service.undo(service.document_id, 1)
    floating.protected_cells.add(('实验','B2'))
    with pytest.raises(GuardError, match='保护'):
        service.redo(service.document_id, 2)
    assert service.revision == 2 and len(service.redo_groups) == 1
    floating.protected_cells.clear()
    service.apply(service.document_id, 2, operation('D2','新编辑',''))
    assert not service.redo_groups and service.state()['undo_count'] == 1
    with pytest.raises(GuardError, match='重做'):
        service.redo(service.document_id, 3)


def test_redo_preserves_suggestion_attribution_but_not_old_approval(sheet_draft):
    service, floating, _ = sheet_draft
    commit(service, prepare(service))
    attribution = {'id':'own-suggestion', 'reason':'保留条件', 'model':{'model_id':'test-only'}}
    service.summaries[0]['ai'] = attribution
    service.undo(service.document_id, 1)
    service.plans['expired'] = {'id':'expired'}
    floating.app.editor.plans['expired'] = {'id':'expired'}
    service.redo(service.document_id, 2)
    assert service.summaries[0]['ai'] == attribution
    assert not service.plans and not floating.app.editor.plans and service.ai.reference is None


def test_numeric_paste_keeps_leading_zero_identifiers_and_empty_is_not_zero(sheet_draft):
    service, _, _ = sheet_draft
    commit(service, prepare(service, '0012\t1.25\n\t+5', 'numbers', 'D2'))
    cells = load_workbook(io.BytesIO(service.data))['实验']
    assert cells['D2'].value == '0012' and cells['D2'].data_type == 's'
    assert cells['E2'].value == 1.25 and cells['E2'].data_type == 'n'
    assert cells['D3'].value is None
    assert cells['E3'].value == 5


@pytest.mark.parametrize('text,expected', [('"第一行\n第二行"\t"a""b"\r\n', [['第一行\n第二行','a"b']]), ('\t\n',[['','']])])
def test_tsv_quotes_line_endings_and_empty_values(text, expected):
    assert parse_tsv(text) == expected


@pytest.mark.parametrize('text', ['x'*65537,'a\x00b','"unclosed','a\tb\nc','\n','\t'.join(['x']*101)],
                         ids=['too-large','nul','unclosed-quote','ragged','empty-row','too-many-cells'])
def test_tsv_limits_and_malformed_shape(text):
    with pytest.raises(GuardError):
        parse_tsv(text)


def test_unrelated_large_merge_does_not_block_small_range(sheet_draft):
    service, _, _ = sheet_draft
    raw = engine.load(service.data, '.xlsx')
    root = engine.xml(raw['xl/worksheets/sheet1.xml'])
    merge = root.find(engine.S+'mergeCells')
    from lxml import etree as ET
    ET.SubElement(merge, engine.S+'mergeCell', ref='Z200:XFD300')
    raw['xl/worksheets/sheet1.xml'] = engine.encoded(root)
    data = engine.pack(raw)
    assert read_range(data, 'sheet-0', 'B2', 'B3')['rows'] == 2
