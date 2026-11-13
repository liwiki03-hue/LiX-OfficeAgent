import io
import json
from pathlib import Path
import shutil

from lxml import etree as ET
from openpyxl import load_workbook
import pytest

from office_agent import canvas_geometry as geometry
from office_agent import office_canvas as engine
from office_agent.diagram_process import _run_owned
from office_agent.writes import GuardError
from tests.test_l3_l4_editing import sheet_draft


def prepare(service, action='row_height', value=36, target='B2'):
    return service.prepare_geometry(service.document_id, service.revision, 'sheet-0', target, action, value)


def commit(service, plan, confirmed=True):
    return service.commit_range(service.document_id, service.revision, plan['plan']['id'], confirmed)


def test_preview_cancel_adopt_undo_redo_and_exact_non_target_parts(sheet_draft):
    service, _, source = sheet_draft
    original = source.read_bytes()
    snapshot = service.view(service.document_id, 0, 'sheet-0')['view']['geometry']
    assert snapshot['rows']['2']['value'] == 25 and snapshot['columns']['2']['value'] == 24
    assert snapshot['rows']['1']['source'] == 'automatic'
    preview = prepare(service)
    assert service.data == original and service.revision == 0
    assert '整行' in preview['plan']['changes'][0]['scope']
    assert commit(service, preview, False)['cancelled'] and service.data == original
    result = commit(service, prepare(service))
    assert result['draft_revision'] == 1 and result['undo_count'] == 1
    book = load_workbook(io.BytesIO(service.data))
    assert book['实验'].row_dimensions[2].height == 36
    assert book['实验']['B2'].value == 120 and book['实验']['C2'].value == '=B2*2'
    assert book['实验']['B2'].font.bold and book['实验']['B2'].border.bottom.style == 'thin'
    before, after = engine.load(original, '.xlsx'), engine.load(service.data, '.xlsx')
    assert [name for name in before if before[name] != after[name]] == ['xl/worksheets/sheet1.xml']
    root_before, root_after = engine.xml(before['xl/worksheets/sheet1.xml']), engine.xml(after['xl/worksheets/sheet1.xml'])
    assert [ET.tostring(node) for node in root_before.findall('.//'+engine.S+'c')] == [ET.tostring(node) for node in root_after.findall('.//'+engine.S+'c')]
    assert engine.verify(original, service.data, '.xlsx', service.operations)['non_target_parts_exact']
    service.undo(service.document_id, 1)
    assert service.data == original
    service.redo(service.document_id, 2)
    assert load_workbook(io.BytesIO(service.data))['实验'].row_dimensions[2].height == 36
    assert source.read_bytes() == original


def test_column_range_split_preserves_neighbors_and_formula_cache(sheet_draft):
    service, _, _ = sheet_draft
    raw = engine.load(service.data, '.xlsx'); root = engine.xml(raw['xl/worksheets/sheet1.xml'])
    columns = root.find(engine.S+'cols'); columns[0].set('max', '4');columns[0].set('style', '1')
    formula = root.find('.//'+engine.S+'c[@r="C2"]');formula.find(engine.S+'v').text = '240'
    raw['xl/worksheets/sheet1.xml'] = engine.encoded(root)
    original = engine.pack(raw)
    op = {'item_id':'sheet-0','target':'C2','action':'column_width','value':30,
          'expected':geometry.snapshot(root, 'column_width', 2, 3)}
    candidate, summary = engine.transform(original, '.xlsx', op)
    parts = engine.load(candidate, '.xlsx'); current = engine.xml(parts['xl/worksheets/sheet1.xml'])
    dims = [dict(node.attrib) for node in current.find(engine.S+'cols')]
    assert [(node['min'], node['max'], node['width']) for node in dims] == [('2','2','24'),('3','3','30'),('4','4','24')]
    assert all(node['style'] == '1' for node in dims)
    assert current.find('.//'+engine.S+'c[@r="C2"]/'+engine.S+'v').text == '240'
    assert parts['xl/workbook.xml'] == raw['xl/workbook.xml'] and summary['unit'] == '字符宽度'
    assert engine.verify(original, candidate, '.xlsx', [op])['operations'] == 1


@pytest.mark.parametrize('action,value', [('row_height',5),('row_height',201),('column_width',1),('column_width',81),('row_height',True),('row_height','NaN'),('column_width',float('inf')),('row_height',{'height':20})])
def test_invalid_sizes_never_commit(sheet_draft, action, value):
    service, _, source = sheet_draft
    original = source.read_bytes()
    with pytest.raises(GuardError):
        prepare(service, action, value)
    assert service.data == original and service.revision == 0 and not service.operations


def test_protected_axis_covers_cells_outside_view_and_late_changes(sheet_draft):
    service, floating, _ = sheet_draft
    floating.protected_cells.add(('实验','XFD2'))
    with pytest.raises(GuardError, match='保护'):
        prepare(service)
    floating.protected_cells.clear()
    preview = prepare(service, 'column_width', 30)
    floating.protected_cells.add(('实验','B1048576'))
    with pytest.raises(GuardError, match='过期'):
        commit(service, preview)
    with pytest.raises(GuardError, match='保护'):
        prepare(service, 'column_width', 30)
    assert service.revision == 0


def test_expiry_duplicate_and_implementation_change_refuse_old_plan(sheet_draft, monkeypatch):
    service, _, _ = sheet_draft
    preview = prepare(service);service.range_plan['deadline'] = 0
    with pytest.raises(GuardError, match='过期'):
        commit(service, preview)
    preview = prepare(service)
    from office_agent.webui import office_canvas as host
    identity = host.canvas_implementation
    monkeypatch.setattr(host, 'canvas_implementation', lambda:['changed'])
    with pytest.raises(GuardError, match='过期'):
        commit(service, preview)
    monkeypatch.setattr(host, 'canvas_implementation', identity)
    commit(service, prepare(service))
    with pytest.raises(GuardError):
        commit(service, preview)
    assert service.revision == 1 and len(service.operations) == 1


def test_reset_and_new_row_column_keep_content_and_restore(sheet_draft):
    service, _, source = sheet_draft
    commit(service, prepare(service, 'row_height', None))
    assert load_workbook(io.BytesIO(service.data))['实验'].row_dimensions[2].height is None
    commit(service, prepare(service, 'column_width', 32, 'D100'))
    commit(service, prepare(service, 'row_height', 40, 'D100'))
    book = load_workbook(io.BytesIO(service.data))
    assert book['实验'].row_dimensions[100].height == 40 and book['实验'].column_dimensions['D'].width == 32
    assert book['实验']['A45'].value == load_workbook(source)['实验']['A45'].value
    record, _ = service.drafts.prepare(service.original, service.data, service.operations)
    data, changes = service._replay_draft(record)
    assert engine.load(data, '.xlsx') == engine.load(service.data, '.xlsx') and len(changes) == 3


@pytest.mark.parametrize('attribute,value', [('hidden','1'),('outlineLevel','1'),('collapsed','1'),('bestFit','1')])
def test_hidden_grouped_or_autofit_dimension_stays_readonly(attribute, value):
    root = ET.fromstring(('<worksheet xmlns="'+engine.S[1:-1]+'"><cols><col min="1" max="2" width="24" '+attribute+'="'+value+'"/></cols><sheetData/></worksheet>').encode())
    assert geometry.describe(root, 'column_width', 1, 2)['editable'] is False
    with pytest.raises(GuardError):
        geometry.snapshot(root, 'column_width', 1, 2)


def test_overlapping_dimensions_and_changed_expected_fail_closed(sheet_draft):
    service, _, _ = sheet_draft
    raw = engine.load(service.data, '.xlsx'); root = engine.xml(raw['xl/worksheets/sheet1.xml'])
    columns = root.find(engine.S+'cols')
    ET.SubElement(columns, engine.S+'col', min='2', max='3', width='12')
    assert geometry.describe(root, 'column_width', 2, 2)['editable'] is False
    op = {'item_id':'sheet-0','target':'B2','action':'row_height','value':36,'expected':None}
    with pytest.raises(GuardError, match='前置条件'):
        engine.transform(service.data, '.xlsx', op)


@pytest.mark.parametrize('identity', ['02','0','1048577'])
def test_invalid_row_identity_cannot_create_duplicate_or_out_of_range_row(identity):
    root = ET.fromstring(('<worksheet xmlns="'+engine.S[1:-1]+'"><sheetData><row r="'+identity+'"/></sheetData></worksheet>').encode())
    with pytest.raises(GuardError, match='行顺序'):
        geometry.apply(root, {'action':'row_height','target':'B2','value':36,'expected':None}, 2, 2)


def test_external_source_change_rejects_prepared_geometry(sheet_draft):
    service, _, source = sheet_draft
    preview = prepare(service)
    book = load_workbook(source);book['实验']['D2'] = '外部修改';book.save(source)
    with pytest.raises(ValueError, match='外部改变'):
        commit(service, preview)
    assert service.revision == 0 and not service.operations
    assert service.floating.document['ready'] is False


def test_axis_protection_blocks_undo_restoration_and_final_plan(sheet_draft):
    service, floating, _ = sheet_draft
    commit(service, prepare(service))
    record, _ = service.drafts.prepare(service.original, service.data, service.operations)
    floating.protected_cells.add(('实验','IV2'))
    with pytest.raises(GuardError, match='保护'):
        service.undo(service.document_id, 1)
    with pytest.raises(GuardError, match='保护'):
        service._replay_draft(record)
    with pytest.raises(GuardError, match='保护'):
        service.plan(service.document_id, 1)
    assert service.revision == 1 and len(service.operations) == 1


def test_copy_plan_includes_axis_and_refuses_changed_implementation(sheet_draft, monkeypatch):
    service, _, source = sheet_draft
    original = source.read_bytes()
    commit(service, prepare(service, 'column_width', 30))
    plan = service.plan(service.document_id, 1)['plan']
    assert plan['confirmation_required'] and '整列' in plan['changes'][0]['scope']
    assert service.publish(service.document_id, 1, plan['id'], False)['cancelled']
    plan = service.plan(service.document_id, 1)['plan']
    from office_agent.webui import office_canvas as host
    monkeypatch.setattr(host, 'canvas_implementation', lambda:['changed'])
    with pytest.raises(GuardError, match='实现已改变'):
        service.publish(service.document_id, 1, plan['id'], True)
    assert source.read_bytes() == original and not service.publishing


@pytest.mark.skipif(shutil.which('node') is None, reason='Existing Node required for bounded layout code regression')
def test_paragraph_layout_transaction_code(tmp_path):
    script = Path(__file__).with_name('probes') / 'paragraph_layout.cjs'
    output = tmp_path / 'layout.json'
    result = _run_owned((str(Path(shutil.which('node')).resolve()), str(script.resolve()), str(output)), timeout_ms=15000)
    payload = json.loads(output.read_text(encoding='utf8')) if output.exists() else {}
    assert result.code == 'process_completed' and result.exit_code == 0 and result.empty_confirmed, payload
    assert payload['passed']


@pytest.mark.skipif(shutil.which('node') is None, reason='Existing Node required for bounded geometry DOM code regression')
def test_geometry_dom_request_isolation(tmp_path):
    script = Path(__file__).with_name('probes') / 'canvas_geometry.cjs'
    output = tmp_path / 'geometry.json'
    result = _run_owned((str(Path(shutil.which('node')).resolve()), str(script.resolve()), str(output)), timeout_ms=15000)
    payload = json.loads(output.read_text(encoding='utf8')) if output.exists() else {}
    assert result.code == 'process_completed' and result.exit_code == 0 and result.empty_confirmed, payload
    assert payload['passed']
