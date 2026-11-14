import importlib.util
import io
import json
from pathlib import Path
import shutil

from lxml import etree as ET
from openpyxl import load_workbook
import pytest

from office_agent import canvas_panes as panes
from office_agent import office_canvas as engine
from office_agent.diagram_process import _run_owned
from office_agent.writes import GuardError
from office_agent.browser_runtime import find_edge
from tests.test_l3_l4_editing import sheet_draft


def prepare(service, rows=2, cols=3):
    return service.prepare_panes(service.document_id,service.revision,'sheet-0',rows,cols)


def commit(service, plan, confirmed=True):
    return service.commit_range(service.document_id,service.revision,plan['plan']['id'],confirmed)


def test_freeze_preview_cancel_reopen_navigation_and_undo_redo(sheet_draft):
    service,_,source=sheet_draft
    original=source.read_bytes()
    plan=prepare(service)
    assert plan['plan']['changes'][0]['scope']=='工作表视图' and service.revision==0 and service.data==original
    assert commit(service,plan,False)['cancelled'] and service.data==original
    adopted=commit(service,prepare(service))
    assert adopted['undo_count']==1 and adopted['operation_count']==1
    assert load_workbook(io.BytesIO(service.data))['实验'].freeze_panes=='D3'
    view=service.view(service.document_id,1,'sheet-0',45,8)['view']
    assert view['row_indices']==[1,2]+list(range(45,73))
    assert view['col_indices']==[1,2,3]+list(range(8,21))
    assert len(view['row_indices'])==30 and len(view['col_indices'])==16
    cells={cell['id']:cell for cell in view['cells']}
    assert cells['B2']['value']=='120' and cells['C2']['calculation']['value']=='240' and 'A45' in cells
    assert 'D2' not in cells and view['geometry']['rows']['2']['value']==25
    edge=service.view(service.document_id,1,'sheet-0',10000,256)['view']
    assert edge['row_indices']==[1,2,10000] and edge['col_indices']==[1,2,3,256]
    service.undo(service.document_id,1)
    assert service.data==original
    service.redo(service.document_id,2)
    assert load_workbook(io.BytesIO(service.data))['实验'].freeze_panes=='D3'
    commit(service,prepare(service,0,0))
    assert load_workbook(io.BytesIO(service.data))['实验'].freeze_panes is None
    assert source.read_bytes()==original


def test_freeze_keeps_cells_caches_styles_and_other_package_parts_exact(sheet_draft):
    service,_,_=sheet_draft
    original=service.data;raw=engine.load(original,'.xlsx');root=engine.xml(raw['xl/worksheets/sheet1.xml'])
    operation={'item_id':'sheet-0','target':'A1','action':panes.ACTION,'value':{'rows':1,'cols':0},'expected':panes.identity(root)}
    candidate,summary=engine.transform(original,'.xlsx',operation)
    after=engine.load(candidate,'.xlsx');current=engine.xml(after['xl/worksheets/sheet1.xml'])
    assert [name for name in raw if raw[name]!=after[name]]==['xl/worksheets/sheet1.xml']
    before_views=root.find(engine.S+'sheetViews');after_views=current.find(engine.S+'sheetViews')
    root.remove(before_views);current.remove(after_views)
    assert ET.tostring(root)==ET.tostring(current)
    assert summary['after']=='前1行 / 前0列'
    assert engine.verify(original,candidate,'.xlsx',[operation])['non_target_parts_exact']


@pytest.mark.parametrize('rows,cols', [(5,0),(0,4),(-1,0),(True,0),(1.5,0),('1',0),(None,1)])
def test_invalid_freeze_counts_do_not_change_draft(sheet_draft,rows,cols):
    service,_,_=sheet_draft
    before=service.data
    with pytest.raises(GuardError):prepare(service,rows,cols)
    assert service.data==before and service.revision==0


def test_merge_boundary_and_expired_or_superseded_plan_rejected(sheet_draft):
    service,_,_=sheet_draft
    with pytest.raises(GuardError,match='合并'):prepare(service,1,1)
    stale=prepare(service);service.range_plan['deadline']=0
    with pytest.raises(GuardError,match='过期'):commit(service,stale)
    old=prepare(service);new=prepare(service,1,0)
    with pytest.raises(GuardError,match='过期'):commit(service,old,False)
    assert service.range_plan['id']==new['plan']['id']
    commit(service,new)
    with pytest.raises(GuardError):commit(service,new)
    assert service.revision==1 and len(service.operations)==1


def test_protected_cells_are_not_modified_and_restore_replays_view_only(sheet_draft):
    service,floating,_=sheet_draft
    floating.protected_cells.add(('实验','A1'))
    commit(service,prepare(service))
    record,_=service.drafts.prepare(service.original,service.data,service.operations)
    replay,changes=service._replay_draft(record)
    assert engine.load(replay,'.xlsx')==engine.load(service.data,'.xlsx') and len(changes)==1
    plan=service.plan(service.document_id,1)['plan']
    assert plan['confirmation_required'] and plan['changes'][0]['action']=='freeze_panes'
    assert service.publish(service.document_id,1,plan['id'],False)['cancelled']


@pytest.mark.parametrize('state,split', [('split','1'),('frozenSplit','1'),('frozen','5')])
def test_unsupported_file_panes_report_unknown_and_stay_unchanged(state,split):
    root=ET.fromstring(('<worksheet xmlns="'+engine.S[1:-1]+'"><sheetViews><sheetView workbookViewId="0"><pane state="'+state+'" ySplit="'+split+'" topLeftCell="A6"/></sheetView></sheetViews><sheetData/></worksheet>').encode())
    original=ET.tostring(root)
    assert not panes.inspect(root)['supported']
    info,rows,cols=panes.window(root,45,8)
    assert not info['supported'] and rows==list(range(45,75)) and cols==list(range(8,24))
    with pytest.raises(GuardError):panes.apply(root,{'target':'A1','expected':panes.identity(root),'value':{'rows':0,'cols':0}})
    assert ET.tostring(root)==original


def test_sheet_protection_and_changed_view_identity_refuse_write(sheet_draft):
    service,_,_=sheet_draft
    raw=engine.load(service.data,'.xlsx');root=engine.xml(raw['xl/worksheets/sheet1.xml'])
    identity=panes.identity(root)
    ET.SubElement(root,engine.S+'sheetProtection')
    with pytest.raises(GuardError,match='保护'):panes.apply(root,{'target':'A1','expected':identity,'value':{'rows':1,'cols':0}})
    root.remove(root.find(engine.S+'sheetProtection'))
    root.find('./'+engine.S+'sheetViews/'+engine.S+'sheetView').set('zoomScale','85')
    with pytest.raises(GuardError,match='前置条件'):panes.apply(root,{'target':'A1','expected':identity,'value':{'rows':1,'cols':0}})


def test_missing_workbook_view_cannot_create_dangling_pane_reference(sheet_draft):
    service,_,_=sheet_draft;raw=engine.load(service.data,'.xlsx')
    book=engine.xml(raw['xl/workbook.xml']);book.remove(book.find(engine.S+'bookViews'));raw['xl/workbook.xml']=engine.encoded(book)
    sheet=engine.xml(raw['xl/worksheets/sheet1.xml'])
    operation={'item_id':'sheet-0','target':'A1','action':panes.ACTION,'value':{'rows':1,'cols':0},'expected':panes.identity(sheet)}
    with pytest.raises(GuardError,match='窗口身份'):engine.transform(engine.pack(raw),'.xlsx',operation)


@pytest.mark.skipif(shutil.which('node') is None,reason='Existing Node required for bounded table transaction code regression')
def test_table_column_policy_does_not_mutate_engine_json(tmp_path):
    script=Path(__file__).with_name('probes')/'table_column_layout.cjs'
    output=tmp_path/'table.json'
    result=_run_owned((str(Path(shutil.which('node')).resolve()),str(script.resolve()),str(output)),timeout_ms=15000)
    payload=json.loads(output.read_text('utf8')) if output.exists() else {}
    assert result.exit_code==0 and result.empty_confirmed,payload
    assert payload['passed']


@pytest.mark.skipif(find_edge() is None,reason='Requires existing Edge; no runtime installed')
def test_fixed_engine_column_export_reopen_code_probe():
    path=Path(__file__).resolve().parents[1]/'deliverables/l3-l4-panes-20261111/engine_probe.py'
    spec=importlib.util.spec_from_file_location('column_engine_probe',path);module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    result=module.run_probe()
    assert result['exit_code']==0 and result['empty_confirmed'],(path.parent/'ENGINE_ERROR.txt').read_text('utf8')
    evidence=json.loads((path.parent/'ENGINE_PROBE.json').read_text('utf8'))
    assert evidence['proof']['width_pt']==120 and evidence['texts_preserved'] and evidence['header_preserved'] and evidence['source_unchanged']
    assert evidence['iframe_cancel_unchanged'] and evidence['preceding_manual_draft_preserved'] and evidence['iframe_code_checks']['redo']['revision']==4


@pytest.mark.skipif(shutil.which('node') is None,reason='Existing Node required for bounded panes DOM code regression')
def test_pinned_grid_navigation_and_late_response_code(tmp_path):
    script=Path(__file__).with_name('probes')/'canvas_panes.cjs';output=tmp_path/'panes.json'
    result=_run_owned((str(Path(shutil.which('node')).resolve()),str(script.resolve()),str(output)),timeout_ms=15000)
    payload=json.loads(output.read_text('utf8')) if output.exists() else {}
    assert result.exit_code==0 and result.empty_confirmed,payload
    assert payload['passed']
