from copy import deepcopy
from decimal import Decimal
from io import BytesIO
import json
from pathlib import Path
import zipfile
from xml.etree.ElementTree import fromstring,tostring

from pptx import Presentation
import pytest

from office_agent.presentation_drafts import build_deck,validate_deck,PresentationError,identifier,make_outline
from office_agent.presentation_charts import chart_values
from office_agent.fidelity import package_projection,FidelityError
from office_agent.webui.presentation_workspace import PresentationWorkspace
from tests.test_presentation_workspace import workspace,material,fake_ready
from tests.test_paged_editor import draft,sync


def chart_spec():
    spec=make_outline({'segments':[{'id':identifier(),'text':'自建测量数据'}]},'实验报告',2)
    spec['slides'][1].update(layout='chart',points=[],rows=[['方案','耗时'],['基线','120'],['优化','105.5'],['复核','98.25']],chart_unit='ms',chart_reviewed=True)
    return spec


def rewrite(raw,changes):
    output=BytesIO()
    with zipfile.ZipFile(BytesIO(raw)) as source,zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():target.writestr(name,changes.get(name,source.read(name)))
        for name,data in changes.items():
            if name not in source.namelist():target.writestr(name,data)
    return output.getvalue()


@pytest.mark.parametrize('value',['','120 ms','31%','=1+1','NaN','1e2','1,200','0.1234567'])
def test_ambiguous_chart_numbers_never_become_zero_or_formula(value):
    with pytest.raises(PresentationError):chart_values([['条件','数值'],['样本',value]])


def test_chart_decimal_rules_preserve_sign_and_refuse_ambiguous_categories():
    assert chart_values([['组别','变化'],['降低','-12.25'],['不变','0'],['升高','0.125']])==[Decimal('-12.25'),Decimal('0'),Decimal('.125')]
    for rows in [[['A','B'],['同名','1'],['同名','2']],[['=A1','B'],['值','1']],[['A','B','C'],['值','1','2']]]:
        with pytest.raises(PresentationError):chart_values(rows)


def test_chart_native_cache_workbook_units_and_notes_match_approved_spec(tmp_path):
    spec=chart_spec();raw=build_deck(spec);checked=validate_deck(raw,spec)
    book=Presentation(BytesIO(raw));chart=next(shape.chart for shape in book.slides[1].shapes if shape.has_chart)
    assert checked['editable_charts']==1 and list(chart.series[0].values)==[120,105.5,98.25]
    assert '图表单位：ms' in book.slides[1].notes_slide.notes_text_frame.text
    assert any('单位：ms' in shape.text for shape in book.slides[1].shapes if shape.has_text_frame)
    path=tmp_path/'native-chart.pptx';path.write_bytes(raw)
    with pytest.raises(FidelityError):package_projection(path,for_render=True)


@pytest.mark.parametrize('change',['cache','workbook','formula','external','orphan','macro','extra_relationship','axis_title','axis_scale'])
def test_chart_tampering_rejected_before_renderer_or_publish(change):
    spec=chart_spec();raw=build_deck(spec)
    with zipfile.ZipFile(BytesIO(raw)) as archive:
        chart_name=next(name for name in archive.namelist() if name.startswith('ppt/charts/chart') and name.endswith('.xml'))
        embedded=next(name for name in archive.namelist() if name.startswith('ppt/embeddings/'))
        changes={}
        if change=='cache':
            node=fromstring(archive.read(chart_name));c='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
            node.find('.//'+c+'val/'+c+'numRef/'+c+'numCache/'+c+'pt/'+c+'v').text='999';changes[chart_name]=tostring(node)
        elif change in {'axis_title','axis_scale'}:
            node=fromstring(archive.read(chart_name));c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';a='{http://schemas.openxmlformats.org/drawingml/2006/main}'
            if change=='axis_title':node.find('.//'+c+'valAx/'+c+'title//'+a+'t').text='其他指标'
            else:node.find('.//'+c+'valAx/'+c+'scaling/'+c+'min').set('val','90')
            changes[chart_name]=tostring(node)
        elif change in {'workbook','formula'}:
            nested=archive.read(embedded)
            with zipfile.ZipFile(BytesIO(nested)) as inner:
                node=fromstring(inner.read('xl/worksheets/sheet1.xml'));ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                cell=next(row for row in node.iter(ns+'c') if row.get('r')=='B2')
                if change=='workbook':cell.find(ns+'v').text='999'
                else:
                    from xml.etree.ElementTree import SubElement
                    SubElement(cell,ns+'f').text='1+1'
                changes[embedded]=rewrite(nested,{'xl/worksheets/sheet1.xml':tostring(node)})
        elif change=='external':
            name='ppt/charts/_rels/'+Path(chart_name).name+'.rels';node=fromstring(archive.read(name));list(node)[0].set('TargetMode','External');list(node)[0].set('Target','https://example.invalid/private.xlsx');changes[name]=tostring(node)
        elif change=='orphan':changes['ppt/embeddings/extra.xlsx']=archive.read(embedded)
        elif change=='extra_relationship':
            name='ppt/charts/_rels/'+Path(chart_name).name+'.rels';node=fromstring(archive.read(name));copy=deepcopy(list(node)[0]);copy.set('Id','rId999');node.append(copy);changes[name]=tostring(node)
        else:changes['ppt/vbaProject.bin']=b'not allowed'
    with pytest.raises(PresentationError):validate_deck(rewrite(raw,changes),spec)


def test_chart_unit_review_required_and_data_edit_invalidates_review(workspace):
    service,api=workspace;spec=chart_spec();service.spec=spec
    for changes in [{'chart_unit':''},{'chart_reviewed':False}]:
        candidate=deepcopy(spec);candidate['slides'][1].update(changes)
        with pytest.raises(PresentationError,match='核对'):build_deck(candidate)
    service.edit(spec['id'],1,'update',spec['slides'][1]['id'],{'rows':[['方案','耗时'],['基线','119']]})
    assert service.spec['slides'][1]['chart_reviewed'] is False
    with pytest.raises(PresentationError,match='核对'):build_deck(service.spec)


def selection(prepared,ids=None,**changes):
    return {'material_id':prepared['material_id'],'selected_ids':ids if ids is not None else [row['id'] for row in prepared['source']['segments']],
            'title':'选定材料','pages':2,'replace_id':prepared['replace_id'],'replace_revision':prepared['replace_revision'],**changes}


def test_material_subset_preserves_host_order_and_explicit_coverage(workspace):
    service,api=workspace
    prepared=service.materials(material(text='第一段。\n第二段。\n第三段。'))
    rows=prepared['source']['segments'];result=service.create(selection(prepared,[rows[2]['id'],rows[0]['id']],pages=3))
    assert [row['text'] for row in result['source']['segments']]==['第一段。','第三段。']
    assert result['source']['unselected_blocks']==[2] and result['source']['limited'] and result['source']['included_blocks']==2
    with pytest.raises(PresentationError,match='过期'):service.create(selection(prepared))


@pytest.mark.parametrize('kind',['empty','foreign','duplicate','expired','newer_outline','privacy','new_capture'])
def test_material_selection_requires_current_identity_and_scope(workspace,kind):
    service,api=workspace;prepared=service.materials(material());request=selection(prepared)
    if kind=='empty':request['selected_ids']=[]
    elif kind=='foreign':request['selected_ids']=[identifier()]
    elif kind=='duplicate':request['selected_ids']*=2
    elif kind=='expired':service.materials_pending['created']-=301
    elif kind=='newer_outline':service.edit(service.spec['id'],1,'update',service.spec['slides'][1]['id'],{'title':'后续人工编辑'})
    elif kind=='privacy':service.invalidate()
    else:service.materials(material())
    before=deepcopy(service.spec)
    with pytest.raises(PresentationError):service.create(request)
    assert service.spec==before


def test_material_capture_source_changes_require_read_again(draft):
    paged,floating,path=draft;sync(paged);floating.api._floating=floating;service=PresentationWorkspace(floating.api)
    prepared=service.materials(material(kind='draft',text='',document_id=paged.document_id,draft_revision=1))
    paged.changed(paged.document_id,2)
    with pytest.raises(PresentationError,match='来源草稿'):service.create(selection(prepared))
    path.write_bytes(b'external change')
    with pytest.raises(ValueError):service.create(selection(prepared))


def test_chart_publish_uses_same_confirmation_and_modified_output_guard(workspace,tmp_path,monkeypatch):
    service,api=workspace;service.spec=chart_spec();fake_ready(service,monkeypatch)
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan']
    assert not Path(plan['path']).exists()
    assert service.confirm(plan['id'],False)['result']['wrote_nothing']
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan'];receipt=service.confirm(plan['id'],True)
    path=Path(receipt['result']['path']);assert validate_deck(path.read_bytes(),service.spec)['editable_charts']==1
    path.write_bytes(b'external file replacement')
    assert not api.presentation_saved_action(plan['id'],'rollback_plan')['ok'] and path.read_bytes()==b'external file replacement'
