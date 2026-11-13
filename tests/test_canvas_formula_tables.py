from copy import deepcopy
from decimal import Decimal
import io

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
import pytest

from office_agent.canvas_formulas import Calculator, parse
from office_agent.office_canvas import A, P, S, encoded, load, pack, transform, verify, view, xml
from office_agent.writes import GuardError
from tests.test_office_canvas import canvas, fixtures, operation


def workbook(tmp_path):
    source, _ = fixtures(tmp_path)
    return source.read_bytes()


@pytest.mark.parametrize('formula,result', [('=B2*3','360'), ('=SUM(B2,30)/2','75'), ('=AVERAGE(B2,100)','110'), ('=-B2+MIN(20,30)*2','-80'), ('=MAX($B$2,1.25E2)','125')])
def test_formula_exact_numeric_preview_and_persistent_expression(tmp_path, formula, result):
    data = workbook(tmp_path)
    op = operation('C2', formula, '=B2*2', action='formula')
    changed, summary = transform(data, '.xlsx', op)
    cell = next(cell for cell in view(changed, '.xlsx')['cells'] if cell['id']=='C2')
    assert Decimal(cell['calculation']['value'])==Decimal(result)
    assert cell['calculation']['precision']==28 and cell['calculation']['rounding']=='ROUND_HALF_EVEN'
    assert cell['calculation']['depends_on']==['B2']
    assert summary['calculation']==cell['calculation']
    assert cell['editable'] is False and cell['formula_editable'] is True
    assert load_workbook(io.BytesIO(changed))['实验']['C2'].value==formula
    assert load_workbook(io.BytesIO(changed),data_only=True)['实验']['C2'].value is None
    assert verify(data, changed, '.xlsx', [op])['non_target_parts_exact']


@pytest.mark.parametrize('formula', ['=WEBSERVICE("https://example.invalid")','=Sheet2!A1','=B2/0','=D99+1','=SUM(B1:B2)','=SUM(A1:IV10000)','=C2+1','=1**2','=__import__(1)','=B2;1'])
def test_formula_refusal_never_changes_original(tmp_path, formula):
    data = workbook(tmp_path)
    with pytest.raises(GuardError):
        transform(data, '.xlsx', operation('C2', formula, '=B2*2', action='formula'))
    assert load_workbook(io.BytesIO(data))['实验']['C2'].value=='=B2*2'


def test_formula_dependency_recalculation_range_and_cycle(tmp_path):
    data = workbook(tmp_path)
    first = operation('D2', '=SUM(B2:C2)', '', action='formula')
    data, _ = transform(data, '.xlsx', first)
    data, _ = transform(data, '.xlsx', operation('B2', '125', '120', action='number'))
    cells = {cell['id']:cell for cell in view(data, '.xlsx')['cells']}
    assert Decimal(cells['D2']['calculation']['value'])==375
    assert cells['D2']['value']=='' and cells['C2']['value']==''
    with pytest.raises(GuardError, match='循环'):
        transform(data, '.xlsx', operation('C2', '=D2+1', '=B2*2', action='formula'))
    data, _ = transform(data, '.xlsx', operation('D3', '=1/3', '', action='formula'))
    assert next(cell for cell in view(data, '.xlsx')['cells'] if cell['id']=='D3')['calculation']['value']=='0.3333333333333333333333333333'
    sheet=xml(load(data,'.xlsx')['xl/worksheets/sheet1.xml']);calculator=Calculator(sheet)
    graph=calculator.graph_report()
    assert graph['formulas']>=2 and graph['limited'] is False
    assert graph['edges']['D2']==['B2','C2']
    results=calculator.results({'D2','D3'})
    assert set(results)=={'D2','D3'} and all(row['status']=='calculated' for row in results.values())
    assert view(data,'.xlsx')['formula_graph']['topological']==graph['topological']


@pytest.mark.parametrize('canvas',['.xlsx'],indirect=True)
def test_plan_formula_preview_uses_final_draft_and_rejects_invalidated_dependency(canvas):
    service, floating, source, result, op = canvas; doc=floating.document['id']
    service.apply(doc,0,operation('C2','=SUM(B2,30)','=B2*2',action='formula'))
    old=service.plan(doc,1)['plan']
    service.apply(doc,1,operation('B2','125','120',action='number'))
    current=service.plan(doc,2)['plan']
    assert Decimal(current['changes'][0]['calculation']['value'])==155
    with pytest.raises(GuardError):service.publish(doc,2,old['id'],True)
    service.apply(doc,2,operation('B2','待核对','125'))
    with pytest.raises(GuardError,match='空白|文本|缺失'):service.plan(doc,3)
    assert not list(source.parent.glob('*.agent-*'))
    service.undo(doc,3)
    service.apply(doc,4,operation('C2','=B2*4','=SUM(B2,30)',action='formula'))
    final=service.plan(doc,5)['plan']
    assert final['changes'][0]['final_formula']=='=B2*4'
    assert final['changes'][0]['calculation']['value']=='500'
    assert final['changes'][0]['calculation_at_adoption']['value']=='150'


@pytest.mark.parametrize('canvas',['.xlsx'],indirect=True)
def test_service_view_reuses_calculation_cache_until_edit(canvas):
    service,floating,source,result,op=canvas;doc=floating.document['id']
    first=service.view(doc,0,'sheet-0',1,1)
    assert first['view']['formula_graph']['formulas']>=1 and service.calculation_cache
    cache=service.calculation_cache
    second=service.view(doc,0,'sheet-0',1,1)
    assert service.calculation_cache is cache and second['view']['cells']
    service.apply(doc,0,operation('A1','标题','温度'))
    assert service.calculation_cache and service.calculation_revision==1
    service.apply(doc,1,operation('B2','125','120',action='number'))
    assert service.calculation_cache=={} and service.calculation_revision==2


@pytest.mark.parametrize('formula,expected', [
    ('=1+2*3', '7'), ('=(1+2)*3', '9'), ('=10/4', '2.5'), ('=-2+5', '3'),
    ('=SUM(1,2,3)', '6'), ('=AVERAGE(2,4,6)', '4'), ('=MIN(3,1,2)', '1'), ('=MAX(3,1,2)', '3'),
    ('=0.1+0.2', '0.3'), ('=SUM(B2:C2)', '360'), ('=1/3', '0.3333333333333333333333333333'),
    ('=ABS(-2.5)', '2.5'), ('=ROUND(2.345,2)', '2.35'), ('=ROUND(2.5,0)', '3'),
    ('=COUNT(1,2,3)', '3'), ('=PRODUCT(2,3,4)', '24'),
    ('=INT(-1.2)', '-2'), ('=MOD(7,3)', '1'), ('=POWER(2,10)', '1024'), ('=SQRT(9)', '3'),
    ('=ROUNDUP(2.341,2)', '2.35'), ('=ROUNDDOWN(2.349,2)', '2.34'),
    ('=DATE(2025,1,1)', '45658'), ('=YEAR(45658)', '2025'), ('=MONTH(45658)', '1'), ('=DAY(45658)', '1'),
    ('=IF(1>0,10,20)', '10'), ('=IF(0,10,20)', '20'), ('=AND(1,0)', '0'), ('=OR(1,0)', '1'), ('=NOT(0)', '1'),
    ('=1=1', '1'), ('=1<>2', '1'), ('=MEDIAN(1,3,2)', '2'), ('=STDEV(1,2,3)', '1'),
    ('=STDEVP(1,1,1)', '0'), ('=SUMPRODUCT(B2:C2,B2:C2)', '72000'),
])
def test_bounded_formula_matrix(tmp_path,formula,expected):
    data=workbook(tmp_path);changed,_=transform(data,'.xlsx',operation('D2',formula,'',action='formula'))
    cell=next(cell for cell in view(changed,'.xlsx')['cells'] if cell['id']=='D2')
    assert cell['calculation']['status']=='calculated'
    assert Decimal(cell['calculation']['value'])==Decimal(expected)


def test_date_formula_1904_serial_and_components(tmp_path):
    data=workbook(tmp_path);sheet=xml(load(data,'.xlsx')['xl/worksheets/sheet1.xml'])
    calculator=Calculator(sheet,date1904=True)
    serial=calculator.describe('D2','=DATE(2025,1,1)')
    assert serial['status']=='calculated' and serial['value']=='44196'
    assert calculator.describe('D2','=YEAR(44196)')['value']=='2025'
    assert calculator.describe('D2','=MONTH(44196)')['value']=='1'
    assert calculator.describe('D2','=DAY(44196)')['value']=='1'
    assert calculator.describe('D2','=DATE(1900,2,29)')['status']=='unavailable'


def test_shared_formula_family_and_unknown_formula_remain_readonly(tmp_path):
    data = workbook(tmp_path); raw = load(data, '.xlsx'); sheet = xml(raw['xl/worksheets/sheet1.xml'])
    formula = sheet.find('.//'+S+'f'); formula.set('t', 'array'); formula.set('ref', 'C2:D3')
    raw['xl/worksheets/sheet1.xml']=encoded(sheet); special=pack(raw)
    cells={cell['id']:cell for cell in view(special,'.xlsx')['cells']}
    assert not cells['C2']['formula_editable'] and not cells['B2']['formula_editable']
    with pytest.raises(GuardError):transform(special,'.xlsx',operation('D3','=1+2','',action='formula'))
    formula.attrib.clear();formula.text='RAND()';raw['xl/worksheets/sheet1.xml']=encoded(sheet)
    cell=next(cell for cell in view(pack(raw),'.xlsx')['cells'] if cell['id']=='C2')
    assert cell['calculation']['status']=='unavailable' and cell['calculation']['code']=='unsupported_formula' and not cell['formula_editable']
    with pytest.raises(GuardError):parse('='+'('*20+'1'+')'*20)


def test_formula_error_code_division(tmp_path):
    data=workbook(tmp_path);raw=load(data,'.xlsx');sheet=xml(raw['xl/worksheets/sheet1.xml'])
    sheet.find('.//'+S+'c[@r="C2"]/'+S+'f').text='B2/0';raw['xl/worksheets/sheet1.xml']=encoded(sheet)
    cell=next(cell for cell in view(pack(raw),'.xlsx')['cells'] if cell['id']=='C2')
    assert cell['calculation']['status']=='unavailable' and cell['calculation']['code']=='division'


def table_fixture(tmp_path):
    _, source = fixtures(tmp_path); deck = Presentation(source)
    table = deck.slides[0].shapes.add_table(3,2,Inches(1), Inches(4.3), Inches(7), Inches(1.5)).table
    for row, values in enumerate([['指标','结果'],['耗时','120 ms'],['条件','同一设备']]):
        for col, value in enumerate(values):
            table.cell(row,col).text=value
    frame=table.cell(1,1).text_frame
    frame.paragraphs[0].runs[0].text='120'
    frame.paragraphs[0].runs[0].font.bold=True
    frame.paragraphs[0].runs[0].font.size=Pt(18)
    frame.paragraphs[0].add_run().text=' ms'
    source=tmp_path/'表格汇报.pptx';deck.save(source)
    return source


def test_ppt_table_run_edit_preserves_grid_style_and_other_runs(tmp_path):
    source = table_fixture(tmp_path); data=source.read_bytes(); before=load(data,'.pptx')
    target=next(run for run in view(data,'.pptx')['runs'] if run['text']=='120')
    assert target['box'] and target['location']=='表格1 · 第2行第2列'
    ops=[operation(target['id'],'115','120','slide-0'),operation(target['id'],{'italic':True},'115','slide-0','format')]
    changed=data
    for op in ops:changed,_=transform(changed,'.pptx',op)
    after=load(changed,'.pptx');assert {key for key in before if before[key]!=after[key]}=={'ppt/slides/slide1.xml'}
    old_table=xml(before['ppt/slides/slide1.xml']).find('.//'+A+'tbl')
    new_table=xml(after['ppt/slides/slide1.xml']).find('.//'+A+'tbl')
    assert encoded(old_table.find(A+'tblGrid'))==encoded(new_table.find(A+'tblGrid'))
    assert [row.get('h') for row in old_table.findall(A+'tr')]==[row.get('h') for row in new_table.findall(A+'tr')]
    assert [encoded(cell.find(A+'tcPr')) for cell in old_table.findall('.//'+A+'tc')]==[encoded(cell.find(A+'tcPr')) for cell in new_table.findall('.//'+A+'tc')]
    table=Presentation(io.BytesIO(changed)).slides[0].shapes[-1].table
    run=table.cell(1,1).text_frame.paragraphs[0].runs[0]
    assert run.text=='115' and run.font.bold and run.font.italic and run.font.size.pt==18
    assert table.cell(1,1).text=='115 ms' and table.cell(2,1).text=='同一设备'
    assert verify(data,changed,'.pptx',ops)['non_target_parts_exact']


def test_merged_and_special_table_cells_are_not_edit_targets(tmp_path):
    source=table_fixture(tmp_path);target=next(run for run in view(source.read_bytes(),'.pptx')['runs'] if run['text']=='120');deck=Presentation(source);deck.slides[0].shapes[-1].table.cell(2,0).merge(deck.slides[0].shapes[-1].table.cell(2,1));deck.save(source)
    result=view(source.read_bytes(),'.pptx')
    assert all(not run['location'].startswith('表格') for run in result['runs'])
    assert any('合并' in reason for reason in result['limitations'])
    with pytest.raises(GuardError):transform(source.read_bytes(),'.pptx',operation(target['id'],'115','120','slide-0'))
    raw=load(table_fixture(tmp_path).read_bytes(),'.pptx');root=xml(raw['ppt/slides/slide1.xml']);para=root.find('.//'+A+'tbl/'+A+'tr/'+A+'tc/'+A+'txBody/'+A+'p')
    field=deepcopy(para.find(A+'r'));field.tag=A+'fld';para.append(field);raw['ppt/slides/slide1.xml']=encoded(root)
    result=view(pack(raw),'.pptx')
    assert not any(run['location']=='表格1 · 第1行第1列' for run in result['runs'])
    assert any('特殊' in reason for reason in result['limitations'])
