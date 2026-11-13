from copy import deepcopy
import io
import json
from pathlib import Path
import time
from types import SimpleNamespace

from lxml import etree as ET
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
import pytest

from office_agent import office_canvas as engine
from office_agent.office_canvas import A, S, GuardError, encoded, load, pack, transform, verify, view, xml
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.document_open import parse_payload, read_local_view
from office_agent.webui.floating_session import FloatingSession
from office_agent.writes import digest
from office_agent.instruction_runtime import InstructionRuntime


def fixtures(folder):
    xlsx = folder/'实验数据.xlsx'
    book = Workbook(); sheet = book.active; sheet.title = '实验'
    sheet['A1'] = '温度'; sheet['B1'] = '耗时(ms)'; sheet['A2'] = '120 ℃'; sheet['B2'] = 120; sheet['C2'] = '=B2*2'
    sheet['A1'].font = Font(color='000000')
    sheet['B2'].font = Font(name='Arial', bold=True, color='002244'); sheet['B2'].fill = PatternFill('solid', fgColor='FFF0CC')
    sheet['B2'].border = Border(bottom=Side(style='thin')); sheet.merge_cells('A4:C4'); sheet['A4'] = '合并说明'
    sheet.column_dimensions['B'].width = 24; sheet.row_dimensions[2].height = 25; sheet['A45'] = '后部数据'
    sheet.print_options.horizontalCentered = True; sheet.print_area = 'A1:C8'
    book.create_sheet('复核')['A1'] = '第二个工作表'; book.save(xlsx)
    pptx = folder/'技术汇报.pptx'
    deck = Presentation()
    for index in range(2):
        slide = deck.slides.add_slide(deck.slide_layouts[6]); frame = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(7), Inches(1)).text_frame
        first = frame.paragraphs[0].add_run(); first.text = '实验汇报' if index == 0 else '复核结论'; first.font.bold = True; first.font.size = Pt(28)
        second = frame.paragraphs[0].add_run(); second.text = ' · 温度120 ℃'; second.font.italic = True
        slide.shapes.add_textbox(Inches(1), Inches(3), Inches(7), Inches(1)).text = '耗时120 ms，保留条件和单位。'
        slide.notes_slide.notes_text_frame.text = '自建验收样例，不代表实验事实。'
    deck.save(pptx)
    return xlsx, pptx


def operation(target, value, expected, item='sheet-0', action='text'):
    return {'item_id': item, 'target': target, 'action': action, 'value': value, 'expected': expected}


@pytest.fixture(params=['.xlsx', '.pptx'])
def canvas(tmp_path, request):
    paths = fixtures(tmp_path); source = paths[0 if request.param == '.xlsx' else 1]
    app = WebviewApp(); api = SimpleNamespace(_app=app, _auth=SimpleNamespace(status=lambda:{'authenticated':True}), _memory_key=lambda:'self-created', _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating = FloatingSession(api)
    assert floating.admit(str(source))['ok']; assert floating.read(floating.read_token, True)['ok']
    service = floating.canvas; result = service.open(floating.document['id'], floating.document['version'])
    op = operation('B2', '125', '120', action='number') if request.param == '.xlsx' else operation('run-0', '实验汇报（复核）', '实验汇报', 'slide-0')
    return service, floating, source, result, op


def test_workbook_exact_parts_formula_cache_format_and_window(tmp_path):
    source, _ = fixtures(tmp_path); original = source.read_bytes()
    before = load(original, '.xlsx'); sheet = xml(before['xl/worksheets/sheet1.xml']); formula = sheet.find('.//'+S+'c[@r="C2"]'); formula.find(S+'v').text = '240'; before['xl/worksheets/sheet1.xml'] = encoded(sheet); original = pack(before)
    op = operation('B2', '125', '120', action='number'); data, changes = transform(original, '.xlsx', op)
    after = load(data, '.xlsx')
    assert changes['derived'] and xml(after['xl/worksheets/sheet1.xml']).find('.//'+S+'c[@r="C2"]/'+S+'v') is None
    assert after['xl/styles.xml'] == before['xl/styles.xml']
    for name in before:
        if name not in changes['changed_parts']: assert before[name] == after[name]
    book = load_workbook(io.BytesIO(data)); assert book['实验']['B2'].value == 125 and book['实验']['C2'].value == '=B2*2'
    assert book['实验']['B2'].font.bold and book['实验']['B2'].border.bottom.style == 'thin' and str(book['实验'].merged_cells) == 'A4:C4'
    assert view(data, '.xlsx', 'sheet-0', 40, 1)['cells'][0]['value'] == '后部数据'
    formatted, _ = transform(data, '.xlsx', operation('B2', {'italic':True,'size':16}, '125', action='format'))
    changed = load_workbook(io.BytesIO(formatted)); assert changed['实验']['B2'].font.bold and changed['实验']['B2'].font.italic and changed['实验']['B2'].font.sz == 16
    assert not changed['实验']['A1'].font.italic
    assert verify(original, data, '.xlsx', [op])['non_target_parts_exact']


def test_ppt_run_identity_non_target_parts_and_reopen(tmp_path):
    _, source = fixtures(tmp_path); original = source.read_bytes(); before = load(original, '.pptx')
    ops = [operation('run-0','实验汇报（复核）','实验汇报','slide-0'), operation('run-0',{'size':32,'italic':True},'实验汇报（复核）','slide-0','format')]
    data = original
    for op in ops: data, _ = transform(data, '.pptx', op)
    after = load(data, '.pptx'); assert {name for name in before if before[name] != after[name]} == {'ppt/slides/slide1.xml'}
    deck = Presentation(io.BytesIO(data)); runs = deck.slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert runs[0].text == '实验汇报（复核）' and runs[0].font.bold and runs[0].font.italic and runs[0].font.size.pt == 32
    assert runs[1].text == ' · 温度120 ℃' and runs[1].font.italic
    assert '复核结论' in deck.slides[1].shapes[0].text
    assert verify(original, data, '.pptx', ops)['changed_parts'] == ['ppt/slides/slide1.xml']
    after['ppt/slides/slide2.xml'] = after['ppt/slides/slide2.xml'].replace('复核结论'.encode(), '未经批准'.encode())
    with pytest.raises(GuardError, match='归因'): verify(original, pack(after), '.pptx', ops)


@pytest.mark.parametrize('target,action,value,expected', [('C2','text','1',''), ('B4','text','覆盖',''), ('B2','number','NaN','120'), ('B2','number','1234567890123456','120'), ('B2','format',{'xml':'x'},'120'), ('B2','text','改动','旧值'), ('ZZ10001','text','越界','')])
def test_reject_unsafe_or_ambiguous_cells(tmp_path, target, action, value, expected):
    source, _ = fixtures(tmp_path)
    with pytest.raises(GuardError): transform(source.read_bytes(), '.xlsx', operation(target,value,expected,action=action))


def test_literal_formula_is_text_and_shared_rich_text_stays_readonly(tmp_path):
    source, _ = fixtures(tmp_path); original = source.read_bytes()
    data, _ = transform(original,'.xlsx',operation('D2','=WEBSERVICE("https://invalid.example")',''))
    book = load_workbook(io.BytesIO(data)); assert book['实验']['D2'].data_type == 's'
    raw = load(original,'.xlsx'); styles = xml(raw['xl/workbook.xml']); protection = styles.find(S+'workbookProtection'); protection.set('lockStructure','1'); raw['xl/workbook.xml'] = encoded(styles)
    with pytest.raises(GuardError,match='保护'):view(pack(raw),'.xlsx')


def test_revision_cancel_conflict_undo_and_switch(canvas):
    service, floating, source, result, op = canvas; original = source.read_bytes(); doc = floating.document['id']
    service.apply(doc,0,op)
    assert digest(source) == floating.document['version'] and service.state()['dirty']
    assert service.blocks_original()
    with pytest.raises(GuardError):service.apply(doc,0,op)
    with pytest.raises(GuardError):service.before_admit()
    plan = service.plan(doc,1)['plan']; assert service.publish(doc,1,plan['id'],False)['cancelled']
    assert not list(source.parent.glob('*.agent-*')) and not service.plans
    plan = service.plan(doc,1)['plan']; service.plans[plan['id']]['deadline'] = time.monotonic()-1
    with pytest.raises(GuardError):service.publish(doc,1,plan['id'],True)
    service.undo(doc,1); assert load(service.data,source.suffix) == load(original,source.suffix) and not service.operations
    assert not service.blocks_original() and not service.state()['needs_copy_for_agent']
    service.apply(doc,2,op); assert service.discard(doc,3,True)['ok']; service.before_admit()
    with pytest.raises(GuardError):service.before_admit()
    source.write_bytes(original+b'changed')
    with pytest.raises(ValueError,match='改变'):service.plan(doc,3)


def wait_finished(service):
    deadline=time.monotonic()+80
    while service.publishing:
        assert time.monotonic()<deadline, 'owned renderer/publication deadline'
        time.sleep(.05)


def test_transaction_cancel_render_failure_and_audit_rollback(canvas, monkeypatch):
    service, floating, source, result, op = canvas; doc=floating.document['id']; original=source.read_bytes()
    service.apply(doc,0,op); plan=service.plan(doc,1)['plan']
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents',lambda *args,**kw:{'inputs_unchanged':True,'render':{'status':'error'}})
    service.publish(doc,1,plan['id'],True);wait_finished(service)
    assert not list(source.parent.glob('*.agent-*')) and not service.uncertain and service.state()['dirty']
    assert json.loads((floating.app.editor.audit/plan['id']/'audit.json').read_text('utf8'))['status']=='failed'
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents',lambda *args,**kw:{'inputs_unchanged':True,'render':{'status':'different'}})
    plan=service.plan(doc,1)['plan'];service.publish(doc,1,plan['id'],True);wait_finished(service)
    output=Path(plan['output']);assert output.is_file() and output.read_bytes()==service.data and source.read_bytes()==original
    record=floating.app.editor.rollback(plan['id'],confirmed=True);assert record['status']=='rolled_back' and not output.exists() and source.read_bytes()==original


def test_real_parser_worker_and_native_drop_same_path(canvas):
    service, floating, source, result, op = canvas
    parsed=read_local_view(source,lambda:False,expected_version=digest(source))
    assert parsed['view'] and parse_payload(source)['view'] and floating.document['ready']
    assert result['view']['engine']=='source-preserving-ooxml-v1'


@pytest.mark.parametrize('suffix', ['.xlsx','.pptx'])
def test_package_rejects_external_relationships_and_active_content(tmp_path,suffix):
    paths=fixtures(tmp_path);source=paths[suffix=='.pptx'];raw=load(source.read_bytes(),suffix)
    raw['unsafe.xml']=b'<!DOCTYPE data [<!ENTITY x "bad">]><data>&x;</data>'
    with pytest.raises(GuardError):load(pack(raw),suffix)
    raw=load(source.read_bytes(),suffix);name=next(name for name in raw if name.endswith('.rels'));tree=xml(raw[name]);next(iter(tree)).set('TargetMode','External');next(iter(tree)).set('Target','https://invalid.example/file');raw[name]=encoded(tree)
    with pytest.raises(GuardError):load(pack(raw),suffix)


def test_manual_edit_invalidates_agent_approval_and_protection(canvas):
    service, floating, source, result, op = canvas
    runtime=InstructionRuntime(floating.document['version']);runtime.submit('user','修改指定位置');runtime.plans['old']={'id':'old'};floating.runtime=runtime
    floating.app.pending_id='old'
    if source.suffix=='.xlsx':
        floating.protected_cells.add(('实验','B2'))
        with pytest.raises(GuardError,match='保护'):service.apply(floating.document['id'],0,op)
        assert not runtime.closed
        floating.protected_cells.clear()
    service.apply(floating.document['id'],0,op)
    assert runtime.closed and not runtime.plans and floating.app.pending_id is None
    with pytest.raises(ValueError,match='草稿'):floating.check_plan({'id':'old'})


def test_privacy_epoch_does_not_repopulate_new_session(canvas, monkeypatch):
    service, floating, source, result, op=canvas;service.apply(floating.document['id'],0,op)
    plan=service.plan(floating.document['id'],1)['plan']
    def render(*args,**kwargs):
        floating.reset()
        return {'inputs_unchanged':True,'render':{'status':'different'}}
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents',render)
    service.publish(floating.document['id'],1,plan['id'],True)
    deadline=time.monotonic()+10
    while floating.app.busy:
        assert time.monotonic()<deadline;time.sleep(.02)
    assert floating.document is None and not floating.outputs and not service.active()
    log=json.loads((floating.app.editor.audit/plan['id']/'audit.json').read_text('utf8'));assert log['status']=='verified'
    floating.app.editor.rollback(plan['id'],confirmed=True)


@pytest.mark.parametrize('suffix',['.xlsx','.pptx'])
def test_incoming_raster_preserved_and_disguised_svg_refused(tmp_path,suffix):
    paths=fixtures(tmp_path);source=paths[suffix=='.pptx'];image=io.BytesIO();Image.new('RGB',(40,30),(70,90,120)).save(image,format='PNG');image.seek(0)
    if suffix=='.xlsx':
        from openpyxl.drawing.image import Image as SheetImage
        book=load_workbook(source);book['实验'].add_image(SheetImage(image),'E5');book.save(source)
        op=operation('B2','125','120',action='number')
    else:
        book=Presentation(source);book.slides[0].shapes.add_picture(image,Inches(1),Inches(5));book.save(source)
        op=operation('run-0','复核标题','实验汇报','slide-0')
    original=source.read_bytes();before=load(original,suffix);candidate,_=transform(original,suffix,op);after=load(candidate,suffix)
    media=[name for name in before if '/media/' in name];assert len(media)==1 and before[media[0]]==after[media[0]]
    before[media[0]]=b'<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.invalid/private"/></svg>'
    with pytest.raises(GuardError,match='PNG/JPEG'):load(pack(before),suffix)


def test_preview_reuses_existing_render_preflight_before_conversion(canvas, monkeypatch):
    service,floating,source,result,op=canvas;calls=[]
    monkeypatch.setattr('office_agent.webui.office_canvas.runtime_path',lambda:source)
    monkeypatch.setattr('office_agent.webui.office_canvas.environment_identity',lambda path:{'complete':False})
    def preflight(path,*,for_render=False):
        calls.append((path.suffix,for_render))
        raise GuardError('existing render preflight rejected')
    def conversion(*args,**kwargs):raise AssertionError('conversion must not run before preflight')
    monkeypatch.setattr('office_agent.webui.office_canvas.package_projection',preflight)
    monkeypatch.setattr('office_agent.webui.office_canvas._convert',conversion)
    with pytest.raises(GuardError,match='preflight'):service.preview(floating.document['id'],0,result['view']['item_id'])
    assert calls==[(source.suffix,True)] and not service.previewing and source.read_bytes()==service.original


def test_unchanged_value_does_not_create_draft_or_plan(canvas):
    service,floating,source,result,op=canvas;op['value']=op['expected']
    with pytest.raises(GuardError,match='没有实际变化'):service.apply(floating.document['id'],0,op)
    assert service.revision==0 and not service.operations and not service.state()['dirty']
    with pytest.raises(GuardError):service.plan(floating.document['id'],0)


@pytest.mark.parametrize('distant',['A20000','A1048576','XFD2'])
def test_window_edit_never_shrinks_original_sheet_extent(tmp_path,distant):
    source,_=fixtures(tmp_path);book=load_workbook(source);book['实验'][distant]='后部内容必须保留';book.save(source)
    original=source.read_bytes();before=load(original,'.xlsx');old_extent=xml(before['xl/worksheets/sheet1.xml']).find(S+'dimension').get('ref')
    changed,_=transform(original,'.xlsx',operation('B2','125','120',action='number'))
    tree=xml(load(changed,'.xlsx')['xl/worksheets/sheet1.xml']);last=tree.find(S+'dimension').get('ref').split(':')[-1]
    assert last==old_extent.split(':')[-1]
    reopened=load_workbook(io.BytesIO(changed));assert reopened['实验'][distant].value=='后部内容必须保留' and reopened['实验']['B2'].value==125
    assert distant not in [cell['id'] for cell in view(changed,'.xlsx')['cells']]


@pytest.mark.parametrize('inherited',['row','column','named'])
def test_inherited_format_refused_but_literal_edit_preserves_it(tmp_path,inherited):
    from openpyxl.styles import NamedStyle
    source,_=fixtures(tmp_path);book=load_workbook(source);sheet=book['实验'];sheet['A2']='原文'
    if inherited=='row':sheet.row_dimensions[2].font=Font(bold=True,italic=True)
    elif inherited=='column':sheet.column_dimensions['A'].font=Font(bold=True,italic=True)
    else:
        named=NamedStyle(name='继承格式',font=Font(bold=True,italic=True));book.add_named_style(named);sheet['A2'].style=named
    book.save(source);original=source.read_bytes();cell=next(cell for cell in view(original,'.xlsx')['cells'] if cell['id']=='A2')
    assert cell['editable'] and not cell['format_editable']
    with pytest.raises(GuardError,match='继承'):transform(original,'.xlsx',operation('A2',{'size':16},'原文',action='format'))
    candidate,_=transform(original,'.xlsx',operation('A2','修改文字','原文'));before=load(original,'.xlsx');after=load(candidate,'.xlsx')
    assert before['xl/styles.xml']==after['xl/styles.xml']
    old_tree,new_tree=xml(before['xl/worksheets/sheet1.xml']),xml(after['xl/worksheets/sheet1.xml'])
    assert old_tree.find('.//'+S+'c[@r="A2"]').get('s')==new_tree.find('.//'+S+'c[@r="A2"]').get('s')
    if inherited=='row':assert old_tree.find('.//'+S+'row[@r="2"]').get('s')==new_tree.find('.//'+S+'row[@r="2"]').get('s')
    if inherited=='column':assert encoded(old_tree.find(S+'cols'))==encoded(new_tree.find(S+'cols'))


def test_ppt_preview_rejects_page_count_mismatch(tmp_path,monkeypatch):
    _,source=fixtures(tmp_path);app=WebviewApp();api=SimpleNamespace(_app=app,_auth=SimpleNamespace(status=lambda:{'authenticated':True}),_memory_key=lambda:'self-created')
    floating=FloatingSession(api);floating.admit(str(source));floating.read(floating.read_token,True);service=floating.canvas;opened=service.open(floating.document['id'],floating.document['version'])
    monkeypatch.setattr('office_agent.webui.office_canvas.runtime_path',lambda:source)
    monkeypatch.setattr('office_agent.webui.office_canvas.environment_identity',lambda path:{'complete':False})
    monkeypatch.setattr('office_agent.webui.office_canvas._convert',lambda *args:tmp_path/'not-read.pdf')
    def worker(command,*args):
        (Path(command[-1])/'page.json').write_text(json.dumps({'page':1,'pages':1}),encoding='utf8')
    monkeypatch.setattr('office_agent.webui.office_canvas._run',worker)
    with pytest.raises(GuardError,match='页数'):service.preview(floating.document['id'],0,opened['view']['item_id'])
    assert not service.previewing and service.preview_cache is None
