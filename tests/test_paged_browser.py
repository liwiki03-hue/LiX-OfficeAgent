from contextlib import contextmanager
from copy import deepcopy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import base64
import json
import mimetypes
from pathlib import Path
import threading
import time
from types import SimpleNamespace
from urllib.parse import urlsplit

from docx import Document
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_model_connections import connection, payload, save
from tests.test_professional_browser import STATIC, EDGE


@contextmanager
def serve_static():
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            target=(STATIC/urlsplit(self.path).path.lstrip('/')).resolve()
            if target==STATIC.resolve():target=STATIC/'index.html'
            if not target.is_relative_to(STATIC.resolve()) or not target.is_file() or target.suffix not in {'.html','.js','.css','.svg','.png','.woff2'}:
                self.send_error(404);return
            self.send_response(200);self.send_header('Content-Type',mimetypes.guess_type(target.name)[0] or 'application/octet-stream');self.end_headers();self.wfile.write(target.read_bytes())
        def log_message(self,*args):pass
    server=ThreadingHTTPServer(('127.0.0.1',0),Handler);worker=threading.Thread(target=server.serve_forever,daemon=True);worker.start()
    try:yield f'http://127.0.0.1:{server.server_port}/'
    finally:server.shutdown();server.server_close();worker.join(5);assert not worker.is_alive()


@pytest.fixture
def paged_browser(tmp_path,monkeypatch):
    if not EDGE.is_file():pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source=tmp_path/'分页验收.docx'
    book=Document();book.add_heading('技术报告',0);book.add_paragraph('实验于2025年完成，温度120 ℃。');book.add_paragraph('非目标段落不变。');table=book.add_table(rows=2,cols=2);table.cell(0,0).text='条件';table.cell(1,1).text='120 ms';book.add_page_break();book.add_heading('结果',1);book.add_paragraph('第二页正文。');book.add_paragraph('换行一\n换行二\t尾部');book.sections[0].header.paragraphs[0].text='自建验收页眉';book.sections[0].footer.paragraphs[0].text='自建验收页脚';book.save(source)
    app=WebviewApp();app.attach(SimpleNamespace(create_file_dialog=lambda *a,**kw:[str(source)]))
    api=FloatingApi(app,settings_root=tmp_path);save(api._connections,payload([connection()]))
    monkeypatch.setenv('TEMP',str(tmp_path));monkeypatch.setenv('TMP',str(tmp_path))
    with serve_static() as url,sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            page=browser.new_page(viewport={'width':1280,'height':820});page.set_default_timeout(20000)
            page.expose_function('testHost',lambda name,args:getattr(api,name)(*args))
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.confirm=()=>true;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js',lambda route:route.fulfill(status=200,body=''))
            errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
            page.goto(url)
            page.locator('#floating-choose').click()
            deadline=time.monotonic()+10
            while not api._floating.document:
                if time.monotonic()>deadline:raise AssertionError({'status':page.locator('#floating-status').text_content(),'errors':errors,'auth':api._auth.status()})
                page.wait_for_timeout(50)
            page.get_by_role('button',name='同意读取并打开',exact=True).click()
            deadline=time.monotonic()+70
            while not api._floating.paged.manifest:
                if api._floating.paged.document_id and not api._floating.paged.editable:
                    raise AssertionError({'readonly_reasons':api._floating.paged.reasons,'feedback':page.locator('#paged-feedback').text_content()})
                if time.monotonic()>deadline:raise AssertionError({'feedback':page.locator('#paged-feedback').text_content() if page.locator('#paged-feedback').count() else '', 'errors':errors})
                page.wait_for_timeout(100)
            yield page,api,source,errors
        finally:browser.close()


def engine_frame(page):
    return next(frame for frame in page.frames if frame.url.endswith('paged_editor_frame.html'))


def test_real_engine_typing_undo_format_snapshot_and_source_unchanged(paged_browser):
    page,api,source,errors=paged_browser
    original=source.read_bytes();frame=engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End');page.keyboard.press('Control+b')
    frame.page.keyboard.insert_text('实验于2026年完成，温度120 ℃。')
    page.wait_for_function('document.getElementById("paged-save-state").textContent.includes("未保存")')
    frame.page.keyboard.press('Control+z');frame.page.keyboard.press('Control+y')
    page.keyboard.press('Home');page.keyboard.press('Shift+End')
    page.get_by_role('button',name='引用选区',exact=True).click()
    deadline=time.monotonic()+10
    while not api._floating.paged.references:
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)
    assert '2026' in api._floating.paged.references[next(iter(api._floating.paged.references))]['context']
    import io
    result=Document(io.BytesIO(api._floating.paged.data))
    assert result.paragraphs[1].text=='实验于2026年完成，温度120 ℃。'
    assert any(run.bold for run in result.paragraphs[1].runs if run.text)
    assert source.read_bytes()==original
    assert errors==[]


def test_ai_sees_latest_draft_blocks_old_result_and_applies_only_current(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    floating=api._floating
    floating.access.set(floating.document['id'],floating.document['version'],floating.epoch,'edit',True)
    entered,release=threading.Event(),threading.Event()
    contexts=[]
    class Model:
        def __init__(self,reference,first):self.reference=deepcopy(reference);self.first=first;self.calls=0
        def bind_tools(self,tools):return self
        def invoke(self,messages):
            self.calls+=1;contexts.append(str(messages))
            if self.first and self.calls==1:entered.set();assert release.wait(15)
            if self.calls==1:
                row=self.reference
                return {'tool_calls':[{'name':'propose_draft_edit','id':'own-call','args':{'reference_id':row['id'],'old':row['old'],'new':row['old'].replace('实验于','实验完成于').replace('年完成','年'),'reason':'保留数字和单位，精简语句。'}}]}
            return {'content':'建议待采用。'}
    built=[]
    def build(route):
        model=Model(next(iter(api._floating.paged.references.values())),not built);built.append(model);return model
    monkeypatch.setattr(api._connections,'build_model',build)
    page.locator('#simple-model-menu > summary').click()
    page.locator('#model-selector').select_option(json.dumps(['official','same-model','agent'],separators=(',',':')))
    page.keyboard.press('Escape')
    frame=engine_frame(page)
    def replace(text,old):
        frame.locator('.superdoc-page').get_by_text(old,exact=True).first.click()
        page.keyboard.press('Home');page.keyboard.press('Shift+End');page.keyboard.insert_text(text)
        page.keyboard.press('Home');page.keyboard.press('Shift+End')
        page.get_by_role('button',name='引用选区',exact=True).click()
        page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('已引用')")
    replace('实验于2026年完成，温度120 ℃。','实验于2025年完成，温度120 ℃。')
    page.locator('#floating-input').fill('精简这句话，保留数字和单位。');page.locator('#floating-send').click()
    deadline=time.monotonic()+10
    while not entered.is_set():
        if time.monotonic()>deadline:raise AssertionError(page.locator('#floating-status').text_content())
        page.wait_for_timeout(50)
    try:
        replace('实验于2027年完成，温度120 ℃。','实验于2026年完成，温度120 ℃。')
        page.locator('#floating-input').fill('使用最新年份，继续精简。');page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
    finally:release.set()
    deadline=time.monotonic()+20
    while not api._floating.paged.proposals or not all(row['status']=='adopted' for row in api._floating.paged.proposals.values()):
        if time.monotonic()>deadline:raise AssertionError({'proposals':api._floating.paged.proposals,'feedback':page.locator('#paged-feedback').text_content()})
        page.wait_for_timeout(50)
    proposals=list(api._floating.paged.proposals.values())
    assert len(proposals)==1 and '2027' in proposals[0]['old'] and '2026' not in proposals[0]['new']
    assert '2026' in contexts[0] and '2027' in contexts[-1]
    deadline=time.monotonic()+10
    while api._floating.paged.proposals[proposals[0]['id']]['status']!='adopted':
        if time.monotonic()>deadline:raise AssertionError({'feedback':page.locator('#paged-feedback').text_content(),'revision':api._floating.paged.revision,'proposal':api._floating.paged.proposals[proposals[0]['id']]['status']})
        page.wait_for_timeout(50)
    text='\n'.join(frame.locator('.superdoc-page').all_text_contents())
    assert '2027' in text and '2026' not in text
    assert errors==[]


def test_fluid_type_sidebar_selection_and_instance_survive_resize(paged_browser):
    page,api,source,errors=paged_browser
    frame=engine_frame(page);frame_id=frame._impl_obj._guid
    page.locator('#floating-input').fill('这条要求尚未发送，保留草稿。')
    for width in [520,820,1280,1600]:
        page.set_viewport_size({'width':width,'height':820})
        for theme in ['light','dark','highcontrast']:
            page.evaluate('(value)=>document.documentElement.dataset.theme=value',theme)
            values=page.evaluate("""()=>{const n=document.getElementById('floating-input'),r=n.getBoundingClientRect();return {font:parseFloat(getComputedStyle(n).fontSize),bottom:r.bottom,hidden:n.hidden,overflow:document.documentElement.scrollWidth>innerWidth}}""")
            expected=max(14,min(16.5,11.37+.00321*width))
            assert values['font']==pytest.approx(expected,abs=.02)
            assert values['bottom']<=820 and not values['overflow']
    page.get_by_role('button',name='AI 对话',exact=True).click();page.get_by_role('button',name='AI 对话',exact=True).click()
    assert engine_frame(page)._impl_obj._guid==frame_id
    assert page.locator('#floating-input').input_value()=='这条要求尚未发送，保留草稿。'
    assert len(api._floating.paged.manifest)>0 and errors==[]
    page.set_viewport_size({'width':360,'height':440})
    page.locator('#floating-send').scroll_into_view_if_needed()
    rectangle=page.locator('#floating-send').bounding_box()
    assert rectangle and rectangle['y']>=0 and rectangle['y']+rectangle['height']<=440
    page.locator('#floating-input').press('Escape')
    assert page.locator('.paged-frame').evaluate('node=>!node.inert')
    page.get_by_role('button',name='AI 对话',exact=True).click()
    assert page.locator('#floating-input').input_value()=='这条要求尚未发送，保留草稿。'


def test_confirmed_engine_copy_real_render_reopen_and_later_draft_is_dirty(paged_browser):
    page,api,source,errors=paged_browser
    original=source.read_bytes();frame=engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End');page.keyboard.insert_text('实验于2026年完成，温度120 ℃。')
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).click()
    assert not list(source.parent.glob('*.agent-*.docx'))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('button',name='确认生成副本',exact=True).click()
    deadline=time.monotonic()+10
    while not api._floating.paged.publishing:
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)
    frozen_revision=api._floating.paged.revision
    frame.locator('.superdoc-page').get_by_text('非目标段落不变。',exact=True).first.click()
    page.keyboard.press('End');page.keyboard.insert_text('后续未发布草稿。')
    deadline=time.monotonic()+180
    while api._floating.paged.publishing:
        if time.monotonic()>deadline:raise AssertionError('bounded real render timed out')
        page.wait_for_timeout(100)
    assert api._floating.outputs, list(api._floating.notifications.queue)
    result=next(iter(api._floating.outputs.values()));output=Path(result['output'])
    reopened=Document(output)
    assert result['status']=='verified' and result['render_diff']['inputs_unchanged']
    assert reopened.paragraphs[1].text=='实验于2026年完成，温度120 ℃。'
    assert reopened.paragraphs[2].text=='非目标段落不变。'
    assert reopened.tables[0].cell(1,1).text=='120 ms'
    assert reopened.sections[0].header.paragraphs[0].text=='自建验收页眉' and reopened.sections[0].footer.paragraphs[0].text=='自建验收页脚'
    assert source.read_bytes()==original
    state=api._floating.paged.state()
    assert state['published_revision']==frozen_revision and state['draft_revision']>frozen_revision and state['dirty']
    page.screenshot(path=str(source.parent/'engine-copy-verified.png'))
    (source.parent/'ENGINE_COPY_REPORT.json').write_text(json.dumps({'output':str(output),'source_unchanged':source.read_bytes()==original,'state':{k:v for k,v in state.items() if k not in {'proposals'}},'render':result['render_diff']},ensure_ascii=False,indent=2),encoding='utf8')
    assert errors==[]


def test_real_font_size_style_and_spacing_controls_export_exact_values(paged_browser):
    import io
    page,api,source,errors=paged_browser;frame=engine_frame(page)
    frame.locator('.ProseMirror[contenteditable=true]').first.wait_for(state='attached')
    original_revision=api._floating.paged.revision
    page.get_by_role('button',name='纸张深色显示',exact=True).click()
    page.get_by_role('button',name='纸张深色显示',exact=True).click()
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End')
    assert api._floating.paged.revision==original_revision==0
    frame.locator('[data-item="btn-fontSize"] input').fill('16')
    frame.locator('[data-item="btn-fontSize"] input').press('Enter')
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).wait_for()
    first=Document(io.BytesIO(api._floating.paged.data)).paragraphs[1]
    assert any(run.font.size and run.font.size.pt==16 for run in first.runs),[run.font.size.pt if run.font.size else None for run in first.runs]
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).click()
    frame.locator('#paragraph-style').select_option('Heading2')
    frame.locator('#paragraph-after').select_option('12')
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).wait_for()
    exported=Document(io.BytesIO(api._floating.paged.data));paragraph=exported.paragraphs[1]
    assert paragraph.style.style_id=='Heading2'
    assert any(run.font.size and run.font.size.pt==16 for run in paragraph.runs)
    assert paragraph.paragraph_format.space_after is not None,frame.locator('#frame-error').text_content()
    assert paragraph.paragraph_format.space_after.pt==12
    assert exported.paragraphs[2].text=='非目标段落不变。' and errors==[]


def test_outline_find_and_paste_use_engine_model_not_dom_rebuild(paged_browser):
    page,api,source,errors=paged_browser;frame=engine_frame(page)
    frame.locator('.ProseMirror[contenteditable=true]').first.wait_for(state='attached')
    page.get_by_role('button',name='查找',exact=True).click()
    page.get_by_role('textbox',name='查找文字',exact=True).fill('第二页正文')
    page.get_by_role('button',name='定位第一个匹配',exact=True).click()
    page.get_by_role('button',name='引用选区',exact=True).click()
    page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('已引用')")
    reference=next(iter(api._floating.paged.references.values()))
    assert reference['old']=='第二页正文。'
    assert api._floating.paged.revision==0 and errors==[]
    requests=[];page.on('request',lambda request:requests.append(request.url))
    frame.evaluate("""()=>{const data=new DataTransfer();data.setData('text/html','<p><strong>实验说明</strong><script>window.bad=1</script><img src="https://invalid.example/unwanted"></p>');data.setData('text/plain','实验说明');document.querySelector('.ProseMirror').dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,bubbles:true,cancelable:true}));}""")
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).wait_for()
    import io
    result=Document(io.BytesIO(api._floating.paged.data))
    assert any('实验说明' in p.text and any(run.bold and '实验说明' in run.text for run in p.runs) for p in result.paragraphs)
    assert result.paragraphs[1].text=='实验于2025年完成，温度120 ℃。'
    assert not frame.evaluate('Boolean(window.bad)') and not any('invalid.example' in url for url in requests)


def test_adapter_contract_and_host_undo_redo(paged_browser):
    page,api,source,errors=paged_browser;frame=engine_frame(page)
    original=source.read_bytes()
    frame.locator('.ProseMirror[contenteditable=true]').first.wait_for(state='attached')
    contract=page.evaluate('window.OfficePaged.contract()')
    assert contract['schema_version']==1 and contract['missing']==[]
    assert set(contract['methods'])=={'open','close','capabilities','getSelection','getSnapshot','subscribeChanges','applyValidatedOperations','applyTableCommand','undo','redo','exportCandidate','validateCandidate','focusTarget','restoreDraft'}
    capabilities=page.evaluate('window.OfficePaged.capabilities()')
    assert capabilities['format']=='docx' and capabilities['engine']=='SuperDoc'
    assert capabilities['operations']['undo'] and capabilities['operations']['redo'] and capabilities['operations']['batch']
    assert capabilities['limits']['max_operations']==8
    valid={'document_id':api._floating.paged.document_id,'session_epoch':api._floating.paged.state()['session_epoch'],
           'source_version':api._floating.paged.source_version,'draft_revision':api._floating.paged.revision,
           'instruction_revision':0,'target_id':'node-1','expected_before':'before','operation_id':'op-1',
           'idempotency_key':'key-1','action':'replace_text','start':0,'end':6,'after':'after'}
    assert page.evaluate('(op)=>window.OfficeEditorContract.validateOperation(op)',valid)
    assert page.evaluate("""(op)=>{try{window.OfficeEditorContract.validateOperation({...op,action:'delete'});return false}catch(error){return error.message.includes('文字替换')}}""",valid)
    page.evaluate("""()=>{window.__adapterEvents=[];window.__adapterUnsubscribe=window.OfficePaged.subscribeChanges((type,value)=>window.__adapterEvents.push({type,revision:value?.revision}));}""")
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
    page.keyboard.press('End');page.keyboard.insert_text('撤销重做')
    page.wait_for_function('document.getElementById("paged-save-state").textContent.includes("未保存")')
    page.wait_for_function("window.__adapterEvents.some(event=>event.type==='changed')")
    snapshot=page.evaluate('window.OfficePaged.snapshot()')
    assert snapshot['revision']==api._floating.paged.revision and snapshot['manifest']
    page.evaluate('window.__adapterUnsubscribe()')
    before=api._floating.paged.revision
    assert page.evaluate('async()=>(await window.OfficePaged.history()).canUndo') is True
    page.evaluate('window.OfficePaged.undo()')
    deadline=time.monotonic()+10
    while '撤销重做' in '\n'.join(frame.locator('.superdoc-page').all_text_contents()):
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)
    while api._floating.paged.revision<=before:
        if time.monotonic()>deadline:raise AssertionError('undo revision not acknowledged')
        page.wait_for_timeout(50)
    assert page.evaluate('async()=>(await window.OfficePaged.history()).canRedo') is True
    page.evaluate('window.OfficePaged.redo()')
    deadline=time.monotonic()+10
    while '撤销重做' not in '\n'.join(frame.locator('.superdoc-page').all_text_contents()):
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成',exact=False).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End')
    selection=page.evaluate('window.OfficePaged.selection()')
    assert selection['node_id'] and selection['formats']['paragraph_styles']==['Normal']
    assert source.read_bytes()==original and errors==[]


def test_batch_operations_non_overlap_id_map_and_stale_conflict(paged_browser):
    page,api,source,errors=paged_browser;frame=engine_frame(page)
    original=source.read_bytes()
    frame.locator('.ProseMirror[contenteditable=true]').first.wait_for(state='attached')
    state=api._floating.paged.state()
    row=next(item for item in api._floating.paged.manifest if '2025' in item['text'] and '120' in item['text'])
    row2=next(item for item in api._floating.paged.manifest if '非目标段落不变' in item['text'])
    first=row['text'].index('2025');second=row['text'].index('120');third=row2['text'].index('不变')
    mapping=page.evaluate('window.OfficePaged.idMap()')
    assert len(mapping['rows'])==len(api._floating.paged.manifest)
    assert any(item['editor_id']==row['id'] and item['paragraph_index'] is not None for item in mapping['rows'])
    mapped_locators={item['editor_id']:item['host_locator_id'] for item in mapping['rows'] if item['host_locator_id']}
    assert mapped_locators
    base={'document_id':state['document_id'],'session_epoch':state['session_epoch'],'source_version':state['source_version'],
          'draft_revision':state['draft_revision'],'instruction_revision':state['instruction_revision'],'target_id':row['id'],
          'operation_id':'batch-op-1','idempotency_key':'batch-key-1','action':'replace_text'}
    operations=[{**base,'start':first,'end':first+4,'expected_before':'2025','after':'2026'},
                {**base,'operation_id':'batch-op-2','idempotency_key':'batch-key-2','start':second,'end':second+3,'expected_before':'120','after':'150'},
                {**base,'target_id':row2['id'],'operation_id':'batch-op-3','idempotency_key':'batch-key-3','start':third,'end':third+2,'expected_before':'不变','after':'保持'}]
    before=state['draft_revision']
    result=page.evaluate('(ops)=>window.OfficePaged.applyOperations(ops)',operations)
    assert result['draft_revision']==before+1
    text='\n'.join(frame.locator('.superdoc-page').all_text_contents())
    assert '2026' in text and '150' in text and '保持' in text and '2025' not in text
    after_map=page.evaluate('window.OfficePaged.idMap()')
    after_rows={item['editor_id']:item for item in after_map['rows']}
    assert set(after_rows)=={item['editor_id'] for item in mapping['rows']}
    for editor_id,locator in mapped_locators.items():
        assert editor_id in after_rows and after_rows[editor_id]['host_locator_id']==locator
    assert source.read_bytes()==original
    stale=page.evaluate('(ops)=>(async()=>{try{await window.OfficePaged.applyOperations(ops);return ""}catch(error){return error.message}})()',operations)
    assert '版本' in stale or '过期' in stale
    current=api._floating.paged.state()
    conflict=[{**base,'session_epoch':current['session_epoch'],'draft_revision':current['draft_revision'],
               'instruction_revision':current['instruction_revision'],'start':first,'end':first+4,
               'expected_before':'9999','after':'2026','operation_id':'batch-op-4','idempotency_key':'batch-key-4'}]
    failed=page.evaluate('(ops)=>(async()=>{try{await window.OfficePaged.applyOperations(ops);return ""}catch(error){return error.message}})()',conflict)
    assert '原文不匹配' in failed
    overlap=[{**base,'session_epoch':current['session_epoch'],'draft_revision':current['draft_revision'],
              'instruction_revision':current['instruction_revision'],'start':first,'end':first+4,
              'expected_before':'2026','after':'2026','operation_id':'batch-op-5','idempotency_key':'batch-key-5'},
             {**base,'session_epoch':current['session_epoch'],'draft_revision':current['draft_revision'],
              'instruction_revision':current['instruction_revision'],'start':first+2,'end':first+4,
              'expected_before':'26','after':'27','operation_id':'batch-op-6','idempotency_key':'batch-key-6'}]
    overlap_error=page.evaluate('(ops)=>(async()=>{try{await window.OfficePaged.applyOperations(ops);return ""}catch(error){return error.message}})()',overlap)
    assert '重叠' in overlap_error
    assert errors==[]


def test_table_structure_toolbar_and_host_commands(paged_browser):
    import io
    page,api,source,errors=paged_browser;frame=engine_frame(page)
    original=source.read_bytes()
    frame.locator('.superdoc-page').get_by_text('条件',exact=True).first.click()
    frame.locator('#table-body-actions').wait_for(state='visible')
    assert frame.get_by_role('button',name='在下方插入表格行',exact=True).is_visible()
    before=api._floating.paged.revision
    frame.get_by_role('button',name='在下方插入表格行',exact=True).click()
    deadline=time.monotonic()+10
    while api._floating.paged.revision<=before:
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)
    def shape():
        snapshot=page.evaluate('window.OfficePaged.snapshot()')
        table=Document(io.BytesIO(base64.b64decode(snapshot['data']))).tables[0]
        return len(table.rows),len(table.columns)
    assert shape()==(3,2)
    frame.locator('.superdoc-page').get_by_text('条件',exact=True).first.click()
    assert page.evaluate("window.OfficePaged.tableCommand('addRowBefore')")['revision']>before
    assert shape()==(4,2)
    frame.locator('.superdoc-page').get_by_text('条件',exact=True).first.click()
    assert page.evaluate("window.OfficePaged.tableCommand('deleteRow')")['revision']>before
    assert shape()==(3,2)
    frame.locator('.superdoc-page').get_by_text('120 ms',exact=True).first.click()
    frame.get_by_role('button',name='删除当前表格行',exact=True).click()
    deadline=time.monotonic()+10
    while shape()!=(2,2):
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)
    assert frame.get_by_role('button',name='在左侧插入表格列',exact=True).count()==0
    assert frame.get_by_role('button',name='删除当前表格列',exact=True).count()==0
    closed=page.evaluate("(async()=>{try{await window.OfficePaged.tableCommand('addColumnAfter');return ''}catch(error){return error.message}})()")
    assert '不支持' in closed
    assert source.read_bytes()==original
    assert errors==[]
