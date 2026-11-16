from datetime import date
import json
import os
from pathlib import Path
import shutil
import time
from types import SimpleNamespace

from openpyxl import Workbook, load_workbook
from pptx import Presentation
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.writes import digest
from tests.test_office_canvas import fixtures
from tests.test_paged_browser import serve_static
from tests.test_professional_browser import EDGE
from tests.ui_confirmation import record_browser_failure, track_browser


@pytest.fixture(params=['.xlsx','.pptx'])
def native_browser(tmp_path, request, monkeypatch):
    if not EDGE.is_file(): pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    paths=fixtures(tmp_path); source=paths[request.param=='.pptx']
    monkeypatch.setenv('TEMP',str(tmp_path));monkeypatch.setenv('TMP',str(tmp_path))
    app=WebviewApp();app.attach(SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(source)]));api=FloatingApi(app,settings_root=tmp_path)
    with serve_static() as url,sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            page=browser.new_page(viewport={'width':1280,'height':820});page.set_default_timeout(60000)
            page.expose_function('testHost',lambda name,args:getattr(api,name)(*args))
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.confirm=()=>true;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js',lambda route:route.fulfill(status=200,body=''))
            errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
            track_browser(page,tmp_path.name,lambda:{'document':api._floating.document,'stage':api._floating.stage,'busy':app.busy,'canvas_active':api._floating.canvas.active()})
            try:
                response=page.goto(url)
                assert response is not None and response.ok
                page.locator('#floating-choose').click();page.get_by_role('button',name='同意读取并打开',exact=True).click()
            except Exception:
                record_browser_failure(page,'canvas page bootstrap and file selection',60)
                raise
            page.locator('#canvas-items button').first.wait_for()
            if request.param=='.pptx':page.locator('#canvas-slide img').wait_for()
            yield page,api,source,errors
        finally:browser.close()


@pytest.fixture
def display_browser(tmp_path, monkeypatch):
    if not EDGE.is_file(): pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source=tmp_path/'显示回归.xlsx'
    book=Workbook();sheet=book.active;sheet.title='显示';sheet['A1']='日期';sheet['A2']=date(2025,1,1);sheet['A2'].number_format='yyyy-mm-dd';sheet['B1']='数值';sheet['B2']=1234.5;sheet['B2'].number_format='#,##0.00';sheet['C2']='=B2+1';sheet['C2'].number_format='#,##0.00';sheet['A45']='后部滚动';book.save(source)
    monkeypatch.setenv('TEMP',str(tmp_path));monkeypatch.setenv('TMP',str(tmp_path))
    app=WebviewApp();app.attach(SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(source)]));api=FloatingApi(app,settings_root=tmp_path)
    with serve_static() as url,sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            page=browser.new_page(viewport={'width':1280,'height':820});page.set_default_timeout(60000)
            page.expose_function('testHost',lambda name,args:getattr(api,name)(*args))
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.confirm=()=>true;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js',lambda route:route.fulfill(status=200,body=''))
            errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
            track_browser(page,tmp_path.name,lambda:{'document':api._floating.document,'stage':api._floating.stage,'busy':app.busy,'canvas_active':api._floating.canvas.active()})
            try:
                response=page.goto(url)
                assert response is not None and response.ok
                page.locator('#floating-choose').click();page.get_by_role('button',name='同意读取并打开',exact=True).click()
            except Exception:
                record_browser_failure(page,'canvas page bootstrap and file selection',60)
                raise
            page.locator('#canvas-items button').first.wait_for()
            yield page,api,source,errors
        finally:browser.close()


def test_formatted_display_keeps_raw_edit_value(display_browser):
    page,api,source,errors=display_browser
    original=source.read_bytes()
    page.locator('.canvas-grid button[data-target="A2"]').click()
    assert page.locator('.canvas-grid button[data-target="A2"]').text_content()=='2025-01-01'
    assert page.locator('#canvas-value').input_value()=='45658'
    assert '显示：2025-01-01' in page.locator('#canvas-feedback').text_content()
    assert '原始值：45658' in page.locator('#canvas-feedback').text_content()
    page.locator('.canvas-grid button[data-target="B2"]').click()
    assert page.locator('.canvas-grid button[data-target="B2"]').text_content()=='1,234.50'
    assert page.locator('#canvas-value').input_value()=='1234.5'
    assert page.locator('.canvas-grid button[data-target="C2"]').text_content()=='1,235.50'
    assert source.read_bytes()==original and errors==[]


def test_virtualized_scroll_loads_next_window(display_browser):
    page,api,source,errors=display_browser
    for _ in range(4):
        page.locator('#canvas-surface').evaluate('(el)=>{el.scrollTop=el.scrollHeight;}')
        page.wait_for_timeout(350)
    page.locator('.canvas-grid button[data-target="A45"]').wait_for(timeout=10000)
    assert page.locator('.canvas-grid button[data-target="A45"]').text_content()=='后部滚动'
    assert page.locator('#canvas-window').input_value()=='A41'
    assert errors==[]


def idle(page):
    page.wait_for_function("document.getElementById('canvas-state') && !document.getElementById('canvas-state').textContent.includes('正在')")


def test_actual_grid_or_slide_edit_publish_reopen_and_rollback(native_browser):
    page,api,source,errors=native_browser; original=source.read_bytes(); suffix=source.suffix;service=api._floating.canvas
    evidence=os.environ.get('OFFICE_CANVAS_EVIDENCE');folder=Path(evidence)/suffix[1:] if evidence else None
    if folder:folder.mkdir(parents=True,exist_ok=True);shutil.copy2(source,folder/source.name);page.screenshot(path=str(folder/'before.png'))
    if suffix=='.xlsx':
        page.locator('.canvas-grid button[data-target="B2"]').dblclick();page.locator('#canvas-value').fill('125');page.locator('#canvas-value').press('Enter')
        page.wait_for_function("document.querySelector('.canvas-grid button[data-target=B2]').textContent==='125'")
        page.get_by_role('button',name='斜体',exact=True).click();page.wait_for_function("document.querySelector('.canvas-grid button[data-target=B2]').style.fontStyle==='italic'")
        page.locator('#canvas-window').fill('A40');page.locator('#canvas-window').press('Enter');page.get_by_role('button',name='后部数据',exact=True).wait_for();idle(page)
        page.locator('#canvas-items button').nth(1).click();page.get_by_role('button',name='第二个工作表',exact=True).wait_for();idle(page)
        page.locator('#canvas-items button').first.click();idle(page)
    else:
        page.get_by_role('button',name='编辑片段 run-0 实验汇报',exact=True).click();page.locator('#canvas-value').fill('实验汇报（复核）');page.locator('#canvas-value').press('Enter')
        page.get_by_role('button',name='编辑片段 run-0 实验汇报（复核）',exact=True).wait_for();idle(page)
        page.get_by_role('button',name='斜体',exact=True).click();page.wait_for_function("document.getElementById('canvas-italic').getAttribute('aria-pressed')==='true'");idle(page)
        page.get_by_label('刷新页图',exact=True).click();page.locator('#canvas-slide img').wait_for();idle(page)
    assert source.read_bytes()==original and service.revision==2
    page.get_by_label('保存副本',exact=True).click();confirm=page.get_by_role('dialog',name='确认保存副本',exact=True);confirm.get_by_role('button',name='取消',exact=True).click();confirm.wait_for(state='detached')
    assert not list(source.parent.glob('*.agent-*'))
    page.get_by_label('保存副本',exact=True).click();confirm.get_by_role('button',name='确认生成副本',exact=True).click();confirm.wait_for(state='detached')
    deadline=time.monotonic()+90
    while service.publishing:
        assert time.monotonic()<deadline,'bounded publication';page.wait_for_timeout(100)
    assert not service.uncertain,service.last_error
    assert api._floating.outputs,service.last_error
    receipt=next(iter(api._floating.outputs.values()));output=Path(receipt['output']);assert output.is_file() and source.read_bytes()==original
    assert receipt['render_diff']['inputs_unchanged'] and receipt['render_diff']['render']['status'] in {'different','equivalent'}
    if suffix=='.xlsx':
        book=load_workbook(output);assert book['实验']['B2'].value==125 and book['实验']['B2'].font.bold and book['实验']['B2'].font.italic and book['实验']['C2'].value=='=B2*2'
    else:
        book=Presentation(output);run=book.slides[0].shapes[0].text_frame.paragraphs[0].runs[0];assert run.text=='实验汇报（复核）' and run.font.bold and run.font.italic
    page.wait_for_function("document.getElementById('canvas-state').textContent==='副本已验证'")
    page.wait_for_function("document.getElementById('canvas-feedback').textContent.startsWith('副本已生成：')")
    assert page.get_by_role('button',name='继续编辑副本',exact=True).is_visible()
    if folder:
        shutil.copy2(output,folder/('编辑副本'+suffix));page.screenshot(path=str(folder/'after.png'))
        shutil.copy2(Path(receipt['backup']),folder/('BACKUP'+suffix))
        (folder/'AUDIT.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2),encoding='utf8')
    page.get_by_role('button',name='继续编辑副本',exact=True).click()
    page.get_by_role('button',name='同意读取并打开',exact=True).wait_for()
    assert api._app.path==output and not api._floating.document['ready'] and api._floating.runtime is None
    page.get_by_role('button',name='同意读取并打开',exact=True).click();page.locator('#canvas-items button').first.wait_for()
    if suffix=='.pptx':page.locator('#canvas-slide img').wait_for();idle(page)
    assert api._floating.canvas.original==output.read_bytes() and api._floating.canvas.revision==0
    record=api._app.editor.rollback(receipt['id'],confirmed=True);assert record['status']=='rolled_back' and not output.exists() and source.read_bytes()==original
    if folder:(folder/'ROLLBACK.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8')
    assert errors==[]


def test_draft_protection_keyboard_themes_and_sizes(native_browser):
    page,api,source,errors=native_browser
    if source.suffix=='.xlsx':page.locator('.canvas-grid button[data-target="B2"]').click()
    else:page.get_by_role('button',name='编辑片段 run-0 实验汇报',exact=True).click()
    page.locator('#canvas-value').fill('未采用草稿');page.locator('#canvas-value').press('Escape');assert '未采用草稿' not in page.locator('#canvas-value').input_value()
    page.locator('#canvas-value').fill('中文输入保留');page.locator('#canvas-value').dispatch_event('keydown',{'key':'Enter','isComposing':True});assert api._floating.canvas.revision==0
    page.locator('#floating-choose').click();dialog=page.get_by_role('dialog',name='处理未保存编辑');dialog.get_by_role('button',name='返回保存副本',exact=True).click();assert page.locator('#canvas-value').input_value()=='中文输入保留'
    page.locator('#canvas-value').press('Escape')
    evidence=os.environ.get('OFFICE_CANVAS_EVIDENCE');measurements=[]
    for width in [1440,1280,1024,520]:
        for theme in ['light','dark','highcontrast']:
            page.set_viewport_size({'width':width,'height':820});page.evaluate('(theme)=>document.documentElement.dataset.theme=theme',theme);page.wait_for_timeout(50)
            assert page.locator('#canvas-value').is_visible() and page.get_by_label('保存副本',exact=True).is_visible()
            assert page.evaluate("document.querySelector('#canvas-surface').clientHeight>100")
            if source.suffix=='.xlsx':
                assert page.locator('.canvas-grid').evaluate("node=>getComputedStyle(node).backgroundColor")=='rgb(255, 255, 255)'
                assert page.locator('.canvas-grid button[data-target="A1"]').evaluate("node=>getComputedStyle(node).color")=='rgb(0, 0, 0)'
            measurements.append({'width':width,'theme':theme,'input':page.locator('#canvas-value').bounding_box(),'surface':page.locator('#canvas-surface').bounding_box()})
            if evidence:
                folder=Path(evidence)/source.suffix[1:];folder.mkdir(parents=True,exist_ok=True);page.screenshot(path=str(folder/f'{width}-{theme}.png'))
    if evidence:(Path(evidence)/source.suffix[1:]/'RESPONSIVE.json').write_text(json.dumps(measurements,ensure_ascii=False,indent=2),encoding='utf8')
    assert errors==[]


@pytest.mark.parametrize('native_browser',['.xlsx'],indirect=True)
def test_unsupported_workbook_keeps_existing_readonly_view(native_browser, monkeypatch):
    page,api,source,errors=native_browser
    book=load_workbook(source);book['实验']['A1'].hyperlink='https://example.invalid/source';protected=source.with_name('含链接数据.xlsx');book.save(protected)
    api._app.window=SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(protected)])
    calls=[];original=api.canvas_open
    def opened(*args):
        calls.append(args)
        return original(*args)
    monkeypatch.setattr(api,'canvas_open',opened)
    page.locator('#floating-choose').click();page.get_by_role('button',name='同意读取并打开',exact=True).click()
    page.wait_for_function("document.getElementById('floating-status').textContent.includes('原有结构视图')")
    assert not page.locator('#floating-root').evaluate("node=>node.classList.contains('office-active')")
    assert '120' in page.locator('#floating-document').text_content() and not api._floating.canvas.active()
    page.wait_for_timeout(1200);assert len(calls)==1 and errors==[]
