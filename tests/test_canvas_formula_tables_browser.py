import base64
import json
import os
from pathlib import Path
import shutil
import time
from types import SimpleNamespace

from openpyxl import load_workbook
from pptx import Presentation
import pytest

from office_agent.office_canvas import view
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_canvas_formula_tables import table_fixture
from tests.test_office_canvas import fixtures
from tests.test_paged_browser import serve_static
from tests.test_professional_browser import EDGE


@pytest.mark.parametrize('kind',['xlsx','pptx'])
def test_formula_or_table_real_ui_copy_reopen_and_rollback(tmp_path, monkeypatch, kind):
    if not EDGE.is_file():pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source=fixtures(tmp_path)[0] if kind=='xlsx' else table_fixture(tmp_path)
    original=source.read_bytes();monkeypatch.setenv('TEMP',str(tmp_path));monkeypatch.setenv('TMP',str(tmp_path))
    app=WebviewApp();app.attach(SimpleNamespace(create_file_dialog=lambda *args,**kw:[str(source)]));api=FloatingApi(app,settings_root=tmp_path)
    evidence=os.environ.get('OFFICE_CANVAS_NEXT_EVIDENCE');folder=Path(evidence)/kind if evidence else None
    if folder:folder.mkdir(parents=True,exist_ok=True);shutil.copy2(source,folder/source.name)
    with serve_static() as url,sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            page=browser.new_page(viewport={'width':1280,'height':820});page.set_default_timeout(60000)
            page.expose_function('testHost',lambda name,args:getattr(api,name)(*args))
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js',lambda route:route.fulfill(status=200,body=''))
            errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
            page.goto(url);page.locator('#floating-choose').click();page.get_by_role('button',name='同意读取并打开',exact=True).click();page.locator('#canvas-items button').first.wait_for()
            if kind=='pptx':page.locator('#canvas-slide img').wait_for()
            if folder:
                page.screenshot(path=str(folder/'before.png'))
                if kind=='pptx':(folder/'page-before.png').write_bytes(base64.b64decode(page.locator('#canvas-slide img').get_attribute('src').split(',',1)[1]))
            service=api._floating.canvas
            if kind=='xlsx':
                page.locator('.canvas-grid button[data-target="C2"]').dblclick()
                assert page.locator('#canvas-type').input_value()=='formula'
                page.locator('#canvas-value').fill('=B2/0');page.locator('#canvas-value').press('Enter')
                page.wait_for_function("document.getElementById('canvas-feedback').textContent.includes('除零')")
                assert service.revision==0 and page.locator('#canvas-value').input_value()=='=B2/0'
                page.locator('#canvas-value').fill('=SUM(B2,30)');page.locator('#canvas-value').press('Enter')
                page.wait_for_function("document.querySelector('.canvas-grid button[data-target=C2]').textContent==='150'")
                page.locator('.canvas-grid button[data-target="B2"]').dblclick();page.locator('#canvas-value').fill('125');page.locator('#canvas-value').press('Enter')
                page.wait_for_function("document.querySelector('.canvas-grid button[data-target=C2]').textContent==='155'")
                assert service.revision==2
                page.get_by_label('刷新页图',exact=True).click();printing=page.get_by_role('dialog',name='实际打印页预览',exact=True);printing.locator('img').wait_for()
                if folder:(folder/'print-after.png').write_bytes(base64.b64decode(printing.locator('img').get_attribute('src').split(',',1)[1]))
                printing.get_by_role('button',name='关闭',exact=True).click();printing.wait_for(state='detached')
            else:
                target=next(run for run in view(original,'.pptx')['runs'] if run['text']=='120')
                page.get_by_role('button',name='编辑片段 '+target['id']+' 120',exact=True).click()
                assert page.locator('#canvas-position').inner_text()=='表格1 · 第2行第2列'
                page.locator('#canvas-value').fill('115');page.locator('#canvas-value').press('Enter')
                page.get_by_role('button',name='编辑片段 '+target['id']+' 115',exact=True).wait_for()
                page.wait_for_function("!document.getElementById('canvas-state').textContent.includes('正在')")
                page.locator('#canvas-italic').click();page.wait_for_function("document.getElementById('canvas-italic').getAttribute('aria-pressed')==='true'")
                page.get_by_label('刷新页图',exact=True).click();page.locator('#canvas-slide img').wait_for();page.wait_for_function("!document.getElementById('canvas-state').textContent.includes('正在')")
            page.get_by_label('保存副本',exact=True).click();confirm=page.get_by_role('dialog',name='确认保存副本',exact=True)
            assert ('受限计算' if kind=='xlsx' else '表格1') in confirm.inner_text()
            if kind=='xlsx':assert '155' in confirm.inner_text()
            confirm.get_by_role('button',name='取消',exact=True).click();confirm.wait_for(state='detached');assert not list(tmp_path.glob('*.agent-*'))
            page.get_by_label('保存副本',exact=True).click();confirm.get_by_role('button',name='确认生成副本',exact=True).click();confirm.wait_for(state='detached')
            deadline=time.monotonic()+90
            while service.publishing:
                assert time.monotonic()<deadline,'bounded publish';page.wait_for_timeout(100)
            assert not service.uncertain and api._floating.outputs,service.last_error
            receipt=next(iter(api._floating.outputs.values()));output=Path(receipt['output'])
            assert source.read_bytes()==original and receipt['status']=='verified'
            if kind=='xlsx':
                book=load_workbook(output);assert book['实验']['B2'].value==125 and book['实验']['C2'].value=='=SUM(B2,30)'
                assert load_workbook(output,data_only=True)['实验']['C2'].value is None
            else:
                table=Presentation(output).slides[0].shapes[-1].table
                run=table.cell(1,1).text_frame.paragraphs[0].runs[0]
                assert run.text=='115' and run.font.bold and run.font.italic and table.cell(1,1).text=='115 ms'
            page.wait_for_function("document.getElementById('canvas-state').textContent==='副本已验证'")
            if folder:
                shutil.copy2(output,folder/('编辑副本'+source.suffix));shutil.copy2(receipt['backup'],folder/('BACKUP'+source.suffix))
                (folder/'AUDIT.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2),encoding='utf8')
                if kind=='pptx':
                    (folder/'page-after.png').write_bytes(base64.b64decode(page.locator('#canvas-slide img').get_attribute('src').split(',',1)[1]))
                    page.locator('#canvas-surface').evaluate('node=>node.scrollTop=300')
                else:page.locator('#canvas-surface').evaluate('node=>node.scrollTop=0')
                page.screenshot(path=str(folder/'after.png'))
            result=app.editor.rollback(receipt['id'],confirmed=True);assert result['status']=='rolled_back' and not output.exists()
            if folder:(folder/'ROLLBACK.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
            assert errors==[]
        finally:browser.close()
