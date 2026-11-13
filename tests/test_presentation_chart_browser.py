import json
import os
from pathlib import Path
import re
import shutil
import time
from types import SimpleNamespace

from docx import Document
from pptx import Presentation
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.presentation_drafts import validate_deck
from tests.test_paged_browser import serve_static,engine_frame
from tests.test_presentation_browser import open_source,wait_ready
from tests.test_model_connections import connection,payload,save
from tests.test_professional_browser import EDGE


@pytest.fixture
def chart_browser(tmp_path,monkeypatch):
    if not EDGE.is_file():pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source=tmp_path/'自建冷却实验.docx'
    book=Document();book.add_heading('循环冷却实验',0)
    for text in ['实验于2025年完成，温度120 ℃。','目的：核对实验耗时。','条件：相同样本、温度和测量方法。','边界：本次测量不代表其他条件。','结论：仅按本表比较。','不纳入：背景介绍，不作为图表依据。']:book.add_paragraph(text)
    table=book.add_table(rows=4,cols=2)
    for row,values in zip(table.rows,[['方案','耗时'],['基线','120'],['优化','105.5'],['复核','98.25']]):
        for cell,value in zip(row.cells,values):cell.text=value
    book.add_paragraph('读者说明：单位为ms。');book.save(source)
    app=WebviewApp();app.attach(SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(source)]))
    api=FloatingApi(app,settings_root=tmp_path);save(api._connections,payload([connection()]))
    monkeypatch.setenv('TEMP',str(tmp_path));monkeypatch.setenv('TMP',str(tmp_path))
    with serve_static() as url,sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            page=browser.new_page(viewport={'width':1280,'height':820});page.set_default_timeout(20000)
            page.expose_function('testHost',lambda name,args:getattr(api,name)(*args))
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.confirm=()=>true;document.addEventListener('DOMContentLoaded',()=>{const node=document.getElementById('workspace');node.hidden=false;node.inert=false;node.setAttribute('aria-hidden','false')})}""")
            page.route('**/auth.js',lambda route:route.fulfill(status=200,body=''))
            errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
            page.goto(url);page.locator('#floating-choose').click();page.get_by_role('button',name='同意读取并打开',exact=True).click()
            deadline=time.monotonic()+70
            while not api._floating.paged.manifest:
                if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
                page.wait_for_timeout(100)
            yield page,api,source,errors
        finally:browser.close()


def test_select_current_material_chart_render_publish_cancel_and_rollback(chart_browser):
    page,api,source,errors=chart_browser;original=source.read_bytes();frame=engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End');page.keyboard.insert_text('实验于2026年完成，温度120 ℃。')
    node=open_source(page);node.get_by_label('演示标题',exact=True).fill('冷却实验数据演示')
    node.get_by_role('button',name='选择材料',exact=True).click()
    node.get_by_role('checkbox',name=re.compile('第6块')).uncheck();node.get_by_role('checkbox',name=re.compile('第7块')).uncheck()
    node.get_by_text('已选 7 / 9 块 · 原正文共 9 块',exact=True).wait_for()
    evidence=os.environ.get('OFFICE_PRESENTATION_CHART_EVIDENCE')
    if evidence:
        folder=Path(evidence);folder.mkdir(parents=True,exist_ok=True);page.screenshot(path=str(folder/'material-selection.png'))
    node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for()
    selected=api._presentations.source
    assert selected['unselected_blocks']==[6,7] and len(selected['segments'])==7 and '2026' in selected['segments'][1]['text']
    work.get_by_role('button',name='7. 普通表格 8',exact=True).click()
    work.get_by_label('页面标题',exact=True).fill('同等条件下的实验耗时')
    work.get_by_label('页面版式',exact=True).select_option('chart')
    work.get_by_label('图表单位',exact=True).fill('ms')
    work.get_by_role('checkbox',name='我已核对本页分类、数值和单位',exact=True).check()
    work.get_by_role('button',name='保存本页到大纲',exact=True).click()
    page.wait_for_function("document.querySelector('[data-presentation-title]').textContent.includes('草稿 2')")
    work.get_by_role('button',name='生成实际预览',exact=True).click();wait_ready(page,api)
    assert api._presentations.rendered['report']['editable_charts']==1
    assert all(row['matched_blocks']==row['expected_blocks'] for row in api._presentations.rendered['report']['text_checks'])
    if evidence:page.screenshot(path=str(folder/'chart-preview.png'))
    api._app.window=SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(source.parent)])
    work.get_by_role('button',name='发布PPTX',exact=True).click()
    confirm=page.get_by_role('dialog',name='确认发布演示稿',exact=True)
    confirm.get_by_role('button',name='取消',exact=True).click();confirm.wait_for(state='detached')
    assert not list(source.parent.glob('.office-presentation-*'))
    work.get_by_role('button',name='发布PPTX',exact=True).click();confirm.get_by_role('button',name='确认发布PPTX',exact=True).click();confirm.wait_for(state='detached')
    saved=next(iter(api._presentations.saved.values()));path=Path(saved['path'])
    assert validate_deck(path.read_bytes(),api._presentations.spec)['editable_charts']==1 and source.read_bytes()==original
    book=Presentation(path);chart=next(shape.chart for slide in book.slides for shape in slide.shapes if shape.has_chart)
    assert list(chart.series[0].values)==[120,105.5,98.25]
    if evidence:
        page.screenshot(path=str(folder/'chart-published.png'))
        for artifact,name in [(source,'自建冷却实验.docx'),(path,'冷却实验数据演示.pptx'),(path.parent/'backup','SOURCE_SPEC_BACKUP.json'),(path.parent/'audit-verified.json','PUBLISH_AUDIT.json')]:shutil.copyfile(artifact,folder/name)
        for index,image in enumerate(api._presentations.rendered['previews']):(folder/f'slide-{index+1}.png').write_bytes(image)
        (folder/'SPEC.json').write_text(json.dumps(api._presentations.spec,ensure_ascii=False,indent=2),encoding='utf8')
        (folder/'VERIFICATION.json').write_text(json.dumps({'source':selected,'report':api._presentations.rendered['report'],'source_unchanged':source.read_bytes()==original,'saved_path':str(path),'native':False,'cloud_calls':0,'expected_chart_values':['120','105.5','98.25'],'unit':'ms'},ensure_ascii=False,indent=2),encoding='utf8')
    work.get_by_role('button',name='撤销本次发布',exact=True).click()
    deadline=time.monotonic()+10
    while path.exists():
        if time.monotonic()>deadline:raise AssertionError('rollback did not remove owned output')
        page.wait_for_timeout(100)
    if evidence:shutil.copyfile(path.parent/'audit-removed.json',folder/'ROLLBACK_AUDIT.json')
    assert source.read_bytes()==original and not errors


def test_chart_data_change_clears_review_and_materials_stay_reachable(chart_browser):
    page,api,source,errors=chart_browser
    node=open_source(page);node.get_by_label('页数（2—12）',exact=True).fill('2')
    node.get_by_role('button',name='选择材料',exact=True).click()
    boxes=node.get_by_role('checkbox');boxes.first.wait_for()
    for index in range(boxes.count()):
        if index!=7:boxes.nth(index).uncheck()
    matrix=[]
    for width in [1440,1280,1024,520]:
        page.set_viewport_size({'width':width,'height':800})
        for theme in ['light','dark','highcontrast']:
            page.evaluate('(value)=>document.documentElement.dataset.theme=value',theme)
            submit=node.get_by_role('button',name='整理大纲',exact=True);submit.scroll_into_view_if_needed()
            assert submit.is_visible() and not node.evaluate('node=>node.scrollWidth>node.clientWidth+2')
            box=node.bounding_box();assert abs(box['x']+box['width']/2-width/2)<2
            matrix.append({'width':width,'theme':theme,'material_controls_reachable':True})
    node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for()
    work.locator('[data-presentation-slides] button').nth(1).click();work.get_by_label('页面版式',exact=True).select_option('chart')
    work.get_by_label('图表单位',exact=True).fill('ms');review=work.get_by_role('checkbox',name='我已核对本页分类、数值和单位',exact=True);review.check()
    work.get_by_label('图表单位',exact=True).fill('s');assert not review.is_checked()
    work.get_by_label('图表单位',exact=True).fill('ms');review.check()
    rows=work.get_by_label('表格（列用Tab分隔，行用换行分隔）',exact=True);rows.fill('方案\t耗时\n基线\t')
    assert not review.is_checked();work.get_by_role('button',name='保存本页到大纲',exact=True).click()
    page.locator('[data-presentation-status]').filter(has_text='校验未通过').wait_for()
    assert rows.input_value().endswith('\t') and api._presentations.spec['revision']==1
    work.get_by_role('button',name='关闭',exact=True).click();work.wait_for(state='detached')
    evidence=os.environ.get('OFFICE_PRESENTATION_CHART_EVIDENCE')
    if evidence:(Path(evidence)/'RESPONSIVE.json').write_text(json.dumps(matrix,indent=2),encoding='utf8')
    assert not errors and not list(source.parent.glob('.office-presentation-*'))


def test_late_material_response_cannot_restore_previous_text_or_privacy(chart_browser):
    page,api,source,errors=chart_browser
    node=open_source(page);node.get_by_label('材料来源',exact=True).select_option('text')
    text=node.get_by_label('材料（按段落换行）',exact=True);text.fill('初始材料。')
    page.evaluate("""()=>{const host=window.testHost;window.testHost=(name,args)=>name==='presentation_materials'?new Promise((resolve,reject)=>{window.releaseMaterials=()=>Promise.resolve(host(name,args)).then(resolve,reject)}):host(name,args)}""")
    node.get_by_role('button',name='选择材料',exact=True).click()
    text.fill('后续材料，保留这一段。');page.evaluate('()=>window.releaseMaterials()')
    node.get_by_role('alert').filter(has_text='材料输入已改变').wait_for()
    assert text.input_value()=='后续材料，保留这一段。' and node.get_by_role('checkbox').count()==0
    node.get_by_role('button',name='选择材料',exact=True).click()
    page.evaluate("()=>window.dispatchEvent(new CustomEvent('officeauth:unauthenticated'))")
    page.evaluate('()=>window.releaseMaterials()');page.wait_for_timeout(100)
    assert page.get_by_role('dialog',name='新建演示稿',exact=True).count()==0
    assert page.locator('.presentation-material-option').count()==0 and not errors
