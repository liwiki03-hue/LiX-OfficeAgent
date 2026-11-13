from pathlib import Path
from types import SimpleNamespace

from pptx import Presentation
from pptx.util import Inches
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.writes import digest
from tests.test_paged_browser import serve_static
from tests.test_professional_browser import EDGE


@pytest.fixture
def ppt_browser(tmp_path, monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge for isolated browser regression')
    from playwright.sync_api import sync_playwright
    source = tmp_path / '页图异步回归.pptx'
    book = Presentation()
    for page_number in range(3):
        slide = book.slides.add_slide(book.slide_layouts[6])
        for index in range(16):
            box = slide.shapes.add_textbox(Inches(.5), Inches(.2 + index * .3), Inches(8), Inches(.3))
            paragraph = box.text_frame.paragraphs[0]
            paragraph.add_run().text = f'第{page_number + 1}页片段{index} 2025年 120℃'
            if index == 0:
                paragraph.add_run().text = ' 保留第二种格式'
    book.save(source)
    monkeypatch.setenv('TEMP', str(tmp_path));monkeypatch.setenv('TMP', str(tmp_path))
    app = WebviewApp();app.attach(SimpleNamespace(create_file_dialog=lambda *a, **kw: [str(source)]))
    api = FloatingApi(app, settings_root=tmp_path)
    calls = []
    def preview(doc, revision, item, page):
        calls.append([doc, revision, item])
        return {'ok': True, 'document_id': doc, 'draft_revision': revision, 'item_id': item,
                'image': 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
                'renderer': 'controlled preview response; not a rendered Office page', 'pages': 3, 'page': page}
    monkeypatch.setattr(api, 'canvas_preview', preview)
    with serve_static() as url, sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1100, 'height': 740});page.set_default_timeout(12000)
            page.expose_function('testHost', lambda name, args: getattr(api, name)(*args))
            page.add_init_script("""if(window===window.top){
              window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};
              window.__hold=true;window.__held=null;window.__fail=false;
              window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args).then(result=>{
                if(name==='canvas_preview'&&window.__fail){window.__fail=false;return {ok:false,error:'受控渲染失败'};}
                if(name==='canvas_preview'&&window.__hold){window.__hold=false;return new Promise(resolve=>window.__held=()=>{window.__held=null;resolve(result);});}return result;
              })})};
              document.addEventListener('DOMContentLoaded',()=>{const node=document.getElementById('workspace');node.hidden=false;node.inert=false;node.setAttribute('aria-hidden','false');});
            }""")
            page.route('**/*', lambda route: route.continue_() if route.request.url.startswith(url) else route.abort())
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            errors = [];page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto(url);page.locator('#floating-choose').click()
            page.get_by_role('button', name='同意读取并打开', exact=True).click()
            page.wait_for_function('window.__held!==null')
            yield page, api, source, calls, errors
        finally:
            browser.close()


def test_pending_preview_keeps_scroll_selection_input_and_real_draft(ppt_browser):
    page, api, source, calls, errors = ppt_browser
    before = digest(source)
    page.get_by_role('button', name='编辑片段 run-10 第1页片段9 2025年 120℃', exact=True).click()
    dialog = page.get_by_role('dialog', name='编辑本页文字', exact=True)
    assert dialog.is_visible() and page.locator('#canvas-value').count() == 1
    page.locator('#canvas-value').fill('第1页片段9 2026年 120℃')
    scroll = page.locator('#canvas-surface').evaluate('(node)=>node.scrollTop')
    page.evaluate('window.__held()')
    page.wait_for_function('document.getElementById("canvas-slide").dataset.revision==="0"')
    assert page.locator('#canvas-value').input_value() == '第1页片段9 2026年 120℃'
    assert page.locator('#canvas-surface').evaluate('(node)=>node.scrollTop') == scroll
    assert page.locator('.canvas-runs [data-target="run-10"]').get_attribute('aria-selected') == 'true'
    page.locator('#canvas-value').press('Enter');dialog.wait_for(state='detached')
    page.wait_for_function('document.getElementById("canvas-state").textContent.includes("未保存")')
    service = api._floating.canvas
    assert service.revision == 1 and '2026年' in service.view(service.document_id, 1, 'slide-0')['view']['runs'][10]['text']
    assert page.locator('#canvas-slide-image').is_hidden() and '尚未更新' in page.locator('#canvas-slide-status').text_content()
    assert len(calls) == 1 and digest(source) == before and errors == []
    page.get_by_label('刷新页图', exact=True).click()
    page.wait_for_function('document.getElementById("canvas-feedback").textContent.includes("页图对应草稿 1")')
    assert page.locator('#canvas-slide').get_attribute('data-revision') == '1'
    assert api._floating.canvas.revision == 1 and digest(source) == before


def test_preview_completion_does_not_clear_later_feedback(ppt_browser):
    page, api, source, calls, errors = ppt_browser
    page.evaluate('document.getElementById("canvas-feedback").textContent="保存失败，请核对输出";window.__held()')
    page.wait_for_function('document.getElementById("canvas-slide").dataset.revision==="0"')
    assert page.locator('#canvas-feedback').inner_text() == '保存失败，请核对输出'
    assert len(calls) == 1 and errors == []


def test_late_preview_does_not_replace_new_slide_and_only_latest_is_requested(ppt_browser):
    page, api, source, calls, errors = ppt_browser
    page.locator('#canvas-items button').nth(1).click()
    page.wait_for_function('document.querySelector("#canvas-items [aria-current=true]").textContent.includes("2")')
    page.locator('#canvas-items button').nth(2).click()
    page.wait_for_function('document.querySelector("#canvas-items [aria-current=true]").textContent.includes("3")')
    assert len(calls) == 1
    page.get_by_role('button', name='编辑片段 run-2 第3页片段1 2025年 120℃', exact=True).click()
    page.locator('#canvas-value').fill('第三页输入必须保留')
    page.evaluate('window.__held()')
    page.wait_for_function('document.getElementById("canvas-slide").dataset.revision==="0"')
    assert [item[2] for item in calls] == ['slide-0', 'slide-2']
    assert page.locator('#canvas-value').input_value() == '第三页输入必须保留'
    assert '3' in page.locator('#canvas-slide-image').get_attribute('alt')
    assert errors == []


def test_hitbox_group_requires_explicit_run_and_failure_is_manual_retry(ppt_browser):
    page, api, source, calls, errors = ppt_browser
    page.evaluate('window.__held()')
    page.locator('.canvas-hit').first.wait_for()
    page.locator('.canvas-hit').first.click()
    dialog = page.get_by_role('dialog', name='编辑本页文字', exact=True)
    assert page.locator('#canvas-value').get_attribute('readonly') is not None
    dialog.get_by_role('button', name='保留第二种格式', exact=True).click()
    page.locator('#canvas-value').fill('第二片段的新输入')
    dialog.get_by_role('button', name='返回页图', exact=True).click()
    page.evaluate('window.__fail=true')
    page.get_by_label('刷新页图', exact=True).click()
    page.wait_for_function('document.getElementById("canvas-slide-status").textContent.includes("受控渲染失败")')
    page.wait_for_timeout(150)
    assert len(calls) == 2 and page.locator('#canvas-value').input_value() == '第二片段的新输入'
    page.get_by_label('刷新页图', exact=True).click()
    page.wait_for_function('!document.getElementById("canvas-slide-status").textContent.includes("失败")')
    assert len(calls) == 3 and page.locator('#canvas-value').input_value() == '第二片段的新输入'
    assert errors == []


def test_close_discards_inflight_image_and_preserves_new_composer(ppt_browser):
    page, api, source, calls, errors = ppt_browser
    page.evaluate('window.OfficeNative.close();document.getElementById("floating-input").value="下一任务要求";window.__held()')
    page.wait_for_timeout(200)
    assert page.locator('#canvas-slide').count() == 0
    assert page.locator('#floating-input').input_value() == '下一任务要求'
    assert len(calls) == 1 and errors == []


def test_narrow_editor_keeps_single_input_and_composition_does_not_apply(ppt_browser):
    page, api, source, calls, errors = ppt_browser
    page.set_viewport_size({'width': 520, 'height': 740})
    page.get_by_role('button', name='编辑片段 run-2 第1页片段1 2025年 120℃', exact=True).click()
    field = page.locator('#canvas-value');field.fill('中文组合输入')
    field.dispatch_event('keydown', {'key': 'Enter', 'isComposing': True, 'keyCode': 229})
    assert api._floating.canvas.revision == 0 and field.input_value() == '中文组合输入'
    assert field.count() == 1
    back = page.get_by_role('button', name='返回页图', exact=True)
    box = back.bounding_box()
    assert box['x'] >= 0 and box['x'] + box['width'] <= 520 and box['y'] + box['height'] <= 740
    field.press('Escape')
    assert field.input_value() == '第1页片段1 2025年 120℃'
    page.get_by_role('dialog', name='编辑本页文字', exact=True).wait_for(state='detached')
    assert len(calls) == 1 and errors == []
