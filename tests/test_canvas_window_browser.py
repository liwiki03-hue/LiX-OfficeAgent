import json
import os
from pathlib import Path
from types import SimpleNamespace

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_paged_browser import serve_static
from tests.test_professional_browser import EDGE


@pytest.fixture
def window_browser(tmp_path, monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source = tmp_path / '锚点与公式回归.xlsx'
    book = Workbook(); sheet = book.active; sheet.title = '变行高'; sheet.freeze_panes = 'B3'
    for row in range(1, 121):
        sheet.row_dimensions[row].height = 58 if row % 4 == 0 else 24
        for col in range(1, 41):
            sheet.cell(row, col, f'记录 {row} / {col}')
    for col in range(1, 41):
        sheet.column_dimensions[get_column_letter(col)].width = 25 if col % 3 == 0 else 14
    sheet['B2'] = '  Ab中文  '
    sheet['C2'] = '=TRIM(B2)'
    sheet['D2'] = '=NUMBERVALUE("1.234,56",",",".")'
    sheet['D2'].number_format = '#,##0.00'
    book.create_sheet('另一页')['A1'] = '另一页'
    book.save(source)
    monkeypatch.setenv('TEMP', str(tmp_path)); monkeypatch.setenv('TMP', str(tmp_path))
    app = WebviewApp(); app.attach(SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(source)]))
    api = FloatingApi(app, settings_root=tmp_path)
    calls = []
    def invoke(name, args):
        calls.append([name, args])
        return getattr(api, name)(*args)
    with serve_static() as url, sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1100, 'height': 740}); page.set_default_timeout(15000)
            page.expose_function('testHost', invoke)
            page.add_init_script("""if(window===window.top){
                window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};
                window.__delay=false;window.__held=null;
                window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args).then(result=>{
                  if(name==='canvas_view'&&window.__delay){window.__delay=false;return new Promise(resolve=>window.__held=()=>resolve(result));}return result;
                })})};
                document.addEventListener('DOMContentLoaded',()=>{const node=document.getElementById('workspace');node.hidden=false;node.inert=false;node.setAttribute('aria-hidden','false');});
            }""")
            page.route('**/*', lambda route: route.continue_() if route.request.url.startswith(url) else route.abort())
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            errors = []; page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto(url)
            page.locator('#floating-choose').click()
            page.get_by_role('button', name='同意读取并打开', exact=True).click()
            page.locator('.canvas-grid button[data-target="B28"]').wait_for()
            yield page, api, source, calls, errors
        finally:
            browser.close()


def stable(page, position):
    page.wait_for_function('(position)=>document.getElementById("canvas-window").value===position&&!document.getElementById("canvas-state").textContent.includes("正在")', arg=position)
    page.wait_for_timeout(150)


def scroll_and_measure(page, target, horizontal=False, reverse=False):
    return page.locator('#canvas-surface').evaluate('''(node,args)=>{
        const [target,horizontal,reverse]=args;
        if(horizontal)node.scrollLeft=reverse?0:node.scrollWidth;else node.scrollTop=reverse?0:node.scrollHeight;
        const cell=node.querySelector('button[data-target="'+target+'"]'),box=cell.getBoundingClientRect(),surface=node.getBoundingClientRect();
        return {top:box.top-surface.top,left:box.left-surface.left};
    }''', [target, horizontal, reverse])


def relative_box(page, target):
    return page.locator('.canvas-grid button[data-target="'+target+'"]').evaluate('''node=>{
        const box=node.getBoundingClientRect(),surface=document.getElementById('canvas-surface').getBoundingClientRect();
        return {top:box.top-surface.top,left:box.left-surface.left};
    }''')


def test_measured_vertical_horizontal_anchor_prefetch_and_frozen_budget(window_browser):
    page, api, source, calls, errors = window_browser; original = source.read_bytes()
    assert page.locator('.canvas-grid button[data-target="C2"]').text_content() == 'Ab中文'
    assert page.locator('.canvas-grid button[data-target="D2"]').text_content() == '1,234.56'
    page.locator('#canvas-surface').evaluate('node=>node.scrollTop=node.scrollHeight-node.clientHeight-100')
    page.wait_for_function('()=>document.getElementById("canvas-window").value==="A1"')
    page.wait_for_timeout(250)
    assert any(name == 'canvas_prefetch' for name, args in calls)
    assert not any(name == 'canvas_view' and args[3] == 11 for name, args in calls)
    before = scroll_and_measure(page, 'B28')
    stable(page, 'A11')
    after = relative_box(page, 'B28')
    assert abs(before['top']-after['top']) < 2
    horizontal_before = scroll_and_measure(page, 'O28', horizontal=True)
    page.wait_for_function('()=>document.getElementById("canvas-window").value!=="A11"&&!document.getElementById("canvas-state").textContent.includes("正在")')
    stable(page, page.locator('#canvas-window').input_value())
    horizontal_after = relative_box(page, 'O28')
    assert abs(horizontal_before['left']-horizontal_after['left']) < 2
    previous_position = page.locator('#canvas-window').input_value()
    reverse_before = scroll_and_measure(page, 'M28', horizontal=True, reverse=True)
    page.wait_for_function('(old)=>document.getElementById("canvas-window").value!==old', arg=previous_position)
    stable(page, page.locator('#canvas-window').input_value())
    reverse_after = relative_box(page, 'M28')
    assert abs(reverse_before['left']-reverse_after['left']) < 2
    up_before = scroll_and_measure(page, 'M14', reverse=True)
    page.wait_for_function('()=>!document.getElementById("canvas-window").value.endsWith("11")')
    stable(page, page.locator('#canvas-window').input_value())
    up_after = relative_box(page, 'M14')
    assert abs(up_before['top']-up_after['top']) < 2
    assert page.locator('.canvas-grid tbody tr').count() == 30
    assert page.locator('.canvas-grid thead th').count() == 17
    assert page.locator('td[data-target="A1"]').get_attribute('data-frozen-row') == 'true'
    assert page.locator('td[data-target="A1"]').get_attribute('data-frozen-col') == 'true'
    assert source.read_bytes() == original and api._floating.canvas.revision == 0 and not errors
    evidence = os.environ.get('L4_WINDOW_EVIDENCE')
    if evidence:
        Path(evidence).mkdir(parents=True, exist_ok=True)
        (Path(evidence) / 'ANCHORS.json').write_text(json.dumps({'vertical_before':before,'vertical_after':after,'horizontal_before':horizontal_before,'horizontal_after':horizontal_after,'left_before':reverse_before,'left_after':reverse_after,'up_before':up_before,'up_after':up_after,'request_names':[name for name, _ in calls]}, ensure_ascii=False, indent=2), encoding='utf8')


def test_late_window_response_keeps_new_input_and_next_document(window_browser):
    page, api, source, calls, errors = window_browser
    page.locator('.canvas-grid button[data-target="B2"]').click()
    page.evaluate('window.__delay=true')
    page.locator('#canvas-surface').evaluate('node=>node.scrollTop=node.scrollHeight')
    page.wait_for_function('window.__held!==null')
    page.locator('#canvas-value').fill('保留新输入')
    page.evaluate('window.__held();window.__held=null')
    page.wait_for_timeout(250)
    assert page.locator('#canvas-value').input_value() == '保留新输入'
    assert page.locator('#canvas-window').input_value() == 'A1'
    page.locator('#canvas-value').press('Escape')
    page.evaluate('window.__delay=true')
    page.locator('#canvas-window').fill('A41'); page.locator('#canvas-window').press('Enter')
    page.wait_for_function('window.__held!==null')
    assert page.evaluate('async()=>{const result=await window.pywebview.api.floating_close_document();window.OfficeNative.close();document.getElementById("floating-input").value="新的对话草稿";window.__held();window.__held=null;return result.ok;}')
    page.wait_for_timeout(250)
    assert page.locator('#canvas-work').count() == 0
    assert page.locator('#floating-input').input_value() == '新的对话草稿'
    assert not errors


def test_last_window_and_reverse_scroll_retain_overlap(window_browser):
    page, api, source, calls, errors = window_browser
    page.locator('#canvas-window').fill('A9961'); page.locator('#canvas-window').press('Enter')
    stable(page, 'A9961')
    scroll_and_measure(page, 'B9985')
    stable(page, 'A9971')
    scroll_and_measure(page, 'B9995')
    stable(page, 'A9973')
    assert page.locator('button[data-target="B10000"]').count() == 1
    count = len([row for row in calls if row[0] == 'canvas_view'])
    page.locator('#canvas-surface').evaluate('node=>node.scrollTop=node.scrollHeight')
    page.wait_for_timeout(250)
    assert len([row for row in calls if row[0] == 'canvas_view']) == count
    scroll_and_measure(page, 'B9973', reverse=True)
    stable(page, 'A9963')
    assert not errors
