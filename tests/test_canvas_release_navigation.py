import json
import os
from pathlib import Path
from types import SimpleNamespace

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_canvas_window_browser import stable
from tests.test_paged_browser import serve_static
from tests.test_professional_browser import EDGE


@pytest.fixture
def release_grid(tmp_path, monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source = tmp_path / '滚动与公式.xlsx'
    book = Workbook(); sheet = book.active; sheet.title = '长表'; sheet.freeze_panes = 'B3'
    for row in range(1, 121):
        sheet.row_dimensions[row].height = 58 if row % 4 == 0 else 24
        for col in range(1, 41):
            sheet.cell(row, col, f'记录{row}/{col}')
    for col in range(1, 41):
        sheet.column_dimensions[get_column_letter(col)].width = 25 if col % 3 == 0 else 14
    sheet['B2'] = 120; sheet['C2'] = '=B2*2'
    sheet['D2'] = '=NUMBERVALUE("1.234,56",",",".")'; sheet['D2'].number_format = '#,##0.00'
    other = book.create_sheet('对照'); other['B2'] = 7; other['C2'] = '=B2*3'
    book.save(source)
    monkeypatch.setenv('TEMP', str(tmp_path)); monkeypatch.setenv('TMP', str(tmp_path))
    app = WebviewApp(); app.attach(SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(source)]))
    api = FloatingApi(app, settings_root=tmp_path); calls = []
    def invoke(name, args):
        calls.append([name, args])
        return getattr(api, name)(*args)
    with serve_static() as url, sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1100, 'height': 740}); page.set_default_timeout(6000)
            page.expose_function('testHost', invoke)
            page.add_init_script('''if(window===window.top){
              window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};
              window.__delay=false;window.__held=null;
              window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args).then(result=>{
                if(name==='canvas_view'&&window.__delay){window.__delay=false;return new Promise(resolve=>window.__held=()=>resolve(result));}return result;
              })})};
              document.addEventListener('DOMContentLoaded',()=>{const node=document.getElementById('workspace');node.hidden=false;node.inert=false;node.setAttribute('aria-hidden','false');});
            }''')
            page.route('**/*', lambda route: route.continue_() if route.request.url.startswith(url) else route.abort())
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            errors = []; page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto(url); page.locator('#floating-choose').click()
            page.get_by_role('button', name='同意读取并打开', exact=True).click()
            page.locator('.canvas-grid button[data-target="C2"]').wait_for(); stable(page, 'A1')
            yield page, api, source, calls
            assert not errors
        finally:
            browser.close()


def save_evidence(page, name, details):
    root = os.environ.get('R1_EVIDENCE')
    if root:
        folder = Path(root); folder.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(folder / (name + '.png')))
        (folder / (name + '.json')).write_text(json.dumps(details, ensure_ascii=False, indent=2), encoding='utf8')


def jump(page, position):
    page.locator('#canvas-window').fill(position); page.locator('#canvas-window').press('Enter'); stable(page, position)


def wheel(page, dx=0, dy=0):
    box = page.locator('#canvas-surface').bounding_box()
    page.mouse.move(box['x'] + box['width'] * .8, box['y'] + box['height'] * .6)
    page.mouse.wheel(dx, dy)
    page.wait_for_timeout(300)


def hit_identity(page, target='C2'):
    return page.evaluate('''target=>{
      const cell=document.querySelector('td[data-target="'+target+'"]'),row=cell.parentElement.querySelector('th');
      const box=cell.getBoundingClientRect(),header=row.getBoundingClientRect();
      const hit=document.elementFromPoint(header.x+header.width/2,box.y+box.height/2);
      return {rowLabel:hit.textContent.trim(),row:cell.parentElement.dataset.row,target:cell.dataset.target,
        text:cell.textContent.trim(),cellTop:box.top,headerTop:header.top,window:document.getElementById('canvas-window').value};
    }''', target)


def test_frozen_visible_header_matches_real_cell(release_grid):
    page, api, source, calls = release_grid; original = source.read_bytes()
    wheel(page, dy=1050); stable(page, 'A11')
    page.locator('#canvas-surface').evaluate('node=>node.scrollLeft=120'); page.wait_for_timeout(150)
    identity = hit_identity(page); save_evidence(page, 'frozen-header', identity)
    assert identity['rowLabel'] == '2', identity
    assert identity['row'] == '2' and identity['target'] == 'C2' and identity['text'] == '240'
    assert abs(identity['cellTop'] - identity['headerTop']) < 1
    page.locator('.canvas-grid button[data-target="C2"]').click()
    assert page.locator('#canvas-position').text_content() == 'C2'
    assert page.locator('#canvas-value').input_value() == '=B2*2'
    assert api._floating.canvas.revision == 0 and source.read_bytes() == original


def test_wheel_at_zero_displacement_can_return_to_start(release_grid):
    page, api, source, calls = release_grid
    jump(page, 'K11')
    assert page.locator('#canvas-surface').evaluate('node=>[node.scrollTop,node.scrollLeft]') == [0, 0]
    wheel(page, dy=-120)
    save_evidence(page, 'boundary-up', {'window': page.locator('#canvas-window').input_value(), 'view_calls': len([c for c in calls if c[0] == 'canvas_view'])})
    stable(page, 'K3')
    positions = []
    for _ in range(8):
        wheel(page, dx=-2000)
        positions.append({'window': page.locator('#canvas-window').input_value(), 'scroll': page.locator('#canvas-surface').evaluate('node=>node.scrollLeft')})
        if page.locator('#canvas-window').input_value() == 'B3':
            break
    save_evidence(page, 'boundary-left', positions)
    stable(page, 'B3')
    wheel(page, dx=-2000); wheel(page, dy=-2000)
    assert page.locator('#canvas-surface').evaluate('node=>[node.scrollTop,node.scrollLeft]') == [0, 0]
    count = len([c for c in calls if c[0] == 'canvas_view'])
    for _ in range(3):
        wheel(page, dy=-120); wheel(page, dx=-120)
    assert len([c for c in calls if c[0] == 'canvas_view']) == count
    assert page.locator('.canvas-grid tbody tr').count() == 30
    assert page.locator('.canvas-grid thead th').count() == 17


@pytest.mark.parametrize('width,zoom', [(420, 1), (1100, 1.25), (1100, 1.5)])
def test_frozen_intersections_resizing_and_zoom(release_grid, width, zoom):
    page, api, source, calls = release_grid
    page.set_viewport_size({'width': width, 'height': 800})
    page.evaluate('(zoom)=>document.documentElement.style.zoom=zoom', zoom)
    jump(page, 'A11')
    page.locator('#canvas-surface').evaluate('node=>{node.scrollTop=450;node.scrollLeft=120}')
    page.wait_for_timeout(250)
    identity = hit_identity(page)
    assert identity['rowLabel'] == '2', identity
    assert abs(identity['cellTop'] - identity['headerTop']) < 1
    headers = page.evaluate('''()=>{
      const surface=document.getElementById('canvas-surface').getBoundingClientRect();
      const frozen=document.querySelector('td[data-target="A2"]').getBoundingClientRect();
      const result=[];
      for(const node of document.querySelectorAll('thead th[data-column]')){
        const box=node.getBoundingClientRect(),x=Math.max(box.left,frozen.right)+4,y=box.top+box.height/2;
        if(x>=Math.min(box.right,surface.right)-2)continue;
        const hit=document.elementFromPoint(x,y)?.closest('th');
        result.push({column:node.dataset.column,actual:hit?.dataset.column,label:hit?.textContent});
      }
      const corner=document.querySelector('thead th').getBoundingClientRect();
      const cornerHit=document.elementFromPoint(corner.x+corner.width/2,corner.y+corner.height/2);
      return {columns:result,corner:cornerHit===document.querySelector('thead th')};
    }''')
    assert headers['corner'] and headers['columns'], headers
    assert all(h['column'] == h['actual'] and h['label'] == get_column_letter(int(h['column'])) for h in headers['columns'])
    save_evidence(page, f'layout-{width}-{zoom}', {'identity': identity, 'headers': headers, 'zoom': 'CSS simulation, not physical DPI'})
    assert api._floating.canvas.revision == 0


def dispatch_wheel(page, **kwargs):
    return page.locator('#canvas-surface').evaluate('''(node,args)=>{
      const event=new WheelEvent('wheel',{bubbles:true,cancelable:true,...args});
      node.dispatchEvent(event);return event.defaultPrevented;
    }''', kwargs)


def test_wheel_modes_shift_and_true_end_limits(release_grid):
    page, api, source, calls = release_grid
    jump(page, 'K21')
    assert dispatch_wheel(page, deltaY=-3, deltaMode=1)
    stable(page, 'K11')
    assert dispatch_wheel(page, deltaY=-1, deltaMode=2, shiftKey=True)
    page.wait_for_timeout(300)
    assert page.locator('#canvas-window').input_value() != 'K11'
    jump(page, 'IH9973')
    page.locator('#canvas-surface').evaluate('node=>{node.scrollTop=node.scrollHeight;node.scrollLeft=node.scrollWidth}')
    page.wait_for_timeout(200)
    count = len([c for c in calls if c[0] == 'canvas_view'])
    assert not dispatch_wheel(page, deltaY=100)
    assert not dispatch_wheel(page, deltaX=100)
    page.wait_for_timeout(250)
    assert len([c for c in calls if c[0] == 'canvas_view']) == count
    assert page.locator('button[data-target="IV10000"]').count() == 1
    wheel(page, dy=-1000)
    wheel(page, dy=-1000)
    assert not page.locator('#canvas-window').input_value().endswith('9973')


def test_queued_reverse_is_bounded_and_new_input_is_not_replaced(release_grid):
    page, api, source, calls = release_grid
    jump(page, 'K21')
    page.evaluate('window.__delay=true')
    assert dispatch_wheel(page, deltaY=-120)
    page.wait_for_function('window.__held!==null')
    page.locator('#canvas-surface').evaluate('node=>node.scrollTop=node.scrollHeight')
    for _ in range(15):
        dispatch_wheel(page, deltaY=-2)
        dispatch_wheel(page, deltaY=2)
    count = len([c for c in calls if c[0] == 'canvas_view'])
    page.evaluate('window.__held();window.__held=null')
    page.wait_for_timeout(500)
    assert page.locator('#canvas-window').input_value() == 'K21'
    assert len([c for c in calls if c[0] == 'canvas_view']) <= count + 1
    page.locator('.canvas-grid button[data-target="A2"]').click()
    page.locator('#canvas-value').fill('保留尚未采用的输入')
    count = len([c for c in calls if c[0] == 'canvas_view'])
    assert not dispatch_wheel(page, deltaY=-120)
    page.wait_for_timeout(250)
    assert page.locator('#canvas-value').input_value() == '保留尚未采用的输入'
    assert len([c for c in calls if c[0] == 'canvas_view']) == count


def test_wheel_does_not_capture_dialog_zoom_or_other_inputs(release_grid):
    page, api, source, calls = release_grid
    jump(page, 'K11')
    assert not dispatch_wheel(page, deltaY=-120, ctrlKey=True)
    page.locator('.canvas-grid button[data-target="A2"]').click()
    page.locator('#canvas-panes').click()
    assert not dispatch_wheel(page, deltaY=-120)
    page.get_by_role('button', name='取消', exact=True).click()
    count = len([c for c in calls if c[0] == 'canvas_view'])
    page.locator('#canvas-value').dispatch_event('wheel', {'deltaY': -120, 'bubbles': True, 'cancelable': True})
    page.wait_for_timeout(250)
    assert len([c for c in calls if c[0] == 'canvas_view']) == count


def test_closing_discards_queued_wheel_and_late_window(release_grid):
    page, api, source, calls = release_grid
    jump(page, 'K21'); page.evaluate('window.__delay=true')
    dispatch_wheel(page, deltaY=-120); page.wait_for_function('window.__held!==null')
    dispatch_wheel(page, deltaX=-120)
    count = len([c for c in calls if c[0] == 'canvas_view'])
    assert page.evaluate('''async()=>{
      const result=await window.pywebview.api.floating_close_document();window.OfficeNative.close();
      document.getElementById('floating-input').value='下一份文档的要求';
      window.__held();window.__held=null;return result.ok;
    }''')
    page.wait_for_timeout(350)
    assert page.locator('#canvas-work').count() == 0
    assert page.locator('#floating-input').input_value() == '下一份文档的要求'
    assert len([c for c in calls if c[0] == 'canvas_view']) == count


def test_frozen_corner_after_horizontal_window_replacement(release_grid):
    page, api, source, calls = release_grid
    wheel(page, dy=1050); stable(page, 'A11')
    wheel(page, dx=1800)
    page.wait_for_function('document.getElementById("canvas-window").value!=="A11"')
    page.wait_for_timeout(200)
    identity = hit_identity(page, 'A2')
    save_evidence(page, 'frozen-after-right', identity)
    assert identity['rowLabel'] == '2', identity
    page.locator('.canvas-grid button[data-target="A2"]').click()
    assert page.locator('#canvas-position').text_content() == 'A2'
