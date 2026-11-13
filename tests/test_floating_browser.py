from pathlib import Path
import re
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_model_connections import connection, payload, save
from tests.test_professional_browser import STATIC, EDGE


@pytest.fixture
def floating_browser(tmp_path, monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    path = tmp_path / '自建报告.docx'
    doc = Document()
    doc.add_paragraph('实验延迟为20 ms，样本为30组。')
    doc.add_paragraph('第二段应保持原样。')
    doc.save(path)
    app = WebviewApp()
    app.attach(SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(path)]))
    api = FloatingApi(app, auth=SimpleNamespace(root=tmp_path/'auth', status=lambda: {'authenticated': True}), settings_root=tmp_path)
    save(api._connections, payload([connection()]))
    calls = []
    monkeypatch.setenv('TEMP', str(tmp_path))
    monkeypatch.setenv('TMP', str(tmp_path))
    def invoke(name, args):
        if name != 'floating_snapshot':
            calls.append(name)
        return getattr(api, name)(*args)
    with sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1120, 'height': 760})
            page.set_default_timeout(15000)
            page.route('**/*', lambda route: route.abort())
            errors = []
            page.on('pageerror', lambda error: errors.append(str(error)))
            page.expose_function('floatingHost', invoke)
            source = (STATIC/'index.html').read_text(encoding='utf8')
            page.set_content(re.sub(r'<script\b[^>]*>.*?</script>', '', source, flags=re.S))
            for css in re.findall(r'<link[^>]+href="([^"]+\.css)"', source):
                page.add_style_tag(content=(STATIC/css).read_text(encoding='utf8'))
            page.evaluate('''() => {
              document.getElementById('workspace').hidden=false; document.getElementById('workspace').inert=false;
              document.getElementById('workspace').setAttribute('aria-hidden','false');
              document.getElementById('auth-screen').hidden=true;
              window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};
              window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.floatingHost(name,args)})};
              window.confirm=()=>true;
            }''')
            for script in re.findall(r'<script src="([^"]+)"', source):
                if script != 'auth.js':
                    page.add_script_tag(path=str(STATIC/script))
            page.evaluate("window.dispatchEvent(new Event('pywebviewready'))")
            page.wait_for_selector('#floating-root')
            yield page, api, path, calls, errors
        finally:
            browser.close()


def test_changed_source_excerpt_is_historical_and_cannot_reuse_controls(floating_browser):
    page, api, path, calls, errors = floating_browser
    page.locator('#floating-choose').click()
    page.get_by_role('button', name='同意读取并打开', exact=True).click()
    page.wait_for_function("!document.getElementById('floating-send').disabled")
    pane = page.locator('#floating-document')
    assert '已读取原稿' in pane.inner_text()
    page.locator('#floating-input').fill('保留这条尚未发送的要求')
    document = Document(path)
    document.paragraphs[0].text = '外部新版本，不应自动读取。'
    document.save(path)
    with pytest.raises(ValueError, match='外部改变'):
        api._floating.check_source()
    page.wait_for_function("document.getElementById('floating-document').textContent.includes('上次读取版本')")
    assert '实验延迟为20 ms' in pane.inner_text()
    assert '外部新版本' not in pane.inner_text()
    assert pane.get_by_role('button', name='引用到对话', exact=True).count() == 0
    assert pane.locator('input[type=checkbox]').count() == 0
    assert page.locator('#floating-send').is_disabled()
    assert '源文件已改变' in page.locator('#floating-identity').inner_text()
    assert page.locator('#floating-input').input_value() == '保留这条尚未发送的要求'
    assert calls.count('floating_read') == 1 and 'floating_submit' not in calls
    assert not list(path.parent.glob('*.agent-*.docx')) and not errors


@pytest.mark.parametrize('width,height,theme', [(1120,760,'light'), (420,640,'dark'), (360,440,'highcontrast')])
def test_candidate_chooser_read_and_reparented_model_controls(floating_browser, width, height, theme):
    page, api, path, calls, errors = floating_browser
    page.set_viewport_size({'width':width, 'height':height})
    page.evaluate('(t)=>document.documentElement.dataset.theme=t', theme)
    page.locator('#floating-input').fill('第一段改得简洁一点，保留20 ms和30组。📄')
    page.locator('#floating-choose').click()
    page.wait_for_function("document.getElementById('floating-title').textContent === '自建报告.docx'")
    assert not api._floating.document['ready']
    page.get_by_role('button', name='同意读取并打开', exact=True).click()
    page.wait_for_function("!document.getElementById('floating-send').disabled")
    if width < 900:
        page.locator('#floating-tabs').get_by_role('button', name='文档', exact=True).click()
    assert page.locator('#floating-document').is_visible()
    assert '实验延迟为20 ms' in page.locator('#floating-document').inner_text()
    assert page.locator('#floating-input').input_value().startswith('第一段改得简洁一点')
    assert page.locator('#floating-send').bounding_box()['y'] < height
    page.locator('#simple-model-menu > summary').click()
    assert page.locator('#model-selector').is_visible()
    assert page.locator('#model-configure').is_visible()
    page.keyboard.press('Escape')
    assert page.locator('#simple-model-menu > summary').evaluate('n=>n===document.activeElement')
    assert page.locator('#model-selector').count() == 1
    page.locator('#floating-root header details > summary').click()
    assert page.get_by_role('button', name='恢复本文件历史要求到草稿', exact=True).is_visible()
    assert page.locator('#floating-send').bounding_box()['y'] < height
    page.keyboard.press('Escape')
    assert page.locator('#floating-root header details > summary').evaluate('n=>n===document.activeElement')
    assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
    assert not [name for name in calls if name in {'ask','floating_submit','confirm_image_generation'}]
    assert errors == []
    folder = STATIC.parents[2]/'deliverables/floating-orb-20261104/browser'
    folder.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(folder/f'candidate-{width}-{theme}.png'))


def test_secondary_image_plan_and_chat_events_share_the_single_event_consumer(floating_browser):
    page, api, path, calls, errors = floating_browser
    from tests.test_model_connections import select
    select(api._connections, 'official', 'image', 'image-model')
    page.evaluate('window.relayed=[]; window.addEventListener("officedaily:event", event=>window.relayed.push(event.detail.type))')
    page.locator('#floating-root header details > summary').click()
    page.get_by_role('button', name='问答、图片与专业工具', exact=True).click()
    page.evaluate('window.OfficeModelSelector.refresh()')
    reply = api.ask('生成一张技术报告的实验流程配图。', {'size': '1024x1024', 'style': 'natural'})
    assert reply['ok'] and reply['generation_not_started']
    generated = []
    api._app._image_session[1]._provider_factory = lambda: generated.append(True)
    api._app._push({'type': 'chat_changed'})
    page.get_by_role('button', name='确认生成（可能产生费用）', exact=True).wait_for()
    page.wait_for_function('window.relayed.includes("chat_changed") && window.relayed.filter(kind=>kind==="image_plan").length===1')
    assert generated == [] and not [name for name in calls if name == 'poll']
    assert errors == []


def test_native_open_waiting_input_and_reference_have_no_implicit_model_call(floating_browser, monkeypatch):
    import threading
    from office_agent.webui.document_open import parse_payload
    page, api, path, calls, errors = floating_browser
    entered, release = threading.Event(), threading.Event()
    started = []
    def gated(source, cancelled, **options):
        entered.set()
        assert release.wait(20)
        return parse_payload(source)
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', gated)
    monkeypatch.setattr(api._app, '_start_task', lambda *a, **kw: started.append(a))
    api._floating.dropped((str(path),))
    try:
        assert entered.wait(10)
        page.wait_for_function("document.getElementById('floating-send').textContent.includes('解析后处理')")
        assert page.locator('#floating-input').is_visible()
        page.locator('#floating-input').fill('第一段语气自然，保留20 ms与30组。')
        page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
        assert len(api._floating.runtime.messages) == 1 and started == []
        page.locator('#floating-input').fill('尚未发送的草稿')
    finally:
        release.set()
        worker = api._floating._open_worker
        if worker:
            worker.join(10)
            assert not worker.is_alive()
    page.wait_for_function("document.getElementById('floating-document').textContent.includes('实验延迟为20 ms')")
    assert len(started) == 1
    assert page.locator('#floating-input').input_value() == '尚未发送的草稿'
    page.locator('#floating-document').get_by_role('button', name='引用到对话', exact=True).first.click()
    assert '尚未发送的草稿' in page.locator('#floating-input').input_value()
    assert len(started) == 1 and calls.count('floating_submit') == 1
    assert 'floating_reference' in calls and 'floating_read' not in calls and 'floating_choose' not in calls
    assert errors == []


def _open_main_workspace(page):
    page.locator('#floating-root header details > summary').click()
    page.get_by_role('button', name='问答、图片与专业工具', exact=True).click()
    page.locator('#floating-return').wait_for(state='visible')


def _set_snapshot(page, **overrides):
    page.evaluate('''async overrides => {
      if (!window.__fabBaseline) window.__fabBaseline = await window.floatingHost('floating_snapshot', [0]);
      window.__fabState = {...window.__fabBaseline, ...overrides};
      window.OfficeAuth.wrapApi = () => ({floating_snapshot: async () => window.__fabState});
      await window.OfficeFloating.refresh();
    }''', overrides)


def _idle_snapshot():
    return {'busy': False, 'open_request': None, 'plans': [], 'outputs': [],
            'needs_clarification': False, 'stage': 'idle'}


def _after_style(page):
    return page.locator('#floating-return').evaluate('''node => ({
      display: getComputedStyle(node, '::after').display,
      background: getComputedStyle(node, '::after').backgroundColor
    })''')


def _boxes_disjoint(first, second):
    return (first['x'] + first['width'] <= second['x'] or
            second['x'] + second['width'] <= first['x'] or
            first['y'] + first['height'] <= second['y'] or
            second['y'] + second['height'] <= first['y'])


def test_document_review_icon_favicon_and_floating_return(floating_browser):
    page, api, path, calls, errors = floating_browser
    assert page.locator('.agent-eyebrow').inner_text() == 'LiX'
    assert page.locator('#floating-title').inner_text() == 'LiX · 文档助手'
    favicon = STATIC / 'brand/lix.png'
    assert favicon.is_file()
    assert page.locator('link[rel="icon"][type="image/png"]').get_attribute('href') == 'brand/lix.png'

    _open_main_workspace(page)
    mark = page.locator('.sidebar .logo .mark')
    assert mark.inner_text() == ''
    assert mark.locator('img').count() == 1
    assert mark.locator('img').get_attribute('src') == 'brand/lix.png'
    assert mark.locator('img').get_attribute('alt') == ''  # 容器统一隐藏装饰图标
    assert mark.get_attribute('aria-hidden') == 'true'

    back = page.locator('#floating-return')
    assert back.get_attribute('aria-label') == '返回悬浮文档助手'
    assert back.get_attribute('title') == '返回悬浮文档助手'
    assert back.locator('svg').get_attribute('aria-hidden') == 'true'
    assert back.locator('svg').get_attribute('focusable') == 'false'
    box = back.bounding_box()
    assert box['width'] == box['height'] == 56
    back.click()
    assert page.locator('#floating-root').is_visible()
    assert back.is_hidden()
    assert errors == []


def test_floating_return_status_tracks_running_confirmation_and_review(floating_browser):
    page, api, path, calls, errors = floating_browser
    _open_main_workspace(page)
    back = page.locator('#floating-return')

    _set_snapshot(page, **_idle_snapshot())
    page.wait_for_function("document.getElementById('floating-return').dataset.state === 'idle'")
    assert _after_style(page)['display'] == 'none'
    assert back.get_attribute('aria-label') == '返回悬浮文档助手'

    _set_snapshot(page, **{**_idle_snapshot(), 'busy': True})
    page.wait_for_function("document.getElementById('floating-return').dataset.state === 'running'")
    assert '任务运行中' in back.get_attribute('aria-label')
    running = _after_style(page)
    assert running['display'] == 'block'
    assert running['background'] == back.evaluate('node => getComputedStyle(node).color')

    plan = {'revision': 1, 'plan': {'id': 'fab-plan', 'target': '第一段', 'before': '原文',
                                     'after': '修改后', 'risks': [], 'location': None}}
    document = {'id': 'fab-doc', 'version': 'source-v1', 'ready': True, 'editable': True,
                'name': 'fab.docx', 'source_origin': 'native_drop'}
    _set_snapshot(page, **{**_idle_snapshot(), 'document': document, 'plans': [plan]})
    page.wait_for_function("document.getElementById('floating-return').dataset.state === 'attention'")
    assert '有待确认方案' in back.get_attribute('aria-label')
    warning = _after_style(page)
    assert warning['display'] == 'block'
    assert warning['background'] == page.locator('#floating-status').evaluate('node => getComputedStyle(node).color')

    output = {'id': 'fab-output', 'undone': False, 'output': 'sample.agent-fab.docx',
              'render_diff': True, 'source_version': 'source-v1', 'instruction_revision': 0}
    _set_snapshot(page, **{**_idle_snapshot(), 'stage': 'verified', 'outputs': [output]})
    page.wait_for_function("document.getElementById('floating-return').dataset.state === 'attention'")
    assert '有待验收副本' in back.get_attribute('aria-label')

    _set_snapshot(page, **_idle_snapshot())
    page.wait_for_function("document.getElementById('floating-return').dataset.state === 'idle'")
    assert _after_style(page)['display'] == 'none'
    assert back.get_attribute('aria-label') == '返回悬浮文档助手'
    assert errors == []


@pytest.mark.parametrize('width,height,theme', [
    (1120, 760, 'light'), (1120, 760, 'dark'), (1120, 760, 'highcontrast'),
    (420, 640, 'dark'), (360, 440, 'highcontrast')])
def test_floating_return_fits_small_windows_and_all_themes(floating_browser, width, height, theme):
    page, api, path, calls, errors = floating_browser
    page.set_viewport_size({'width': width, 'height': height})
    page.evaluate('(value) => document.documentElement.dataset.theme = value', theme)
    _open_main_workspace(page)
    back = page.locator('#floating-return')
    expected = 56 if width >= 900 else 48
    box = back.bounding_box()
    assert box['width'] == box['height'] == expected
    assert box['x'] >= 0 and box['y'] >= 0
    assert box['x'] + box['width'] <= width
    assert box['y'] + box['height'] <= height
    for selector in ('#simple-compose-area', '#prompt', '#btn-ask'):
        target = page.locator(selector)
        if target.is_visible():
            assert _boxes_disjoint(box, target.bounding_box()), f'悬浮球遮挡 {selector}'

    _set_snapshot(page, **{**_idle_snapshot(), 'busy': True})
    page.wait_for_function("document.getElementById('floating-return').dataset.state === 'running'")
    assert _after_style(page)['display'] == 'block'
    assert back.evaluate('node => getComputedStyle(node).backgroundColor') != 'rgba(0, 0, 0, 0)'
    mark = page.locator('.sidebar .logo .mark')
    assert mark.count() == 1
    assert mark.evaluate('node => getComputedStyle(node).backgroundColor') != 'rgba(0, 0, 0, 0)'
    assert mark.locator('img').evaluate('node => getComputedStyle(node).objectFit') == 'contain'
    folder = STATIC.parents[2] / 'deliverables' / 'fab-icon-20261109' / 'browser'
    folder.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(folder / f'fab-{width}-{theme}.png'))
    assert errors == []
