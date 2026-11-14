import hashlib
import json
import mimetypes
import os
from types import SimpleNamespace
from urllib.parse import urlsplit

import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_model_connections import connection, payload, save
from tests.test_office_workflow_browser import make_document, wait_for
from tests.test_paged_browser import engine_frame, serve_static
from tests.test_professional_browser import EDGE, STATIC
from tests.test_floating_browser import open_floating_menu


OUT = STATIC.parents[2]/'deliverables/ui-style-unification-20261113'


def palette(page):
    return page.evaluate('''() => {
      const s=getComputedStyle(document.documentElement);
      return Object.fromEntries(['--bg','--card','--ink','--accent','--primary-bg','--on-primary','--font','--fs-body'].map(k=>[k,s.getPropertyValue(k).trim()]));
    }''')


def open_settings(page):
    page.locator('#floating-root:not([hidden]) header > details > summary').click()
    page.locator('.floating-menu').get_by_role('button',name='设置',exact=True).click()
    page.locator('#tab-machine').click()
    page.locator('#settings-theme-choice').wait_for(state='visible')


def test_shared_style_two_themes_migration_and_actual_editor(tmp_path, monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    OUT.mkdir(parents=True,exist_ok=True)
    source = tmp_path/'全局样式样例.docx'
    make_document(source)
    original = source.read_bytes()
    monkeypatch.setenv('TEMP',str(tmp_path))
    monkeypatch.setenv('TMP',str(tmp_path))
    captures, checks = [], []
    with serve_static() as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            variants = ['before','after'] if os.environ.get('LIX_TEST_STYLE_BASELINE')=='1' else ['after']
            for variant in variants:
                context = browser.new_context(viewport={'width':1280,'height':820})
                page = context.new_page()
                page.set_default_timeout(15000)
                app = WebviewApp()
                app.attach(SimpleNamespace(create_file_dialog=lambda *a,**kw:[str(source)]))
                settings_root = tmp_path/variant
                settings_root.mkdir(exist_ok=True)
                api = FloatingApi(app,settings_root=settings_root)
                save(api._connections,payload([connection('style-fixture')]))
                calls, errors = [], []
                def invoke(name,args):
                    if name not in {'floating_poll','poll'}:calls.append(name)
                    assert name not in {'ask','paged_submit','floating_submit','confirm_image_generation','paged_publish'}
                    return getattr(api,name)(*args)
                def route(request):
                    name = urlsplit(request.request.url).path.lstrip('/') or 'index.html'
                    if name=='auth.js':request.fulfill(status=200,body='');return
                    saved = OUT/'before'/name
                    if variant=='before' and name and saved.is_file() and saved.resolve().is_relative_to((OUT/'before').resolve()):
                        request.fulfill(path=str(saved),content_type=mimetypes.guess_type(name)[0] or 'application/octet-stream');return
                    request.continue_()
                page.route('**/*',route)
                page.expose_function('styleHost',invoke)
                page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.styleHost(name,args)})};window.confirm=()=>false;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
                page.on('pageerror',lambda error:errors.append(str(error)))
                page.goto(url)
                page.locator('#floating-input').fill('请精简所选段落，保留数字、单位和引文。')
                page.screenshot(path=str(OUT/f'{variant}-workspace-dark-1280.png'))
                open_settings(page)
                choices = page.locator('#settings-theme-choice option').evaluate_all('(rows)=>rows.map(row=>row.value)')
                assert choices==(['dark','light','highcontrast'] if variant=='before' else ['dark','light'])
                for theme in ['dark','light']:
                    page.locator('#settings-theme-choice').select_option(theme)
                    page.wait_for_function('(t)=>document.documentElement.dataset.theme===t',arg=theme)
                    widths = [1280,520] if variant=='before' else [1440,1280,1024,520]
                    for width in widths:
                        page.set_viewport_size({'width':width,'height':820})
                        assert page.locator('#settings-theme-choice').is_visible()
                        bounds = page.locator('#settings-dialog').bounding_box()
                        assert bounds['x']>=-1 and bounds['x']+bounds['width']<=width+1
                        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
                        file = f'{variant}-settings-{theme}-{width}.png'
                        page.screenshot(path=str(OUT/file))
                        captures.append({'file':file,'variant':variant,'theme':theme,'width':width,'palette':palette(page)})
                assert page.locator('#floating-input').input_value()=='请精简所选段落，保留数字、单位和引文。'
                page.locator('#settings-close').click()
                page.set_viewport_size({'width':1280,'height':820})
                open_floating_menu(page, '工具与记录')
                page.locator('.floating-menu').get_by_role('button',name='历史记录',exact=True).click()
                page.locator('#simple-history').wait_for(state='visible')
                page.wait_for_function("typeof historyLoading==='undefined'||!historyLoading")
                page.screenshot(path=str(OUT/f'{variant}-history-light-1280.png'))
                page.locator('#nav-work').click()
                page.locator('#model-mode').select_option('chat')
                page.locator('#prompt').fill('请用简短的语言介绍技术报告的常见结构。')
                if variant=='after':
                    assert page.locator('.composer').evaluate('(node)=>getComputedStyle(node).borderRadius')=='18px'
                    assert page.locator('#prompt').evaluate('(node)=>getComputedStyle(node).outlineStyle')=='none'
                page.screenshot(path=str(OUT/f'{variant}-chat-light-1280.png'))
                page.locator('#model-mode').select_option('image')
                page.locator('#prompt').fill('一张简洁的技术报告配图，橙色点缀，白色背景。')
                page.screenshot(path=str(OUT/f'{variant}-image-light-1280.png'))
                page.locator('#floating-return').click()
                page.locator('#floating-choose').click()
                wait_for(page,lambda:bool(api._floating.document),'document selection')
                page.get_by_role('button',name='同意读取并打开',exact=True).click()
                wait_for(page,lambda:bool(api._floating.paged.manifest),'actual editor import',70)
                frame = engine_frame(page)
                frame.locator('.ProseMirror').first.wait_for(state='attached')
                text = frame.locator('.ProseMirror').first.text_content()
                for theme in ['dark','light']:
                    page.evaluate('(theme)=>applyTheme(theme)',theme)
                    frame.wait_for_function('(theme)=>document.body.dataset.theme===theme',arg=theme)
                    assert frame.locator('.ProseMirror').first.text_content()==text
                    ui = frame.locator('#format-bar').evaluate('(node)=>({background:getComputedStyle(node).backgroundColor,color:getComputedStyle(node).color,font:getComputedStyle(node).fontFamily})')
                    if variant=='after':
                        expected = page.locator('#floating-composer').evaluate('(node)=>getComputedStyle(node).backgroundColor')
                        assert ui['background']==expected
                        assert frame.locator('#format-bar').evaluate('(node)=>getComputedStyle(node).getPropertyValue("--font").trim()')==palette(page)['--font']
                    page.screenshot(path=str(OUT/f'{variant}-document-{theme}-1280.png'))
                    checks.append({'variant':variant,'theme':theme,'frame_ui':ui,'draft_revision':api._floating.paged.revision,'text_unchanged':True})
                assert api._floating.paged.revision==0 and not api._floating.outputs
                assert source.read_bytes()==original
                if variant=='after':
                    page.evaluate("localStorage.setItem('office-agent.theme','highcontrast')")
                    page.reload()
                    page.wait_for_function("document.documentElement.dataset.theme==='dark'")
                    assert page.evaluate("localStorage.getItem('office-agent.theme')")=='dark'
                    assert page.locator('#settings-theme-choice option').count()==2
                assert errors==[],errors
                (OUT/f'{variant}-calls.json').write_text(json.dumps({'calls':calls,'page_errors':errors,'new_model_or_write_task':False},ensure_ascii=False,indent=2),encoding='utf-8')
                context.close()
        finally:
            browser.close()
    (OUT/'BROWSER.json').write_text(json.dumps({'native':False,'real_service':False,'real_editor':'bundled SuperDoc in headless Edge','source_sha256':hashlib.sha256(original).hexdigest(),'source_unchanged':True,'captures':captures,'document_checks':checks,'legacy_preference_migrated_to_dark':True},ensure_ascii=False,indent=2),encoding='utf-8')
