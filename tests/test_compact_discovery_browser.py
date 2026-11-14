import json
from urllib.parse import urlsplit

from tests.test_floating_browser import floating_browser, open_floating_menu
from tests.test_professional_browser import STATIC
from tests.test_skill_discovery import response


OUT = STATIC.parents[2] / 'deliverables/compact-discovery-20261113'


def test_compact_menu_history_and_public_search(floating_browser):
    page, api, path, calls, errors = floating_browser
    queries = []
    def fetch(query):
        queries.append(query)
        return response()
    api._skill_discovery._fetch = fetch
    OUT.mkdir(parents=True, exist_ok=True)
    def asset(route):
        target = (STATIC / urlsplit(route.request.url).path.lstrip('/')).resolve()
        if target.is_relative_to(STATIC) and target.is_file() and target.suffix in {'.png','.svg','.ico'}:
            route.fulfill(path=str(target))
        else:
            route.abort()
    page.route('http://lix.invalid/**', asset)
    page.evaluate("() => {const base=document.createElement('base');base.href='http://lix.invalid/';document.head.prepend(base);}")
    page.evaluate("document.querySelectorAll('img[src]').forEach(n=>{n.src=new URL(n.getAttribute('src'),document.baseURI).href;})")
    page.evaluate("document.getElementById('floating-input').value='保留数字，精简说明'")
    summary = page.locator('#floating-root > header > details > summary')
    page.set_viewport_size({'width':520, 'height':800})
    summary.click()
    assert page.locator('.floating-menu > section, .floating-menu > details').count() == 4
    assert page.locator('.floating-menu button').count() == 13
    assert page.locator('.floating-menu button').first.bounding_box()['height'] <= 40
    page.screenshot(path=str(OUT / 'menu-520.png'))
    summary.focus()
    page.keyboard.press('ArrowDown')
    assert page.locator('.floating-menu button').first.evaluate('(n)=>n===document.activeElement')
    page.keyboard.press('Escape')
    assert summary.evaluate('(n)=>n===document.activeElement&&!n.parentElement.open')
    page.set_viewport_size({'width':1120, 'height':800})
    open_floating_menu(page, '工具与记录')
    page.get_by_role('button', name='历史记录', exact=True).click()
    page.wait_for_function('!document.getElementById("simple-history").hidden')
    page.locator('#simple-session-list > summary').click()
    page.locator('#chat-search').fill('论文')
    assert page.locator('#simple-session-list').get_attribute('open') is not None
    page.keyboard.press('Escape')
    page.set_viewport_size({'width':1120,'height':800})
    assert page.locator('#simple-history-chats #chat-search').count() == 1
    page.screenshot(path=str(OUT / 'history-1120.png'))
    page.evaluate('window.OfficeSettings.show()')
    page.locator('#tab-skills').click()
    page.locator('#skills-discover').click()
    page.locator('#skills-search').fill('document')
    assert queries == []
    page.locator('#skills-search').press('Enter')
    page.wait_for_function('document.getElementById("skills-public-status").textContent.includes("已获取 1")')
    assert queries == ['document']
    assert api._app.path is None
    page.locator('#skills-public-results .skill-entry').click()
    assert '尚未核对许可' in page.locator('#skills-public-detail').inner_text()
    assert page.locator('#skills-public-detail input').input_value() == 'https://skills.sh/example/workflows/document-review'
    page.screenshot(path=str(OUT / 'discovery-detail-1120.png'))
    page.evaluate('''() => {
      const original=window.pywebview.api;
      window.delayedCalls=0;
      window.pywebview.api=new Proxy({}, {get:(_,name)=>name==='search_public_skills'?()=>{window.delayedCalls++;return new Promise(resolve=>window.resolveSearch=resolve);}:original[name]});
    }''')
    page.locator('#skills-search').fill('draft')
    page.locator('#skills-public-search').click()
    page.locator('#skills-search').press('Enter')
    assert page.evaluate('window.delayedCalls') == 1
    page.locator('#skills-public-cancel').click()
    page.locator('#skills-search').fill('new query')
    page.evaluate("window.resolveSearch({ok:true,query:'draft',items:[]})")
    page.wait_for_function('!document.getElementById("skills-public-search").disabled')
    assert '已获取 0' not in page.locator('#skills-public-status').inner_text()
    assert page.locator('#skills-search').input_value() == 'new query'
    page.keyboard.press('Escape')
    assert page.locator('#floating-input').input_value() == '保留数字，精简说明'
    assert not errors
    assert not [name for name in calls if name in {'confirm_skill_change','floating_submit','confirm_image_generation'}]
    (OUT / 'BROWSER.json').write_text(json.dumps({'queries':queries,'page_errors':errors,'native':False,
        'real_remote':False,'fixture':'controlled skills.sh response through real host service',
        'menu_actions':13,'duplicate_requests':0,'writes':0}, ensure_ascii=False, indent=2), 'utf-8')
