import json

from tests.test_floating_browser import floating_browser
from tests.test_professional_browser import STATIC


OUT = STATIC.parents[2] / 'deliverables/settings-compact-20261113'


def test_narrow_settings_navigation_and_window_expansion(floating_browser):
    page, api, path, calls, errors = floating_browser
    OUT.mkdir(parents=True, exist_ok=True)
    page.evaluate("document.getElementById('floating-input').value='保留这条未发送要求'")
    page.set_viewport_size({'width':420,'height':640})
    page.evaluate('window.OfficeSettings.show()')
    header = page.locator('#settings-dialog .settings-head')
    for width in [320, 420, 520]:
        page.set_viewport_size({'width':width,'height':640})
        assert page.locator('#settings-section').is_visible()
        assert page.locator('.settings-nav').is_hidden()
        assert header.bounding_box()['height'] <= 58
        assert page.locator('#settings-expand').is_visible()
        assert page.locator('#settings-close').is_visible()
        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
        assert page.locator('#settings-dialog .settings-panes').bounding_box()['height'] >= 530
        page.locator('#settings-section').select_option('tab-model')
        assert page.locator('#settings-pane-model').is_visible()
        pane = page.locator('#settings-dialog .settings-panes')
        assert pane.evaluate('(n)=>n.scrollHeight>n.clientHeight')
        page.screenshot(path=str(OUT / f'settings-{width}.png'))
    page.locator('#settings-section').select_option('tab-skills')
    assert page.locator('#settings-pane-skills').is_visible()
    page.evaluate('window.savedSkills=window.OfficeSkills;window.OfficeSkills={requestClose:()=>false}')
    page.locator('#settings-section').select_option('tab-machine')
    assert page.locator('#settings-section').input_value() == 'tab-skills'
    assert page.locator('#settings-pane-skills').is_visible()
    page.evaluate('window.OfficeSkills=window.savedSkills')
    page.set_viewport_size({'width':1120,'height':800})
    assert page.locator('#settings-section').is_hidden()
    page.locator('#tab-model').click()
    before = page.locator('#settings-dialog').bounding_box()
    page.locator('#settings-expand').click()
    assert page.locator('#settings-dialog').bounding_box()['width'] > before['width']
    assert page.locator('#settings-expand').get_attribute('aria-label') == '还原设置大小'
    page.locator('#settings-expand').click()
    page.evaluate('''() => {
      const previous=window.pywebview.api;
      window.windowActions=[];
      window.pywebview.api=new Proxy({}, {get:(_,name)=>name==='floating_window_action' ? action=>{
        window.windowActions.push(action);return new Promise(resolve=>window.resolveWindowAction=resolve);
      }: previous[name]});
      document.body.dataset.desktopWindow='true'; document.body.dataset.windowMaximized='false';
    }''')
    page.locator('#settings-expand').click()
    page.locator('#settings-expand').evaluate('(n)=>n.click()')
    assert page.evaluate('window.windowActions') == ['maximize']
    page.evaluate('window.resolveWindowAction({ok:true,maximized:true})')
    page.wait_for_function('document.getElementById("settings-expand").getAttribute("aria-label")==="还原窗口"')
    page.locator('#settings-expand').click()
    page.evaluate('window.resolveWindowAction({ok:false})')
    page.wait_for_function('document.getElementById("settings-window-status").textContent.includes("窗口未能放大")')
    assert page.locator('#settings-window-status').is_visible()
    page.keyboard.press('Escape')
    assert page.locator('#settings-dialog').is_hidden()
    assert page.locator('#floating-input').input_value() == '保留这条未发送要求'
    assert not errors
    (OUT / 'BROWSER.json').write_text(json.dumps({'page_errors':errors,'native':False,
        'window_action_fixture':True,'widths':[320,420,520,1120],
        'window_actions':page.evaluate('window.windowActions'),
        'draft_preserved':True,'dirty_tab_change_blocked':True},ensure_ascii=False,indent=2),'utf-8')
