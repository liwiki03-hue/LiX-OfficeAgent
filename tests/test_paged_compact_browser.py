import pytest

from tests.test_paged_browser import engine_frame, paged_browser


@pytest.fixture
def visible_scrollbar_browser(tmp_path, monkeypatch):
    from playwright.sync_api import BrowserType
    launch = BrowserType.launch
    def with_scrollbars(browser_type, *args, **kwargs):
        kwargs['ignore_default_args'] = ['--hide-scrollbars']
        return launch(browser_type, *args, **kwargs)
    monkeypatch.setattr(BrowserType, 'launch', with_scrollbars)
    yield from paged_browser.__wrapped__(tmp_path, monkeypatch)


def test_save_state_is_unique_at_bottom_without_extra_header_or_losing_draft(paged_browser):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    frame = engine_frame(page)
    frame_id = frame._impl_obj._guid
    page.locator('#floating-input').fill('请保留数字和单位，稍后继续。')
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。', exact=True).first.click()
    page.keyboard.press('End')
    page.keyboard.insert_text('新草稿。')
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('未保存')")
    assert page.locator('#paged-save-state').count() == 1
    assert page.locator('#paged-statusbar #paged-save-state').count() == 1
    assert page.locator('#paged-top #paged-filebar').count() == 1
    assert not page.locator('#paged-filebar #paged-save-state').count()
    assert page.locator('#paged-save-state').get_attribute('aria-live') == 'off'
    page.wait_for_function("document.getElementById('paged-state-announcement').textContent==='有未保存编辑'")
    announcement = page.locator('#paged-state-announcement').inner_text()
    for width, height in [(1189, 789), (1280, 800), (1024, 800), (520, 800), (360, 440)]:
        page.set_viewport_size({'width': width, 'height': height})
        for theme in ['light', 'dark', 'highcontrast']:
            page.evaluate('(theme)=>document.documentElement.dataset.theme=theme', theme)
            assert page.locator('#paged-top').bounding_box()['height'] <= 46
            if width >= 1024:
                assert page.locator('#paged-statusbar').bounding_box()['height'] <= 34
            for control in [page.locator('#paged-save-state'), page.locator('#paged-zoom'), *page.locator('#paged-statusbar').get_by_role('button').all()]:
                box = control.bounding_box()
                assert box and 0 <= box['x'] < box['x'] + box['width'] <= width
                assert 0 <= box['y'] < box['y'] + box['height'] <= height
            page.get_by_role('button', name='收起对话侧栏', exact=True).click()
            page.get_by_role('button', name='展开对话侧栏', exact=True).click()
            assert page.locator('#floating-input').input_value() == '请保留数字和单位，稍后继续。'
            assert engine_frame(page)._impl_obj._guid == frame_id
    assert page.locator('#paged-state-announcement').inner_text() == announcement
    assert source.read_bytes() == original and errors == []


def test_real_iframe_scrollbar_keyboard_scroll_and_fit_width_preserve_source(visible_scrollbar_browser):
    page, api, source, errors = visible_scrollbar_browser
    original = source.read_bytes()
    page.set_viewport_size({'width': 1189, 'height': 789})
    frame = engine_frame(page)
    scroll = frame.locator('#page-scroll')
    frame.locator('.superdoc-page[data-page-number="2"]').wait_for(state='visible')
    assert scroll.evaluate("node=>getComputedStyle(node,'::-webkit-scrollbar').width") == '6px'
    assert scroll.evaluate("node=>getComputedStyle(node).scrollbarWidth") == 'auto'
    assert scroll.evaluate('node=>node.offsetWidth-node.clientWidth') == 6, frame.locator('body').evaluate("node=>[node,...node.querySelectorAll('*')].filter(n=>n.scrollHeight>n.clientHeight+2&&n.clientHeight>100).slice(0,12).map(n=>({id:n.id,cls:n.className,h:n.clientHeight,sh:n.scrollHeight,ow:n.offsetWidth,cw:n.clientWidth,overflow:getComputedStyle(n).overflow}))")
    scroll.focus()
    page.keyboard.press('PageDown')
    frame.wait_for_function("()=>document.getElementById('page-scroll').scrollTop>100")
    page.keyboard.press('Home')
    frame.wait_for_function("()=>document.getElementById('page-scroll').scrollTop===0")
    page.get_by_role('button', name='适合页宽', exact=True).click()
    frame.wait_for_function("()=>document.getElementById('page-scroll').scrollWidth<=document.getElementById('page-scroll').clientWidth+1")
    assert source.read_bytes() == original
    assert page.get_by_role('button', name='保存副本', exact=True).is_disabled()
    assert errors == []
