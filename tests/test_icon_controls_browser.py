from tests.test_floating_browser import floating_browser, open_floating_menu
from tests.test_paged_browser import engine_frame, paged_browser


def test_icon_controls_have_accessible_labels_and_bounded_keyboard_tooltips(floating_browser):
    page, api, source, calls, errors = floating_browser
    page.locator('#floating-input').fill('保留这条未发送的要求。')
    choose = page.get_by_role('button', name='选择文档', exact=True).filter(has=page.locator('.agent-visually-hidden'))
    for width in [1280, 1024, 520, 360]:
        page.set_viewport_size({'width': width, 'height': 640})
        choose.focus()
        tooltip = page.get_by_role('tooltip')
        assert tooltip.inner_text() == '选择文档'
        bounds = tooltip.bounding_box()
        assert bounds and 0 <= bounds['x'] < bounds['x'] + bounds['width'] <= width
        assert bounds['width'] >= 60
        assert 0 <= bounds['y'] < bounds['y'] + bounds['height'] <= 640
        assert choose.bounding_box()['width'] == 32
        page.keyboard.press('Escape')
        assert not tooltip.is_visible()
        assert page.locator('#floating-input').input_value() == '保留这条未发送的要求。'
    page.locator('#floating-input').focus()
    choose.hover()
    page.get_by_role('tooltip').wait_for(state='visible')
    assert page.get_by_role('tooltip').inner_text() == '选择文档'
    assert not [name for name in calls if name in {'floating_submit', 'floating_read', 'paged_publish'}]
    assert api._floating.document is None and errors == []


def test_sidebar_closes_and_reopens_without_recreating_editor_or_losing_draft(paged_browser):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    frame_id = engine_frame(page)._impl_obj._guid
    page.locator('#floating-input').fill('压缩结论，保留温度与限定条件。')
    toggle = page.get_by_role('button', name='AI 对话', exact=True)
    for width in [1280, 1024, 520, 360]:
        page.set_viewport_size({'width': width, 'height': 700})
        if toggle.get_attribute('aria-expanded') == 'false':
            toggle.click()
        page.get_by_role('button', name='收起对话侧栏', exact=True).click()
        assert toggle.get_attribute('aria-expanded') == 'false'
        assert not page.locator('#floating-conversation').is_visible()
        assert toggle.evaluate('n=>n===document.activeElement')
        assert not page.locator('.paged-frame').evaluate('n=>n.inert')
        toggle.click()
        assert toggle.get_attribute('aria-expanded') == 'true'
        assert page.locator('#floating-input').input_value() == '压缩结论，保留温度与限定条件。'
        assert engine_frame(page)._impl_obj._guid == frame_id
        page.locator('#simple-model-menu > summary').click()
        page.keyboard.press('Escape')
        assert page.locator('#floating-conversation').is_visible()
        assert toggle.get_attribute('aria-expanded') == 'true'
    assert page.get_by_role('button', name='保存副本', exact=True).is_disabled()
    assert source.read_bytes() == original and not api._floating.paged.proposals and errors == []


def test_window_controls_remain_single_and_available_in_secondary_views(floating_browser):
    page, api, source, calls, errors = floating_browser
    actions = []
    class Native:
        maximum = False
        def state(self):
            return {'ok': True, 'available': True, 'maximized': self.maximum}
        def action(self, action):
            actions.append(action)
            self.maximum = not self.maximum if action == 'maximize' else self.maximum
            return self.state()
    api._window_chrome = Native()
    page.set_viewport_size({'width': 1024, 'height': 700})
    page.get_by_role('button', name='最大化窗口', exact=True).wait_for(state='visible')
    page.get_by_role('button', name='最大化窗口', exact=True).click()
    page.get_by_role('button', name='还原窗口', exact=True).wait_for(state='visible')
    page.locator('#floating-input').hover()
    page.get_by_role('button', name='还原窗口', exact=True).hover()
    page.get_by_role('tooltip').wait_for(state='visible')
    assert page.get_by_role('tooltip').inner_text() == '还原窗口'
    page.get_by_role('button', name='还原窗口', exact=True).click()
    page.get_by_role('button', name='最大化窗口', exact=True).wait_for(state='visible')
    page.locator('#floating-input').hover()
    page.get_by_role('button', name='最大化窗口', exact=True).hover()
    page.get_by_role('tooltip').wait_for(state='visible')
    assert page.get_by_role('tooltip').inner_text() == '最大化窗口'
    page.locator('#floating-input').fill('保留原输入。')
    open_floating_menu(page, '工具与记录')
    page.get_by_role('button', name='历史记录', exact=True).click()
    assert page.locator('#desktop-window-bar').is_visible()
    assert page.locator('.desktop-window-controls').count() == 1
    assert page.get_by_role('button', name='最小化窗口', exact=True).is_visible()
    page.get_by_role('button', name='返回悬浮文档助手', exact=True).click()
    assert not page.locator('#desktop-window-bar').is_visible()
    assert page.locator('#floating-input').input_value() == '保留原输入。'
    assert api.floating_window_action('destroy')['ok'] is False
    assert api.floating_window_action({'action': 'maximize'})['ok'] is False
    assert actions == ['maximize', 'maximize'] and errors == []
