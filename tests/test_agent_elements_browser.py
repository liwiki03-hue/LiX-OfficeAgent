from tests.test_floating_browser import floating_browser
from tests.test_paged_browser import engine_frame, paged_browser
from copy import deepcopy


def test_reference_composer_preserves_draft_and_has_one_active_popover(floating_browser):
    page, api, source, calls, errors = floating_browser
    draft = '请压缩实验报告的结论，保留30组样本与20 ms。'
    page.locator('#floating-input').fill(draft)
    page.locator('#floating-task-details > summary').click()
    assert page.locator('.agent-context-body').is_visible()
    page.locator('#simple-model-menu > summary').click()
    assert not page.locator('.agent-context-body').is_visible()
    page.keyboard.press('Escape')
    assert page.locator('#simple-model-menu > summary').evaluate('n=>n===document.activeElement')
    for width, height in [(1280, 800), (820, 700), (520, 800), (360, 440)]:
        page.set_viewport_size({'width': width, 'height': height})
        assert page.locator('#floating-input').input_value() == draft
        send = page.get_by_role('button', name='发送要求', exact=True)
        bounds = send.bounding_box()
        assert bounds and bounds['width'] == bounds['height'] == 36
        assert 0 <= bounds['y'] < bounds['y'] + bounds['height'] <= height
        assert page.locator('#floating-controls').count() == 1
        assert page.locator('#simple-model-menu').count() == 1
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        title_size = page.locator('#floating-welcome h1').evaluate('n=>parseFloat(getComputedStyle(n).fontSize)')
        assert 20 <= title_size <= 26
        page.locator('#simple-model-menu > summary').click()
        popup = page.locator('#simple-model-menu > .simple-menu-body')
        popup.wait_for(state='visible')
        bounds = popup.bounding_box()
        assert bounds and bounds['x'] >= 0 and bounds['y'] >= 0
        assert bounds['x'] + bounds['width'] <= width and bounds['y'] + bounds['height'] <= height
        page.keyboard.press('Escape')
        assert page.locator('#simple-model-menu > summary').evaluate('n=>n===document.activeElement')
    assert api._floating.document is None and api._floating.runtime is None
    assert not [name for name in calls if name in {'floating_submit', 'floating_read', 'paged_publish'}]
    assert errors == []


def test_edit_card_keeps_untrusted_text_literal_and_unicode_diff_counts(floating_browser):
    page, api, source, calls, errors = floating_browser
    result = page.evaluate('''() => {
      const node = document.createElement('article');
      const actions = OfficeAgentElements.edit(node, {
        before:'结果🙂：<img src=x onerror=alert(1)>',
        after:'结果✅：<img src=x onerror=alert(1)>',
        title:'文字修改', status:'待审阅', reason:'保留数字与单位。'
      });
      return {html:node.querySelectorAll('img,script').length,
        before:node.querySelector('.agent-edit-before pre').textContent,
        counts:node.querySelector('.agent-edit-counts').getAttribute('aria-label'),
        actions:actions.className};
    }''')
    assert result == {'html': 0, 'before': '结果🙂：<img src=x onerror=alert(1)>',
                      'counts': '删除 1 字，新增 1 字', 'actions': 'agent-tool-actions'}
    assert page.evaluate('''() => {
      const node = document.createElement('article');
      const actions = OfficeAgentElements.edit(node, {before:'旧文', after:'新文', status:'待审阅', risks:['仅写副本，原稿保持。']});
      return !!(node.querySelector('.agent-tool-warning').compareDocumentPosition(actions) & Node.DOCUMENT_POSITION_FOLLOWING);
    }''')
    assert not [name for name in calls if name in {'floating_submit', 'floating_read', 'paged_publish'}]
    assert errors == []


def test_new_valid_reference_replaces_its_own_previous_failure(paged_browser):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    frame = engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。', exact=True).first.click()
    page.get_by_role('button', name='引用选区', exact=True).click()
    page.wait_for_function("document.getElementById('paged-feedback').textContent.includes('明确文字范围')")
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。', exact=True).first.click()
    page.keyboard.press('Home')
    page.keyboard.press('Shift+End')
    page.get_by_role('button', name='引用选区', exact=True).click()
    page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('已引用') && document.getElementById('paged-feedback').textContent === ''")
    assert source.read_bytes() == original and not api._floating.paged.proposals
    assert errors == []


def test_wrapped_page_controls_remain_inside_narrow_window(paged_browser):
    page, api, source, errors = paged_browser
    frame = engine_frame(page)
    page.locator('#floating-input').fill('保留尚未发送的修改要求。')
    for width, height in [(520, 800), (360, 440), (1280, 800)]:
        page.set_viewport_size({'width': width, 'height': height})
        frame.get_by_role('button', name='加粗', exact=True).wait_for(state='visible')
        for control in page.locator('#paged-statusbar').get_by_role('button').all():
            box = control.bounding_box()
            assert box and 0 <= box['y'] < box['y'] + box['height'] <= height
        assert page.locator('#floating-input').input_value() == '保留尚未发送的修改要求。'
    assert errors == []


def test_progress_uses_the_actual_host_publication_phase(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    snapshot = api.floating_snapshot

    def publishing_snapshot(*args):
        state = deepcopy(snapshot(*args))
        state['busy'] = True
        state['stage'] = 'drafting'
        state['paged']['publishing'] = True
        return state

    monkeypatch.setattr(api, 'floating_snapshot', publishing_snapshot)
    page.wait_for_function("document.getElementById('floating-progress').textContent.includes('正在写副本与验证')")
    assert source.read_bytes() == original and not api._floating.outputs
    assert errors == []
