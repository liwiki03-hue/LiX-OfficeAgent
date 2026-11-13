from pathlib import Path

import pytest

from test_professional_browser import STATIC, _browser


def setup(page, stylesheet=None):
    if stylesheet:
        page.locator('style').evaluate_all('nodes => nodes.forEach(node => node.remove())')
        page.add_style_tag(content=Path(stylesheet).read_text(encoding='utf8'))
    page.evaluate('window.OfficeAuth.ensure = async () => false')
    page.add_script_tag(path=str(STATIC / 'app.js'))
    page.add_script_tag(path=str(STATIC / 'document_tasks.js'))


def position(page):
    return page.locator('#work-view').evaluate('node => ({top:node.scrollTop, height:node.clientHeight, total:node.scrollHeight, overflow:getComputedStyle(node).overflowY})')


def in_view(page, selector):
    return page.locator(selector).evaluate('node => {const box=node.getBoundingClientRect();const host=document.getElementById("work-view").getBoundingClientRect();return box.top>=host.top && box.bottom<=host.bottom && box.bottom<=innerHeight;}')


def point_into_work(page):
    box = page.locator('#work-view').bounding_box()
    page.mouse.move(box['x'] + box['width'] - 24, box['y'] + box['height'] * .55)


def wheel_to(page, selector, direction=1):
    for _ in range(30):
        if in_view(page, selector):
            return
        point_into_work(page)
        page.mouse.wheel(0, direction * 580)
        page.wait_for_timeout(50)
    raise AssertionError('target_not_wheel_reachable:' + selector)


def exercise(page, *, expected=True, capture=None):
    point_into_work(page)
    before = position(page)
    page.mouse.wheel(0, 620)
    page.wait_for_timeout(150)
    after = position(page)
    result = {'before': before, 'after': after, 'wheel_moved': after['top'] > before['top']}
    if not expected:
        return result
    assert result['wheel_moved'], result
    wheel_to(page, '#doc-task-return-panel > summary')
    page.locator('#doc-task-return-panel > summary').click()
    wheel_to(page, '#doc-task-return-compare')
    assert in_view(page, '#doc-task-return-compare')
    assert page.locator('#doc-task-return-baseline').input_value() == ''
    if capture:
        page.screenshot(path=str(capture))
    wheel_to(page, '#professional-panel > summary', -1)
    page.locator('#professional-panel > summary').click()
    assert page.locator('#professional-panel').get_attribute('open') is None
    assert position(page)['overflow'] == 'visible' or position(page)['top'] == 0
    assert in_view(page, '#prompt')
    assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
    page.locator('#professional-panel > summary').focus()
    page.keyboard.press('Space')
    assert page.locator('#professional-panel').get_attribute('open') is not None
    page.keyboard.press('Tab')
    assert page.locator('#doc-task-profile').evaluate('node => node === document.activeElement')
    result.update(return_controls_wheel_reachable=True, collapse_restores_prompt=True, keyboard_preserved=True)
    return result


@pytest.mark.parametrize('width,theme', [(520, 'light'), (1100, 'dark'), (520, 'highcontrast')])
def test_full_frontend_long_task_is_wheel_reachable_without_auto_scroll(tmp_path, monkeypatch, width, theme):
    session = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, replies = next(session)
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.set_viewport_size({'width': width, 'height': 800})
        setup(page)
        page.evaluate('theme => document.documentElement.dataset.theme=theme', theme)
        exercise(page)
        assert not calls and not app.editor.audit.exists() and not errors
    finally:
        session.close()
