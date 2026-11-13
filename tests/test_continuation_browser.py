from pathlib import Path

from office_agent.writes import digest
from test_continuation_receipts import saved_fixture
from test_professional_browser import _browser, STATIC


def exercise(tmp_path, monkeypatch, real_render=False, capture=None):
    selected, services, first, receipt, original = saved_fixture(tmp_path, monkeypatch, real_render)
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        app.path = selected[0]
        before = digest(app.path)
        allowed = {'read_structure_anchors', 'read_citation_workspace', 'plan_structure_edit', 'confirm_professional_edit', 'rollback_professional_edit'}
        def invoke(method, args):
            assert method in allowed
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('continuationHost', invoke)
        page.evaluate("window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.continuationHost(method,args)}); void 0")
        errors = []
        page.on('pageerror', lambda e: errors.append(str(e)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-write').click()
        page.locator('#cr-receipt').fill(receipt)
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('已恢复编号')")
        assert page.locator('#cr-number-map select').input_value()
        assert page.locator('#cr-scope input[data-scope="bibliography"]:checked').count() == 1
        assert page.locator('#cr-scope input[data-scope="authorized"]:checked').count() == 0
        page.locator('#cr-text').fill('Additional context.')
        page.locator('#cr-add-insert').click()
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.evaluate('window.approved = false')
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消')")
        assert len(list(tmp_path.glob('*.agent-*.docx'))) == 1
        page.evaluate('window.approved = true')
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('副本已保存')")
        saved = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_professional_edit')
        assert digest(app.path) == before and Path(saved['output']).exists()
        if capture:
            capture(page)
        page.locator('#cr-undo').click()
        page.wait_for_function("document.getElementById('cr-output').textContent.includes('已撤销')")
        assert not Path(saved['output']).exists() and digest(app.path) == before
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#cr-receipt').input_value() == ''
        assert not errors
        return {'browser': 'Edge headless', 'restore_candidates_only': True, 'cancel_save_undo': True,
                'source_unchanged': True, 'logout_cleared': True, 'render': saved['render'], 'errors': errors}
    finally:
        browser.close()


def test_structure_receipt_browser_restore_cancel_save_undo(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
