import shutil
from pathlib import Path

from office_agent.writes import digest
from tests.test_professional_browser import STATIC, _browser


def setup_range(page, api, app, calls, receipts):
    source = Path(__file__).resolve().parents[1] / 'deliverables/docx-fidelity-20261020/challenges/development.docx'
    shutil.copyfile(source, app.path)
    def invoke(method, args):
        assert method in {'read_edit_ranges', 'plan_range_edit', 'confirm_professional_edit', 'rollback_professional_edit'}
        calls.append(method)
        reply = getattr(api, method)(*args)
        receipts.append((method, reply))
        return reply
    page.expose_function('rangeHost', invoke)
    page.evaluate("() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.rangeHost(method, args)}); }")
    page.add_script_tag(path=str(STATIC / 'range_edits.js'))


def select_edit(page):
    page.locator('#range-load').click()
    page.wait_for_function("document.getElementById('range-target').options.length > 1")
    page.locator('#range-target').select_option(index=1)
    page.locator('#range-original').evaluate('(node) => {node.focus(); node.setSelectionRange(7, 13);}')
    page.locator('#range-text').fill('second')
    page.locator('#range-reason').fill('Replace only the second occurrence')
    page.locator('#range-add').click()
    page.locator('#range-plan').click()
    page.wait_for_function("!document.getElementById('range-confirm').disabled")


def test_browser_selection_cancel_save_undo_and_logout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup_range(page, api, app, calls, receipts)
        version = digest(app.path)
        select_edit(page)
        assert not app.editor.audit.exists()
        assert 'repeat second' not in page.locator('#range-original').input_value()
        page.evaluate('window.approved = false')
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-status').textContent.includes('已取消')")
        assert digest(app.path) == version and not app.editor.audit.exists()
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        page.evaluate('window.approved = true')
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已保存副本')")
        assert digest(app.path) == version
        page.locator('#range-undo').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已撤销')")
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#range-original').input_value() == ''
        assert page.locator('#range-plan').is_disabled()
    finally:
        fixture.close()


def test_browser_busy_late_read_and_narrow_layout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup_range(page, api, app, calls, receipts)
        page.evaluate('''() => {
          window.oldRangeApi = window.pywebview.api;
          window.rangeReads = 0;
          window.pywebview.api = {read_edit_ranges: () => {window.rangeReads++; return new Promise(resolve => window.finishRangeRead = resolve);}};
        }''')
        page.locator('#range-load').click()
        assert page.locator('#range-load').is_disabled()
        assert page.evaluate('window.rangeReads') == 1
        reply = api.read_edit_ranges(0)
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate('(reply) => window.finishRangeRead(reply)', reply)
        page.wait_for_timeout(100)
        assert page.locator('#range-target option').count() == 0
        page.set_viewport_size({'width': 520, 'height': 900})
        assert page.locator('#range-panel').evaluate('(node) => node.scrollWidth <= node.clientWidth + 1')
    finally:
        fixture.close()
