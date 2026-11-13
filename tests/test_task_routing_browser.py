from pathlib import Path
import shutil

from tests.test_professional_browser import _browser, STATIC
from office_agent.writes import digest


def test_recommended_route_compare_cancel_two_transactions_and_logout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        shutil.copyfile(Path(__file__).resolve().parents[1] / 'deliverables/docx-tasks-20261021/development.docx', app.path)
        original = digest(app.path)
        def invoke(method, args):
            assert method in {'create_document_task', 'plan_document_task_group', 'confirm_document_task_group'}
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('taskHost', invoke)
        page.evaluate("() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.taskHost(method, args)}); }")
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        page.locator('#doc-task-request').fill('将全文“旧术语”替换为“规范术语”')
        page.locator('#doc-task-create').click()
        page.wait_for_function("!document.getElementById('doc-task-route').disabled")
        page.locator('#doc-task-route').click()
        page.wait_for_function("document.getElementById('doc-task-preview').textContent.includes('word_batch_replace')")
        assert '账本30项' in page.locator('#doc-task-preview').inner_text()
        assert '计划页图上下文：12/30' in page.locator('#doc-task-preview').inner_text()
        assert not app.editor.audit.exists()
        page.locator('#doc-task-plan').click()
        page.wait_for_function("document.getElementById('doc-task-preview').textContent.includes('bounded_ranges')")
        page.locator('#doc-task-route').click()
        page.wait_for_function("document.getElementById('doc-task-preview').textContent.includes('word_batch_replace')")
        page.evaluate('window.approved = false')
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('已取消')")
        assert not app.editor.audit.exists()
        page.evaluate('window.approved = true')
        for primitive, completed in [('word_batch_replace', 30), ('word_table_cell_replace', 31)]:
            page.locator('#doc-task-route').click()
            page.wait_for_function(f"document.getElementById('doc-task-preview').textContent.includes('{primitive}')")
            page.locator('#doc-task-confirm').click()
            page.wait_for_function(f"document.getElementById('doc-task-summary').textContent.includes('已执行并验证 {completed}')")
        assert len(list(app.editor.audit.iterdir())) == 2
        assert digest(app.path) == original
        assert all(reply['ok'] for _, reply in receipts)
        page.set_viewport_size({'width': 520, 'height': 900})
        assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth + 1')
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#doc-task-route').is_disabled()
        assert page.locator('#doc-task-request').input_value() == ''
    finally:
        fixture.close()
