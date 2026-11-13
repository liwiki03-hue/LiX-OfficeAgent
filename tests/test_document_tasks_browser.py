from pathlib import Path
import shutil

from tests.test_professional_browser import _browser, STATIC
from office_agent.writes import digest
from docx import Document


def test_browser_complete_task_cancel_three_groups_archive_and_logout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        shutil.copyfile(Path(__file__).resolve().parents[1] / 'deliverables/docx-tasks-20261021/development.docx', app.path)
        original = digest(app.path)
        def invoke(method, args):
            assert method in {'navigate_document_task', 'create_document_task', 'plan_document_task_group', 'confirm_document_task_group',
                              'get_document_task', 'plan_document_task_archive', 'confirm_document_task_archive'}
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('taskHost', invoke)
        page.evaluate("() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.taskHost(method, args)}); }")
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        assert calls == []
        page.locator('#doc-task-navigate').click()
        page.wait_for_function("document.getElementById('doc-task-section').options.length === 4")
        page.locator('#doc-task-request').fill('将全文所有“旧术语”替换为“规范术语”，保留数字、单位、公式和引文')
        page.locator('#doc-task-create').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('0/3')")
        assert page.locator('#doc-task-ledger > div').count() == 20
        page.locator('#doc-task-plan').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.evaluate('window.approved = false')
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('已取消')")
        assert not app.editor.audit.exists()
        page.evaluate('window.approved = true')
        for count in (1, 2, 3):
            page.locator('#doc-task-plan').click()
            page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
            page.locator('#doc-task-confirm').click()
            page.wait_for_function(f"document.getElementById('doc-task-summary').textContent.includes('{count}/3')")
        assert '支持范围已执行且检查通过' in page.locator('#doc-task-summary').inner_text()
        assert digest(app.path) == original
        page.locator('#doc-task-save').click()
        page.wait_for_function("document.getElementById('doc-task-receipt').value.length === 32")
        assert all(reply['ok'] for _, reply in receipts), receipts
        page.set_viewport_size({'width': 520, 'height': 900})
        assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth + 1')
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#doc-task-request').input_value() == ''
        assert page.locator('#doc-task-receipt').input_value() == ''
        assert page.locator('#doc-task-confirm').is_disabled()
    finally:
        fixture.close()


def test_browser_cross_paragraph_clean_and_read_only_return(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch, real_render=True)
    try:
        page, api, app, calls, receipts = next(fixture)
        source, original = app.path, digest(app.path)
        def invoke(method, args):
            assert method in {'navigate_document_task', 'create_docx_block_task', 'plan_document_task_group', 'confirm_document_task_group', 'compare_document_return'}
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('taskHost', invoke)
        page.evaluate("() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.taskHost(method, args)}); }")
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        page.locator('#doc-task-navigate').click()
        page.wait_for_function("document.getElementById('doc-task-block-first').options.length > 2")
        page.locator('#doc-task-block-action').evaluate("element => element.closest('details').open = true")
        page.locator('#doc-task-block-first').select_option(index=1)
        page.locator('#doc-task-block-last').select_option(index=2)
        page.locator('#doc-task-block-texts').fill('样本30个，测得误差为12.5% [1]，计算采用 $x=2$。\n---\nUnchanged paragraph with preserved styling.')
        page.locator('#doc-task-block-create').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('0/1')")
        assert not app.editor.audit.exists()
        page.locator('#doc-task-plan').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('1/1')")
        saved = next(reply['result'] for method, reply in receipts if method == 'confirm_document_task_group')
        assert saved['status'] == 'saved' and digest(source) == original
        assert saved['render']['status'] == 'different' and saved['alignment']
        delivered = tmp_path / saved['task']['current_name']
        doc = Document(delivered)
        assert doc.paragraphs[1].runs[0].bold
        assert doc.tables[0].cell(1, 0).text == '20 ms'
        doc.paragraphs[0].text = '用户重新提供的实验报告'
        returned = tmp_path / 'returned.docx'
        doc.save(returned)
        app.path = returned
        api._document_tasks.invalidate()
        page.locator('#file-name').evaluate("element => { element.textContent = 'returned.docx'; }")
        page.locator('#doc-task-return-baseline').evaluate("element => element.closest('details').open = true")
        page.locator('#doc-task-return-baseline').fill(source.name)
        page.locator('#doc-task-return-delivered').fill(delivered.name)
        page.locator('#doc-task-return-receipts').fill(saved['save_id'])
        hashes = [digest(path) for path in (source, delivered, returned)]
        page.locator('#doc-task-return-compare').click()
        page.wait_for_function("document.getElementById('doc-task-return-output').textContent.includes('用户新增文字修改')")
        assert hashes == [digest(path) for path in (source, delivered, returned)]
        assert all(reply['ok'] for _, reply in receipts)
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#doc-task-return-receipts').input_value() == ''
    finally:
        fixture.close()
