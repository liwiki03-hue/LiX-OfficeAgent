import json

from docx import Document

from office_agent.writes import digest
from test_pdf_header_rows import header_pdf
from test_professional_browser import _browser, STATIC
from test_project_workbench import CSL


def exercise_workbench(tmp_path, monkeypatch, *, real_render=False, capture=None):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        doc = Document()
        doc.add_paragraph('Measured latency is 20 ms.')
        doc.add_paragraph('Non-target context remains unchanged.')
        doc.add_paragraph('References follow.')
        doc.save(app.path)
        original = digest(app.path)
        (tmp_path / 'table.pdf').write_bytes(header_pdf([['Sample', 'Latency'], ['Identifier', '(ms)'], ['A', '20 ms']]))
        allowed = {'workbench_snapshot', 'workbench_action', 'read_structure_anchors', 'plan_project_save', 'confirm_project_save',
                   'restore_project', 'plan_project_removal', 'confirm_project_removal', 'plan_workbench_queue',
                   'confirm_workbench_queue', 'rollback_professional_edit', 'read_citation_workspace', 'read_evidence_page'}
        def invoke(method, args):
            assert method in allowed, method
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('workbenchHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.workbenchHost(method,args).then(reply=>{window.workbenchLastMethod=method;return reply})});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        for name in ['pdf_preview.js', 'continuous_research.js', 'evidence_page.js', 'workbench.js']:
            page.add_script_tag(path=str(STATIC / name))
        assert not calls
        page.locator('#wb-load').click()
        page.wait_for_function("document.getElementById('wb-target').options.length===3")
        page.locator('#workbench-panel summary').filter(has_text='CSL-JSON导入与导出').click()
        page.locator('#wb-import').fill(json.dumps(CSL))
        page.locator('#wb-import-run').click()
        page.wait_for_function("document.querySelectorAll('#wb-records input[type=radio]').length===1")
        page.locator('#wb-records input[type=radio]').check()
        page.locator('#wb-decision').select_option('include')
        page.locator('#wb-reason').fill('Owned developer fixture for provenance checks.')
        page.locator('#wb-review').click()
        page.wait_for_function("document.getElementById('wb-records').textContent.includes('纳入')")
        page.locator('#wb-file').fill('table.pdf')
        page.locator('#wb-associate').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('已关联当前文件身份')")
        page.locator('#wb-capture').click()
        page.wait_for_function("document.querySelectorAll('#wb-evidence input[type=checkbox]').length>0")
        first = page.locator('#wb-evidence input[type=checkbox]').first
        first.check()
        page.locator('#wb-operation').select_option('insert')
        page.locator('#wb-text').fill('The selected experiment reports latency of 20 ms.')
        page.locator('#wb-queue-add').click()
        page.wait_for_function("document.getElementById('wb-queue').textContent.includes('待审阅')")
        page.locator('#wb-relation').select_option('supports')
        page.locator('#wb-applicability').fill('Owned sample A only.')
        page.locator('#wb-limitation').fill('Developer fixture; not independent scientific validation.')
        page.locator('#wb-reviewed').check()
        page.locator('#wb-accept').click()
        page.wait_for_function("document.getElementById('wb-queue').textContent.includes('已接受待写')")
        last = page.locator('#wb-bib-anchor option').last.get_attribute('value')
        page.locator('#wb-bib-anchor').select_option(last)
        page.locator('#wb-plan').click()
        page.wait_for_function("window.workbenchLastMethod==='plan_workbench_queue'")
        planned = next(reply for method, reply in reversed(replies) if method == 'plan_workbench_queue')
        assert planned['ok'], planned
        page.wait_for_function("!document.getElementById('wb-confirm').disabled")
        assert 'bibliography_insert' in page.locator('#wb-preview').inner_text()
        page.evaluate('window.approved=false')
        page.locator('#wb-confirm').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('已取消本次文档写入')")
        assert not list(tmp_path.glob('*.agent-*.docx'))
        page.evaluate('window.approved=true')
        page.locator('#wb-plan').click()
        page.wait_for_function("!document.getElementById('wb-confirm').disabled")
        page.locator('#wb-confirm').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('副本已保存并验证')")
        written = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_workbench_queue')
        if capture:
            capture(page)
        page.locator('#wb-undo').click()
        page.wait_for_function("document.getElementById('wb-output').textContent.includes('rolled_back')")
        assert digest(app.path) == original
        page.locator('#wb-save').click()
        page.wait_for_function("!document.getElementById('wb-save-confirm').disabled")
        page.locator('#wb-save-confirm').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('项目快照已保存')")
        receipt = page.locator('#wb-receipt').input_value()
        assert len(receipt) == 32
        api._workbench.invalidate()
        page.locator('#wb-restore').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('项目已恢复')")
        assert '需重新核对' in page.locator('#wb-queue').inner_text()
        assert page.locator('#wb-evidence input').count() == 0
        assert not errors
        return {'insert_bibliography_cancel_save_render_undo': True, 'source_unchanged': True,
                'library_restore_keeps_metadata_not_permissions': True, 'host_ids_not_manually_copied': True,
                'project_receipt': receipt, 'render': written['render'], 'calls': calls, 'errors': errors}
    finally:
        browser.close()


def test_workbench_library_to_reviewed_copy_and_restore(tmp_path, monkeypatch):
    exercise_workbench(tmp_path, monkeypatch)
