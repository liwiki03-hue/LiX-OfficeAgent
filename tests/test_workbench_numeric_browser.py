import json

from docx import Document

from test_pdf_header_rows import header_pdf
from test_professional_browser import _browser, STATIC
from test_project_workbench import CSL
from office_agent.writes import digest


def exercise_numeric(tmp_path, monkeypatch, real_render=False):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        doc = Document()
        doc.add_paragraph('Selected value is 20.')
        doc.add_paragraph('Non-target context remains unchanged.')
        doc.save(app.path)
        original = digest(app.path)
        (tmp_path / 'table.pdf').write_bytes(header_pdf())
        allowed = {'workbench_snapshot','workbench_action','read_structure_anchors','propose_pdf_layout','read_pdf_numeric',
            'acknowledge_pdf_numeric','verify_paragraph_numeric','plan_workbench_queue','confirm_workbench_queue',
            'rollback_professional_edit','read_citation_workspace'}
        def invoke(method, args):
            assert method in allowed, method
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('numericWorkbench', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.numericWorkbench(method,args).then(reply=>{window.numericWorkbenchLastReply={method,reply};return reply})});void 0")
        for script in ('pdf_preview.js','continuous_research.js','workbench.js'):
            page.add_script_tag(path=str(STATIC / script))
        page.locator('#wb-load').click()
        page.wait_for_function("document.getElementById('wb-target').options.length===2")
        page.locator('#workbench-panel summary').filter(has_text='CSL-JSON导入与导出').click()
        page.locator('#wb-import').fill(json.dumps(CSL))
        page.locator('#wb-import-run').click()
        page.wait_for_function("document.querySelectorAll('#wb-records input[type=radio]').length===1")
        page.locator('#wb-records input[type=radio]').check()
        page.locator('#wb-decision').select_option('include')
        page.locator('#wb-reason').fill('Owned numeric fixture.')
        page.locator('#wb-review').click()
        page.wait_for_function("document.getElementById('wb-records').textContent.includes('纳入')")
        page.locator('#wb-file').fill('table.pdf')
        page.locator('#wb-associate').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('已关联当前文件身份')")
        page.locator('#wb-actions button[data-action=numeric]').click()
        page.wait_for_function("document.getElementById('wb-numeric-slot').contains(document.getElementById('cr-check'))")
        assert page.locator('#cr-pdf-table-file').input_value() == 'table.pdf'
        assert page.locator('#cr-pdf-table-toggle').get_attribute('aria-expanded') == 'true'
        page.locator('#cr-pdf-table-region').fill('2,16,98,62')
        page.locator('#cr-pdf-table-propose').click()
        page.wait_for_function("document.getElementById('cr-pdf-table-rows').value.split(',').length===4")
        page.locator('#cr-pdf-table-header-rows').select_option('2')
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        parsed = next(reply['result'] for method, reply in reversed(replies) if method=='read_pdf_numeric')
        row = next(row for row in parsed['observations'] if row['raw']=='20')
        page.locator('#cr-pdf-table-ack').click()
        page.wait_for_function("document.getElementById('cr-observation').options.length>0")
        page.locator('#cr-observation').select_option(row['id'])
        page.locator('#cr-value').fill('20')
        page.locator('#cr-unit').fill('kg')
        page.locator('#cr-verify').click()
        page.wait_for_function("!document.getElementById('wb-numeric-queue').disabled")
        assert 'unit_mismatch' in page.locator('#cr-check-result').inner_text()
        page.locator('#wb-numeric-queue').click()
        page.wait_for_function("document.getElementById('wb-queue').textContent.includes('待审阅')")
        page.locator('#wb-reviewed').check()
        page.locator('#wb-accept').click()
        page.wait_for_function("document.getElementById('wb-queue').textContent.includes('已接受待写')")
        page.locator('#wb-plan').click()
        page.wait_for_function("!document.getElementById('wb-confirm').disabled")
        page.locator('#wb-confirm').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('副本已保存并验证')||(window.numericWorkbenchLastReply?.method==='confirm_workbench_queue'&&!window.numericWorkbenchLastReply.reply.ok)")
        assert page.locator('#wb-status').inner_text().find('副本已保存并验证') >= 0, replies[-1]
        saved = next(reply['result'] for method, reply in reversed(replies) if method=='confirm_workbench_queue')
        page.locator('#wb-undo').click()
        page.wait_for_function("document.getElementById('wb-output').textContent.includes('rolled_back')")
        assert digest(app.path)==original and not list(tmp_path.glob('*.agent-*.docx'))
        return {'unit_mismatch_detected':True,'numeric_source_autofilled':True,'host_numeric_proof_bound':True,
            'candidate_not_fact_review':True,'comment_copy_render_undo':True,'source_unchanged':True,'render':saved['render']}
    finally:
        browser.close()


def test_numeric_candidate_to_workbench_comment_transaction(tmp_path, monkeypatch):
    exercise_numeric(tmp_path, monkeypatch)
