from docx import Document

from office_agent.writes import digest
from test_pdf_header_rows import header_pdf, GRID, ROWS
from test_professional_browser import _browser, STATIC


def exercise_headers(tmp_path, monkeypatch, *, real_render=False, capture=None):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        document = Document()
        document.add_paragraph('Measured latency is 20 ms.')
        document.add_paragraph('Non-target content remains unchanged.')
        document.save(app.path)
        original = digest(app.path)
        (tmp_path / 'table.pdf').write_bytes(header_pdf())
        def invoke(method, args):
            assert method in {'read_structure_anchors', 'read_citation_workspace', 'preview_pdf_numeric', 'read_pdf_numeric',
                              'acknowledge_pdf_numeric', 'verify_paragraph_numeric', 'plan_structure_edit',
                              'confirm_professional_edit', 'rollback_professional_edit'}
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('headerGridHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.headerGridHost(method,args)});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-check').click()
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-anchor').options.length===2")
        page.locator('#cr-pdf-table-toggle').click()
        page.locator('#cr-pdf-table-file').fill('table.pdf')
        page.locator('#cr-pdf-table-preview').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-image').hidden")
        for field in ('rows', 'columns'):
            page.locator('#cr-pdf-table-' + field).fill(','.join(str(value * 100) for value in GRID[field]))
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        single = next(reply['result'] for method, reply in reversed(replies) if method == 'read_pdf_numeric')
        old_row = single['observations'][-1]
        assert old_row['unit'] is None and single['header_rows'] == 1
        before = len(calls)
        page.locator('#cr-pdf-table-header-rows').select_option('2')
        assert page.locator('#cr-pdf-table-ack').is_disabled()
        assert page.locator('#cr-pdf-table-cells').inner_text() == ''
        assert len(calls) == before
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        double = next(reply['result'] for method, reply in reversed(replies) if method == 'read_pdf_numeric')
        assert double['rows'] == ROWS and double['header_rows'] == 2
        row = double['observations'][-1]
        assert row['unit'] == 'ms' and row['sample'] == 'A' and row['id'] != old_row['id']
        assert page.locator('#cr-pdf-table-cells th').count() == 4
        assert page.locator('#cr-pdf-table-cells td').count() == 2
        if capture:
            capture(page)
        page.locator('#cr-pdf-table-ack').click()
        page.wait_for_function("document.getElementById('cr-observation').options.length===6")
        page.locator('#cr-observation').select_option(row['id'])
        page.locator('#cr-value').fill('20')
        page.locator('#cr-unit').fill('ms')
        page.locator('#cr-check summary').click()
        page.locator('#cr-sample').fill('A')
        page.locator('#cr-verify').click()
        page.wait_for_function("document.getElementById('cr-check-result').textContent.includes('needs_manual_judgment')")
        checked = next(reply['result'] for method, reply in reversed(replies) if method == 'verify_paragraph_numeric')
        assert checked['numeric_status'] == 'match' and not checked['user_reviewed']
        page.locator('#cr-check-comment').click()
        page.locator('#cr-add-comment').click()
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        plan = next(reply['result'] for method, reply in reversed(replies) if method == 'plan_structure_edit')
        assert 'pdf_numeric_review' in plan['after']
        page.evaluate('window.approved=false')
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消')")
        assert not list(tmp_path.glob('*.agent-*.docx'))
        page.evaluate('window.approved=true')
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('副本已保存')")
        saved = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_professional_edit')
        page.locator('#cr-undo').click()
        page.wait_for_function("document.getElementById('cr-output').textContent.includes('已撤销')")
        assert digest(app.path) == original and not list(tmp_path.glob('*.agent-*.docx'))
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#cr-pdf-table-header-rows').input_value() == '1'
        assert not errors
        return {'rows': double['rows'], 'single_header_unit': old_row['unit'], 'double_header_unit': row['unit'],
                'header_sources': row['header_sources'], 'row_location': row['location'],
                'interpretation_changes_observation_id': True, 'numeric_status': checked['numeric_status'],
                'user_reviewed': False, 'source_unchanged': True, 'cancel_save_render_audit_undo': True,
                'render': saved['render'], 'errors': errors}
    finally:
        browser.close()


def test_explicit_double_header_browser_to_comment(tmp_path, monkeypatch):
    exercise_headers(tmp_path, monkeypatch)
