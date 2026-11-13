from pathlib import Path

from docx import Document

from office_agent.writes import digest
from test_professional_browser import _browser, STATIC
from test_table_context import material


def exercise(tmp_path, monkeypatch, real_render=False, workbook=None, capture=None, retain=None):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        value = '0.3334' if workbook else '20'
        text = 'The mass fraction is 0.3334%.' if workbook else 'The mean latency is 20 ms.'
        document = Document()
        document.add_paragraph(text)
        document.add_paragraph('Original non-target content remains unchanged.')
        document.save(app.path)
        original = digest(app.path)
        if workbook:
            (tmp_path / 'table.xlsx').write_bytes(Path(workbook).read_bytes())
        else:
            material(tmp_path, '.xlsx')
        methods = {'read_structure_anchors', 'read_citation_workspace', 'read_numeric_source', 'verify_paragraph_numeric',
                   'plan_structure_edit', 'confirm_professional_edit', 'rollback_professional_edit'}
        def invoke(method, args):
            assert method in methods
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('contextHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.contextHost(method,args)});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-check').click()
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-anchor').options.length===2")
        page.locator('#cr-table-file').fill('table.xlsx')
        page.locator('#cr-table-load').click()
        page.wait_for_function("document.getElementById('cr-observation').options.length>1")
        records = next(reply['result']['observations'] for method, reply in replies if method == 'read_numeric_source')
        row = next(row for row in records if row['raw'] == ('0.3334' if workbook else '20 ± 2[a]'))
        page.locator('#cr-observation').select_option(row['id'])
        page.locator('#cr-value').fill(value)
        page.locator('#cr-unit').fill('%' if workbook else 'ms')
        page.locator('#cr-statistic').select_option('value' if workbook else 'mean')
        page.locator('#cr-check details').evaluate('(node)=>node.open=true')
        page.locator('#cr-conditions').fill('unsupported condition' if workbook else 'wet')
        page.locator('#cr-verify').click()
        page.wait_for_function("document.getElementById('cr-check-result').textContent.includes('context_issues')")
        first = next(reply['result'] for method, reply in reversed(replies) if method == 'verify_paragraph_numeric')
        assert first['status'] == ('insufficient' if workbook else 'condition_mismatch')
        assert first['numeric_status'] == 'insufficient' and not first['user_reviewed']
        page.locator('#cr-conditions').fill('')
        page.locator('#cr-unit').fill('kg')
        page.locator('#cr-verify').click()
        page.wait_for_function("document.getElementById('cr-check-result').textContent.includes('unit_mismatch')")
        if capture:
            capture(page)
        page.locator('#cr-check-comment').click()
        page.locator('#cr-add-comment').click()
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.evaluate('window.approved=false')
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消')")
        assert not list(tmp_path.glob('*.agent-*.docx'))
        page.evaluate('window.approved=true')
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('副本已保存')")
        assert digest(app.path) == original
        saved = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_professional_edit')
        if retain:
            retain(app.path, next(tmp_path.glob('*.agent-*.docx')))
        page.locator('#cr-undo').click()
        page.wait_for_function("document.getElementById('cr-output').textContent.includes('已撤销')")
        assert not list(tmp_path.glob('*.agent-*.docx')) and digest(app.path) == original
        assert not errors
        return {'context_status': first['status'], 'unit_mismatch': True, 'comment_plan_cancel_save_undo': True,
                'source_unchanged': True, 'render': saved['render'], 'errors': errors}
    finally:
        browser.close()


def test_table_context_browser(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
