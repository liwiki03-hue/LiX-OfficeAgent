from docx import Document

from office_agent.writes import digest
from test_pdf_table import grid_pdf, GRID
from test_professional_browser import _browser, STATIC


def exercise(tmp_path, monkeypatch, source_pdf=None, real_render=False, capture=None, automatic=False, unit_choices=False, unit_capture=None):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        value = '0.3334' if source_pdf else '20'
        doc = Document()
        doc.add_paragraph('Selected value is ' + value + '.')
        doc.add_paragraph('Non-target context remains unchanged.')
        doc.save(app.path)
        before = digest(app.path)
        (tmp_path / 'table.pdf').write_bytes(source_pdf.read_bytes() if source_pdf else grid_pdf())
        grid = {'page': 94, 'rows': [.209, .227, .244], 'columns': [.28, .38, .46, .55, .65, .75], 'units': [[2, 1], [4, 1]]} if source_pdf else GRID
        methods = {'read_structure_anchors', 'read_citation_workspace', 'preview_pdf_numeric', 'read_pdf_numeric',
            'acknowledge_pdf_numeric', 'verify_paragraph_numeric', 'plan_structure_edit', 'confirm_professional_edit', 'rollback_professional_edit'}
        if automatic:
            methods.add('propose_pdf_layout')
        def invoke(method, args):
            assert method in methods
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('pdfGridHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.pdfGridHost(method,args)});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-check').click()
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-anchor').options.length===2")
        page.locator('#cr-pdf-table-toggle').click()
        page.locator('#cr-pdf-table-file').fill('table.pdf')
        page.locator('#cr-pdf-table-page').fill(str(grid['page']))
        assert not any(method.startswith('preview_pdf') for method in calls)
        page.evaluate("document.getElementById('cr-pdf-table-preview').click();document.getElementById('cr-pdf-table-preview').click()")
        page.wait_for_function("!document.getElementById('cr-pdf-table-image').hidden")
        assert calls.count('preview_pdf_numeric') == 1
        if automatic:
            page.locator('#cr-pdf-table-region').fill('28,20.9,75,24.4')
            page.locator('#cr-pdf-table-propose').click()
            page.wait_for_function("document.getElementById('cr-pdf-table-columns').value.split(',').length===6")
        else:
            for field in ('rows', 'columns'):
                page.locator('#cr-pdf-table-' + field).fill(','.join(str(round(number * 100, 4)) for number in grid[field]))
        if grid['units']:
            if unit_choices:
                assert page.locator('#cr-pdf-table-units').input_value()==''
                assert page.locator('#cr-unit-column-4').input_value()==''
                page.locator('#cr-unit-column-3').select_option('2')
                page.locator('#cr-unit-column-5').select_option('2')
                if unit_capture:
                    unit_capture(page)
                page.locator('#cr-unit-review-apply').click()
                assert page.locator('#cr-pdf-table-units').input_value()=='3:2,5:2'
                assert page.locator('#cr-observation option').count()==0
            else:
                page.locator('#cr-pdf-table-units').fill('3:2,5:2')
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        assert page.locator('#cr-observation option').count() == 0
        parsed = next(reply['result'] for method, reply in reversed(replies) if method == 'read_pdf_numeric')
        assert parsed['rows'] == ([['Measurand', 'Units', 'w', 'k', 'U'], ['S', '%', '0.3334', '2.19', '0.0078']] if source_pdf else [['Sample', 'Value (ms)'], ['A', '20']])
        assert all(not row['layout_confirmed'] for row in parsed['observations'])
        if capture:
            capture(page)
        page.locator('#cr-pdf-table-ack').click()
        page.wait_for_function("document.getElementById('cr-observation').options.length>0")
        row = next(row for row in parsed['observations'] if row['raw'] == value)
        page.locator('#cr-observation').select_option(row['id'])
        page.locator('#cr-value').fill(value)
        page.locator('#cr-unit').fill('kg')
        page.locator('#cr-verify').click()
        page.wait_for_function("document.getElementById('cr-check-result').textContent.includes('unit_mismatch')")
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
        saved = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_professional_edit')
        page.locator('#cr-undo').click()
        page.wait_for_function("document.getElementById('cr-output').textContent.includes('已撤销')")
        assert digest(app.path) == before and not list(tmp_path.glob('*.agent-*.docx'))
        page.locator('#cr-nav-check').click()
        page.evaluate("""window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.pdfGridHost(method,args).then(reply=>{
            if(method!=='preview_pdf_numeric')return reply;
            return new Promise(resolve=>window.releasePdfGrid=()=>resolve(reply)).then(value=>{window.pdfGridDelivered=true;return value});
        })});void 0""")
        page.locator('#cr-pdf-table-preview').click()
        page.wait_for_function("typeof window.releasePdfGrid==='function'")
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));window.releasePdfGrid()")
        page.wait_for_function('window.pdfGridDelivered===true')
        assert page.locator('#cr-pdf-table-image').is_hidden() and not page.locator('#cr-pdf-table-image').get_attribute('src')
        assert page.locator('#cr-pdf-table-cells').inner_text() == '' and not errors
        return {'rows': parsed['rows'], 'layout_required': True, 'layout_is_not_fact_review': True,
            'raw_candidate': next(({key: reply['result'][key] for key in ('candidate_rows', 'proposal', 'unit_candidates', 'uncertainties')}
                for method, reply in replies if method == 'propose_pdf_layout' and reply.get('ok')), None),
            'no_auto_network': True, 'duplicate_click_guard': True, 'source_unchanged': True,
            'comment_cancel_save_undo': True, 'late_response_ignored_after_logout': True, 'render': saved['render'], 'errors': errors}
    finally:
        browser.close()


def test_pdf_grid_browser(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
