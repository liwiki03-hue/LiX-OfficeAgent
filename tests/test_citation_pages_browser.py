from docx import Document

from office_agent.writes import digest
from test_citations_browser import citations_browser


def long_browser(tmp_path, monkeypatch, *, real_render=False):
    for page, api, app, calls, receipts, first, second, archived in citations_browser(tmp_path, monkeypatch, real_render=real_render):
        doc = Document()
        doc.add_paragraph('The measured latency was 20 ms [4].')
        for index in range(79):
            doc.add_paragraph('Unchanged technical discussion ' + str(index) + '.')
        doc.add_paragraph('Independent result [2], repeated [7].')
        doc.add_paragraph('[2] First experiment. doi: 10.1234/expt.01')
        doc.add_paragraph('[4] Second experiment. doi: 10.1234/expt.02')
        doc.add_paragraph('[7] Duplicate first experiment. doi: 10.1234/expt.01')
        doc.add_paragraph('Reference placeholder')
        doc.save(app.path)
        yield page, api, app, calls, receipts, first, second, archived


def prepare_long(page, first, second):
    page.locator('#citations-load').click()
    page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
    for index, record in enumerate([first, second, first]):
        page.locator('#citations-mapping select').nth(index).select_option(record)
    page.locator('#citations-targets input.citations-authorized').first.check()
    page.locator('#citations-next').click()
    page.wait_for_function("document.getElementById('citations-page').textContent.includes('41—80')")
    page.locator('#citations-next').click()
    page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 5")
    page.locator('#citations-targets input.citations-authorized').evaluate_all('(nodes) => nodes.forEach(node => node.click())')
    for index in range(1, 5):
        page.locator('#citations-targets article').nth(index).locator('select.citations-role').select_option('bibliography')


def test_cross_page_ui_selection_plan_cancel_save_and_rollback(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in long_browser(tmp_path, monkeypatch):
        original = digest(app.path)
        prepare_long(page, first, second)
        assert '6/12' in page.locator('#citations-selected').inner_text()
        page.locator('#citations-selected button').first.click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('1—40')")
        assert page.locator('#citations-targets input.citations-authorized').first.is_checked()
        assert page.locator('#citations-mapping select').nth(1).input_value() == second
        page.locator('#citations-targets article').first.locator('label input:not([type="checkbox"])').fill('4')
        page.locator('#citations-prev').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 5")
        page.locator('#citations-selected button').first.click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('1—40')")
        assert page.locator('#citations-targets article').first.locator('label input:not([type="checkbox"])').input_value() == '4'
        page.locator('#citations-targets article').first.locator('label input:not([type="checkbox"])').fill('')
        before = calls.count('read_citation_workspace')
        page.evaluate('window.approved = false')
        page.locator('#citations-load').click()
        assert calls.count('read_citation_workspace') == before
        assert '6/12' in page.locator('#citations-selected').inner_text()
        page.locator('#citations-reviewed').check()
        page.locator('#citations-plan').click()
        page.wait_for_function("!document.getElementById('citations-confirm').disabled")
        page.locator('#citations-position').fill('81')
        page.locator('#citations-jump').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 5")
        assert page.locator('#citations-confirm').is_enabled()
        assert page.locator('#citations-targets select.citations-role').evaluate_all('(nodes) => nodes.slice(1).every(node => node.value === "bibliography")')
        page.locator('#citations-confirm').click()
        page.wait_for_function("document.getElementById('citations-status').textContent.includes('已取消')")
        assert len(list(tmp_path.glob('*.docx'))) == 1 and digest(app.path) == original
        page.locator('#citations-plan').click()
        page.wait_for_function("!document.getElementById('citations-confirm').disabled")
        page.evaluate('window.approved = true')
        page.locator('#citations-confirm').click()
        page.wait_for_function("document.getElementById('citations-output').textContent.includes('已核验副本')")
        planned = [reply['result'] for method, reply in receipts if method == 'plan_citation_edit'][-1]
        assert planned['manifest']['number_changes'] == {'2': 2, '4': 1, '7': 2}
        saved = next(reply['result'] for method, reply in receipts if method == 'confirm_citation_edit' and reply.get('result', {}).get('status') == 'saved')
        assert Document(saved['output']).paragraphs[80].text == 'Independent result [2], repeated [2].'
        assert len(planned['items']) <= 12 and len(planned['manifest']['number_to_key']) == 2
        page.locator('#citations-rollback').click()
        page.wait_for_function("document.getElementById('citations-output').textContent.includes('已撤销')")
        assert digest(app.path) == original and len(list(tmp_path.glob('*.docx'))) == 1


def test_page_response_duplicate_click_and_logout_clear(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in long_browser(tmp_path, monkeypatch):
        prepare_long(page, first, second)
        expected = api.read_citation_workspace('', 0, digest(app.path))
        page.evaluate('''() => {
          window.readCount = 0;
          window.pywebview.api = {read_citation_workspace: () => {window.readCount++; return new Promise(resolve => {window.finishPage = resolve;});}};
        }''')
        page.locator('#citations-prev').click()
        page.evaluate("document.getElementById('citations-prev').click()")
        assert page.evaluate('window.readCount') == 1
        page.evaluate('''reply => {
          window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'));
          window.finishPage(reply);
        }''', expected)
        assert page.locator('#citations-targets').inner_text() == ''
        assert page.locator('#citations-selected').inner_text() == ''
        assert page.locator('#citations-mapping').inner_text() == ''
        assert page.locator('#citations-next').is_disabled() and page.locator('#citations-plan').is_disabled()


def test_changed_source_blocks_paging_keeps_drafts_for_review_only(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in long_browser(tmp_path, monkeypatch):
        page.locator('#citations-load').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
        page.locator('#citations-targets input.citations-authorized').first.check()
        doc = Document(app.path)
        doc.paragraphs[50].text = 'External revision.'
        doc.save(app.path)
        page.locator('#citations-next').click()
        page.wait_for_function("document.getElementById('citations-status').textContent.includes('DOCX 已改变')")
        assert '1/12' in page.locator('#citations-selected').inner_text()
        assert page.locator('#citations-plan').is_disabled() and page.locator('#citations-jump').is_disabled()
        assert page.locator('#citations-targets input.citations-authorized').first.is_disabled()
        assert page.locator('#citations-load').is_enabled()
        page.locator('#citations-load').click()
        page.wait_for_function("document.getElementById('citations-status').textContent.includes('支持翻页与跳转')")
        assert '0/12' in page.locator('#citations-selected').inner_text()
        assert page.locator('#citations-next').is_enabled()


def test_reading_other_pages_does_not_authorize_and_selection_has_bound(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in long_browser(tmp_path, monkeypatch):
        page.locator('#citations-load').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
        controls = page.locator('#citations-targets input.citations-authorized')
        for index in range(12):
            controls.nth(index).check()
        assert controls.nth(12).is_disabled()
        page.locator('#citations-next').click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('41—80')")
        assert page.locator('#citations-targets input.citations-authorized').first.is_disabled()
        assert page.locator('#citations-targets select.citations-role').first.is_disabled()
        page.locator('#citations-selected button').first.click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('1—40')")
        page.locator('#citations-targets input.citations-authorized').first.uncheck()
        assert '11/12' in page.locator('#citations-selected').inner_text()
        assert page.locator('#citations-targets input.citations-authorized').nth(12).is_enabled()
        assert calls.count('plan_citation_edit') == 0 and len(list(tmp_path.glob('*.docx'))) == 1


def test_evidence_review_draft_survives_pages_and_explicit_discard(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in long_browser(tmp_path, monkeypatch):
        page.locator('#grounded-archive').fill(archived['save_id'])
        page.locator('#grounded-archive-capture').click()
        page.wait_for_function("document.getElementById('grounded-evidence').children.length > 0")
        page.locator('#citations-load').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
        card = page.locator('#citations-targets article').first
        card.locator('input.citations-authorized').check()
        card.locator('select.citations-role').select_option('claim')
        card.get_by_role('button', name='增加一篇原文支持来源').click()
        capture = next(iter(api._grounded._captures))
        card.locator('fieldset select').first.select_option(capture)
        card.locator('fieldset > div input[type="checkbox"]').first.check()
        card.locator('fieldset select').nth(1).select_option('supports')
        card.locator('fieldset textarea').nth(0).fill('Laboratory only')
        card.locator('fieldset textarea').nth(1).fill('Synthetic fixture')
        card.locator('fieldset input[type="checkbox"]').nth(-2).check()
        card.locator('fieldset input[type="checkbox"]').nth(-1).check()
        page.locator('#citations-next').click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('41—80')")
        page.locator('#citations-prev').click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('1—40')")
        assert card.locator('select.citations-role').input_value() == 'claim'
        assert card.locator('fieldset select').first.input_value() == capture
        assert card.locator('fieldset textarea').nth(0).input_value() == 'Laboratory only'
        assert card.locator('fieldset input[type="checkbox"]').nth(0).is_checked()
        assert card.locator('fieldset input[type="checkbox"]').nth(-1).is_checked()
        assert page.evaluate('document.activeElement.id') == 'citations-targets'
        page.evaluate('window.approved = false')
        card.locator('input.citations-authorized').click()
        assert card.locator('input.citations-authorized').is_checked() and card.locator('fieldset').count() == 1
        page.evaluate('window.approved = true')
        card.locator('input.citations-authorized').uncheck()
        assert card.locator('fieldset').count() == 0
        assert '0/12' in page.locator('#citations-selected').inner_text()
        assert calls.count('plan_citation_edit') == 0
