from docx import Document

from office_agent.writes import digest
from test_citation_pages_browser import long_browser, prepare_long


def location_browser(tmp_path, monkeypatch, *, real_render=False):
    for page, api, app, calls, receipts, first, second, archived in long_browser(tmp_path, monkeypatch, real_render=real_render):
        def locate(args):
            calls.append('read_citation_locations')
            reply = api.read_citation_locations(*args)
            receipts.append(('read_citation_locations', reply))
            return reply
        page.expose_function('locationsHost', locate)
        page.evaluate('''() => {
          const previous = window.pywebview.api;
          window.pywebview.api = new Proxy(previous, {get: (target, key) => key === 'read_citation_locations' ? (...args) => window.locationsHost(args) : target[key]});
        }''')
        yield page, api, app, calls, receipts, first, second, archived


def find_seven(page):
    page.locator('#citations-find-number').select_option('7')
    page.locator('#citations-find').click()
    page.wait_for_function("document.querySelectorAll('#citations-locations article').length === 2")


def test_location_jump_preserves_plan_and_does_not_authorize(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in location_browser(tmp_path, monkeypatch):
        prepare_long(page, first, second)
        page.locator('#citations-reviewed').check()
        page.locator('#citations-plan').click()
        page.wait_for_function("!document.getElementById('citations-confirm').disabled")
        before = page.locator('#citations-preview').inner_text()
        find_seven(page)
        assert '不是完整重排' not in page.locator('#citations-preview').inner_text()
        page.locator('#citations-locations button').last.click()
        page.wait_for_function("document.querySelector('#citations-targets article.citations-found') !== null")
        assert page.locator('#citations-targets article.citations-found').get_attribute('data-locator')
        assert page.locator('#citations-targets article.citations-found strong').inner_text() == 'body:83/paragraph'
        assert page.locator('#citations-preview').inner_text() == before and page.locator('#citations-confirm').is_enabled()
        assert '6/12' in page.locator('#citations-selected').inner_text()
        assert not any(method == 'confirm_citation_edit' for method in calls)
        page.locator('#citations-locations button').first.click()
        page.wait_for_function("document.querySelector('#citations-targets article.citations-found strong')?.textContent === 'body:80/paragraph'")
        assert page.locator('#citations-targets article.citations-found').count() == 1


def test_jump_does_not_select_and_changed_source_invalidates_results(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in location_browser(tmp_path, monkeypatch):
        original = digest(app.path)
        page.locator('#citations-load').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
        find_seven(page)
        page.locator('#citations-locations button').first.click()
        page.wait_for_function("document.querySelector('#citations-targets article.citations-found') !== null")
        assert '0/12' in page.locator('#citations-selected').inner_text()
        assert not page.locator('#citations-targets input.citations-authorized').first.is_checked()
        assert page.locator('#citations-plan').is_disabled() and digest(app.path) == original
        page.locator('#citations-locations button').first.click()
        page.wait_for_function("!document.getElementById('citations-prev').disabled")
        page.locator('#citations-prev').click()
        page.wait_for_function("document.getElementById('citations-page').textContent.includes('1—40')")
        assert page.locator('#citations-prev').is_disabled()
        doc = Document(app.path)
        doc.paragraphs[30].text = 'External revision.'
        doc.save(app.path)
        page.locator('#citations-locations button').last.click()
        page.wait_for_function("document.getElementById('citations-status').textContent.includes('DOCX 已改变')")
        assert page.locator('#citations-find').is_disabled() and page.locator('#citations-locations button').first.is_disabled()
        assert not app.editor.audit.exists()


def test_late_location_response_logout_duplicate_and_error_retry(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in location_browser(tmp_path, monkeypatch):
        page.locator('#citations-load').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
        page.evaluate('''() => {
          window.locationCalls = 0;
          window.pywebview.api = {read_citation_locations: () => {window.locationCalls++; return new Promise(resolve => {window.finishLocations = resolve;});}};
        }''')
        page.locator('#citations-find').click()
        page.evaluate("document.getElementById('citations-find').click()")
        assert page.evaluate('window.locationCalls') == 1
        page.evaluate("window.finishLocations({ok:false,reason:'professional_failed'})")
        page.wait_for_function("!document.getElementById('citations-find').disabled")
        page.locator('#citations-find').click()
        assert page.evaluate('window.locationCalls') == 2
        reply = api.read_citation_locations(4, digest(app.path))
        page.evaluate('''reply => {
          window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'));
          window.finishLocations(reply);
        }''', reply)
        assert page.locator('#citations-locations').inner_text() == '' and page.locator('#citations-find-number').inner_text() == ''
        assert page.locator('#citations-find').is_disabled()


def test_result_paging_readonly_reason_and_html_are_safe(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in location_browser(tmp_path, monkeypatch):
        doc = Document()
        for index in range(42):
            doc.add_paragraph('Observation [4].')
        doc.add_heading('Styled [4] <img src=x onerror=alert(1)>', 1)
        doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Table [4]'
        doc.save(app.path)
        page.locator('#citations-load').click()
        page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 40")
        page.locator('#citations-find').click()
        page.wait_for_function("document.querySelectorAll('#citations-locations article').length === 40")
        page.locator('#citations-find-more').click()
        page.wait_for_function("document.querySelectorAll('#citations-locations article').length === 4")
        assert page.locator('#citations-find-more').is_disabled()
        assert '显式段落样式' in page.locator('#citations-locations').inner_text()
        assert '表格等非普通正文' in page.locator('#citations-locations').inner_text()
        assert page.locator('#citations-locations img').count() == 0
        page.locator('#citations-locations button').nth(2).click()
        page.wait_for_function("document.querySelector('#citations-targets article.citations-found') !== null")
        assert '只读原因' in page.locator('#citations-targets article').first.inner_text()
        assert page.locator('#citations-targets article').first.locator('input.citations-authorized').count() == 0
        assert '0/12' in page.locator('#citations-selected').inner_text()
