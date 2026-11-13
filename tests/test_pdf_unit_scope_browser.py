from test_professional_browser import _browser, STATIC
from test_pdf_header_rows import header_pdf


def test_explicit_unit_choices_invalidate_old_extraction_and_source_state(tmp_path, monkeypatch):
    browser = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, replies = next(browser)
        (tmp_path/'table.pdf').write_bytes(header_pdf([['Value','Units'],['20','ms']]))
        def invoke(method,args):
            assert method in {'propose_pdf_layout','read_pdf_numeric','acknowledge_pdf_numeric'}
            calls.append(method)
            return getattr(api,method)(*args)
        page.expose_function('unitScopeHost',invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.unitScopeHost(method,args)});void 0")
        page.add_script_tag(path=str(STATIC/'continuous_research.js'))
        page.locator('#cr-nav-check').click()
        page.locator('#cr-pdf-table-toggle').click()
        page.locator('#cr-pdf-table-file').fill('table.pdf')
        page.locator('#cr-pdf-table-region').fill('2,16,98,62')
        def propose():
            page.locator('#cr-pdf-table-propose').click()
            page.wait_for_function("!document.getElementById('cr-unit-review').hidden&&!document.getElementById('cr-unit-review-apply').disabled")
            assert page.locator('#cr-unit-column-1').input_value()==''
        propose()
        assert page.locator('#cr-pdf-table-units').input_value()==''
        page.locator('#cr-unit-column-1').select_option('2')
        page.locator('#cr-unit-review-apply').click()
        assert page.locator('#cr-pdf-table-units').input_value()=='1:2'
        assert page.locator('#cr-unit-review').is_hidden()
        assert not api._professional._observations
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        page.locator('#cr-pdf-table-ack').click()
        page.wait_for_function("document.getElementById('cr-observation').options.length>0")
        propose()
        assert page.locator('#cr-observation option').count()==0
        page.locator('#cr-unit-column-1').select_option('2')
        page.locator('#cr-pdf-table-header-rows').select_option('2')
        assert page.locator('#cr-unit-review').is_hidden()
        assert page.locator('#cr-unit-review-apply').is_disabled()
        propose()
        page.locator('#cr-unit-column-1').select_option('2')
        page.locator('#cr-unit-review-apply').click()
        assert page.locator('#cr-pdf-table-units').input_value()=='1:2'
        page.locator('#cr-pdf-table-file').fill('another.pdf')
        assert page.locator('#cr-unit-review').is_hidden()
        assert page.locator('#cr-pdf-table-units').input_value()==''
        page.locator('#cr-pdf-table-file').fill('table.pdf')
        propose()
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#cr-unit-review').is_hidden()
        assert page.locator('#cr-unit-review-fields select').count()==0
        assert calls.count('read_pdf_numeric')==1
    finally:
        browser.close()
