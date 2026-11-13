from pathlib import Path

from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import digest
from test_professional_browser import _browser, STATIC
from test_rag_workspace import SyntheticEmbedder
from test_research_reading import long_pdf


def exercise(browser, capture=None):
    page, api, app, calls, receipts = browser
    methods = {'get_pdf_ranges', 'plan_pdf_ranges', 'query_pdf_ranges', 'confirm_rag_index', 'confirm_rag_query',
               'plan_rag_rollback', 'confirm_rag_rollback', 'read_structure_anchors', 'plan_structure_edit',
               'confirm_professional_edit', 'rollback_professional_edit', 'read_citation_workspace',
               'read_numeric_source', 'verify_paragraph_numeric'}
    def invoke(method, args):
        assert method in methods
        calls.append(method)
        reply = getattr(api, method)(*args)
        receipts.append((method, reply))
        return reply
    page.expose_function('continuousHost', invoke)
    page.evaluate("window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.continuousHost(method,args)}); void 0")
    page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
    errors = []
    page.on('pageerror', lambda e: errors.append(str(e)))
    original = digest(app.path)
    (app.path.parent / 'paper.pdf').write_bytes(long_pdf())
    embedding = SyntheticEmbedder()
    api._rag._factory = lambda check: WorkspaceRag(app.editor.space, embedding, scope_check=check)
    page.locator('#cr-pdf-file').fill('paper.pdf')
    page.locator('#cr-range-load').click()
    page.wait_for_function("document.getElementById('cr-range-coverage').textContent.includes('missing')")
    assert not embedding.calls
    page.locator('#cr-range-input').fill('1-2,6-7')
    page.locator('#cr-range-plan').click()
    page.wait_for_function("!document.getElementById('cr-range-confirm').disabled")
    assert not embedding.calls
    page.evaluate('window.approved = false')
    page.locator('#cr-range-confirm').click()
    page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消')")
    assert not embedding.calls
    page.evaluate('window.approved = true')
    page.locator('#cr-range-plan').click()
    page.wait_for_function("!document.getElementById('cr-range-confirm').disabled")
    page.locator('#cr-range-confirm').click()
    page.wait_for_function("document.getElementById('cr-status').textContent.includes('整组发布')")
    assert '3, 4, 5' in page.locator('#cr-range-coverage').inner_text()
    page.locator('#cr-query').fill('latency')
    page.locator('#cr-query-run').click()
    page.wait_for_function("document.getElementById('cr-status').textContent.includes('证据已保存')")
    assert 'pdf-page:7' in page.locator('#cr-range-results').inner_text()
    assert page.locator('#assist-receipt').input_value()
    page.locator('#cr-nav-write').click()
    page.locator('#cr-load').click()
    page.wait_for_function("document.getElementById('cr-anchor').options.length > 1")
    page.wait_for_function("!document.getElementById('cr-load').disabled")
    source = api.read_structure_anchors()['result']
    row = next(r for r in source['anchors'] if 'Unchanged paragraph' in r['text'])
    page.locator('#cr-anchor').select_option(row['locator_id'])
    page.locator('#cr-text').fill('An additional paragraph without a placeholder.')
    page.locator('#cr-add-insert').click()
    page.locator('#cr-comment-quote').fill('Unchanged paragraph')
    page.locator('#cr-comment-text').fill('Please verify the applicability of this statement.')
    page.locator('#cr-add-comment').click()
    page.locator('#cr-plan').click()
    page.wait_for_function("!document.getElementById('cr-confirm').disabled")
    assert 'clean_paragraph_insertion' in page.locator('#cr-preview').inner_text()
    assert not list(app.path.parent.glob('*.agent-*.docx'))
    page.evaluate('window.approved = false')
    page.locator('#cr-confirm').click()
    page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消')")
    assert not list(app.path.parent.glob('*.agent-*.docx'))
    page.evaluate('window.approved = true')
    page.locator('#cr-plan').click()
    page.wait_for_function("!document.getElementById('cr-confirm').disabled")
    page.locator('#cr-confirm').click()
    page.wait_for_function("document.getElementById('cr-status').textContent.includes('副本已保存')")
    assert digest(app.path) == original and len(list(app.path.parent.glob('*.agent-*.docx'))) == 1
    if capture:
        capture(page, api, app, calls, receipts)
    page.locator('#cr-undo').click()
    page.wait_for_function("document.getElementById('cr-output').textContent.includes('已撤销')")
    assert not list(app.path.parent.glob('*.agent-*.docx')) and digest(app.path) == original
    page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert page.locator('#cr-text').input_value() == '' and page.locator('#cr-plan').is_disabled()
    assert not errors
    return {'cancel_save_undo': True, 'range_nonadjacent': True, 'kept_docx_selected': True,
            'receipt_handoff': True, 'logout_clear': True, 'source_unchanged': True, 'browser_errors': errors}


def test_browser_collection_and_structure_review(tmp_path, monkeypatch):
    browser = _browser(tmp_path, monkeypatch)
    try:
        exercise(next(browser))
    finally:
        browser.close()
