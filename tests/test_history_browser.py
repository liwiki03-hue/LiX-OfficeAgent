from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import digest
from test_professional_browser import _browser, STATIC
from test_rag_workspace import SyntheticEmbedder
from test_research_reading import long_pdf


def exercise(tmp_path, monkeypatch, capture=None, restart=None):
    browser = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, replies = next(browser)
        original = digest(app.path)
        (tmp_path / 'paper.pdf').write_bytes(long_pdf())
        embedding = SyntheticEmbedder()
        api._rag._factory = lambda check: WorkspaceRag(app.editor.space, embedding, scope_check=check)
        methods = {'get_pdf_ranges', 'plan_pdf_ranges', 'plan_pdf_attempt_history', 'confirm_pdf_attempt_history',
                   'restore_pdf_attempt_history', 'plan_remove_pdf_attempt_history', 'confirm_remove_pdf_attempt_history', 'preview_output_dependencies'}
        def invoke(method, args):
            assert method in methods
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('historyHost', invoke)
        page.evaluate("window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.historyHost(method,args)}); void 0")
        errors = []
        page.on('pageerror', lambda e: errors.append(str(e)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-pdf-file').fill('paper.pdf')
        page.locator('#cr-range-input').fill('99-100')
        page.locator('#cr-range-plan').click()
        page.wait_for_function("document.getElementById('cr-range-coverage').textContent.includes('unprocessed_pages')")
        page.evaluate("document.getElementById('cr-history-plan').click();document.getElementById('cr-history-plan').click()")
        page.wait_for_function("!document.getElementById('cr-history-confirm').disabled")
        assert calls.count('plan_pdf_attempt_history') == 1
        page.evaluate('window.approved = false')
        page.locator('#cr-history-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消')")
        assert not list(tmp_path.glob('.office-rag-evidence-*'))
        page.locator('#cr-history-plan').click()
        page.wait_for_function("!document.getElementById('cr-history-confirm').disabled")
        page.evaluate('window.approved = true')
        page.locator('#cr-history-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('记录已保存')")
        receipt = page.locator('#cr-history-receipt').input_value()
        artifact = tmp_path / ('.office-rag-evidence-' + receipt) / 'pdf-range-attempt.json'
        assert artifact.exists()
        if restart:
            restart(receipt)
        page.locator('#cr-history-restore').click()
        page.wait_for_function("document.getElementById('cr-history-result').textContent.includes('historical_only')")
        assert 'matches' in page.locator('#cr-history-result').inner_text()
        if capture:
            capture(page)
        page.evaluate("""window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.historyHost(method,args).then(reply => {
          if (method !== 'restore_pdf_attempt_history') return reply;
          return new Promise(resolve => {window.releaseHistory = () => resolve(reply)}).then(value => {window.historyDelivered = true;return value});
        })}); void 0""")
        page.locator('#cr-history-restore').click()
        page.wait_for_function("typeof window.releaseHistory === 'function'")
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));window.releaseHistory()")
        page.wait_for_function('window.historyDelivered === true')
        assert page.locator('#cr-history-result').inner_text() == ''
        page.evaluate("window.authenticated=true;window.dispatchEvent(new Event('officeauth:authenticated'))")
        page.locator('#cr-history-receipt').fill(receipt)
        page.locator('#cr-nav-write').click()
        page.locator('#cr-dependency-target').fill('a' * 32)
        page.locator('#cr-dependency-check').click()
        page.wait_for_function("document.getElementById('cr-dependency-result').textContent.includes('complete_dependency_graph')")
        assert 'false' in page.locator('#cr-dependency-result').inner_text()
        page.locator('#cr-nav-ranges').click()
        page.locator('#cr-history-remove-plan').click()
        page.wait_for_function("!document.getElementById('cr-history-remove-confirm').disabled")
        page.locator('#cr-history-remove-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('记录已撤销')")
        assert not artifact.exists() and artifact.with_name('audit-removed.json').exists()
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert not page.locator('#cr-history-receipt').input_value() and not page.locator('#cr-history-result').inner_text()
        assert not page.locator('#cr-dependency-target').input_value()
        assert not embedding.calls and digest(app.path) == original and not errors
        return {'browser': 'actual Edge headless', 'failure_preview': True, 'double_click_guard': True, 'save_cancel_confirm_restore_remove': True,
                'diagnostic_only': True, 'source_unchanged': True, 'embedding_calls': 0, 'late_response_ignored_after_logout': True, 'errors': errors}
    finally:
        browser.close()


def test_history_browser_lifecycle(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
