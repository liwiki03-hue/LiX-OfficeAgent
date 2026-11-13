from docx import Document

from office_agent.writes import digest
from test_professional_browser import _browser, STATIC


def exercise(tmp_path, monkeypatch, real_render=False, capture=None, restart=None):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        doc = Document()
        doc.add_paragraph('Owned report for persistent review.')
        doc.add_paragraph('Unchanged original paragraph.')
        doc.save(app.path)
        before = digest(app.path)
        methods = {'read_structure_anchors', 'read_citation_workspace', 'plan_structure_edit', 'confirm_professional_edit',
                   'start_output_review', 'plan_output_review_archive', 'confirm_output_review_archive',
                   'restore_output_review_archive', 'plan_remove_output_review_archive', 'confirm_remove_output_review_archive'}
        def invoke(method, args):
            assert method in methods
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('archiveHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.archiveHost(method,args)});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-write').click()
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-anchor').options.length === 2")
        page.locator('#cr-text').fill('New context for archive verification.')
        page.locator('#cr-add-insert').click()
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('副本已保存')")
        generated = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_professional_edit')
        page.locator('#cr-review-start').click()
        page.wait_for_function("!document.getElementById('cr-review-save-plan').disabled")
        page.locator('#cr-review-save-plan').click()
        page.wait_for_function("!document.getElementById('cr-review-save-confirm').disabled")
        assert not list(tmp_path.glob('.office-rag-evidence-*'))
        page.evaluate('window.approved=false')
        page.locator('#cr-review-save-confirm').click()
        page.wait_for_function("document.getElementById('cr-review-save-confirm').disabled")
        assert not list(tmp_path.glob('.office-rag-evidence-*'))
        page.locator('#cr-review-save-plan').click()
        page.wait_for_function("!document.getElementById('cr-review-save-confirm').disabled")
        page.evaluate('window.approved=true')
        page.evaluate("document.getElementById('cr-review-save-confirm').click();document.getElementById('cr-review-save-confirm').click()")
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('集合已保存')")
        receipt = page.locator('#cr-review-archive-receipt').input_value()
        assert len(receipt) == 32 and calls.count('confirm_output_review_archive') == 2
        if restart:
            restart(receipt)
        page.locator('#cr-review-restore').click()
        page.wait_for_function("document.getElementById('cr-review-result').textContent.includes('恢复的集合，仅供只读审阅')")
        assert page.locator('#cr-review-undo').is_disabled()
        if capture:
            capture(page)
        page.locator('#cr-review-remove-plan').click()
        page.wait_for_function("!document.getElementById('cr-review-remove-confirm').disabled")
        page.evaluate('window.approved=false')
        page.locator('#cr-review-remove-confirm').click()
        page.wait_for_function("document.getElementById('cr-review-remove-confirm').disabled")
        artifact = tmp_path / ('.office-rag-evidence-' + receipt) / 'output-review-set.json'
        assert artifact.exists()
        page.locator('#cr-review-remove-plan').click()
        page.wait_for_function("!document.getElementById('cr-review-remove-confirm').disabled")
        page.evaluate('window.approved=true')
        page.locator('#cr-review-remove-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('集合记录已删除')")
        assert not artifact.exists() and artifact.with_name('audit-removed.json').exists()
        assert len(list(tmp_path.glob('*.agent-*.docx'))) == 1 and digest(app.path) == before
        page.evaluate("""window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.archiveHost(method,args).then(reply=>{
          if(method!=='restore_output_review_archive')return reply;
          return new Promise(resolve=>{window.releaseArchive=()=>resolve(reply)}).then(value=>{window.archiveDelivered=true;return value});
        })});void 0""")
        page.locator('#cr-review-restore').click()
        page.wait_for_function("typeof window.releaseArchive === 'function'")
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));window.releaseArchive()")
        page.wait_for_function('window.archiveDelivered === true')
        assert page.locator('#cr-review-result').inner_text() == ''
        assert page.locator('#cr-review-archive-result').inner_text() == ''
        assert page.locator('#cr-review-archive-receipt').input_value() == ''
        assert not errors
        return {'cancel_no_write': True, 'duplicate_submit_blocked': True, 'restored_read_only': True,
                'independent_removal': True, 'documents_unchanged_by_archive': True,
                'late_response_ignored_after_logout': True, 'render': generated['render'], 'errors': errors}
    finally:
        browser.close()


def test_review_archive_browser(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
