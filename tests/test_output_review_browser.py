from docx import Document

from office_agent.writes import digest
from test_professional_browser import _browser, STATIC


def exercise(tmp_path, monkeypatch, real_render=False, capture=None):
    browser = _browser(tmp_path, monkeypatch, real_render=real_render)
    try:
        page, api, app, calls, replies = next(browser)
        assert app.path.parent == tmp_path
        doc = Document()
        doc.add_paragraph('Owned experimental report.')
        doc.add_paragraph('Unchanged context.')
        doc.save(app.path)
        before = digest(app.path)
        methods = {'read_structure_anchors', 'read_citation_workspace', 'plan_structure_edit', 'confirm_professional_edit',
                   'start_output_review', 'refresh_output_review', 'acknowledge_output_review', 'cancel_output_review', 'rollback_reviewed_output'}
        def invoke(method, args):
            assert method in methods
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('reviewHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.reviewHost(method,args)});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-write').click()
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-anchor').options.length === 2")
        page.locator('#cr-text').fill('Supplementary context for review.')
        page.locator('#cr-add-insert').click()
        page.locator('#cr-plan').click()
        page.wait_for_function("!document.getElementById('cr-confirm').disabled")
        page.locator('#cr-confirm').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('副本已保存')")
        saved = next(reply['result'] for method, reply in reversed(replies) if method == 'confirm_professional_edit')
        receipt = saved['save_id']
        assert page.locator('#cr-dependency-target').input_value() == receipt
        page.locator('#cr-dependency-receipts').fill('f' * 32)
        page.evaluate("document.getElementById('cr-review-start').click();document.getElementById('cr-review-start').click()")
        page.wait_for_function("document.getElementById('cr-review-result').textContent.includes('存在阻塞')")
        assert calls.count('start_output_review') == 1 and len(api._output_dependencies.known) == 1
        page.locator('#cr-dependency-receipts').fill('')
        page.locator('#cr-review-start').click()
        page.wait_for_function("!document.getElementById('cr-review-ack').disabled")
        page.locator('#cr-review-cancel').click()
        page.wait_for_function("document.getElementById('cr-review-result').textContent.includes('会话已结束')")
        assert page.locator('#cr-review-refresh').is_disabled() and len(list(tmp_path.glob('*.agent-*.docx'))) == 1
        page.locator('#cr-review-start').click()
        page.wait_for_function("!document.getElementById('cr-review-ack').disabled")
        assert page.locator('#cr-review-undo').is_disabled()
        page.locator('#cr-review-ack').click()
        page.wait_for_function("!document.getElementById('cr-review-undo').disabled")
        if capture:
            capture(page)
        page.evaluate('window.approved=false')
        page.locator('#cr-review-undo').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('已取消撤销')")
        assert page.locator('#cr-review-undo').is_disabled() and len(list(tmp_path.glob('*.agent-*.docx'))) == 1
        page.locator('#cr-review-ack').click()
        page.wait_for_function("!document.getElementById('cr-review-undo').disabled")
        page.evaluate('window.approved=true')
        page.locator('#cr-review-undo').click()
        page.wait_for_function("document.getElementById('cr-status').textContent.includes('独立撤销单个副本')")
        assert not list(tmp_path.glob('*.agent-*.docx')) and digest(app.path) == before
        page.locator('#cr-review-refresh').click()
        page.wait_for_function("document.getElementById('cr-review-result').textContent.includes('已撤销')")
        assert page.locator('#cr-review-ack').is_disabled()
        page.evaluate("""window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.reviewHost(method,args).then(reply=>{
          if(method!=='refresh_output_review')return reply;
          return new Promise(resolve=>{window.releaseReview=()=>resolve(reply)}).then(value=>{window.reviewDelivered=true;return value});
        })});void 0""")
        page.locator('#cr-review-refresh').click()
        page.wait_for_function("typeof window.releaseReview === 'function'")
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));window.releaseReview()")
        page.wait_for_function('window.reviewDelivered === true')
        assert page.locator('#cr-review-result').inner_text() == '' and page.locator('#cr-review-undo').is_disabled()
        assert not errors
        return {'browser': 'actual controlled Edge', 'duplicate_click_guard': True, 'bad_candidate_isolated': True,
                'separate_ack_and_confirmation': True, 'cancel_then_ack_then_single_undo': True,
                'source_unchanged': True, 'late_response_ignored_after_logout': True, 'render': saved['render'], 'errors': errors}
    finally:
        browser.close()


def test_output_review_browser(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
