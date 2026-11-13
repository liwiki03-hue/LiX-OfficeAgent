import shutil
from pathlib import Path

from tests.test_professional_browser import STATIC, _browser
from office_agent.writes import digest


def test_history_states_visible_when_document_read_is_declined(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        page.add_script_tag(path=str(STATIC / 'range_edits.js'))
        page.evaluate('''() => {
          window.approved = false; window.bodyReads = 0;
          window.pywebview.api = {
            restore_range_review_archive: async () => ({ok: true, result: {review_id: 'historical', revision: 1,
              items: [{index: 0, state: 'viewed', reason: '上次已查看'}, {index: 1, state: 'needs_review', reason: '需要核对页码'}]}}),
            open_range_review: () => {window.bodyReads++; throw new Error('unexpected document read');}
          };
        }''')
        page.locator('#range-restore-review').click()
        page.wait_for_function("document.getElementById('range-review-item').options.length === 2")
        assert '已查看' in page.locator('#range-review-item').inner_text()
        assert page.locator('#range-review-reason').input_value() == '上次已查看'
        page.locator('#range-next-review').click()
        page.wait_for_function("document.getElementById('range-review-reason').value === '需要核对页码'")
        assert page.evaluate('window.bodyReads') == 0
        assert page.locator('#range-mark-review').is_disabled()
        assert page.locator('#range-save-review').is_disabled()
    finally:
        fixture.close()


def test_real_render_page_review_archive_restore_and_logout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch, real_render=True)
    try:
        page, api, app, calls, receipts = next(fixture)
        source = Path(__file__).resolve().parents[1] / 'deliverables/docx-review-20261021/development.docx'
        shutil.copyfile(source, app.path)
        version = digest(app.path)
        def invoke(method, args):
            assert method in {'read_edit_ranges', 'plan_range_edit', 'confirm_professional_edit', 'start_range_review',
                'open_range_review', 'range_review_page', 'mark_range_review', 'plan_range_review_archive',
                'confirm_range_review_archive', 'restore_range_review_archive'}
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('deliveryHost', invoke)
        page.evaluate("() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.deliveryHost(method, args)}); }")
        page.add_script_tag(path=str(STATIC / 'range_edits.js'))
        assert calls == []
        page.locator('#range-load').click()
        page.wait_for_function("document.getElementById('range-target').options.length > 4")
        page.locator('#range-target').select_option(index=1)
        page.locator('#range-original').evaluate('(node) => {node.focus(); node.setSelectionRange(0, 4);}')
        page.locator('#range-text').fill('明确结论')
        page.locator('#range-reason').fill('仅调整本处措辞')
        page.locator('#range-add').click()
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        assert '可能影响分页' in page.locator('#range-preview').inner_text()
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已保存副本')")
        page.locator('#range-open-review').click()
        page.wait_for_function("document.querySelectorAll('#range-before-pages img, #range-after-pages img').length === 2")
        assert page.locator('#range-review-state').input_value() == 'unreviewed'
        page.locator('#range-review-state').select_option('needs_review')
        page.locator('#range-review-reason').fill('在Word中核对PAGEREF')
        page.locator('#range-mark-review').click()
        page.wait_for_function("document.getElementById('range-review-note').textContent.includes('尚未存档')")
        page.locator('#range-save-review').click()
        page.wait_for_function("document.getElementById('range-archive-id').value.length === 32")
        page.locator('#range-restore-review').click()
        page.wait_for_function("document.getElementById('range-review-state').value === 'needs_review' && !document.getElementById('range-open-review').disabled")
        assert page.locator('#range-undo').is_disabled()
        assert digest(app.path) == version
        assert len([call for call in calls if call == 'confirm_professional_edit']) == 1
        assert all(reply['ok'] for _, reply in receipts), receipts
        for method, reply in receipts:
            if method == 'range_review_page':
                assert all(p['empty_confirmed'] for p in reply['result']['processes'] + reply['result']['conversion_processes'])
        page.set_viewport_size({'width': 520, 'height': 900})
        assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth + 1')
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#range-before-pages img').count() == 0
        assert page.locator('#range-archive-id').input_value() == ''
    finally:
        if 'api' in locals():
            api._range_review.invalidate()
        fixture.close()
