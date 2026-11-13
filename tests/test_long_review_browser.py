import json
from pathlib import Path
import shutil

from tests.test_professional_browser import STATIC, _browser
from office_agent.writes import Editor, digest


def exercise(tmp_path, monkeypatch, capture=None):
    fixture = _browser(tmp_path, monkeypatch, real_render=True)
    try:
        page, api, app, calls, receipts = next(fixture)
        root = Path(__file__).resolve().parents[1]
        shutil.copyfile(root / 'deliverables/long-review-20261025/MATERIAL.docx', app.path)
        original = digest(app.path)
        chosen = [app.path]
        def select_file():
            app.editor = Editor(chosen[0].parent, require_confirmation=True, render_validation=True)
            app.path = app.editor.space.file(str(chosen[0]))
            return {'ok': True, 'name': app.path.name, 'dir': str(app.editor.space.root)}
        app.select_file = select_file
        def invoke(method, args):
            assert method in {'create_document_task', 'plan_document_task_group', 'confirm_document_task_group',
                             'start_range_review', 'open_range_review', 'range_review_page', 'select_file'}
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('longHost', invoke)
        page.evaluate('''() => {window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.longHost(method, args)});}''')
        page.add_script_tag(path=str(STATIC / 'range_edits.js'))
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        page.locator('#doc-task-request').fill('将全文所有“旧术语”替换为“规范术语”，保留数字、单位、公式和引文')
        page.locator('#doc-task-create').click()
        page.wait_for_function("!document.getElementById('doc-task-route').disabled")
        for primitive, completed in [('word_batch_replace', 30), ('word_table_cell_replace', 31)]:
            page.locator('#doc-task-route').click()
            page.wait_for_function(f"document.getElementById('doc-task-preview').textContent.includes('{primitive}')")
            page.locator('#doc-task-confirm').click()
            page.wait_for_function(f"document.getElementById('doc-task-summary').textContent.includes('已执行并验证 {completed}')")
        task = [reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group'][-1]
        assert task['deterministic_complete'] and len(task['receipts']) == 2
        assert page.locator('#doc-task-deliveries button').count() == 2
        page.locator('#doc-task-deliveries button').nth(0).click()
        assert page.locator('#range-delivery-id').input_value() == task['receipts'][0]
        assert not [method for method in calls if method == 'start_range_review']
        page.locator('#range-open-review').click()
        page.wait_for_function("document.querySelectorAll('#range-before-pages img, #range-after-pages img').length === 2")
        page.locator('#range-review-item').select_option('7')
        page.wait_for_function("document.getElementById('range-before-caption').textContent.includes('段落内严格定位') && document.querySelectorAll('#range-before-pages img, #range-after-pages img').length === 2")
        assert '段落内严格定位' in page.locator('#range-after-caption').inner_text()
        images = []
        for theme in ('dark', 'light', 'highcontrast'):
            for width in (1100, 520):
                page.evaluate('(value) => {document.documentElement.dataset.theme=value;}', theme)
                page.set_viewport_size({'width': width, 'height': 900})
                page.locator('#range-before-pages').scroll_into_view_if_needed()
                assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth + 1')
                if capture:
                    name = f'{theme}-{width}.png'
                    page.screenshot(path=str(capture / name))
                    images.append(name)
        page.locator('#doc-task-deliveries button').nth(1).click()
        page.locator('#range-open-review').click()
        page.wait_for_function("document.getElementById('range-status').textContent.includes('本回执')")
        assert page.locator('#range-before-pages img').count() == 0
        chosen[0] = tmp_path / task['current_name']
        selected = page.evaluate('window.pywebview.api.select_file()')
        assert selected['ok']
        page.evaluate('(name) => {document.getElementById("file-name").textContent=name;}', selected['name'])
        page.wait_for_function("document.getElementById('doc-task-deliveries').children.length === 0")
        assert page.locator('#range-delivery-id').input_value() == task['receipts'][1]
        assert not api._range_review.reviews and not api._document_tasks.tasks
        page.locator('#range-open-review').click()
        page.wait_for_function("document.querySelectorAll('#range-before-pages img, #range-after-pages img').length === 2")
        pages = [reply['result'] for method, reply in receipts if method == 'range_review_page']
        assert len(pages) == 6 and all(item['conversion_reused'] and not item['conversion_processes'] for item in pages)
        assert all(process['empty_confirmed'] for item in pages for process in item['processes'])
        report = {'screenshots': images, 'engine': 'headless Edge, real host, LibreOffice/PDFium; fixture authentication and file chooser',
                  'task_items': 31, 'transactions': 2, 'read_consent_required': True,
                  'wrong_selected_file_refused': True, 'stale_images_cleared': True,
                  'same_root_select_preserved_only_artifacts_and_receipt_draft': True,
                  'old_tasks_and_read_grants_invalidated': True, 'all_six_page_requests_reused_conversion': True,
                  'registered_failure_located_on_page': 2, 'page_worker_jobs_empty': True,
                  'source_unchanged': digest(tmp_path / 'experiment.docx') == original,
                  'word': 'not_verified', 'production_auth': 'not_verified', 'native_file_dialog': 'not_verified', 'native_webview2': 'not_verified'}
        assert report['source_unchanged']
        api._clear_private_state()
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#range-delivery-id').input_value() == ''
        assert page.locator('#range-before-pages img').count() == 0
        from office_agent.render_handoff import _entries
        assert not _entries()
        if capture:
            (capture / 'REPORT.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
        return report
    finally:
        if 'api' in locals():
            api._range_review.invalidate()
        fixture.close()


def test_two_group_review_selection_fallback_and_logout(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
