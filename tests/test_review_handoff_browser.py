import json
from pathlib import Path
import shutil

from tests.test_professional_browser import STATIC, _browser
from office_agent.writes import digest


def exercise(tmp_path, monkeypatch, capture=None):
    fixture = _browser(tmp_path, monkeypatch, real_render=True)
    try:
        page, api, app, calls, receipts = next(fixture)
        material = Path(__file__).resolve().parents[1] / 'deliverables/review-reuse-20261025/MATERIAL-table.docx'
        shutil.copyfile(material, app.path)
        source_hash = digest(app.path)
        def invoke(method, args):
            assert method in {'read_edit_ranges', 'plan_range_edit', 'confirm_professional_edit',
                             'start_range_review', 'open_range_review', 'range_review_page', 'rollback_professional_edit'}
            calls.append(method)
            response = getattr(api, method)(*args)
            receipts.append((method, response))
            return response
        page.expose_function('handoffHost', invoke)
        page.evaluate('''() => {window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.handoffHost(method, args)});}''')
        page.add_script_tag(path=str(STATIC / 'range_edits.js'))
        page.locator('#range-load').click()
        page.wait_for_function("document.getElementById('range-target').options.length === 7")
        row = next(reply['result']['items'] for method, reply in receipts if method == 'read_edit_ranges')
        target = next(item for item in row if item['text'] == '旧术语')
        page.locator('#range-target').select_option(target['locator_id'])
        page.locator('#range-original').evaluate('(node) => {node.focus();node.setSelectionRange(0,3);}')
        page.locator('#range-text').fill('规范术语')
        page.locator('#range-reason').fill('仅调整选定单元格术语')
        page.locator('#range-add').click()
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已保存副本')")
        page.locator('#range-open-review').click()
        page.wait_for_function("document.querySelectorAll('#range-before-pages img, #range-after-pages img').length === 2")
        assert '复用本进程' in page.locator('#range-before-pages').inner_text()
        assert '复用本进程' in page.locator('#range-after-pages').inner_text()
        pages = [reply['result'] for method, reply in receipts if method == 'range_review_page']
        assert len(pages) == 2 and all(item['conversion_reused'] and not item['conversion_processes'] for item in pages)
        assert all(item['empty_confirmed'] for response in pages for item in response['processes'])
        assert digest(app.path) == source_hash
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
        report = {'screenshots': images, 'engine': 'headless Edge, real host, LibreOffice/PDFium, fixture authentication',
                  'source_unchanged': True, 'table_anchor_unique': True, 'review_conversion_reused': True,
                  'review_conversion_processes': 0, 'page_workers_empty': True,
                  'native_webview2': 'not_verified', 'word': 'not_verified', 'production_auth': 'not_verified'}
        page.locator('#range-undo').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已撤销')")
        undo = next(reply for method, reply in reversed(receipts) if method == 'rollback_professional_edit')
        assert undo['ok'] and undo['result']['status'] == 'rolled_back', undo
        report['independent_undo'] = True
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#range-before-pages img').count() == 0
        if capture:
            (capture / 'REPORT.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
        return report
    finally:
        if 'api' in locals():
            api._range_review.invalidate()
        fixture.close()


def test_real_table_review_handoff_themes_and_undo(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
