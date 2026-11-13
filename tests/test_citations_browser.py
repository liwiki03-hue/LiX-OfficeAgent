from docx import Document

from office_agent.literature import normalize
from office_agent.writes import digest
from test_archive_grounded import archive_source
from test_grounded_browser import grounded_browser
from test_literature import crossref
from test_professional_browser import STATIC


def citations_browser(tmp_path, monkeypatch, *, real_render=False):
    for page, api, app, calls, receipts in grounded_browser(tmp_path, monkeypatch, real_render=real_render):
        first = next(iter(api._literature.records))
        second = normalize(crossref('Independent technical experiment', '10.1234/EXPT.02'), 'crossref')
        api._literature.records[second['id']] = second
        api._literature.decisions[second['id']] = {'decision': 'include', 'reason': 'Synthetic test'}
        document = Document()
        document.add_paragraph('The measured latency was 20 ms [4].')
        document.add_paragraph('Independent result [2], repeated [7].')
        document.add_paragraph('[2] First experiment. doi: 10.1234/expt.01')
        document.add_paragraph('[4] Second experiment. doi: 10.1234/expt.02')
        document.add_paragraph('[7] Duplicate first experiment. doi: 10.1234/expt.01')
        document.add_paragraph('Reference placeholder')
        document.save(app.path)
        archived = archive_source(api._grounded, first)
        allowed = {'read_citation_workspace', 'read_citation_capture', 'plan_citation_edit', 'confirm_citation_edit',
                   'rollback_citation_edit', 'capture_grounded_archive', 'read_professional_structure', 'get_grounded_sources',
                   'capture_grounded_evidence'}
        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('citationsHost', invoke)
        page.evaluate('''() => {
          document.getElementById('citations-panel').open = true;
          window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.citationsHost(method, args)});
        }''')
        page.add_script_tag(path=str(STATIC / 'citations.js'))
        yield page, api, app, calls, receipts, first, second['id'], archived


def prepare_citations(page, first, second):
    page.locator('#citations-load').click()
    page.wait_for_function("document.querySelectorAll('#citations-targets article').length === 6")
    for index, record in enumerate([first, second, first]):
        page.locator('#citations-mapping select').nth(index).select_option(record)
    page.locator('#citations-targets input.citations-authorized').evaluate_all('(nodes) => nodes.forEach(node => node.click())')
    for index in range(2, 6):
        page.locator('#citations-targets article').nth(index).locator('select.citations-role').select_option('bibliography')


def test_browser_archive_capture_multi_citation_group_cancel_save_rollback(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in citations_browser(tmp_path, monkeypatch):
        original = digest(app.path)
        page.locator('#grounded-archive').fill(archived['save_id'])
        page.locator('#grounded-archive-capture').click()
        page.wait_for_function("document.getElementById('grounded-evidence').children.length > 0")
        assert page.locator('#grounded-source').input_value() == first
        prepare_citations(page, first, second)
        page.locator('#citations-reviewed').check()
        page.locator('#citations-plan').click()
        page.wait_for_function("!document.getElementById('citations-confirm').disabled")
        assert 'doi:10.1234/expt.02' in page.locator('#citations-preview').inner_text()
        page.evaluate('window.approved = false')
        page.locator('#citations-confirm').click()
        page.wait_for_function("document.getElementById('citations-status').textContent.includes('已取消')")
        assert len(list(tmp_path.glob('*.docx'))) == 1
        page.locator('#citations-plan').click()
        page.wait_for_function("!document.getElementById('citations-confirm').disabled")
        page.evaluate('window.approved = true')
        page.locator('#citations-confirm').click()
        page.wait_for_function("document.getElementById('citations-output').textContent.includes('已核验副本')")
        assert calls.count('confirm_citation_edit') == 2 and digest(app.path) == original
        page.set_viewport_size({'width': 520, 'height': 850})
        for theme in ('dark', 'light', 'highcontrast'):
            page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
            assert page.locator('#citations-panel').evaluate('(node) => node.scrollWidth <= node.clientWidth')
        page.locator('#citations-rollback').click()
        page.wait_for_function("document.getElementById('citations-output').textContent.includes('已撤销')")
        assert digest(app.path) == original and len(list(tmp_path.glob('*.docx'))) == 1


def test_browser_late_read_after_logout_cannot_restore_citation_metadata(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, first, second, archived in citations_browser(tmp_path, monkeypatch):
        page.evaluate('''() => {
          window.pendingReads = 0;
          window.pywebview.api = {read_citation_workspace: () => {window.pendingReads++; return new Promise(resolve => { window.finishRead = resolve; });}};
        }''')
        page.locator('#citations-load').click()
        page.evaluate("document.getElementById('citations-load').click()")
        assert page.evaluate('window.pendingReads') == 1
        page.evaluate('''() => {
          window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'));
          window.finishRead({ok: true, result: {items: [{text: 'late'}]}});
        }''')
        assert page.locator('#citations-targets').inner_text() == ''
        assert page.locator('#citations-confirm').is_disabled()
