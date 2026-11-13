from docx import Document

from office_agent.literature import normalize
from office_agent.writes import digest
from test_fidelity import pdf_bytes
from test_literature import crossref
from test_professional_browser import _browser, STATIC


def grounded_browser(tmp_path, monkeypatch, *, real_render=False):
    for page, api, app, calls, receipts in _browser(tmp_path, monkeypatch, real_render=real_render):
        doc = Document()
        paragraph = doc.add_paragraph()
        paragraph.add_run('The experiment ').bold = True
        paragraph.add_run('measured latency.').italic = True
        doc.add_paragraph('Reference placeholder')
        doc.add_paragraph('Untouched control paragraph.')
        doc.save(app.path)
        (tmp_path / 'paper.pdf').write_bytes(pdf_bytes(text='The measured latency was 20 ms under laboratory conditions.', width=600))
        api._literature._bind()
        record = normalize(crossref(), 'crossref')
        api._literature.records[record['id']] = record
        api._literature.decisions[record['id']] = {'decision': 'include', 'reason': 'Synthetic fixture, not human relevance labels'}
        allowed = {'get_grounded_sources', 'capture_grounded_evidence', 'plan_grounded_edit',
                   'confirm_grounded_edit', 'rollback_grounded_edit', 'read_professional_structure'}
        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('groundedHost', invoke)
        page.evaluate('''() => {
          document.getElementById('professional-panel').open = false;
          document.getElementById('grounded-panel').open = true;
          window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.groundedHost(method, args)});
        }''')
        page.add_script_tag(path=str(STATIC / 'grounded.js'))
        yield page, api, app, calls, receipts


def prepare_grounded(page):
    page.locator('#grounded-load').click()
    page.wait_for_function("document.getElementById('grounded-source').options.length === 1")
    page.locator('#grounded-file').fill('paper.pdf')
    page.locator('#grounded-capture').click()
    page.locator('#grounded-evidence input').first.check()
    page.locator('#grounded-claim').fill('The measured latency was 20 ms.')
    page.locator('#grounded-relation').select_option('supports')
    page.locator('#grounded-applicability').fill('Laboratory conditions only.')
    page.locator('#grounded-limitation').fill('Synthetic source; no external validity established.')
    page.locator('#grounded-identity').check()
    page.locator('#grounded-reviewed').check()
    page.locator('#grounded-plan').click()
    page.wait_for_function("!document.getElementById('grounded-confirm').disabled")


def test_browser_independent_confirm_cancel_save_rollback_and_themes(tmp_path, monkeypatch):
    for page, api, app, calls, receipts in grounded_browser(tmp_path, monkeypatch):
        original = digest(app.path)
        assert calls == []
        prepare_grounded(page)
        assert calls.count('plan_grounded_edit') == 1
        assert 'user_reviewed_not_machine_proven' in page.locator('#grounded-preview').inner_text()
        assert '修改前：' in page.locator('#grounded-preview').inner_text()
        assert '限制与缺口：' in page.locator('#grounded-preview').inner_text()
        assert '"record_type"' not in page.locator('#grounded-preview').inner_text()
        page.evaluate('window.approved = false')
        page.locator('#grounded-confirm').click()
        page.wait_for_function("document.getElementById('grounded-status').textContent.includes('已取消')")
        assert len(list(tmp_path.glob('*.docx'))) == 1
        page.locator('#grounded-plan').click()
        page.wait_for_function("!document.getElementById('grounded-confirm').disabled")
        page.evaluate('window.approved = true')
        page.locator('#grounded-confirm').click()
        page.wait_for_function("document.getElementById('grounded-output').textContent.includes('已保存并核验')")
        assert digest(app.path) == original
        assert calls.count('confirm_grounded_edit') == 2
        page.set_viewport_size({'width': 520, 'height': 850})
        for theme in ('dark', 'light', 'highcontrast'):
            page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
            assert page.locator('#grounded-panel').evaluate('(node) => node.scrollWidth <= node.clientWidth')
        page.locator('#grounded-rollback').click()
        page.wait_for_function("document.getElementById('grounded-output').textContent.includes('已撤销')")
        assert digest(app.path) == original and len(list(tmp_path.glob('*.docx'))) == 1


def test_browser_logout_drops_delayed_reply_and_prevents_duplicate(tmp_path, monkeypatch):
    for page, api, app, calls, receipts in grounded_browser(tmp_path, monkeypatch):
        page.evaluate('''() => {
          window.sourceCalls = 0;
          window.pywebview.api.get_grounded_sources = undefined;
          window.pywebview.api = {get_grounded_sources: () => {
            window.sourceCalls++; return new Promise(resolve => { window.finishSources = resolve; });
          }};
        }''')
        page.locator('#grounded-load').click()
        page.evaluate("document.getElementById('grounded-load').click()")
        assert page.evaluate('window.sourceCalls') == 1
        page.evaluate('''() => {
          window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'));
          window.finishSources({ok: true, result: {records: [{id: 'late', title: 'late', doi: 'late'}]}});
        }''')
        assert page.locator('#grounded-source option').count() == 0
        assert page.locator('#grounded-confirm').is_disabled()
        assert page.locator('#grounded-claim').input_value() == ''
