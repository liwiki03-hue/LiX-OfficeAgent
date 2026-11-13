from test_grounded_browser import grounded_browser
from test_professional_browser import STATIC
from test_research_assist import Model


def test_transfer_preserves_other_drafts_and_requires_new_source_review(tmp_path, monkeypatch):
    for page, api, app, calls, receipts in grounded_browser(tmp_path, monkeypatch):
        allowed = {'get_grounded_sources', 'read_professional_structure', 'capture_grounded_evidence', 'read_research_assist', 'start_research_assist',
                   'research_assist_status', 'read_citation_workspace', 'read_citation_capture', 'plan_citation_edit', 'confirm_citation_edit', 'rollback_citation_edit'}

        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            result = getattr(api, method)(*args)
            receipts.append((method, result))
            return result

        class EvidenceModel(Model):
            def suggest(self, payload, cancelled):
                result = super().suggest(payload, cancelled)
                result['suggestions'][0].update(new_text='The measured latency was 20 ms.', evidence_ids=[payload['evidence'][0]['id']])
                return result

        api._research_assist.model_factory = lambda: EvidenceModel()
        page.expose_function('transferHost', invoke)
        page.evaluate('''() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.transferHost(method, args)}); document.getElementById('assist-panel').open = true; document.getElementById('citations-panel').open = true; }''')
        page.add_script_tag(path=str(STATIC / 'citations.js'))
        page.add_script_tag(path=str(STATIC / 'research_assist.js'))
        page.locator('#grounded-load').click()
        page.locator('#grounded-file').fill('paper.pdf')
        page.locator('#grounded-capture').click()
        page.locator('#grounded-evidence input').first.wait_for()
        page.locator('#citations-load').click()
        page.locator('#citations-targets article').nth(2).locator('.citations-authorized').check()
        page.locator('#assist-load').click()
        page.locator('#assist-targets input').first.check()
        page.locator('#assist-captures input').first.check()
        page.locator('#assist-mode').select_option('supplement')
        page.locator('#assist-question').fill('Use the selected source to qualify the laboratory finding.')
        page.locator('#assist-start').click()
        page.locator('#assist-suggestions fieldset textarea').wait_for()
        page.locator('#assist-suggestions input[value=accept]').check()
        page.locator('#assist-suggestions input[value=review]').check()
        page.get_by_role('button', name='转入多引文草稿并重新核对…').click()
        page.wait_for_function('document.getElementById("citations-status").textContent.includes("研究建议已转为")')
        claim = page.locator('#citations-targets article').nth(0)
        assert claim.locator('textarea').first.input_value() == 'The measured latency was 20 ms.'
        assert page.locator('#citations-targets article').nth(2).locator('.citations-authorized').is_checked()
        attribution = claim.locator('fieldset')
        assert not attribution.locator('input[type=checkbox]').nth(-2).is_checked()
        assert not attribution.locator('input[type=checkbox]').last.is_checked()
        assert attribution.locator('select').nth(1).input_value() == 'insufficient'
        reference = page.locator('#citations-targets article').nth(1)
        reference.locator('.citations-authorized').check()
        reference.locator('.citations-role').select_option('bibliography')
        attribution.locator('select').nth(1).select_option('supports')
        attribution.locator('textarea').nth(0).fill('Laboratory conditions only.')
        attribution.locator('textarea').nth(1).fill('Self-authored fixture, no independent professional quality claim.')
        attribution.locator('input[type=checkbox]').nth(-2).check()
        attribution.locator('input[type=checkbox]').last.check()
        page.locator('#citations-reviewed').check()
        page.locator('#citations-plan').click()
        page.wait_for_function('!document.getElementById("citations-confirm").disabled')
        page.locator('#citations-confirm').click()
        page.wait_for_function('document.getElementById("citations-output").textContent.includes("已核验副本")')
        assert next(reply for method, reply in reversed(receipts) if method == 'confirm_citation_edit')['result']['status'] == 'saved'
        page.locator('#citations-rollback').click()
        page.wait_for_function('document.getElementById("citations-output").textContent.includes("已撤销")')
