from pathlib import Path

from lxml import etree as ET

from office_agent.professional_edits import W
from office_agent.writes import digest
from test_professional_edits import parts
from test_professional_browser import _browser, STATIC
from test_research_assist import Model


def exercise_word_revision(browser, capture=None):
    page, api, app, calls, receipts = browser
    model = Model()
    api._research_assist.model_factory = lambda: model
    methods = {'read_research_assist','start_research_assist','research_assist_status','plan_assisted_edit','confirm_assisted_edit','rollback_assisted_edit'}
    def invoke(method, args):
        assert method in methods
        calls.append(method)
        reply = getattr(api, method)(*args)
        receipts.append((method, reply))
        return reply
    page.expose_function('revisionHost', invoke)
    page.evaluate("window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.revisionHost(method,args)}); document.getElementById('assist-panel').open = true")
    page.add_script_tag(path=str(STATIC / 'research_assist.js'))
    source = digest(app.path)
    page.locator('#assist-load').click()
    page.locator('#assist-targets input').first.check()
    page.locator('#assist-question').fill('Clarify the original observation.')
    page.locator('#assist-start').click()
    page.locator('#assist-suggestions fieldset textarea').wait_for()
    page.locator('#assist-suggestions input[value=accept]').first.check()
    page.locator('#assist-suggestions input[value=review]').first.check()
    assert page.locator('#assist-granular').is_disabled()
    page.locator('#assist-tracked').check()
    page.locator('#assist-granular').check()
    page.locator('#assist-plan').click()
    page.wait_for_function("!document.getElementById('assist-confirm').disabled")
    assert receipts[-1][1]['result']['items'][0]['native_revisions']['scope'] == 'word_diff_with_character_style_projection'
    page.evaluate('window.approved = false')
    page.locator('#assist-confirm').click()
    page.wait_for_function("document.getElementById('assist-status').textContent.includes('已取消')")
    page.evaluate('window.approved = true')
    page.locator('#assist-plan').click()
    page.wait_for_function("!document.getElementById('assist-confirm').disabled")
    page.locator('#assist-confirm').click()
    page.wait_for_function("document.getElementById('assist-output').textContent.includes('已核验副本')")
    result = next(reply['result'] for method, reply in reversed(receipts) if method == 'confirm_assisted_edit' and reply.get('result',{}).get('status') == 'saved')
    root = ET.fromstring(parts(Path(result['output']))['word/document.xml'])
    assert len(list(root.iter(W+'ins'))) == 1 and not list(root.iter(W+'del'))
    xml = ET.tostring(next(root.iter(W+'ins')),encoding='unicode')
    assert 'The observation is reported.' in xml
    if capture:
        capture(page, result, xml)
    page.locator('#assist-undo').click()
    page.wait_for_function("document.getElementById('assist-output').textContent.includes('已撤销')")
    assert digest(app.path) == source and not Path(result['output']).exists()
    page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert not page.locator('#assist-granular').is_checked()


def test_word_revision_real_browser_local_change_cancel_save_undo(tmp_path, monkeypatch):
    browser = _browser(tmp_path, monkeypatch)
    try:
        exercise_word_revision(next(browser))
    finally:
        browser.close()
