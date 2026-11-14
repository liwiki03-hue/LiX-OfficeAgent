import json

from tests.test_floating_browser import floating_browser
from tests.test_professional_browser import STATIC


def test_diagnostic_preview_export_uses_theme_and_only_sanitized_fields(floating_browser, monkeypatch, tmp_path):
    from office_agent.diagnostics import Diagnostics
    page, api, source, calls, errors = floating_browser
    logger = Diagnostics(tmp_path)
    logger.record(module='model', action='request', result='refused', code='private_scope', document='synthetic-private-body', api_key='synthetic-private-key')
    assert logger.close()
    monkeypatch.setattr('office_agent.diagnostics.get', lambda: logger)
    target = tmp_path / 'diagnostic-export.json'
    api._app.window.create_file_dialog = lambda *args, **kwargs: [str(target)]
    page.evaluate("async()=>{OfficeSettings.show();await OfficeSettings.selectTab('tab-machine');}")
    show = page.get_by_role('button', name='查看诊断', exact=True)
    assert 'btn' in show.get_attribute('class') and show.bounding_box()['height'] >= 30
    show.click()
    dialog = page.get_by_role('dialog', name='诊断摘要', exact=True)
    dialog.get_by_role('button', name='导出摘要…', exact=True).click()
    page.wait_for_function("!document.querySelector('dialog[aria-label=诊断摘要]')")
    data = json.loads(target.read_text('utf8'))
    assert data['events'] and 'path' not in data
    assert 'synthetic-private' not in target.read_text('utf8')
    assert not data['includes_documents'] and not data['includes_configuration'] and not errors


def test_compact_menu_bounds_and_idle_poll_budget(floating_browser, monkeypatch):
    page, api, source, calls, errors = floating_browser
    rows = []
    for width, height in [(360, 520), (420, 800), (520, 820)]:
        page.set_viewport_size({'width': width, 'height': height})
        page.locator('#floating-root header > details > summary').click()
        menu = page.locator('.floating-menu')
        box = menu.bounding_box(); composer = page.locator('#floating-input').bounding_box()
        assert box and box['x'] >= 0 and box['x'] + box['width'] <= width and box['y'] + box['height'] <= min(height, composer['y'])
        assert page.locator('#floating-send').bounding_box()['y'] < height
        rows.append({'width': width, 'height': height, 'menu': box})
        page.keyboard.press('Escape')
    samples = []
    original = api.floating_snapshot
    def measured(*args):
        samples.append(1)
        return original(*args)
    monkeypatch.setattr(api, 'floating_snapshot', measured)
    page.wait_for_timeout(3200)
    sampled = len(samples)
    assert 1 <= sampled <= 4 and not errors
    out = STATIC.parents[2] / 'deliverables/issue-repair-completion-20261114'
    out.mkdir(parents=True, exist_ok=True)
    (out / 'Q2_MENU_POLL.json').write_text(json.dumps({'bounds': rows, 'idle_snapshots_3_2_seconds': sampled, 'native': False}, ensure_ascii=False, indent=2), 'utf8')


def test_cancel_file_preserves_state_and_async_confirmation_focus(floating_browser):
    page, api, source, calls, errors = floating_browser
    out = STATIC.parents[2] / 'deliverables/issue-repair-completion-20261114'
    out.mkdir(parents=True, exist_ok=True)
    page.set_viewport_size({'width': 420, 'height': 800})
    page.locator('#floating-input').fill('未发送要求保持。')
    before = api._floating.collect()
    api._app.window.create_file_dialog = lambda *args, **kwargs: None
    prior = calls.count('floating_present')
    page.locator('#floating-choose').click()
    page.wait_for_timeout(300)
    assert calls.count('floating_present') == prior
    assert api._floating.document == before['document']
    assert page.locator('#floating-input').input_value() == '未发送要求保持。'
    page.locator('#floating-choose').focus()
    page.evaluate("() => {window.__decision=null; window.OfficeConfirm.ask('取消不修改文件。').then(value=>window.__decision=value);}")
    dialog = page.get_by_role('dialog', name='LiX', exact=True)
    dialog.wait_for()
    assert dialog.get_by_role('button', name='取消', exact=True).evaluate('(node)=>node===document.activeElement')
    page.screenshot(path=str(out / 'Q2-confirm-420.png'))
    page.keyboard.press('Escape')
    page.wait_for_function('window.__decision===false')
    assert page.locator('#floating-choose').evaluate('(node)=>node===document.activeElement')
    page.evaluate("()=>{window.OfficeSettings.show();window.OfficeSettingsEditor={requestClose:()=>window.OfficeConfirm.ask('关闭会放弃草稿。')};window.OfficeModelSelector.requestClose=()=>true;window.OfficeSkills.requestClose=()=>true;}")
    page.evaluate('()=>{window.OfficeSettings.close()}')
    dialog.wait_for()
    assert page.evaluate('window.OfficeSettings.isOpen()')
    dialog.get_by_role('button', name='取消', exact=True).click()
    page.wait_for_timeout(50)
    assert page.evaluate('window.OfficeSettings.isOpen()')
    page.evaluate('()=>{window.OfficeSettings.close()}')
    dialog.get_by_role('button', name='确认', exact=True).click()
    page.wait_for_function('!window.OfficeSettings.isOpen()')
    assert page.locator('#floating-input').input_value() == '未发送要求保持。' and not errors
    (out / 'Q2_BROWSER.json').write_text(json.dumps({'file_cancel_no_present': True, 'input_preserved': True,
        'async_guard_cancel_keeps_open': True, 'focus_returned': True, 'width': 420,
        'native_drag': 'not_verified', 'errors': errors}, ensure_ascii=False, indent=2), encoding='utf-8')


def test_draft_models_fill_and_save_once_without_losing_manual_names(floating_browser, monkeypatch):
    page, api, source, calls, errors = floating_browser
    probes = []
    def probe(prepared, **kwargs):
        probes.append(prepared[1])
        return {'status': 'reachable', 'models': ['candidate-model'], 'model_listed': False}
    monkeypatch.setattr('office_agent.settings_probe.probe_connection', probe)
    page.evaluate('window.OfficeSettings.show()')
    page.get_by_role('button', name='编辑', exact=True).first.click()
    page.wait_for_function("document.getElementById('connection-model-candidates').children.length===1")
    page.locator('#connection-model-name').fill('candidate-model')
    page.locator('#connection-model-name').dispatch_event('change')
    assert 'same-model' in page.locator('#connection-models').input_value()
    page.locator('#connection-save').click()
    page.wait_for_timeout(200)
    out = STATIC.parents[2] / 'deliverables/issue-repair-completion-20261114'
    out.mkdir(parents=True, exist_ok=True)
    (out / 'Q3_SAVE_DIAGNOSTIC.json').write_text(json.dumps({'status':page.locator('#connection-status').text_content(),
        'review_count':page.locator('#connection-review-confirm').count(),'calls':calls,'errors':errors}, ensure_ascii=False, indent=2), encoding='utf-8')
    page.locator('#connection-review-confirm').click()
    page.wait_for_function("window.OfficeModelSelector.selection().model==='candidate-model'")
    catalog = api._connections.snapshot()['catalog']
    assert catalog['current']['agent']['model_id'] == 'candidate-model'
    assert any(model['id'] == 'same-model' for model in catalog['connections'][0]['models'])
    assert len(probes) == 1 and not errors
    out = STATIC.parents[2] / 'deliverables/issue-repair-completion-20261114'
    out.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(out / 'Q3-model-candidates.png'))
    (out / 'Q3_BROWSER.json').write_text(json.dumps({'draft_model_list_reads': len(probes), 'manual_models_retained': True,
        'single_save_applied_selection': True, 'simulated_probe': True, 'actual_model_requests': 0,
        'errors': errors}, ensure_ascii=False, indent=2), encoding='utf-8')


def test_confirmation_duplicate_and_document_change_cannot_reuse_approval(floating_browser):
    page, api, source, calls, errors = floating_browser
    page.evaluate("""()=>{
      window.__decisions=[];
      OfficeConfirm.ask('一次确认').then(value=>window.__decisions.push(['first',value]));
      OfficeConfirm.ask('一次确认').then(value=>window.__decisions.push(['duplicate',value]));
    }""")
    dialog = page.locator('dialog.lix-confirm-dialog[open]')
    dialog.wait_for()
    assert dialog.count() == 1
    page.wait_for_function("window.__decisions.some(row=>row[0]==='duplicate'&&row[1]===false)")
    dialog.get_by_role('button', name='确认', exact=True).click()
    page.wait_for_function("window.__decisions.some(row=>row[0]==='first'&&row[1]===true)")
    page.evaluate("()=>{window.__stale=null;OfficeConfirm.ask('旧文档的确认').then(value=>window.__stale=value);}")
    dialog.wait_for()
    page.evaluate("document.getElementById('file-name').textContent='另一份文档.docx'")
    page.wait_for_function('window.__stale===false')
    assert dialog.count() == 0 and not errors
