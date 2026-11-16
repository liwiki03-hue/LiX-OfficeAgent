import hashlib
import json
from pathlib import Path
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_editable_copies import make_readonly
from tests.test_editable_copy_browser import wait
from tests.test_paged_browser import engine_frame, paged_browser, serve_static
from tests.test_professional_browser import EDGE, STATIC
from tests.ui_confirmation import record_browser_failure, track_browser


def permission(page, level):
    page.locator('#floating-access').click()
    dialog = page.get_by_role('dialog', name='编辑方式', exact=True)
    dialog.locator('input[value="' + level + '"]').check()
    assert dialog.locator('input[type="radio"]').count() == 2
    assert dialog.locator('input[value="review"]').count() == 0
    dialog.get_by_role('button', name='应用', exact=True).click()
    dialog.wait_for(state='detached')


def test_chat_edits_and_saves_copies_without_selecting_permissions(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    requests = []
    class Model:
        def __init__(self): self.calls = 0
        def bind_tools(self, tools): return self
        def invoke(self, messages):
            self.calls += 1
            if self.calls == 1:
                row = next(iter(api._floating.paged.references.values()))
                changed = row['old'].replace('实验开展于', '实验实施于') if '实验开展于' in row['old'] else row['old'].replace('实验完成于', '实验开展于') if '实验完成于' in row['old'] else row['old'].replace('实验于', '实验完成于').replace('年完成', '年')
                requests.append({'old': row['old'], 'new': changed})
                return {'tool_calls': [{'name': 'propose_draft_edit', 'id': 'bounded-edit', 'args':
                        {'reference_id': row['id'], 'old': row['old'], 'new': changed, 'reason': '按要求润色，保留数字和单位。'}}]}
            return {'content': '已形成文字修改。'}
    monkeypatch.setattr(api._connections, 'build_model', lambda route: Model())
    page.locator('#simple-model-menu > summary').click()
    page.locator('#model-selector').select_option(json.dumps(['official', 'same-model', 'agent'], separators=(',', ':')))
    page.keyboard.press('Escape')
    out = STATIC.parents[2] / 'deliverables/default-document-edit-20261114'
    out.mkdir(parents=True, exist_ok=True)
    frame = engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。', exact=True).first.click()
    page.keyboard.press('End')
    assert page.locator('#floating-access').text_content() == '自动存副本'
    assert not api._floating.access.state()['active']
    assert not page.locator('dialog[open]').count()
    page.screenshot(path=str(out / '01-default-mode.png'))
    page.evaluate('()=>{window.__editPrompts=[];window.confirm=text=>{window.__editPrompts.push(text);return true};}')
    page.locator('#floating-input').fill('润色光标所在段落，保留数字和单位。')
    page.locator('#floating-send').click()
    wait(page, lambda: api._floating.paged.revision == 1 and len(api._floating.outputs) == 1 and not api._floating.app.busy,
         lambda: {'state': api._floating.paged.state(), 'feedback': page.locator('#paged-feedback').text_content()}, 100)
    wait(page, lambda: not page.locator('#floating-send').is_disabled(), lambda: errors)
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('副本已验证 · 草稿 1')")
    page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('继续修改')")
    first = next(iter(api._floating.outputs.values()))
    assert source.read_bytes() == original and Path(first['backup']).read_bytes() == original
    assert Document(first['output']).paragraphs[1].text == '实验完成于2025年，温度120 ℃。'
    assert first['authorization']['origin'] == 'user_instruction'
    assert any(row['status'] == 'adopted' for row in api._floating.paged.proposals.values())
    page.screenshot(path=str(out / '02-first-automatic-copy.png'))
    page.locator('#floating-input').fill('将“实验完成于”改成“实验开展于”，其他内容保留。')
    page.locator('#floating-send').click()
    wait(page, lambda: len(api._floating.outputs) == 2 and not api._floating.app.busy,
         lambda: {'state': api._floating.paged.state(), 'feedback': page.locator('#paged-feedback').text_content()}, 100)
    receipt = list(api._floating.outputs.values())[-1]
    assert source.read_bytes() == original and Path(receipt['backup']).read_bytes() == original
    assert Document(receipt['output']).paragraphs[1].text == '实验开展于2025年，温度120 ℃。'
    assert receipt['authorization']['level'] == 'save_copy'
    assert len(requests) == 2 and page.evaluate('window.__editPrompts') == []
    assert not page.locator('dialog[open]').count()
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('副本已验证 · 草稿 2')")
    try:
        page.wait_for_function("document.getElementById('paged-feedback').textContent==='草稿 2 的副本已通过验证；原文件未改动。'&&document.getElementById('floating-progress').textContent==='副本已验证，继续输入即可修改'")
    except Exception:
        failure = {'feedback': page.locator('#paged-feedback').text_content(), 'progress': page.locator('#floating-progress').text_content(),
                   'paged': api._floating.paged.state(), 'stage': api._floating.stage,
                   'revision': api._floating.runtime.revision, 'errors': errors,
                   'outputs': [{'id': row['id'], 'revision': row.get('instruction_revision')} for row in api._floating.outputs.values()]}
        (out / ('FAILED-' + source.parent.parent.name + '.json')).write_text(json.dumps(failure, ensure_ascii=False, indent=2), encoding='utf-8')
        page.screenshot(path=str(out / ('FAILED-' + source.parent.parent.name + '.png')))
        raise
    page.screenshot(path=str(out / '03-second-automatic-copy.png'))
    permission(page, 'edit')
    page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('继续修改')")
    page.locator('#floating-input').fill('将“实验开展于”改成“实验实施于”，保留数字和单位。')
    page.locator('#floating-send').click()
    wait(page, lambda: api._floating.paged.revision == 3 and not api._floating.app.busy,
         lambda: api._floating.paged.state())
    page.wait_for_timeout(750)
    assert len(api._floating.outputs) == 2 and len(requests) == 3
    grant = api._floating.access.state()
    page.locator('#floating-access').click()
    mode = page.get_by_role('dialog', name='编辑方式', exact=True)
    page.screenshot(path=str(out / '04-two-edit-modes.png'))
    mode.get_by_role('button', name='停止本次修改', exact=True).click()
    mode.wait_for(state='detached')
    page.wait_for_timeout(750)
    assert not api._floating.access.state()['active'] and api._floating.runtime.closed
    with pytest.raises(ValueError, match='授权已失效'):
        api._floating.access.require(grant['id'], 'edit')
    assert api._floating.paged.revision == 3 and len(api._floating.outputs) == 2
    assert page.evaluate('window.__editPrompts') == []
    assert errors == [] and source.read_bytes() == original
    artifacts = {}
    for name, path in [('样例原稿.docx', source), ('第一次自动副本.docx', Path(first['output'])), ('第二次自动副本.docx', Path(receipt['output']))]:
        data = path.read_bytes(); (out / name).write_bytes(data)
        artifacts[name] = hashlib.sha256(data).hexdigest()
    (out / 'AI_BROWSER.json').write_text(json.dumps({'simulated_model': True, 'edit_requests': requests,
        'default_without_permission_selection': True, 'automatic_copies': [first, receipt], 'source_unchanged': True,
        'draft_only_third_edit': True, 'artifacts': artifacts, 'no_repeat_dialogs': True,
        'stop_invalidates_grant': True, 'final_ui_matches_verified_copy': True, 'errors': errors,
        'native_webview': False, 'real_service': False}, ensure_ascii=False, indent=2), encoding='utf-8')


def test_readonly_copy_opens_for_editing_with_one_confirmation(tmp_path, monkeypatch):
    if not EDGE.is_file(): pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source = tmp_path / '只读报告.docx'
    make_readonly(source)
    original = source.read_bytes()
    app = WebviewApp(); app.attach(SimpleNamespace(create_file_dialog=lambda *a, **kw: [str(source)]))
    api = FloatingApi(app, settings_root=tmp_path)
    monkeypatch.setenv('TEMP', str(tmp_path)); monkeypatch.setenv('TMP', str(tmp_path))
    calls, errors = [], []
    def invoke(name, args):
        calls.append(name)
        assert name not in {'ask', 'paged_submit', 'floating_submit', 'confirm_image_generation'}
        return getattr(api, name)(*args)
    with serve_static() as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 520, 'height': 800}); page.set_default_timeout(15000)
            page.expose_function('testHost', invoke)
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.__prompts=[];window.confirm=text=>{window.__prompts.push(text);return true};document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            page.on('pageerror', lambda error: errors.append(str(error)))
            track_browser(page,tmp_path.name,lambda:{'document':api._floating.document,'paged':api._floating.paged.state(),'busy':app.busy})
            page.goto(url)
            page.locator('#floating-choose').click()
            page.get_by_role('button', name='同意读取并打开', exact=True).click()
            create = page.get_by_role('button', name='创建可编辑副本', exact=True)
            create.wait_for(state='visible')
            old_id = api._floating.document['id']
            try:create.click()
            except Exception:
                record_browser_failure(page,'readonly copy action receives pointer',15)
                raise
            page.get_by_role('dialog', name='创建可编辑副本').get_by_role('button', name='确认创建简化副本', exact=True).click()
            wait(page, lambda: api._floating.document['id'] != old_id and api._floating.paged.editable and bool(api._floating.paged.manifest),
                 lambda: {'document': api._floating.document, 'state': api._floating.paged.state(), 'errors': errors}, 100)
            assert api._floating.document['ready'] and api._floating.access.state()['level'] == 'save_copy'
            assert source.read_bytes() == original and len(api._floating.outputs) == 1
            assert calls.count('floating_read') == 1 and calls.count('floating_select_output') == 1
            assert page.evaluate('window.__prompts') == [] and errors == []
            page.get_by_role('heading', name='当前编辑副本', exact=True).wait_for()
            assert page.get_by_role('button', name='打开副本编辑…', exact=True).count() == 0
            out = STATIC.parents[2] / 'deliverables/default-document-edit-20261114'; out.mkdir(parents=True, exist_ok=True)
            page.screenshot(path=str(out / '05-copy-opened-520.png'))
            (out / 'COPY_BROWSER.json').write_text(json.dumps({'single_creation_confirmation': True, 'automatically_opened': True,
                'source_unchanged': True, 'model_requests': 0, 'read_calls': calls.count('floating_read'),
                'permission': api._floating.access.state(), 'errors': errors, 'native_webview': False}, ensure_ascii=False, indent=2), encoding='utf-8')
        finally:
            browser.close()
