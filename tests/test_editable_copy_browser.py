import hashlib
import json
from pathlib import Path
from types import SimpleNamespace
import time

from docx import Document
import pytest

from office_agent.editable_copies import body_blocks
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_editable_copies import make_readonly
from tests.test_paged_browser import engine_frame, serve_static
from tests.test_professional_browser import EDGE, STATIC


def wait(page, predicate, message, seconds=75):
    deadline = time.monotonic() + seconds
    while not predicate():
        if time.monotonic() >= deadline:
            raise AssertionError(message())
        page.wait_for_timeout(100)


def test_readonly_to_copy_actual_editor_publish_and_rollback(tmp_path, monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    out = STATIC.parents[2] / 'deliverables/default-document-edit-20261114/manual-optout'
    out.mkdir(parents=True, exist_ok=True)
    source = tmp_path / '只读开题报告.docx'
    make_readonly(source)
    original = source.read_bytes()
    monkeypatch.setenv('TEMP', str(tmp_path))
    monkeypatch.setenv('TMP', str(tmp_path))
    app = WebviewApp()
    app.attach(SimpleNamespace(create_file_dialog=lambda *a, **kw: [str(source)]))
    api = FloatingApi(app, settings_root=tmp_path)
    calls, errors = [], []
    def invoke(name, args):
        if name not in {'floating_snapshot', 'poll'}:
            calls.append(name)
        assert name not in {'ask', 'paged_submit', 'floating_submit', 'confirm_image_generation'}
        return getattr(api, name)(*args)
    with serve_static() as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1280, 'height': 820})
            page.set_default_timeout(15000)
            page.expose_function('testHost', invoke)
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.confirm=()=>true;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto(url)
            page.locator('#floating-choose').click()
            page.get_by_role('button', name='同意读取并打开', exact=True).click()
            create = page.get_by_role('button', name='创建可编辑副本', exact=True)
            create.wait_for(state='visible')
            wait(page, lambda: create.is_enabled(), lambda: page.locator('#paged-feedback').text_content())
            page.locator('#floating-input').fill('请稍后检查这一段，保留数字和单位。')
            page.screenshot(path=str(out / '01-readonly.png'))
            create.click()
            plan_dialog = page.get_by_role('dialog', name='创建可编辑副本')
            plan_dialog.get_by_role('button', name='确认创建简化副本', exact=True).wait_for(state='visible')
            assert '原样式、自动编号、颜色和分页将重建' in plan_dialog.inner_text()
            assert '25 ℃' not in plan_dialog.inner_text()
            page.set_viewport_size({'width': 520, 'height': 800})
            bounds = plan_dialog.bounding_box()
            assert bounds['x'] >= 0 and bounds['x'] + bounds['width'] <= 521
            page.screenshot(path=str(out / '02-copy-plan-520.png'))
            plan_dialog.get_by_role('button', name='取消', exact=True).click()
            assert not list(tmp_path.glob('*.agent-*.docx'))
            assert page.locator('#floating-input').input_value() == '请稍后检查这一段，保留数字和单位。'
            page.set_viewport_size({'width': 1280, 'height': 820})
            create.click()
            plan_dialog.get_by_role('checkbox', name='创建后直接打开，按当前方式修改副本').uncheck()
            plan_dialog.get_by_role('button', name='确认创建简化副本', exact=True).click()
            wait(page, lambda: bool(api._floating.outputs) and not app.busy,
                 lambda: {'feedback': page.locator('#paged-feedback').text_content(), 'outputs': api._floating.outputs})
            row = next(iter(api._floating.outputs.values()))
            output = Path(row['output'])
            assert row['kind'] == 'paged_editable_copy' and row['render_diff']['inputs_unchanged']
            assert row['render_diff']['render']['status'] in {'equivalent', 'different'}
            assert source.read_bytes() == original and Path(row['backup']).read_bytes() == original
            assert body_blocks(output.read_bytes()) == body_blocks(original)
            page.evaluate('OfficeFloating.history()')
            page.get_by_role('button', name='打开副本编辑…', exact=True).wait_for(state='visible')
            page.screenshot(path=str(out / '03-copy-generated.png'))
            old_id = api._floating.document['id']
            page.get_by_role('button', name='打开副本编辑…', exact=True).click()
            wait(page, lambda: api._floating.document['id'] != old_id, lambda: api._floating.document)
            if page.locator('#floating-return').is_visible():page.locator('#floating-return').click()
            assert api._floating.document['ready'] and not api._floating.runtime
            assert not api._floating.access.state()['active'] and not api._floating.access.state()['id']
            assert page.locator('#floating-read').is_hidden()
            wait(page, lambda: bool(api._floating.paged.manifest) or bool(api._floating.paged.document_id and not api._floating.paged.editable), lambda: api._floating.paged.state())
            assert api._floating.paged.editable and not api._floating.paged.proposals
            frame = engine_frame(page)
            target = frame.locator('.superdoc-page').get_by_text('实验条件为25 ℃，30个样本的延迟为120 ms[1]。', exact=True).first
            target.click()
            page.keyboard.press('Home')
            page.keyboard.press('Shift+End')
            page.keyboard.insert_text('实验条件保持25 ℃，30个样本的延迟为120 ms[1]。')
            wait(page, lambda: api._floating.paged.revision > 0, lambda: api._floating.paged.state())
            page.get_by_role('button', name='保存副本', exact=True).click()
            publish = page.get_by_role('dialog', name='生成验证副本')
            publish.get_by_role('button', name='确认生成副本', exact=True).click()
            wait(page, lambda: len(api._floating.outputs) == 2 and not app.busy,
                 lambda: {'feedback': page.locator('#paged-feedback').text_content(), 'state': api._floating.paged.state()})
            final = next(r for r in api._floating.outputs.values() if r['id'] != row['id'])
            assert Document(final['output']).paragraphs[1].text == '实验条件保持25 ℃，30个样本的延迟为120 ms[1]。'
            assert source.read_bytes() == original and body_blocks(output.read_bytes()) == body_blocks(original)
            wait(page, lambda: '副本已验证' in page.locator('#paged-save-state').text_content(), lambda: page.locator('#paged-save-state').text_content())
            page.screenshot(path=str(out / '04-copy-edited.png'))
            artifacts = {}
            for name, path in [('原稿.docx', source), ('简化可编辑副本.docx', output), ('修改后副本.docx', Path(final['output']))]:
                data = path.read_bytes()
                (out / name).write_bytes(data)
                artifacts[name] = hashlib.sha256(data).hexdigest()
            assert api.floating_rollback(final['id'], True)['ok']
            wait(page, lambda: not Path(final['output']).exists() and not app.busy, lambda: app.busy)
            api.floating_snapshot(0)
            assert api.floating_rollback(row['id'], True)['ok']
            wait(page, lambda: not output.exists() and not app.busy, lambda: app.busy)
            assert source.read_bytes() == original and errors == []
            (out / 'BROWSER.json').write_text(json.dumps({'source_sha256': hashlib.sha256(original).hexdigest(),
                'source_unchanged': True, 'copy': row, 'edited': final, 'calls': calls, 'errors': errors,
                'model_requests': 0, 'cancel_no_output': True, 'copy_reopened_editable': True,
                'open_click_authorizes_read': True, 'grant_not_inherited_without_consent': True,
                'two_outputs_rolled_back': True, 'artifacts': artifacts,
                'native_webview': False, 'office': False}, ensure_ascii=False, indent=2), encoding='utf-8')
        finally:
            browser.close()
