import json
import io
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image
from docx import Document

from tests.test_daily_browser import setup_daily
from tests.test_model_connections import connection, payload, save
from tests.test_professional_browser import STATIC, _browser
from tests.test_daily_flows import finish
from tests.test_chat_results_browser import extend
from office_agent.writes import digest


OUT = STATIC.parents[2] / 'deliverables/codex-workspace-20261103/regression'


def simplify(page):
    page.add_script_tag(path=str(STATIC / 'chat_markdown.js'))
    page.add_style_tag(content=(STATIC / 'workbench_ui.css').read_text(encoding='utf8'))
    page.add_script_tag(path=str(STATIC / 'workbench_ui.js'))
    page.wait_for_function('window.OfficeWorkbench && document.getElementById("workspace").dataset.uiMode === window.OfficeModelSelector.mode()')


def mode(page, value):
    page.locator('#model-mode').select_option(value)
    page.wait_for_function('(v) => document.getElementById("workspace").dataset.uiMode === v', arg=value)


def open_model(page):
    page.locator('#simple-model-menu > summary').click()


def open_document(page, view):
    page.locator('#simple-document-menu > summary').click()
    page.locator('#workspace-view-' + view).click()
    page.keyboard.press('Escape')


def pick_model(page, capability, name):
    mode(page, capability)
    open_model(page)
    option = page.locator('#model-selector option').filter(has_text=name + ' · official')
    page.locator('#model-selector').select_option(option.get_attribute('value'))
    page.wait_for_function('window.OfficeModelSelector.selection().valid')
    page.keyboard.press('Escape')


@pytest.mark.parametrize('width', [1440, 1280, 1024, 520])
@pytest.mark.parametrize('theme', ['light', 'dark', 'highcontrast'])
def test_workspace_reading_model_drafts_focus_and_density(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        save(api._connections, payload([connection(), connection('custom')]))
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        page.set_viewport_size({'width': width, 'height': 900 if width == 1440 else 800})
        page.evaluate('(v) => document.documentElement.dataset.theme = v', theme)
        page.locator('#model-mode').select_option('chat')
        page.wait_for_function('document.getElementById("model-selector").options.length > 1')
        option = page.locator('#model-selector option').filter(has_text='same-model · official')
        page.locator('#model-selector').select_option(option.get_attribute('value'))
        page.wait_for_function('window.OfficeModelSelector.selection().valid')
        page.locator('#model-open-input').click()
        page.locator('#prompt').fill('请解释技术报告摘要应该包含哪些内容。')
        folder = OUT / 'screenshots'; folder.mkdir(parents=True, exist_ok=True)
        before = page.locator('button,input,select,summary').evaluate_all('nodes => nodes.filter(n=>n.checkVisibility() && n.getBoundingClientRect().top>=0 && n.getBoundingClientRect().bottom<=innerHeight).length')
        page.screenshot(path=str(folder / f'before-chat-{width}-{theme}.png'))
        simplify(page)
        assert page.locator('#prompt').is_visible()
        assert page.locator('#simple-model-menu').is_visible()
        assert page.locator('#workspace-document-bar').is_hidden()
        assert page.locator('#workspace-stage').is_hidden()
        assert page.locator('#chat-save-history').is_hidden()
        assert page.locator('#chat-export-controls').is_hidden()
        assert page.locator('#model-selection-status').is_hidden()
        assert page.locator('#simple-model-state').inner_text() == '未验证'
        assert page.locator('#prompt').input_value().startswith('请解释')
        open_model(page)
        page.locator('#model-search').fill('custom')
        assert page.locator('#model-selector optgroup').count() == 1
        page.locator('#model-search').fill('')
        page.keyboard.press('Escape')
        assert not page.locator('#simple-model-menu').evaluate('node=>node.open')
        assert page.locator('#simple-model-menu > summary').evaluate('node=>node===document.activeElement')
        assert page.locator('#model-search').is_hidden()
        assert page.locator('#prompt').bounding_box()['y'] < page.viewport_size['height']
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        assert page.evaluate('new Set([...document.querySelectorAll("[id]")].map(n=>n.id)).size===document.querySelectorAll("[id]").length')
        after = page.locator('button,input,select,summary').evaluate_all('nodes=>nodes.filter(n=>n.checkVisibility() && n.getBoundingClientRect().top>=0 && n.getBoundingClientRect().bottom<=innerHeight).length')
        page.screenshot(path=str(folder / f'after-chat-{width}-{theme}.png'))
        (folder / f'density-{width}-{theme}.json').write_text(json.dumps({'width':width,'theme':theme,'before_visible_controls':before,'after_visible_controls':after,'scope':'same empty chat and unsent draft, all visible controls including navigation, not a usability score'},ensure_ascii=False,indent=2),encoding='utf8')
        mode(page, 'image')
        page.locator('#prompt').fill('为技术报告生成一张简洁的实验流程配图。')
        assert page.locator('#simple-start').is_visible()
        assert page.locator('#simple-image-options').is_visible()
        assert page.locator('#simple-budget-menu').is_hidden()
        page.screenshot(path=str(folder / f'after-image-{width}-{theme}.png'))
        mode(page, 'chat')
        assert page.locator('#prompt').input_value().startswith('请解释')
        mode(page, 'agent')
        assert page.locator('#simple-start-document').is_visible()
        assert page.locator('#workspace-stage').is_hidden()
        open_document(page, 'task')
        assert page.locator('#doc-task-request').is_visible()
        assert page.locator('#prompt').is_hidden()
        page.locator('#workspace-mode-dialogue').click()
        assert page.locator('#prompt').is_visible()
        page.locator('#prompt').fill('请润色论文中的这一段，保留数字和引文。')
        page.screenshot(path=str(folder / f'after-document-{width}-{theme}.png'))
        open_document(page, 'task')
        assert page.locator('#prompt').is_hidden()
        assert page.locator('#doc-task-request').is_visible()
        assert not errors
        assert not [call for call in calls if call in {'ask','test_model_connection','confirm_image_generation'}]
        assert app.path is None
    finally:
        fixture.close()


def test_model_route_single_send_and_mode_switch_readonly(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections, payload([connection(), connection('custom')]))
        setup_daily(page, api, original, calls, receipts, monkeypatch)
        routes = []
        def build(route):
            routes.append(route)
            return SimpleNamespace(invoke=lambda messages: SimpleNamespace(content='## 摘要结构\n\n- 研究目的\n- 方法与主要结果\n\n保留结论和限定条件。'))
        monkeypatch.setattr(api._connections, 'build_model', build)
        simplify(page)
        mode(page, 'chat')
        open_model(page)
        option = page.locator('#model-selector option').filter(has_text='same-model · custom')
        page.locator('#model-selector').select_option(option.get_attribute('value'))
        page.wait_for_function('document.getElementById("simple-model-menu").textContent.includes("custom")')
        page.keyboard.press('Escape')
        page.locator('#prompt').fill('技术报告的摘要如何组织？')
        page.locator('#prompt').dispatch_event('keydown', {'key':'Enter','isComposing':True})
        assert len(routes) == 0
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("摘要结构")')
        assert len(routes) == 1 and routes[0].connection_id == 'custom'
        assert page.locator('#chat-transcript h2').inner_text() == '摘要结构'
        assert page.locator('#simple-model-menu > summary').get_attribute('title').startswith('same-model · custom')
        page.locator('#prompt').fill('保留草稿')
        mode(page,'image'); mode(page,'chat')
        assert page.locator('#prompt').input_value() == '保留草稿'
        assert len(routes) == 1
        mode(page,'agent')
        assert page.locator('#simple-start-document').is_visible()
        assert page.locator('#content-panes').is_hidden()
        mode(page,'chat')
        assert page.locator('#prompt').input_value() == '保留草稿'
        page.evaluate('window.authenticated=false;window.dispatchEvent(new Event("officeauth:unauthenticated"))')
        assert page.locator('#prompt').input_value() == ''
        assert page.locator('#simple-recent-list').inner_text() == ''
    finally:
        fixture.close()


def test_migrated_image_confirmation_save_history_and_no_repeat(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        simplify(page); pick_model(page, 'image', 'image-model')
        page.locator('#simple-image-options > summary').click()
        page.locator('#daily-image-size').select_option('1536x1024')
        page.keyboard.press('Escape')
        page.locator('#prompt').fill('生成一张用于技术报告的蓝色实验流程配图。')
        page.locator('#btn-ask').click()
        generate = page.get_by_role('button', name='确认生成（可能产生费用）', exact=True)
        generate.wait_for()
        data = io.BytesIO(); Image.new('RGB',(96,64),'#2866aa').save(data, format='PNG')
        requests = []
        app._image_session[1]._provider_factory = lambda: SimpleNamespace(generate=lambda req,cfg:requests.append((req.size,cfg.model)) or data.getvalue())
        generate.click()
        page.get_by_role('dialog').get_by_role('button', name='取消', exact=True).click()
        assert not requests
        generate.click()
        page.get_by_role('dialog').get_by_role('button', name='确认生成（可能产生费用）', exact=True).click()
        save_button = page.get_by_role('button',name='保存成新的图片文件…',exact=True)
        save_button.wait_for()
        target = tmp_path / 'picture-output'; target.mkdir()
        app.window = SimpleNamespace(create_file_dialog=lambda *a,**kw:[str(target)])
        save_button.click()
        page.get_by_role('button',name='确认保存新文件',exact=True).click()
        page.get_by_role('dialog').get_by_role('button',name='确认保存新文件',exact=True).click()
        page.get_by_role('button',name='打开保存位置',exact=True).wait_for()
        assert len(list(target.glob('*.png'))) == 1
        assert requests == [('1536x1024','image-model')] and app.path is None
        page.screenshot(path=str(OUT / 'screenshots/flow-image-saved.png'))
        page.locator('#nav-history').click()
        page.locator('#simple-history-image').click()
        page.wait_for_function('document.getElementById("daily-image-records").textContent.includes("image-model")')
        assert page.locator('#daily-image-history').is_visible()
        assert page.locator('#simple-history').is_visible()
        page.screenshot(path=str(OUT / 'screenshots/flow-image-history.png'))
        page.locator('#nav-work').click()
        assert page.locator('#prompt').is_visible()
        assert len(requests) == 1
    finally:
        fixture.close()


def test_migrated_doc_range_review_cancel_save_and_undo(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        app.path, app.editor = original.path, original.editor
        names = ['read_edit_ranges','plan_range_edit','confirm_professional_edit','rollback_professional_edit']
        def invoke(name,args):
            assert name in names
            result = getattr(api,name)(*args); receipts.append((name,result)); return result
        page.expose_function('simplifyDocHost', invoke)
        page.evaluate('''names => {
          const previous=window.pywebview.api;
          window.pywebview.api=new Proxy({}, {get:(_,name)=>names.includes(name)?(...args)=>window.simplifyDocHost(name,args):previous[name]});
        }''', names)
        simplify(page)
        source = digest(app.path)
        open_document(page, 'document')
        page.locator('#range-load').click()
        page.wait_for_function('document.getElementById("range-target").options.length > 1')
        page.locator('#range-target').select_option(index=1)
        assert page.locator('#range-original').input_value().startswith('repeat repeat')
        page.locator('#range-original').evaluate('n=>{n.focus();n.setSelectionRange(7,13)}')
        open_document(page, 'task')
        page.locator('#workspace-mode-range').click()
        page.locator('#range-text').fill('second')
        page.locator('#range-reason').fill('仅调整第二处重复词，保留其他内容。')
        page.locator('#range-add').click()
        page.locator('#workspace-see-range-review').click()
        assert page.locator('#workspace-review').is_visible()
        page.locator('#range-plan').click()
        page.wait_for_function('!document.getElementById("range-confirm").disabled')
        page.evaluate('window.approved=false')
        page.locator('#range-confirm').click()
        page.wait_for_function('document.getElementById("range-status").textContent.includes("已取消")')
        assert digest(app.path) == source
        page.locator('#range-plan').click()
        page.wait_for_function('!document.getElementById("range-confirm").disabled')
        page.evaluate('window.approved=true')
        page.locator('#range-confirm').click()
        page.wait_for_function('document.getElementById("range-output").textContent.includes("已保存副本")')
        output = Path(next(result['result']['output'] for name,result in receipts if name == 'confirm_professional_edit' and result.get('result',{}).get('status') == 'saved'))
        assert 'repeat second' in '\n'.join(p.text for p in Document(output).paragraphs)
        assert digest(app.path) == source
        page.screenshot(path=str(OUT / 'screenshots/flow-document-copy.png'))
        page.locator('#range-undo').click()
        page.wait_for_function('document.getElementById("range-output").textContent.includes("已撤销")')
        assert not output.exists()
    finally:
        fixture.close()


def test_chat_more_export_new_session_history_and_reading_scroll(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections,payload([connection()]))
        app=setup_daily(page,api,original,calls,receipts,monkeypatch)
        extend(page,api,calls); simplify(page); pick_model(page,'chat','same-model')
        raw='## 技术报告摘要\n\n' + '\n\n'.join('研究目的、方法、结果和结论应互相对应，并保留适用条件。' for _ in range(32))
        monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content=raw,response_metadata={'finish_reason':'stop'})))
        assert api.ask('请说明技术报告摘要的结构。',{'stream':False})['ok']; finish(app)
        page.evaluate('window.OfficeDaily.refresh()')
        page.wait_for_function('document.querySelector(".simple-message-menu")')
        assert page.get_by_role('button',name='复制回答',exact=True).is_visible()
        assert page.get_by_role('button',name='导出此回答',exact=True).is_hidden()
        app.window=SimpleNamespace(create_file_dialog=lambda *a,**kw:[str(tmp_path)])
        page.locator('.simple-message-menu > summary').click()
        page.evaluate('window.approved=false')
        page.get_by_role('button',name='导出此回答',exact=True).click()
        page.wait_for_function('document.getElementById("chat-export-result").textContent.includes("已取消")')
        assert not list(tmp_path.glob('.office-chat-export-*'))
        page.keyboard.press('Escape')
        page.evaluate('window.approved=true')
        page.locator('.simple-message-menu > summary').click()
        page.get_by_role('button',name='导出此回答',exact=True).click()
        page.wait_for_function('document.getElementById("chat-export-result").textContent.includes("已导出")')
        output=next(tmp_path.glob('.office-chat-export-*/conversation.md'))
        assert raw in output.read_text(encoding='utf8')
        page.locator('#simple-reading').evaluate('n=>n.scrollTop=0')
        page.wait_for_function('!document.getElementById("chat-latest").hidden')
        page.screenshot(path=str(OUT/'screenshots/flow-answer-reading.png'))
        page.locator('#prompt').fill('第一会话草稿')
        first=api.conversation_snapshot()['result']['current']
        page.locator('#simple-new-task > summary').click(); page.locator('#chat-new').click()
        page.wait_for_function('(id)=>document.getElementById("chat-session").value!==id',arg=first)
        assert page.locator('#prompt').input_value()==''
        page.locator('#simple-recent-list button').filter(has_text='新会话').first.click()
        page.locator('#simple-session-list > summary').click()
        page.locator('#simple-history-chat-list button[data-session-id="' + first + '"]').click()
        page.wait_for_function('document.getElementById("prompt").value==="第一会话草稿"')
        assert len(api.conversation_snapshot()['result']['sessions'])==2
        assert calls.count('ask')==0
    finally:
        fixture.close()


def test_same_turn_feedback_collapses_but_new_uncertain_error_stays_visible(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        save(api._connections,payload([connection()]))
        setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page); pick_model(page,'chat','same-model')
        monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='摘要应保留研究条件。',response_metadata={'finish_reason':'length'})))
        page.locator('#prompt').fill('说明摘要要求')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("尚未完整")')
        page.wait_for_function('document.getElementById("status-card").hidden')
        page.wait_for_function('document.querySelector("[data-turn-notice]")?.hidden')
        assert page.locator('#chat-notice').is_visible()
        assert page.locator('[data-turn-notice]').is_hidden()
        turn=api.conversation_snapshot()['result']['sessions'][0]['turns'][0]
        page.evaluate('setStatus("warn","新的操作结果不确定，请先核对，不要重复提交。")')
        page.wait_for_function('!document.getElementById("status-card").hidden')
        assert page.locator('#error-line').is_visible()
        assert '结果不确定' in page.locator('#error-line').inner_text()
        page.evaluate('turn=>window.dispatchEvent(new CustomEvent("officedaily:event",{detail:{type:"answer_draft_start",stream_id:turn.id,session_id:turn.session_id}}))',turn)
        assert '结果不确定' in page.locator('#error-line').inner_text()
    finally:
        fixture.close()
