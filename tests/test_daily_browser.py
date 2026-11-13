import io
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from office_agent.webui.bridge import WebviewApp
from tests.test_model_connections import connection, payload, save, select
from tests.test_model_connections_browser import initialize
from tests.test_professional_browser import _browser, STATIC


def setup_daily(page, api, original, calls, receipts, monkeypatch):
    actual = WebviewApp()
    api._app = actual
    api._connections.app = actual
    actual._connections = api._connections
    api._auth.account_key = lambda: 'd' * 64
    api._connections.conversations.key = api._memory_key
    initialize(page, api, original, calls, receipts)
    allowed = {'ask', 'poll', 'list_sessions', 'snapshot', 'cancel_task', 'conversation_snapshot',
               'new_conversation', 'save_conversation_history', 'reload_conversations', 'select_conversation', 'delete_conversation', 'clear_saved_conversations',
               'get_model_connections', 'select_session_model', 'set_model_mode', 'plan_settings_update',
               'confirm_settings_update', 'confirm_image_generation', 'plan_image_save', 'confirm_image_save',
               'plan_image_rollback', 'confirm_image_rollback', 'open_saved_image_folder', 'list_independent_image_history'}
    def invoke(method, args):
        assert method in allowed, method
        if method != 'poll':
            calls.append(method)
        return getattr(api, method)(*args)
    page.expose_function('dailyHost', invoke)
    page.evaluate("() => {window.pywebview.api = new Proxy({}, {get: (_, name) => (...args) => window.dailyHost(name,args)}); window.OfficeAuth.ensure = async () => true; document.getElementById('file-name').textContent = '未选择文件';}")
    page.add_script_tag(path=str(STATIC / 'images.js'))
    page.add_script_tag(path=str(STATIC / 'imagegen.js'))
    page.add_script_tag(path=str(STATIC / 'daily.js'))
    page.evaluate("window.dispatchEvent(new Event('officeauth:authenticated'))")
    page.wait_for_function('window.OfficeDaily && !document.getElementById("model-mode").disabled')
    return actual


def choose(page, capability, model):
    page.locator('#model-mode').select_option(capability)
    page.wait_for_function('(cap) => window.OfficeModelSelector.mode() === cap', arg=capability)
    page.wait_for_function('([cap,model]) => Array.from(document.getElementById("model-selector").options).some(option => {try {const value=JSON.parse(option.value);return value[1]===model && value[2]===cap;}catch (_) {return false;}})', arg=[capability,model])
    option = page.locator('#model-selector option').filter(has_text=model + ' · official')
    page.locator('#model-selector').select_option(option.get_attribute('value'))
    page.wait_for_function('(model) => document.getElementById("model-selection-status").textContent.includes("当前：" + model)', arg=model)
    page.locator('#model-open-input').click()


@pytest.mark.parametrize('width,theme', [(1440,'light'), (520,'dark')])
def test_chat_sessions_mode_drafts_and_history_restore_ui(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        seen = []
        def invoke(messages):
            seen.append(messages)
            return SimpleNamespace(content='回答：' + messages[-1]['content'])
        monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=invoke))
        page.set_viewport_size({'width':width,'height':800})
        page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
        choose(page, 'chat', 'same-model')
        page.locator('#prompt').fill('请编写一段生图提示词，不生成图片')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("回答：请编写")')
        assert len(seen) == 1 and app.path is None and app._image_session is None
        page.wait_for_function('document.getElementById("prompt").value === ""')
        page.locator('#prompt').fill('第一会话草稿 📄')
        first = api.conversation_snapshot()['result']['current']
        page.locator('#chat-new').click()
        page.wait_for_function('(id) => document.getElementById("chat-session").value !== id', arg=first)
        assert page.locator('#prompt').input_value() == ''
        assert '回答：请编写' not in page.locator('#chat-transcript').inner_text()
        page.locator('#prompt').fill('第二会话草稿')
        page.locator('#chat-session').select_option(first)
        page.wait_for_function('document.getElementById("prompt").value.includes("第一会话")')
        choose(page, 'image', 'image-model')
        page.locator('#prompt').fill('生图未发送草稿')
        choose(page, 'chat', 'same-model')
        assert page.locator('#prompt').input_value() == '第一会话草稿 📄'
        assert len(seen) == 1
        page.locator('#chat-save-history').check()
        page.wait_for_function('document.getElementById("chat-save-history").checked && !document.getElementById("chat-save-history").disabled')
        page.locator('#prompt').fill('继续第二轮')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("回答：继续第二轮")')
        assert [m['role'] for m in seen[-1]] == ['user','assistant','user']
        assert Path(api.conversation_snapshot()['result']['location']).exists()
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        page.locator('#prompt').fill('重新读取时保留的草稿')
        page.locator('#chat-refresh').click()
        page.wait_for_function('(id) => document.getElementById("chat-session").value === id', arg=first)
        page.wait_for_function('document.getElementById("prompt").value === "重新读取时保留的草稿"')
        page.screenshot(path=str(tmp_path / f'daily-chat-{width}-{theme}.png'))
        page.locator('#prompt').fill('中文输入法测试')
        count = len(seen)
        page.locator('#prompt').dispatch_event('keydown', {'key':'Enter','isComposing':True})
        assert len(seen) == count
        page.evaluate("window.authenticated=false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#prompt').input_value() == ''
        assert page.locator('#chat-transcript').inner_text() == ''
    finally:
        fixture.close()


def test_standalone_image_real_ui_confirm_save_and_undo(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        folder = tmp_path / 'selected-output'; folder.mkdir()
        app.window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(folder)])
        choose(page, 'image', 'image-model')
        page.locator('#daily-image-size').select_option('1536x1024')
        page.locator('#prompt').fill('一张蓝色技术示意图')
        page.locator('#btn-ask').click()
        page.get_by_role('button', name='确认生成（可能产生费用）', exact=True).wait_for()
        service = app._image_session[1]
        stream = io.BytesIO(); Image.new('RGB',(48,32),'blue').save(stream,format='PNG')
        requests = []
        def generate(request, config):
            requests.append((request.size,config.model))
            return stream.getvalue()
        service._provider_factory = lambda: SimpleNamespace(generate=generate)
        page.get_by_role('button', name='确认生成（可能产生费用）', exact=True).click()
        page.get_by_role('dialog', name='确认联网生成图片？').get_by_role('button', name='取消', exact=True).click()
        assert not requests
        page.get_by_role('button', name='确认生成（可能产生费用）', exact=True).click()
        page.get_by_role('dialog', name='确认联网生成图片？').get_by_role('button', name='确认生成（可能产生费用）', exact=True).click()
        page.get_by_role('button', name='保存成新的图片文件…', exact=True).wait_for()
        assert requests == [('1536x1024','image-model')] and app.path is None
        app.window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: None)
        page.get_by_role('button', name='保存成新的图片文件…', exact=True).click()
        page.wait_for_function('Array.from(document.querySelectorAll("button")).some(button => button.textContent === "保存成新的图片文件…" && !button.disabled)')
        assert not list(folder.glob('*.png')) and len(requests) == 1
        app.window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(folder)])
        page.get_by_role('button', name='保存成新的图片文件…', exact=True).click()
        page.get_by_role('button', name='不保存', exact=True).click()
        page.wait_for_function('Array.from(document.querySelectorAll("button")).some(button => button.textContent === "保存成新的图片文件…" && !button.disabled)')
        assert not list(folder.glob('*.png')) and len(requests) == 1
        page.get_by_role('button', name='保存成新的图片文件…', exact=True).click()
        page.get_by_role('button', name='确认保存新文件', exact=True).click()
        page.get_by_role('dialog').get_by_role('button', name='确认保存新文件', exact=True).click()
        page.get_by_role('button', name='撤销这个文件…', exact=True).wait_for()
        output = next(folder.glob('*.png'))
        assert output.is_file() and len(requests) == 1
        page.wait_for_function('document.getElementById("workspace-current-status").textContent.includes("图片已保存")')
        page.screenshot(path=str(tmp_path / 'daily-image-saved.png'))
        page.get_by_role('button', name='撤销这个文件…', exact=True).click()
        page.get_by_role('button', name='确认删除这个文件', exact=True).click()
        page.get_by_role('dialog').get_by_role('button', name='确认删除这个文件', exact=True).click()
        page.wait_for_function('document.getElementById("log-pane").textContent.includes("图片已撤销")')
        assert not output.exists() and len(requests) == 1
        page.locator('#daily-open-images').click()
        page.wait_for_function('document.getElementById("daily-image-records").textContent.includes("已撤销输出")')
    finally:
        fixture.close()


def test_empty_configuration_preserves_prompt_and_offers_configuration(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        page.locator('#model-mode').select_option('chat')
        page.wait_for_function('window.OfficeModelSelector.mode() === "chat"')
        page.locator('#model-open-input').click()
        page.locator('#prompt').fill('保留尚未发送的中文问题')
        page.locator('#btn-ask').click()
        assert page.locator('#prompt').input_value() == '保留尚未发送的中文问题'
        assert 'ask' not in calls
        page.locator('#model-configure').click()
        page.locator('#connection-add').wait_for(state='visible')
        page.locator('#settings-close').click()
        assert page.locator('#prompt').input_value() == '保留尚未发送的中文问题'
        assert app.path is None and app.editor is None
    finally:
        fixture.close()


def test_background_history_read_cannot_reset_pending_save_preference(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        setup_daily(page, api, original, calls, receipts, monkeypatch)
        choose(page,'chat','same-model')
        page.evaluate("""() => {
          const previous=window.pywebview.api;
          window.pywebview.api=new Proxy({}, {get:(_,name)=>name==='plan_settings_update'
            ? (...args)=>new Promise(resolve=>{window.releasePreference=()=>previous[name](...args).then(resolve);})
            : previous[name]});
        }""")
        page.locator('#chat-save-history').check()
        page.evaluate('window.OfficeDaily.refresh()')
        assert page.locator('#chat-save-history').is_checked()
        assert page.locator('#chat-save-history').is_disabled()
        page.evaluate('window.releasePreference()')
        page.wait_for_function('document.getElementById("chat-save-history").checked && !document.getElementById("chat-save-history").disabled')
        assert api.conversation_snapshot()['result']['saving_enabled'] is True
    finally:
        fixture.close()
