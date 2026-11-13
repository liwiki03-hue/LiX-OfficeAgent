import json
import threading
from types import SimpleNamespace

import httpx2
import pytest

from tests.test_daily_browser import setup_daily, choose
from tests.test_daily_flows import finish
from tests.test_model_connections import connection, payload, save
from tests.test_professional_browser import _browser
from tests.test_chat_context import sse


def initialize(page, api, original, calls, receipts, monkeypatch):
    save(api._connections, payload([connection(),connection('custom','fixture-private-two')]))
    app = setup_daily(page, api, original, calls, receipts, monkeypatch)
    page.expose_function('contextHost', api.preview_chat_context)
    page.evaluate('''() => {
      const previous = window.pywebview.api;
      window.pywebview.api = new Proxy({}, {get: (_,name) => name === 'preview_chat_context'
        ? (...args) => window.contextHost(...args) : previous[name]});
    }''')
    choose(page, 'chat', 'same-model')
    return app


@pytest.mark.parametrize('width,theme', [(1440,'light'), (1280,'dark'), (520,'highcontrast')])
def test_context_preview_and_stale_send_preserve_input_ui(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        app = initialize(page, api, original, calls, receipts, monkeypatch)
        seen = []
        monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=lambda messages:seen.append(messages) or SimpleNamespace(content='原回答',response_metadata={'finish_reason':'stop'})))
        assert api.ask('已有问题', {'stream':False})['ok']; finish(app)
        page.evaluate('window.OfficeDaily.refresh()')
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("原回答")')
        page.set_viewport_size({'width':width, 'height':800})
        page.evaluate('(theme) => document.documentElement.dataset.theme=theme', theme)
        page.locator('#prompt').fill('保留数字20 ms的下一问题 📄')
        page.locator('#chat-preview-context').click()
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("预计发送 1 组历史")')
        assert len(seen) == 1 and '已有问题' in page.locator('#chat-context-content').inner_text()
        assert '内置tokenizer估算' in page.locator('#chat-context-content').inner_text()
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        assert not app.busy and app.path is None
        page.locator('#chat-context-preview').scroll_into_view_if_needed()
        page.screenshot(path=str(tmp_path / f'chat-context-{width}.png'))
        previous = app._push
        with monkeypatch.context() as delayed:
            delayed.setattr(app, '_push', lambda event:None if event['type'] == 'chat_changed' else previous(event))
            assert api.new_conversation()['ok']
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("重新预览核对")')
        assert len(seen) == 1 and page.locator('#prompt').input_value() == '保留数字20 ms的下一问题 📄'
    finally:
        fixture.close()


@pytest.mark.parametrize('change', ['input','mode','logout'])
def test_late_preview_cannot_restore_old_content(tmp_path, monkeypatch, change):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        initialize(page, api, original, calls, receipts, monkeypatch)
        monkeypatch.setattr(api._connections, 'build_model', lambda *args:pytest.fail('preview called model'))
        page.evaluate('''() => {
          const previous = window.pywebview.api;
          window.pywebview.api = new Proxy({}, {get: (_,name) => name === 'preview_chat_context'
            ? (...args) => previous[name](...args).then(value => new Promise(resolve => {window.releasePreview=()=>resolve(value);})):previous[name]});
        }''')
        page.locator('#prompt').fill('旧预览问题')
        page.locator('#chat-preview-context').click()
        page.wait_for_function('typeof window.releasePreview === "function"')
        if change == 'input':
            page.locator('#prompt').fill('新的输入')
        elif change == 'mode':
            page.locator('#model-mode').select_option('image')
            page.wait_for_function('window.OfficeModelSelector.mode() === "image"')
        else:
            page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate('window.releasePreview()')
        page.wait_for_timeout(80)
        assert page.locator('#chat-context-preview').is_hidden()
        assert page.locator('#chat-context-content').inner_text() == ''
        assert not [method for method in calls if method == 'ask']
    finally:
        fixture.close()


def test_submit_lock_and_transfer_reply_after_logout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        initialize(page, api, original, calls, receipts, monkeypatch)
        monkeypatch.setattr(api._connections, 'build_model', lambda *args:pytest.fail('race fixture called model'))
        page.evaluate('''() => {
          const previous = window.pywebview.api; window.sendCalls=[];
          window.pywebview.api = new Proxy({}, {get: (_,name) => name === 'ask' ? (...args) => {
            window.sendCalls.push(args);
            if (window.sendCalls.length===1) return Promise.resolve({ok:false,reason:'chat_transfer_required',route:{connection_name:'新接口',model_id:'same-model'}});
            return new Promise(resolve=>{window.releaseSend=resolve;});
          }:previous[name]});
          window.pendingSend=window.OfficeDaily.send('原问题');
        }''')
        page.wait_for_function('typeof window.releaseSend === "function"')
        assert page.evaluate('window.OfficeDaily.send("重复点击")') == {'ok':False,'reason':'busy'}
        assert page.evaluate('window.sendCalls.length') == 2
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));document.getElementById('prompt').value='新的窗口草稿';document.getElementById('chat-notice').textContent='新的状态';window.releaseSend({ok:true,context_turns:1,omitted_turns:0});")
        assert page.evaluate('window.pendingSend')['reason'] == 'stale'
        assert page.locator('#prompt').input_value() == '新的窗口草稿'
        assert page.locator('#chat-notice').inner_text() == '新的状态'
        assert not [method for method in calls if method == 'ask']
    finally:
        fixture.close()


def test_rejected_host_submission_is_uncertain_and_not_retried(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        initialize(page, api, original, calls, receipts, monkeypatch)
        page.evaluate('''() => {
          const previous = window.pywebview.api;window.sendCalls=0;
          window.pywebview.api=new Proxy({}, {get:(_,name)=>name==='ask'?()=>{window.sendCalls++;return Promise.reject(new Error('private host detail'));}:previous[name]});
        }''')
        page.locator('#prompt').fill('结果不确定时保留 📄')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("不要重复发送")')
        assert page.evaluate('window.sendCalls') == 1
        assert page.locator('#prompt').input_value() == '结果不确定时保留 📄'
        assert 'private host detail' not in page.content()
    finally:
        fixture.close()


def test_logout_during_post_send_refresh_does_not_return_success(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        initialize(page, api, original, calls, receipts, monkeypatch)
        state = api.conversation_snapshot()
        page.evaluate('''() => {
          const previous = window.pywebview.api;
          window.pywebview.api = new Proxy({}, {get:(_,name)=>name==='ask'
            ?()=>Promise.resolve({ok:true,context_turns:0,omitted_turns:0})
            :name==='conversation_snapshot'?()=>new Promise(resolve=>{window.releaseSnapshot=resolve;}):previous[name]});
          window.pendingSend=window.OfficeDaily.send('原问题');
        }''')
        page.wait_for_function('typeof window.releaseSnapshot === "function"')
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));document.getElementById('prompt').value='退出后的新草稿';")
        page.evaluate('(value)=>window.releaseSnapshot(value)', state)
        assert page.evaluate('window.pendingSend')['reason'] == 'stale'
        assert page.locator('#prompt').input_value() == '退出后的新草稿'
        assert page.locator('#chat-transcript').inner_text() == ''
    finally:
        fixture.close()


def test_actual_sdk_live_stream_stop_and_frozen_route_ui(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    release = threading.Event()
    try:
        page, api, original, calls, receipts = next(fixture)
        app = initialize(page, api, original, calls, receipts, monkeypatch)
        requests = []
        class Stream(httpx2.SyncByteStream):
            def __iter__(self):
                yield sse('这是已经接收的真实SDK流式片段，尚未结束。' * 4, None)
                assert release.wait(10)
                yield sse('后续内容', 'stop')
        def handle(request):
            requests.append((str(request.url), request.headers['authorization']))
            return httpx2.Response(200, headers={'content-type':'text/event-stream'}, stream=Stream())
        monkeypatch.setattr('office_agent.webui.model_connections._http_client', lambda:httpx2.Client(transport=httpx2.MockTransport(handle), follow_redirects=False, trust_env=False))
        page.locator('#prompt').fill('检查流式响应，不读文档')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-live-answer").textContent.includes("真实SDK流式片段")')
        custom = page.locator('#model-selector option').filter(has_text='same-model · custom')
        page.locator('#model-selector').select_option(custom.get_attribute('value'))
        page.wait_for_function('document.getElementById("model-selection-status").textContent.includes("custom")')
        assert page.locator('#prompt').is_disabled()
        page.locator('#btn-stop').click()
        page.wait_for_function('document.getElementById("chat-live-answer").hidden')
        release.set()
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("已停止本地显示")')
        page.wait_for_function('!document.getElementById("prompt").disabled')
        page.locator('#prompt').fill('新模型用于后续请求')
        assert requests == [('https://official.example/v1/chat/completions','Bearer fixture-private-one')]
        turn = api.conversation_snapshot()['result']['sessions'][0]['turns'][0]
        assert turn['status'] == 'stopped' and turn['route']['connection_id'] == 'official'
        assert page.locator('#prompt').input_value() == '新模型用于后续请求'
        assert not app.busy
    finally:
        release.set()
        fixture.close()
