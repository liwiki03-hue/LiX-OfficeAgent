from types import SimpleNamespace

import pytest

from tests.test_chat_context_browser import initialize
from tests.test_model_connections import connection,payload,save,select
from tests.test_professional_browser import _browser
from tests.test_chat_budget import local_connection


def prepare(page,api,original,calls,receipts,monkeypatch):
    app = initialize(page,api,original,calls,receipts,monkeypatch)
    row = connection(); row['context_window'] = 1024
    save(api._connections,payload([row])); select(api._connections,'official')
    page.evaluate('window.OfficeModelSelector.refresh()')
    page.wait_for_function('window.OfficeModelSelector.readyForRequest()')
    return app


@pytest.mark.parametrize('width,theme',[(1440,'light'),(1280,'dark'),(520,'highcontrast')])
def test_budget_rejection_recovery_and_model_stays_selected(tmp_path,monkeypatch,width,theme):
    fixture = _browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts = next(fixture)
        app = prepare(page,api,original,calls,receipts,monkeypatch)
        seen = []
        monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:seen.append(messages) or SimpleNamespace(content='缩短后正常回答',response_metadata={'finish_reason':'stop'})))
        page.set_viewport_size({'width':width,'height':800})
        page.evaluate('(theme)=>document.documentElement.dataset.theme=theme',theme)
        text = '保留严格限定条件、数字与单位。' * 300
        page.locator('#prompt').fill(text)
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("上下文预算")')
        page.wait_for_function('document.getElementById("status-line").textContent.includes("模型选择") || document.getElementById("status-line").textContent.includes("预算")')
        assert '模型选择已失效' not in page.locator('#status-line').text_content()
        assert '预算' in page.locator('#status-line').text_content()
        assert page.locator('#prompt').input_value() == text and len(seen) == 0
        assert not app.busy and api.get_model_connections()['catalog']['current']['chat']['valid']
        page.wait_for_function('!document.getElementById("chat-budget-feedback").hidden')
        assert '1024' in page.locator('#chat-budget-message').inner_text()
        page.locator('#chat-budget-feedback').scroll_into_view_if_needed()
        page.screenshot(path=str(tmp_path/f'chat-budget-rejected-{width}.png'))
        page.locator('#chat-budget-shorten').click()
        assert page.locator('#prompt').evaluate('node=>document.activeElement===node')
        page.locator('#prompt').fill('简短问题，保留20 ms。')
        page.locator('#chat-preview-context').click()
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("预计发送")')
        assert len(seen) == 0 and page.locator('#chat-budget-feedback').is_hidden()
        page.locator('#chat-context-preview').scroll_into_view_if_needed()
        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
        page.screenshot(path=str(tmp_path/f'chat-budget-{width}.png'))
        page.locator('#chat-stream').uncheck()
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("缩短后正常回答")')
        assert len(seen) == 1
    finally:
        fixture.close()


def test_ollama_preview_names_runtime_parameter_without_calling_model(tmp_path,monkeypatch):
    fixture = _browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts = next(fixture)
        prepare(page,api,original,calls,receipts,monkeypatch)
        save(api._connections,payload([local_connection(window=2048)])); select(api._connections,'official')
        page.evaluate('window.OfficeModelSelector.refresh()')
        page.wait_for_function('window.OfficeModelSelector.readyForRequest()')
        monkeypatch.setattr(api._connections,'build_model',lambda *args:pytest.fail('preview called model'))
        page.locator('#prompt').fill('不发送的本机问题')
        page.locator('#chat-preview-context').click()
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("num_ctx=2048")')
        assert '未验证服务端实际窗口' in page.locator('#chat-context-content').inner_text()
        assert page.locator('#prompt').input_value() == '不发送的本机问题'
        assert not [method for method in calls if method=='ask']
        page.locator('#chat-context-preview').scroll_into_view_if_needed()
        page.screenshot(path=str(tmp_path/'chat-budget-ollama.png'))
    finally:
        fixture.close()


@pytest.mark.parametrize('change',['input','mode','logout'])
def test_budget_feedback_clears_on_context_change(tmp_path,monkeypatch,change):
    fixture = _browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts = next(fixture)
        prepare(page,api,original,calls,receipts,monkeypatch)
        monkeypatch.setattr(api._connections,'build_model',lambda *args:pytest.fail('overflow called model'))
        page.locator('#prompt').fill('保留严格限定条件、数字与单位。'*300)
        page.locator('#chat-preview-context').click()
        page.wait_for_function('!document.getElementById("chat-budget-feedback").hidden')
        if change=='input':
            page.locator('#prompt').fill('new')
        elif change=='mode':
            page.locator('#model-mode').select_option('image')
            page.wait_for_function('window.OfficeModelSelector.mode()==="image"')
        else:
            page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#chat-budget-feedback').is_hidden()
        assert page.locator('#chat-budget-message').inner_text()==''
    finally:
        fixture.close()


def test_budget_configuration_entry_keeps_original_draft(tmp_path,monkeypatch):
    fixture = _browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts = next(fixture)
        prepare(page,api,original,calls,receipts,monkeypatch)
        text='保留严格限定条件、数字与单位。'*300
        page.locator('#prompt').fill(text)
        page.locator('#chat-preview-context').click()
        page.wait_for_function('!document.getElementById("chat-budget-feedback").hidden')
        page.locator('#chat-budget-configure').click()
        page.wait_for_function('!document.getElementById("settings-dialog").hidden')
        page.locator('#settings-close').click()
        assert page.locator('#prompt').input_value()==text
        assert api.get_model_connections()['catalog']['connections'][0]['context_window']==1024
        assert not [method for method in calls if method=='ask']
    finally:
        fixture.close()
