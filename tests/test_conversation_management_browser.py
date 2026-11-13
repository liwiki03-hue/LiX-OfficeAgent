import pytest

from tests.test_daily_browser import setup_daily, choose
from tests.test_model_connections import connection, payload, save
from tests.test_professional_browser import _browser


def initialize(page, api, app, calls, receipts, monkeypatch):
    save(api._connections, payload([connection()]))
    setup_daily(page, api, app, calls, receipts, monkeypatch)
    page.expose_function('renameHost', api.rename_conversation)
    page.evaluate('''() => {
      const previous = window.pywebview.api;
      window.pywebview.api = new Proxy({}, {get: (_, name) => name === 'rename_conversation'
        ? (...args) => window.renameHost(...args) : previous[name]});
    }''')
    monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('management invoked model'))
    choose(page, 'chat', 'same-model')


@pytest.mark.parametrize('width,theme', [(1440,'light'), (1280,'dark'), (520,'highcontrast')])
def test_search_rename_draft_and_logout_management(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        initialize(page, api, app, calls, receipts, monkeypatch)
        page.set_viewport_size({'width':width,'height':800})
        page.evaluate('(value) => document.documentElement.dataset.theme = value', theme)
        page.locator('#prompt').fill('保留未发送草稿 📄')
        page.locator('#chat-title').fill('中文技术报告')
        assert '名称草稿尚未采用' in page.locator('#chat-title-hint').inner_text()
        page.evaluate('window.OfficeDaily.refresh()')
        page.wait_for_timeout(80)
        assert page.locator('#chat-title').input_value() == '中文技术报告'
        page.locator('#chat-rename').click()
        page.wait_for_function('document.getElementById("chat-session").selectedOptions[0].textContent.includes("中文技术报告")')
        assert page.locator('#prompt').input_value() == '保留未发送草稿 📄'
        first = api.conversation_snapshot()['result']['current']
        page.locator('#chat-new').click()
        page.wait_for_function('(id) => document.getElementById("chat-session").value !== id', arg=first)
        page.locator('#chat-search').fill('技术报告')
        assert page.locator('#chat-session option').count() == 2
        page.locator('#chat-session').select_option(first)
        page.wait_for_function('document.getElementById("prompt").value.includes("未发送草稿")')
        assert page.locator('#chat-session option').count() == 1
        assert '本机保存已关闭' in page.locator('#chat-save-status').inner_text()
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        composer = page.locator('.composer').bounding_box()
        conversation = page.locator('#daily-conversations').bounding_box()
        assert conversation['y'] >= composer['y'] + composer['height'] - 1
        assert page.locator('#content-panes').bounding_box()['height'] >= conversation['height'] - 1
        page.screenshot(path=str(tmp_path / f'conversation-management-{width}.png'))
        page.evaluate("window.authenticated=false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#chat-title').input_value() == '' and page.locator('#chat-search').input_value() == ''
        assert not [value for value in calls if value == 'ask']
    finally:
        fixture.close()


def test_capacity_delete_frees_slot_and_keeps_input_ui(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        initialize(page, api, app, calls, receipts, monkeypatch)
        for _ in range(31):
            assert api.new_conversation()['ok']
        page.evaluate('window.OfficeDaily.refresh()')
        page.wait_for_function('document.getElementById("chat-new").disabled')
        assert '32/32' in page.locator('#chat-capacity').inner_text()
        page.locator('#prompt').fill('删除会话仍保留输入')
        page.locator('#chat-delete').click()
        page.wait_for_function('!document.getElementById("chat-new").disabled')
        assert '31/32' in page.locator('#chat-capacity').inner_text()
        assert page.locator('#prompt').input_value() == '删除会话仍保留输入'
        assert not [value for value in calls if value == 'ask']
    finally:
        fixture.close()


def test_write_failure_manual_retry_and_late_rename_logout(tmp_path, monkeypatch):
    from office_agent.webui import conversations
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        initialize(page, api, app, calls, receipts, monkeypatch)
        page.locator('#chat-save-history').check()
        page.wait_for_function('document.getElementById("chat-save-history").checked && !document.getElementById("chat-save-history").disabled')
        page.locator('#prompt').fill('失败后仍保留的输入')
        page.locator('#chat-title').fill('待保存标题')
        original = conversations.os.fsync
        monkeypatch.setattr(conversations.os, 'fsync', lambda *args: (_ for _ in ()).throw(OSError('private failure')))
        page.locator('#chat-rename').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("写入失败")')
        assert page.locator('#chat-title').input_value() == '待保存标题'
        assert page.locator('#prompt').input_value() == '失败后仍保留的输入'
        monkeypatch.setattr(conversations.os, 'fsync', original)
        page.locator('#chat-save-now').click()
        page.wait_for_function('document.getElementById("chat-save-status").textContent.includes("已保存到本机")')
        page.evaluate('''() => {
          const previous = window.pywebview.api;
          window.pywebview.api = new Proxy({}, {get: (_,name) => name === 'rename_conversation'
            ? (...args) => previous[name](...args).then(value => new Promise(resolve => {window.releaseRename=()=>resolve(value);}))
            : previous[name]});
        }''')
        page.locator('#chat-rename').click()
        page.wait_for_function('typeof window.releaseRename === "function"')
        assert page.locator('#chat-title').is_disabled()
        page.evaluate("window.authenticated=false; window.dispatchEvent(new Event('officeauth:unauthenticated')); window.releaseRename()")
        page.wait_for_timeout(80)
        assert page.locator('#chat-title').input_value() == ''
        assert page.locator('#prompt').input_value() == ''
        assert page.locator('#chat-save-status').inner_text() == ''
        assert 'private failure' not in page.content()
        assert not [value for value in calls if value == 'ask']
    finally:
        fixture.close()
