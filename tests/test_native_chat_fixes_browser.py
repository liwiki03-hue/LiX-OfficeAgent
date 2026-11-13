import json
import threading
from types import SimpleNamespace

import pytest

from tests.test_chat_context_browser import initialize as chat_initialize
from tests.test_model_connections_browser import initialize
from tests.test_model_connections import connection, payload, save
from tests.test_professional_browser import _browser


@pytest.mark.parametrize('reason,expected', [
    ('stop','本轮回答已返回'),
    ('length','尚未完整'),
    ('unknown','完整性未确认'),
    ('content_filter','服务过滤'),
    ('local_limit','达到本地'),
    ('stopped','已停止本地显示'),
    ('timeout','等待超时'),
])
def test_terminal_stream_notice_and_late_events(tmp_path, monkeypatch, reason, expected):
    fixture = _browser(tmp_path, monkeypatch)
    release = threading.Event()
    try:
        page, api, original, calls, receipts = next(fixture)
        app = chat_initialize(page, api, original, calls, receipts, monkeypatch)
        def stream(messages):
            yield SimpleNamespace(content='已接收的合成片段。' * 30, response_metadata={})
            assert release.wait(10)
            if reason == 'timeout':
                raise TimeoutError('fixture-private-one must not be shown')
            yield SimpleNamespace(content='尾段' if reason != 'local_limit' else '限' * 32001,
                                  response_metadata={'finish_reason':reason})
        monkeypatch.setattr(api._connections, 'build_model', lambda route:SimpleNamespace(stream=stream))
        page.locator('#prompt').fill('仅验证流式终态')
        page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("正在接收")')
        if reason == 'stopped':
            page.locator('#btn-stop').click()
        release.set()
        page.wait_for_function('(text)=>document.getElementById("chat-notice").textContent.includes(text)', arg=expected)
        page.wait_for_function('!document.getElementById("prompt").disabled')
        turn = api.conversation_snapshot()['result']['sessions'][0]['turns'][0]
        before = page.locator('#chat-notice').inner_text()
        page.evaluate('''turn => {
          for (const event of [
            {type:'answer_draft_start',stream_id:turn.id},
            {type:'answer_draft_delta',stream_id:turn.id,text:'旧响应不能恢复'},
            {type:'chat_error',turn_id:turn.id,code:'authentication_failed',prompt:'旧请求'}
          ]) window.dispatchEvent(new CustomEvent('officedaily:event',{detail:event}));
        }''', turn)
        assert page.locator('#chat-notice').inner_text() == before
        assert page.locator('#chat-live-answer').is_hidden()
        assert 'fixture-private-' not in page.locator('#chat-transcript').inner_text()
        assert len([name for name in calls if name == 'ask']) == 1
        assert not app.busy
    finally:
        release.set()
        fixture.close()


def test_old_success_snapshot_keeps_new_error(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    release = threading.Event()
    try:
        page, api, original, calls, receipts = next(fixture)
        chat_initialize(page, api, original, calls, receipts, monkeypatch)
        def stream(messages):
            yield SimpleNamespace(content='合成片段' * 40, response_metadata={})
            assert release.wait(10)
            yield SimpleNamespace(content='尾段', response_metadata={'finish_reason':'stop'})
        monkeypatch.setattr(api._connections, 'build_model', lambda route:SimpleNamespace(stream=stream))
        page.locator('#prompt').fill('原请求'); page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("正在接收")')
        page.evaluate('window.dispatchEvent(new CustomEvent("officedaily:event",{detail:{type:"chat_error",code:"authentication_failed",prompt:"新的错误"}}))')
        release.set()
        page.wait_for_function('!document.getElementById("prompt").disabled')
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("尾段")')
        assert '鉴权失败' in page.locator('#chat-notice').inner_text()
    finally:
        release.set()
        fixture.close()


@pytest.mark.parametrize('width,theme', [(1440,'light'),(1280,'dark'),(1100,'highcontrast'),(520,'dark')])
def test_complete_configuration_review_and_output_field_round_trip(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        row = connection()
        row['models'] = [{'id':'model-' + str(i), 'capabilities':['chat','agent'], 'output_parameter':'max_tokens'} for i in range(64)]
        save(api._connections, payload([row]))
        before = (tmp_path / '.env').read_bytes()
        initialize(page, api, app, calls, receipts)
        page.set_viewport_size({'width':width,'height':800})
        page.evaluate('(theme)=>document.documentElement.dataset.theme=theme', theme)
        page.locator('#model-configure').click()
        page.locator('#connection-list').get_by_role('button',name='编辑',exact=True).click()
        assert page.locator('#connection-models').input_value().count('max_tokens') == 64
        page.locator('#connection-name').fill('明确的连接名称')
        page.locator('#connection-models').fill(page.locator('#connection-models').input_value().replace('max_tokens','max_completion_tokens'))
        page.locator('#connection-save').click()
        dialog = page.get_by_role('dialog',name='确认接口配置改动',exact=True)
        dialog.wait_for()
        content = page.locator('#connection-review-content')
        assert 'model-63' in content.inner_text() and 'max_completion_tokens' in content.inner_text()
        assert '密钥处理：保持' in content.inner_text()
        assert 'fixture-private-' not in dialog.inner_text()
        assert (tmp_path / '.env').read_bytes() == before
        assert not [name for name in calls if name == 'confirm_settings_update']
        assert page.evaluate('document.activeElement.id') == 'connection-review-cancel'
        page.keyboard.press('Shift+Tab')
        assert dialog.evaluate('(node)=>node.contains(document.activeElement)')
        page.keyboard.press('Tab')
        content.evaluate('(node)=>node.scrollTop=node.scrollHeight')
        assert content.evaluate('(node)=>node.scrollHeight > node.clientHeight && node.scrollTop > 0')
        assert dialog.evaluate('(node)=>node.getBoundingClientRect().width <= innerWidth && node.getBoundingClientRect().bottom <= innerHeight')
        page.screenshot(path=str(tmp_path / f'connection-review-{width}-{theme}.png'))
        page.keyboard.press('Escape')
        page.wait_for_function('!document.getElementById("connection-save-review")')
        assert page.locator('#settings-dialog').is_visible()
        assert (tmp_path / '.env').read_bytes() == before
        assert page.evaluate('document.activeElement.id') == 'connection-save'
        page.locator('#connection-save').click()
        page.locator('#connection-review-confirm').click()
        page.wait_for_function('document.getElementById("connection-status").textContent.includes("接口已保存")')
        restored = api.get_model_connections()['catalog']['connections'][0]['models']
        assert len(restored) == 64 and all(model['output_parameter'] == 'max_completion_tokens' for model in restored)
        page.locator('#connection-list').get_by_role('button',name='编辑',exact=True).click()
        assert page.locator('#connection-models').input_value().count('max_completion_tokens') == 64
        page.locator('#connection-name').fill('只修改名称仍保留输出字段')
        page.locator('#connection-save').click(); page.locator('#connection-review-confirm').click()
        page.wait_for_function('document.getElementById("connection-status").textContent.includes("接口已保存")')
        assert all(model['output_parameter'] == 'max_completion_tokens' for model in api.get_model_connections()['catalog']['connections'][0]['models'])
        assert not [name for name in calls if name in {'ask','test_model_connection','confirm_image_generation'}]
    finally:
        fixture.close()


def test_logout_cancels_configuration_review_without_writing(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        before = (tmp_path / '.env').read_bytes()
        initialize(page, api, app, calls, receipts)
        page.locator('#model-configure').click()
        page.locator('#connection-list').get_by_role('button',name='编辑',exact=True).click()
        page.locator('#connection-name').fill('退出后不得保存')
        page.locator('#connection-save').click()
        page.locator('#connection-review-confirm').wait_for()
        page.evaluate('window.authenticated=false;window.dispatchEvent(new Event("officeauth:unauthenticated"))')
        page.wait_for_function('!document.getElementById("connection-save-review")')
        assert (tmp_path / '.env').read_bytes() == before
        assert not [name for name in calls if name == 'confirm_settings_update']
        assert page.locator('#connection-key').input_value() == ''
    finally:
        fixture.close()
