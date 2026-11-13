import pytest

from tests.test_chat_budget_browser import prepare
from tests.test_professional_browser import _browser


LONG = '保留严格限定条件、数字与单位。' * 300
SHORT = '缩短后的中文问题，保留 20 ms。📄'


def reject(page):
    page.locator('#prompt').fill(LONG)
    page.locator('#btn-ask').click()
    page.wait_for_function('document.getElementById("status-line").textContent.includes("上下文预算")')
    assert page.locator('#prompt').input_value() == LONG
    assert '预算' in page.locator('#chat-notice').inner_text()


def hold_preview(page):
    page.evaluate('''() => {
      const previous = window.pywebview.api;
      window.previewAttempts = 0;
      window.pywebview.api = new Proxy({}, {get: (_, name) => name === 'preview_chat_context'
        ? (...args) => { window.previewAttempts++; return previous[name](...args).then(value =>
          new Promise(resolve => {window.releaseBudgetPreview = () => resolve(value);})); }
        : previous[name]});
    }''')
    page.locator('#chat-preview-context').click()
    page.wait_for_function('typeof window.releaseBudgetPreview === "function"')


@pytest.mark.parametrize('width,theme', [(1440, 'light'), (1280, 'dark'), (520, 'highcontrast')])
def test_current_preview_clears_only_recovered_budget_feedback(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        app = prepare(page, api, original, calls, receipts, monkeypatch)
        monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('preview called model'))
        page.set_viewport_size({'width': width, 'height': 800})
        page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
        reject(page)
        page.locator('#prompt').fill(SHORT)
        page.locator('#chat-preview-context').click()
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("输入估算")')
        assert '尚未调用模型' in page.locator('#chat-notice').inner_text()
        assert '尚未调用模型' in page.locator('#status-line').text_content()
        assert page.locator('#error-line').is_hidden()
        assert page.locator('#error-detail').text_content() == ''
        assert page.locator('#chat-budget-feedback').is_hidden()
        assert page.locator('#prompt').input_value() == SHORT
        assert api.get_model_connections()['catalog']['current']['chat']['valid']
        assert not app.busy and calls.count('ask') == 1
        assert 'fixture-private-' not in page.content()
        page.locator('#chat-context-preview').scroll_into_view_if_needed()
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        page.screenshot(path=str(tmp_path / f'budget-feedback-recovered-{width}.png'))
    finally:
        fixture.close()


@pytest.mark.parametrize('when', ['before', 'during'])
def test_preview_preserves_other_failure_and_uncertain_status(tmp_path, monkeypatch, when):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        prepare(page, api, original, calls, receipts, monkeypatch)
        monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('preview called model'))
        reject(page)
        page.locator('#prompt').fill(SHORT)
        update = '''() => {
          window.dispatchEvent(new CustomEvent('officedaily:event', {detail:{type:'chat_error',code:'timeout',prompt:'其他请求'}}));
          setStatus('warn', '保存结果不确定，请先核对记录，不要重复写入。');
        }'''
        if when == 'before':
            page.evaluate(update)
        hold_preview(page)
        if when == 'during':
            page.evaluate(update)
        page.evaluate('window.releaseBudgetPreview()')
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("输入估算")')
        assert '服务端可能已受理' in page.locator('#chat-notice').inner_text()
        assert '结果不确定' in page.locator('#status-line').text_content()
        assert '结果不确定' in page.locator('#error-line').inner_text()
        assert page.locator('#prompt').input_value() == SHORT
        assert page.evaluate('window.previewAttempts') == 1
        assert calls.count('ask') == 1
    finally:
        fixture.close()


def test_old_preview_does_not_clear_newer_budget_rejection(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        prepare(page, api, original, calls, receipts, monkeypatch)
        monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('budget called model'))
        reject(page)
        page.locator('#prompt').fill(SHORT)
        hold_preview(page)
        result = page.evaluate('(question) => window.OfficeDaily.send(question)', LONG)
        assert result['reason'] == 'model_context_budget'
        page.evaluate("setStatus('warn', '新一轮预算拒绝，请重新预览。', 'chat-budget')")
        page.evaluate('window.releaseBudgetPreview()')
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("输入估算")')
        assert '预算' in page.locator('#chat-notice').inner_text()
        assert page.locator('#chat-budget-feedback').is_visible()
        assert '1024' in page.locator('#chat-budget-message').inner_text()
        assert page.locator('#status-line').text_content() == '新一轮预算拒绝，请重新预览。'
        assert calls.count('ask') == 2
        assert page.locator('#prompt').input_value() == SHORT
    finally:
        fixture.close()


@pytest.mark.parametrize('change', ['input', 'mode', 'logout'])
def test_stale_recovery_does_not_replace_new_state(tmp_path, monkeypatch, change):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        prepare(page, api, original, calls, receipts, monkeypatch)
        monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('preview called model'))
        reject(page)
        page.locator('#prompt').fill(SHORT)
        hold_preview(page)
        if change == 'input':
            page.locator('#prompt').fill('更新后的输入')
        elif change == 'mode':
            page.locator('#model-mode').select_option('image')
            page.wait_for_function('window.OfficeModelSelector.mode() === "image"')
        else:
            page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate("setStatus('warn', '新的状态需核对');window.releaseBudgetPreview()")
        page.wait_for_timeout(80)
        assert page.locator('#chat-context-preview').is_hidden()
        assert page.locator('#status-line').text_content() == '新的状态需核对'
        assert calls.count('ask') == 1
    finally:
        fixture.close()
