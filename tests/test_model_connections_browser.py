import json

import pytest

from tests.test_professional_browser import STATIC, _browser
from tests.test_workspace_layout_browser import setup
from tests.test_model_connections import connection, payload, save


def initialize(page, api, app, calls, receipts):
    setup(page, api, app, calls, receipts)
    allowed = {'get_model_connections', 'select_session_model', 'set_model_mode', 'plan_model_connections',
               'test_model_connection', 'confirm_settings_update', 'plan_settings_rollback', 'confirm_settings_rollback'}

    def invoke(method, args):
        assert method in allowed
        calls.append(method)
        result = getattr(api, method)(*args)
        receipts.append((method, result))
        return result

    page.expose_function('connectionHost', invoke)
    page.evaluate('''() => {
      const previous = window.pywebview.api;
      const names = new Set(['get_model_connections','select_session_model','set_model_mode','plan_model_connections',
        'test_model_connection','confirm_settings_update','plan_settings_rollback','confirm_settings_rollback']);
      window.pywebview.api = new Proxy({}, {get: (_, name) => names.has(name) ? (...args) => window.connectionHost(name, args) : previous[name]});
    }''')
    page.add_script_tag(path=str(STATIC / 'settings_dialog.js'))
    page.add_script_tag(path=str(STATIC / 'model_connections.js'))
    page.evaluate("window.dispatchEvent(new Event('officeauth:authenticated'))")
    page.wait_for_function("document.getElementById('connection-add').disabled === false")


def add(page, name, url, model, key):
    page.locator('#connection-add').click()
    page.locator('#connection-name').fill(name)
    page.locator('#connection-url').fill(url)
    page.locator('#connection-models').fill(model + ' | chat,agent\nimage-model | image')
    page.locator('#connection-key-action').select_option('replace')
    page.locator('#connection-key').fill(key)
    page.locator('#connection-save').click()
    page.locator('#connection-review-confirm').click()
    page.wait_for_function("document.getElementById('connection-status').textContent.includes('接口已保存')")


def test_saved_connection_feedback_immediate_selection_and_mode_isolation(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        errors = []
        page.on('pageerror', lambda exc: errors.append(str(exc)))
        initialize(page, api, app, calls, receipts)
        page.locator('#model-open-input').click()
        page.locator('#prompt').fill('保留尚未发送的中文输入')
        page.locator('#model-configure').click()
        add(page, '官方接口', 'https://official.example/v1', 'same-model', 'fixture-private-one')
        assert '模型列表尚未获取' in page.locator('#connection-feedback').inner_text()
        page.locator('#connection-feedback').get_by_role('button', name='用于当前对话：same-model', exact=True).click()
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('官方接口')")
        assert page.locator('#settings-dialog').is_hidden()
        page.locator('#model-default').click()
        page.locator('#connection-review-confirm').click()
        page.wait_for_function("document.getElementById('model-default').textContent === '当前为本用途默认'")
        assert api.get_model_connections()['catalog']['defaults']['chat']['model_id'] == 'same-model'
        page.locator('#model-configure').click()
        add(page, '自定义接口', 'https://custom.example/v1', 'same-model', 'fixture-private-two')
        page.locator('#settings-close').click()
        choices = page.locator('#model-selector option').all_text_contents()
        assert any('same-model · 官方接口' in text for text in choices)
        assert any('same-model · 自定义接口' in text for text in choices)
        page.locator('#model-search').fill('自定义')
        assert page.locator('#model-selector optgroup').count() == 1
        page.locator('#model-search').fill('')
        page.locator('#model-mode').select_option('image')
        page.wait_for_function("document.getElementById('model-selector').textContent.includes('image-model')")
        image_option = page.locator('#model-selector option').filter(has_text='image-model · 自定义接口')
        page.locator('#model-selector').select_option(image_option.get_attribute('value'))
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('image-model')")
        page.locator('#model-mode').select_option('chat')
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('same-model')")
        assert '官方接口' in page.locator('#model-selection-status').inner_text()
        assert '未验证' in page.locator('#model-selection-status').inner_text()
        assert page.locator('#prompt').input_value() == '保留尚未发送的中文输入'
        assert not [name for name in calls if name in {'ask', 'test_model_connection', 'confirm_image_generation'}]
        assert 'fixture-private-' not in json.dumps(receipts, ensure_ascii=False)
        assert not errors
    finally:
        fixture.close()


@pytest.mark.parametrize('width,theme', [(1440, 'light'), (1280, 'dark'), (1100, 'highcontrast'), (520, 'light')])
def test_model_selector_layout_and_disabled_connection_preserves_input(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        save(api._connections, payload([connection(), connection('custom')]))
        initialize(page, api, app, calls, receipts)
        page.set_viewport_size({'width': width, 'height': 800})
        page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
        page.locator('#model-mode').select_option('chat')
        page.wait_for_function("document.getElementById('model-selector').options.length === 3")
        option = page.locator('#model-selector option').filter(has_text='same-model · official')
        page.locator('#model-selector').select_option(option.get_attribute('value'))
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('当前：same-model')")
        page.locator('#model-open-input').click()
        page.locator('#prompt').fill('keep 中文 📄')
        changed = connection(); changed['enabled'] = False
        save(api._connections, payload([changed, connection('custom')]))
        page.locator('#model-refresh').click()
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('失效')")
        assert page.locator('#prompt').input_value() == 'keep 中文 📄'
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        assert page.locator('#model-selector').evaluate('node => node.getBoundingClientRect().right <= innerWidth')
        assert not page.evaluate('window.OfficeModelSelector.readyForRequest()')
        page.screenshot(path=str(tmp_path / 'selector-invalid.png'))
    finally:
        fixture.close()


def test_connection_draft_cancel_close_guard_and_logout(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        initialize(page, api, app, calls, receipts)
        page.locator('#model-configure').click()
        page.locator('#connection-add').click()
        page.locator('#connection-name').fill('unsaved')
        page.locator('#connection-key-action').select_option('replace')
        page.locator('#connection-key').fill('fixture-new-private')
        page.evaluate('window.approved = false')
        page.keyboard.press('Escape')
        assert page.locator('#settings-dialog').is_visible()
        assert page.locator('#connection-key').input_value() == 'fixture-new-private'
        assert not (tmp_path / '.env').exists()
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#connection-key').input_value() == ''
        assert not page.locator('#model-selector option').filter(has_text='unsaved').count()
        assert not [r for name, r in receipts if name == 'confirm_settings_update']
    finally:
        fixture.close()


def test_late_catalog_response_cannot_restore_logged_out_selection(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        save(api._connections, payload([connection()]))
        initialize(page, api, app, calls, receipts)
        result = api.get_model_connections()
        page.evaluate('''() => {
          window.pywebview.api = {get_model_connections: () => new Promise(resolve => {window.lateCatalog = resolve;})};
          window.OfficeModelSelector.refresh();
        }''')
        page.wait_for_function('typeof window.lateCatalog === "function"')
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate('(result) => window.lateCatalog(result)', result)
        assert page.locator('#model-selector option').count() == 1
        assert not page.evaluate('window.OfficeModelSelector.readyForRequest()')
        assert page.locator('#connection-key').input_value() == ''
    finally:
        fixture.close()


@pytest.mark.parametrize('failure', ['rejected', 'unknown', 'mode'])
def test_failed_selection_cannot_send_with_mismatched_display(tmp_path, monkeypatch, failure):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        save(api._connections, payload([connection(), connection('custom')]))
        initialize(page, api, app, calls, receipts)
        page.locator('#model-mode').select_option('chat')
        page.wait_for_function("document.getElementById('model-selector').options.length === 3")
        official = page.locator('#model-selector option').filter(has_text='same-model · official').get_attribute('value')
        page.locator('#model-selector').select_option(official)
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('当前：same-model')")
        page.locator('#model-open-input').click(); page.locator('#prompt').fill('保留输入，不向错误接口发送')
        page.evaluate('''failure => {
          const previous = window.pywebview.api;
          window.pywebview.api = new Proxy({}, {get: (_, name) => {
            if (name === (failure === 'mode' ? 'set_model_mode' : 'select_session_model')) return async () => {
              if (failure === 'unknown') throw new Error('fixture-private-failure');
              return {ok:false,reason:'model_configuration_changed'};
            };
            return previous[name];
          }});
        }''', failure)
        if failure == 'mode':
            page.locator('#model-mode').select_option('image')
        else:
            option = page.locator('#model-selector option').filter(has_text='same-model · custom')
            page.locator('#model-selector').select_option(option.get_attribute('value'))
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('切换未确认')")
        assert page.locator('#model-mode').input_value() == 'chat'
        assert page.locator('#model-selector').input_value() == official
        assert not page.evaluate('window.OfficeModelSelector.readyForRequest()')
        assert api._connections.capture('chat').connection_id == 'official'
        assert page.locator('#prompt').input_value() == '保留输入，不向错误接口发送'
        assert 'fixture-private-failure' not in page.locator('body').inner_text()
        page.locator('#model-refresh').click()
        page.wait_for_function("document.getElementById('model-selection-status').textContent.includes('当前：same-model')")
        assert page.evaluate('window.OfficeModelSelector.readyForRequest()')
    finally:
        fixture.close()
