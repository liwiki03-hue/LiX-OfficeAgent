import json

import pytest

from test_settings_editor import _editor, FIXTURE


@pytest.fixture
def setup_editor(tmp_path, monkeypatch, request):
    yield from _editor(tmp_path, monkeypatch, initial=request.param)


@pytest.mark.parametrize('setup_editor', [None, b''], indirect=True, ids=['missing-env', 'empty-env'])
def test_first_setup_draft_test_and_save(setup_editor):
    page, facade, calls = setup_editor
    env = facade._root / '.env'
    existed = env.exists()
    assert calls == ['get_settings']
    assert page.locator('#settings-setup-toggle').get_attribute('aria-expanded') == 'true'
    assert '未指定提供方' in page.locator('#settings-setup-provider-state').inner_text()
    assert page.locator('#settings-save').is_disabled()
    defaults = facade._store.snapshot()['defaults']['deepseek']
    page.locator('#settings-setup-defaults').click()
    assert page.locator('#settings-model').input_value() == defaults['model']
    assert page.locator('#settings-url').input_value() == defaults['base_url']
    assert page.locator('#settings-context').input_value() == ''
    assert page.locator('#settings-key').input_value() == ''
    assert env.exists() is existed
    if existed:
        assert env.read_bytes() == b''
    page.locator('#settings-key').fill('sk-fixture-not-real-1234567890')
    assert 'sk-fixture' not in page.locator('#settings-setup').inner_text()
    page.locator('#settings-setup-test').click()
    assert page.locator('#settings-test').evaluate('(node) => node === document.activeElement')
    assert calls == ['get_settings']
    page.locator('#settings-test').click()
    page.wait_for_function("document.getElementById('settings-setup-test-state').textContent.includes('无法通过模型列表验证')")
    assert calls == ['get_settings', 'test_settings_connection']
    page.locator('#settings-model').fill('manual-model')
    assert '尚未检测' in page.locator('#settings-setup-test-state').inner_text()
    page.locator('#settings-setup-save').click()
    assert calls == ['get_settings', 'test_settings_connection']
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'saved_restart'")
    assert '已保存，重启后生效' in page.locator('#settings-setup-save-state').inner_text()
    assert page.locator('#settings-rollback').is_enabled()
    assert page.locator('#settings-key').input_value() == ''
    saved = facade._store._values(facade._store._read(env))
    assert saved['MODEL_PROVIDER'] == 'deepseek'
    assert saved['DEEPSEEK_MODEL'] == 'manual-model'
    assert 'OFFICE_CONTEXT_WINDOW' not in saved
    assert 'OFFICE_EMBEDDING_MODEL' not in saved
    assert 'DASHSCOPE_API_KEY' not in saved
    assert calls == ['get_settings', 'test_settings_connection', 'plan_settings_edit', 'confirm_settings_update', 'get_settings']
    assert 'sk-fixture' not in json.dumps(page.evaluate('window.modalTexts'))


@pytest.mark.parametrize('setup_editor', [b'MODEL_PROVIDER=deepseek\nDEEPSEEK_MODEL=\nDEEPSEEK_BASE_URL=\nDEEPSEEK_API_KEY=fixture-private-only\n'], indirect=True)
def test_explicit_empty_provider_drafts_and_discard(setup_editor):
    page, facade, calls = setup_editor
    original = (facade._root / '.env').read_bytes()
    assert '显式为空' in page.locator('#settings-setup-model-state').inner_text()
    assert '显式为空' in page.locator('#settings-setup-url-state').inner_text()
    page.locator('#settings-setup-defaults').click()
    alpha = page.locator('#settings-model').input_value()
    page.locator('#settings-provider').select_option('dashscope')
    page.locator('#settings-setup-defaults').click()
    assert page.locator('#settings-model').input_value() != alpha
    assert '尚未配置密钥' in page.locator('#settings-setup-key-state').inner_text()
    page.locator('#settings-provider').select_option('deepseek')
    assert page.locator('#settings-model').input_value() == alpha
    assert '磁盘密钥已配置' in page.locator('#settings-setup-key-state').inner_text()
    assert 'fixture-private' not in page.locator('#settings-setup').inner_text()
    page.locator('#settings-discard').click()
    assert page.locator('#settings-model').input_value() == ''
    assert page.locator('#settings-save').is_disabled()
    assert '显式为空' in page.locator('#settings-setup-model-state').inner_text()
    assert (facade._root / '.env').read_bytes() == original
    assert calls == ['get_settings']


@pytest.mark.parametrize('setup_editor', [FIXTURE], indirect=True)
def test_existing_settings_stay_collapsed_and_navigation_is_readonly(setup_editor):
    page, facade, calls = setup_editor
    assert page.locator('#settings-setup-toggle').get_attribute('aria-expanded') == 'false'
    assert page.locator('#settings-setup-body').is_hidden()
    page.locator('#settings-setup-toggle').click()
    assert page.locator('#settings-setup-defaults').is_disabled()
    assert page.locator('#settings-setup-provider').is_disabled()
    assert '尚未证明认证有效' in page.locator('#settings-setup-key-state').inner_text()
    page.locator('#settings-setup-context').locator('xpath=..').evaluate('(node) => node.open = true')
    page.locator('#settings-setup-context').click()
    assert page.locator('#settings-context').evaluate('(node) => node === document.activeElement')
    assert page.locator('#settings-context').input_value() == ''
    page.locator('#settings-setup-retrieval').click()
    assert page.locator('#settings-embedding-channel').evaluate('(node) => node === document.activeElement')
    assert calls == ['get_settings']
    assert (facade._root / '.env').read_bytes() == FIXTURE


@pytest.mark.parametrize('setup_editor', [FIXTURE.replace(b'MODEL_PROVIDER=deepseek\n', b'')], indirect=True)
def test_only_missing_provider_can_be_explicitly_saved(setup_editor):
    page, facade, calls = setup_editor
    original = (facade._root / '.env').read_bytes()
    assert page.locator('#settings-save').is_disabled()
    page.locator('#settings-setup-provider').click()
    assert page.locator('#settings-save').is_enabled()
    assert calls == ['get_settings']
    assert (facade._root / '.env').read_bytes() == original
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'saved_restart'")
    assert (facade._root / '.env').read_bytes() == original + b'MODEL_PROVIDER=deepseek\n'
    assert '未设置 → deepseek' in page.evaluate('window.modalTexts[0]')
    assert 'fixture-private' not in page.evaluate('window.modalTexts[0]')
    assert 'test_settings_connection' not in calls


@pytest.mark.parametrize('setup_editor', [b'MODEL_PROVIDER=deepseek\nDEEPSEEK_MODEL=custom-model\n'], indirect=True)
def test_defaults_preserve_custom_model_and_secret_logout_clears_guide(setup_editor):
    page, facade, calls = setup_editor
    original = (facade._root / '.env').read_bytes()
    page.locator('#settings-key').fill('sk-new-fixture-not-real-1234567890')
    page.locator('#settings-setup-defaults').click()
    assert page.locator('#settings-model').input_value() == 'custom-model'
    assert page.locator('#settings-key').input_value() == 'sk-new-fixture-not-real-1234567890'
    page.evaluate('window.approved = false')
    page.keyboard.press('Escape')
    assert page.locator('#settings-dialog').is_visible()
    page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert page.locator('#settings-key').input_value() == ''
    assert page.locator('#settings-setup').is_hidden()
    assert page.locator('#settings-setup-key-state').inner_text() == ''
    assert (facade._root / '.env').read_bytes() == original
    assert calls == ['get_settings']


@pytest.mark.parametrize('setup_editor', [None], indirect=True)
def test_onboarding_cancel_and_external_conflict_never_claim_saved(setup_editor):
    page, facade, calls = setup_editor
    page.locator('#settings-setup-defaults').click()
    page.locator('#settings-key').fill('sk-fixture-not-real-1234567890')
    page.evaluate('window.approved = false')
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'denied'")
    assert not (facade._root / '.env').exists()
    assert page.locator('#settings-key').input_value() == ''
    assert '未保存' in page.locator('#settings-setup-save-state').inner_text()
    (facade._root / '.env').write_bytes(b'EXTERNAL=fixture\n')
    page.evaluate('window.approved = true')
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').textContent.includes('外部修改')")
    assert (facade._root / '.env').read_bytes() == b'EXTERNAL=fixture\n'
    assert page.locator('#settings-rollback').is_disabled()
    assert '已保存，重启后生效' not in page.locator('#settings-setup-save-state').inner_text()
    assert 'test_settings_connection' not in calls


@pytest.mark.parametrize('setup_editor', [None], indirect=True)
def test_ollama_setup_needs_no_key_and_never_starts_service(setup_editor):
    page, facade, calls = setup_editor
    page.locator('#settings-provider').select_option('ollama')
    page.locator('#settings-setup-defaults').click()
    assert page.locator('#settings-setup-key').is_hidden()
    assert '无需此密钥' in page.locator('#settings-setup-key-state').inner_text()
    assert '不会启动服务或下载模型' in page.locator('#settings-setup-key-state').inner_text()
    assert page.locator('#settings-key').is_disabled()
    page.locator('#settings-save').click()
    page.wait_for_function("document.getElementById('settings-status').dataset.state === 'saved_restart'")
    values = facade._store._values(facade._store._read(facade._root / '.env'))
    assert values['MODEL_PROVIDER'] == 'ollama'
    assert not any(name.endswith('_API_KEY') for name in values)
    assert calls == ['get_settings', 'plan_settings_edit', 'confirm_settings_update', 'get_settings']


@pytest.mark.parametrize('setup_editor', [None], indirect=True)
def test_ambient_values_are_not_misrepresented_as_saved(setup_editor, monkeypatch):
    page, facade, calls = setup_editor
    monkeypatch.setenv('DEEPSEEK_MODEL', 'ambient-model-fixture')
    monkeypatch.setenv('DEEPSEEK_BASE_URL', 'https://ambient.invalid/v1')
    monkeypatch.setenv('OFFICE_CONTEXT_WINDOW', '65536')
    page.locator('#settings-load').click()
    page.wait_for_function("document.getElementById('settings-status').textContent.includes('已读取')")
    assert page.locator('#settings-model').input_value() == ''
    assert page.locator('#settings-url').input_value() == ''
    assert page.locator('#settings-context').input_value() == ''
    assert '可能采用进程值或默认值' in page.locator('#settings-setup-model-state').inner_text()
    page.locator('#settings-setup-defaults').click()
    assert page.locator('#settings-model').input_value() != 'ambient-model-fixture'
    assert page.locator('#settings-context').input_value() == ''
    assert not (facade._root / '.env').exists()
    assert calls == ['get_settings', 'get_settings']
