"""Host-only facade tests. Temporary credentials only; no actual project .env."""
import json
import threading

import pytest

from office_agent.webui.settings import Settings


def setup_facade(tmp_path):
    (tmp_path / '.env').write_text(
        'MODEL_PROVIDER=deepseek\nDEEPSEEK_MODEL=fixture-model\n'
        'DEEPSEEK_API_KEY=FixtureSecretAlpha\nDEEPSEEK_BASE_URL=https://example.invalid/v1\n'
        'OFFICE_CONTEXT_WINDOW=8192\n', encoding='utf-8')
    state = {'authenticated': True, 'busy': False}
    facade = Settings(threading.RLock(), lambda: (state['authenticated'], state['busy']), tmp_path)
    return facade, state


def assert_no_secret(value):
    text = json.dumps(value)
    assert 'FixtureSecretAlpha' not in text
    assert 'ReplacementSecretBeta' not in text


def test_masked_read_plan_separate_save_and_rollback(tmp_path):
    facade, _ = setup_facade(tmp_path)
    before = (tmp_path / '.env').read_bytes()
    snapshot = facade.snapshot()
    assert snapshot['ok']
    assert snapshot['settings']['api_key'] == '********'
    plan = facade.plan_update({'context_window': 16384, 'api_key': 'ReplacementSecretBeta'})
    assert plan['ok'] and plan['plan']['wrote_nothing']
    assert (tmp_path / '.env').read_bytes() == before
    saved = facade.confirm_update(plan['plan']['plan_id'], True)
    assert saved['result']['status'] == 'saved'
    assert b'ReplacementSecretBeta' in (tmp_path / '.env').read_bytes()
    rollback = facade.plan_rollback(saved['result']['save_id'])
    assert rollback['ok'] and rollback['plan']['confirmation_required']
    restored = facade.confirm_rollback(rollback['plan']['rollback_id'], True)
    assert restored['result']['restored'] is True
    assert (tmp_path / '.env').read_bytes() == before
    for result in (snapshot, plan, saved, rollback, restored):
        assert_no_secret(result)


@pytest.mark.parametrize('approval', [False, 1, 'true', None, {}], ids=['deny','int','text','none','dict'])
def test_nonliteral_approval_is_consumed_not_latched(tmp_path, approval):
    facade, _ = setup_facade(tmp_path)
    before = (tmp_path / '.env').read_bytes()
    plan_id = facade.plan_update({'context_window': 16384})['plan']['plan_id']
    result = facade.confirm_update(plan_id, approval)
    assert result['result']['status'] == 'rejected'
    assert result['result']['wrote_nothing']
    assert facade.confirm_update(plan_id, True)['result']['status'] == 'rejected'
    assert (tmp_path / '.env').read_bytes() == before
    assert facade.snapshot()['ok']


@pytest.mark.parametrize('fence', ['authenticated', 'busy'])
def test_auth_and_busy_retire_plans(tmp_path, fence):
    facade, state = setup_facade(tmp_path)
    plan_id = facade.plan_update({'context_window': 16384})['plan']['plan_id']
    state[fence] = fence == 'busy'
    denied = facade.confirm_update(plan_id, True)
    assert not denied['ok']
    state[fence] = fence == 'authenticated'
    assert facade.confirm_update(plan_id, True)['result']['status'] == 'rejected'


def test_invalid_input_is_known_error_without_permanent_disable(tmp_path):
    facade, _ = setup_facade(tmp_path)
    result = facade.plan_update({'api_key': 'bad\nReplacementSecretBeta'})
    assert result == {'ok': False, 'reason': 'settings_invalid'}
    assert_no_secret(result)
    assert facade.plan_update({'context_window': 16384})['ok']


def test_invalidation_retires_history_authority(tmp_path):
    facade, _ = setup_facade(tmp_path)
    plan = facade.plan_update({'context_window': 16384})
    saved = facade.confirm_update(plan['plan']['plan_id'], True)
    facade.invalidate()
    assert facade.plan_rollback(saved['result']['save_id']) == {'ok': False, 'reason': 'settings_plan_invalid'}


def test_uncertain_after_admitted_io_survives_invalidation(tmp_path, monkeypatch):
    facade, _ = setup_facade(tmp_path)
    plan = facade.plan_update({'context_window': 16384})
    def fail(*args):
        raise OSError('ReplacementSecretBeta')
    monkeypatch.setattr(facade._store, '_replace', fail)
    result = facade.confirm_update(plan['plan']['plan_id'], True)
    assert result['result']['status'] == 'uncertain'
    assert result['result']['applied'] is None
    assert_no_secret(result)
    facade.invalidate()
    assert facade.snapshot() == {'ok': False, 'reason': 'settings_service_disabled'}


def test_unknown_exception_is_redacted_and_sticky(tmp_path, monkeypatch):
    facade, _ = setup_facade(tmp_path)
    assert facade.snapshot()['ok']
    def fail():
        raise RuntimeError('FixtureSecretAlpha')
    monkeypatch.setattr(facade._store, 'snapshot', fail)
    result = facade.snapshot()
    assert result == {'ok': False, 'reason': 'settings_failed', 'outcome_uncertain': True}
    assert_no_secret(result)
    facade.invalidate()
    assert facade.snapshot()['reason'] == 'settings_service_disabled'
