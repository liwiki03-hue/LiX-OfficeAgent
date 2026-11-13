"""Authenticated API integration, temporary config only, no providers/Office."""
from types import SimpleNamespace
import threading

from office_agent.webui.bridge import Api


class Auth:
    authenticated = True
    def status(self):
        return {'authenticated': self.authenticated}
    def logout(self):
        self.authenticated = False
        return {'ok': True}
    def login(self, *args):
        self.authenticated = True
        return {'ok': True, 'authenticated': True}
    def reset_accounts(self, confirmed=False):
        self.authenticated = not confirmed
        return {'ok': confirmed}


def make(tmp_path):
    app = SimpleNamespace(path=None, busy=False, _build_model=lambda: None,
                          _clear_private_state=lambda: None)
    auth = Auth()
    api = Api(app, auth=auth, settings_root=tmp_path)
    (tmp_path / '.env').write_text('MODEL_PROVIDER=ollama\nOLLAMA_MODEL=fixture\nOFFICE_CONTEXT_WINDOW=8192\n', encoding='utf-8')
    return api, app, auth


def test_authenticated_fixed_root_configuration_and_mask(tmp_path):
    api, _, auth = make(tmp_path)
    assert api.get_settings()['settings']['model'] == 'fixture'
    assert not api.plan_settings_update({'root': '../foreign'})['ok']
    plan = api.plan_settings_update({'context_window': 16384})['plan']
    saved = api.confirm_settings_update(plan['plan_id'], True)['result']
    assert saved['status'] == 'saved'
    rollback = api.plan_settings_rollback(saved['save_id'])['plan']
    assert api.confirm_settings_rollback(rollback['rollback_id'], True)['result']['restored']
    api.logout()
    for result in (api.get_settings(), api.plan_settings_update({'context_window': 8192}),
                   api.confirm_settings_update(plan['plan_id'], True),
                   api.plan_settings_rollback(saved['save_id']),
                   api.confirm_settings_rollback(rollback['rollback_id'], True)):
        assert result['reason'] == 'not_authenticated'


def test_relogin_never_restores_old_plan_or_saved_authority(tmp_path):
    api, _, _ = make(tmp_path)
    plan = api.plan_settings_update({'context_window': 16384})['plan']
    api.logout()
    api.login('fake', 'fake')
    assert api.confirm_settings_update(plan['plan_id'], True)['result']['status'] == 'rejected'
    saved = api.confirm_settings_update(api.plan_settings_update({'context_window': 16384})['plan']['plan_id'], True)['result']
    api.reset_accounts(True)
    api.login('new', 'fake')
    assert api.plan_settings_rollback(saved['save_id'])['reason'] == 'settings_plan_invalid'


def test_busy_application_and_diagram_retire_settings_plan(tmp_path):
    api, app, _ = make(tmp_path)
    for owner, field in ((app, 'busy'), (api._diagram, '_slot')):
        plan = api.plan_settings_update({'context_window': 16384})['plan']
        setattr(owner, field, True)
        assert api.confirm_settings_update(plan['plan_id'], True)['reason'] == 'busy'
        setattr(owner, field, False if field == 'busy' else None)
        assert api.confirm_settings_update(plan['plan_id'], True)['result']['status'] == 'rejected'


def test_synchronous_settings_io_serializes_logout(tmp_path, monkeypatch):
    api, _, auth = make(tmp_path)
    plan = api.plan_settings_update({'context_window': 16384})['plan']
    entered, release, logout_attempted, logout_finished = [threading.Event() for _ in range(4)]
    original = api._settings._store._replace
    def replace(*args):
        entered.set()
        assert release.wait(5)
        return original(*args)
    monkeypatch.setattr(api._settings._store, '_replace', replace)
    results = []
    saving = threading.Thread(target=lambda: results.append(api.confirm_settings_update(plan['plan_id'], True)))
    def logout():
        logout_attempted.set()
        results.append(api.logout())
        logout_finished.set()
    leaving = threading.Thread(target=logout)
    try:
        saving.start()
        assert entered.wait(5)
        leaving.start()
        assert logout_attempted.wait(5)
        assert not logout_finished.is_set()
        assert auth.authenticated
    finally:
        release.set()
        saving.join(5)
        if leaving.ident is not None:
            leaving.join(5)
    assert not saving.is_alive() and not leaving.is_alive()
    assert results[0]['result']['status'] == 'saved'
    assert not auth.authenticated
    api.login('fake', 'fake')
    assert api.plan_settings_rollback(results[0]['result']['save_id'])['reason'] == 'settings_plan_invalid'
