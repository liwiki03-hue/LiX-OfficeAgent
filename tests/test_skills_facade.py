"""Authentication/epoch/import protocol tests; no native dialogs or real files."""
from threading import RLock

import pytest

from office_agent.webui.skills import Skills


class KnownError(ValueError):
    def __init__(self, code):
        self.code = code


class Backend:
    error_type = KnownError
    error_codes = frozenset({'skills_invalid', 'skills_uncertain'})
    def __init__(self):
        self.calls = []
        self.failure = None
    def invalidate(self):
        self.calls.append(('invalidate',))
    def _run(self, name, *args):
        self.calls.append((name, *args))
        if self.failure:
            raise self.failure
        return {'status': 'saved', 'applied': True}
    def catalog(self):
        self._run('catalog')
        return []
    def read(self, name, source):
        return self._run('read', name, source)
    def plan_change(self, operation, changes, session):
        return self._run('plan_change', operation, changes, session)
    def plan_import(self, directory, session):
        return self._run('plan_import', directory, session)
    def confirm_change(self, plan, session, confirmed):
        return self._run('confirm_change', plan, session, confirmed)
    def plan_rollback(self, save_id, session):
        return self._run('plan_rollback', save_id, session)
    def confirm_rollback(self, plan, session, confirmed):
        return self._run('confirm_rollback', plan, session, confirmed)


def make(*, picker=None):
    scope = [True, False]
    backend = Backend()
    facade = Skills(RLock(), lambda: tuple(scope), lambda: backend, import_picker=picker)
    return scope, backend, facade


@pytest.mark.parametrize('method,args', [('catalog',()), ('read',('valid','user')), ('plan_change',('create',{})),
    ('plan_import',()), ('confirm_change',('id',True)), ('plan_rollback',('id',)), ('confirm_rollback',('id',True))])
def test_authentication_covers_every_operation_before_factory_or_picker(method, args):
    calls = []
    facade = Skills(RLock(), lambda: (False,False), lambda: calls.append('factory'),
                    import_picker=lambda: calls.append('picker'))
    assert getattr(facade,method)(*args) == {'ok':False,'reason':'not_authenticated'}
    assert calls == []


def test_busy_invalidates_old_authority_and_relogin_changes_epoch():
    scope, backend, facade = make()
    assert facade.plan_change('create',{})['ok']
    old = backend.calls[-1][-1]
    scope[1] = True
    assert facade.confirm_change('id',True)['reason'] == 'busy'
    assert backend.calls[-1] == ('invalidate',)
    scope[:] = [False,False]
    assert facade.catalog()['reason'] == 'not_authenticated'
    scope[0] = True
    assert facade.plan_change('create',{})['ok']
    assert backend.calls[-1][-1] != old


def test_import_picker_is_host_only_and_cancellation_does_not_plan():
    scope, backend, facade = make(picker=lambda: None)
    assert facade.plan_import() == {'ok':True,'cancelled':True}
    assert backend.calls == []
    with pytest.raises(TypeError):
        facade.plan_import('client-selected-path')
    scope, backend, facade = make(picker=lambda: 'host-selected-directory')
    assert facade.plan_import()['ok']
    assert backend.calls[-1][0:2] == ('plan_import','host-selected-directory')


@pytest.mark.parametrize('confirmed',[False, True, 1, 'true', None])
def test_facade_preserves_literal_confirmation_for_core_validation(confirmed):
    scope, backend, facade = make()
    assert facade.confirm_change('plan',confirmed)['ok']
    assert backend.calls[-1][-1] is confirmed
    assert facade.confirm_rollback('rollback',confirmed)['ok']
    assert backend.calls[-1][-1] is confirmed


def test_error_redaction_and_fuse_survives_epoch_change():
    scope, backend, facade = make()
    backend.failure = RuntimeError('temporary-fixture-secret')
    result = facade.catalog()
    assert result == {'ok':False,'reason':'skills_failed','outcome_uncertain':True}
    assert 'temporary-fixture-secret' not in str(result)
    backend.failure = None
    facade.invalidate()
    assert facade.catalog()['reason'] == 'skills_service_disabled'


def test_known_validation_failure_does_not_disable_service():
    scope, backend, facade = make()
    backend.failure = KnownError('skills_invalid')
    assert facade.catalog() == {'ok':False,'reason':'skills_invalid'}
    backend.failure = None
    assert facade.catalog() == {'ok':True,'skills':[]}


def test_unknown_typed_code_never_escapes():
    scope, backend, facade = make()
    backend.failure = KnownError('temporary-fixture-secret')
    assert facade.catalog()['reason'] == 'skills_failed'


def test_reentrant_native_picker_retires_import_before_plan():
    scope, backend, facade = make()
    def picker():
        facade.invalidate()
        return 'host-selected-directory'
    facade._picker = picker
    assert facade.plan_import() == {'ok':True,'cancelled':True,'reason':'skills_scope_changed'}
    assert backend.calls == [('invalidate',)]


def test_invalidation_recreates_backend_for_current_host_workspace():
    scope = [True,False]
    first, second = Backend(), Backend()
    current = [first]
    facade = Skills(RLock(), lambda: tuple(scope), lambda: current[0])
    assert facade.catalog()['ok']
    current[0] = second
    facade.invalidate()
    assert facade.catalog()['ok']
    assert first.calls == [('catalog',),('invalidate',)]
    assert second.calls == [('catalog',)]


def test_reported_uncertain_outcome_disables_without_claiming_success():
    scope, backend, facade = make()
    backend.confirm_change = lambda *args: {'status':'uncertain','applied':None}
    result = facade.confirm_change('id',True)
    assert result['ok'] is True and result['result']['applied'] is None
    assert facade.catalog()['reason'] == 'skills_service_disabled'
