"""Authenticated host settings facade; no model tools or client-selected roots.

Only a trusted Python Api constructor chooses a root. All methods run under its
shared RLock. Approved synchronous file IO serializes with logout; it cannot be
actively preempted. Never load config into history or os.environ here.
"""
from pathlib import Path
import uuid

from ..settings_store import SettingsError

_CODES = frozenset({
    'settings_invalid', 'settings_path_changed', 'settings_file_invalid',
    'settings_source_changed', 'settings_io_failed', 'settings_session_invalid',
    'settings_uncertain', 'settings_capacity', 'settings_id_collision',
    'settings_plan_invalid', 'settings_plan_expired', 'settings_denied',
    'confirmation_required', 'settings_verification_failed',
    'settings_no_changes',
})


class Settings:
    def __init__(self, lock, scope, root):
        self._lock, self._scope = lock, scope
        self._root = Path(root)
        self._store = None
        self._session = uuid.uuid4().hex
        self._disabled = False
        self._probing = False

    def invalidate(self):
        with self._lock:
            self._session = uuid.uuid4().hex
            if self._store is not None:
                self._store.invalidate()

    def _call(self, operation):
        with self._lock:
            try:
                authenticated, busy = self._scope()
                if not authenticated:
                    self.invalidate()
                    return {'ok': False, 'reason': 'not_authenticated'}
                if busy:
                    self.invalidate()
                    return {'ok': False, 'reason': 'busy'}
                if self._disabled:
                    return {'ok': False, 'reason': 'settings_service_disabled'}
                if self._store is None:
                    from ..settings_store import SettingsStore
                    self._store = SettingsStore(self._root)
                result = operation(self._store, self._session)
                if type(result) is not dict:
                    raise TypeError('Invalid host settings response')
                outcome = result.get('result', {})
                if (type(outcome) is dict and
                        (outcome.get('code') == 'settings_uncertain' or outcome.get('status') == 'uncertain')):
                    self._disabled = True
                return {'ok': True, **result}
            except SettingsError as exc:
                if exc.code == 'settings_uncertain':
                    self._disabled = True
                reason = exc.code if type(exc.code) is str and exc.code in _CODES else 'settings_failed'
                result = {'ok': False, 'reason': reason}
                if exc.field is not None:
                    result['field'] = exc.field
                return result
            except Exception:
                # No raw exception text or configuration is part of this channel.
                # An unclassified exception could follow IO: fail closed rather
                # than saying no write occurred. Restart does not restore tokens.
                self._disabled = True
                self.invalidate()
                return {'ok': False, 'reason': 'settings_failed', 'outcome_uncertain': True}

    def snapshot(self):
        return self._call(lambda store, session: {'settings': store.snapshot()})

    def plan_update(self, changes):
        return self._call(lambda store, session: {'plan': store.plan_update(changes, session)})

    def plan_edit(self, changes, revision):
        return self._call(lambda store, session: {'plan': store.plan_edit(changes, session, revision)})

    def test_connection(self, draft):
        from ..settings_probe import prepare_probe, probe_connection
        with self._lock:
            if self._probing:
                return {'ok': False, 'reason': 'busy'}
            prepared = None
            def prepare(store, session):
                nonlocal prepared
                prepared = prepare_probe(store, draft)
                return {'prepared': True}
            result = self._call(prepare)
            if not result['ok']:
                return result
            session = self._session
            self._probing = True
        try:
            outcome = probe_connection(prepared)
            with self._lock:
                authenticated, busy = self._scope()
                if not authenticated:
                    return {'ok': False, 'reason': 'not_authenticated'}
                if busy:
                    return {'ok': False, 'reason': 'busy'}
                if session != self._session:
                    return {'ok': False, 'reason': 'settings_session_invalid'}
                return {'ok': True, 'result': outcome}
        finally:
            prepared = None
            with self._lock:
                self._probing = False

    def confirm_update(self, plan_id, confirmed=False):
        return self._call(lambda store, session: {'result': store.confirm_update(plan_id, session, confirmed)})

    def plan_rollback(self, save_id):
        return self._call(lambda store, session: {'plan': store.plan_rollback(save_id, session)})

    def confirm_rollback(self, rollback_id, confirmed=False):
        return self._call(lambda store, session: {'result': store.confirm_rollback(rollback_id, session, confirmed)})
