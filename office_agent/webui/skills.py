"""Authenticated skill management protocol, independent of filesystem backend.

Host supplies a lazy backend factory and optional native directory picker. Client
never supplies roots/import paths/session authority. All operations serialize with
logout. Unclassified failures disable management across account epochs.
"""
import uuid


class Skills:
    def __init__(self, lock, scope, factory, *, import_picker=None, workspace_available=lambda: True):
        self._lock, self._scope, self._factory = lock, scope, factory
        self._picker = import_picker
        self._workspace_available = workspace_available
        self._store = None
        self._session = uuid.uuid4().hex
        self._disabled = False

    def invalidate(self):
        with self._lock:
            self._session = uuid.uuid4().hex
            if self._store is not None:
                self._store.invalidate()
                self._store = None

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
                if not self._workspace_available():
                    self.invalidate()
                    return {'ok': False, 'reason': 'no_file'}
                if self._disabled:
                    return {'ok': False, 'reason': 'skills_service_disabled'}
                if self._store is None:
                    self._store = self._factory()
                result = operation(self._store, self._session)
                if type(result) is not dict:
                    raise TypeError('Invalid skill backend envelope')
                outcome = result.get('result', {})
                if type(outcome) is dict and outcome.get('status') == 'uncertain':
                    self._disabled = True
                return {'ok': True, **result}
            except Exception as exc:
                # A backend exposes only its exact typed safe code allowlist.
                # Never echo arbitrary ValueError/string messages into UI/history.
                known = getattr(self._store, 'error_type', None)
                codes = getattr(self._store, 'error_codes', frozenset())
                if (isinstance(known, type) and isinstance(exc, known) and
                        type(getattr(exc, 'code', None)) is str and exc.code in codes):
                    if exc.code == 'skills_uncertain':
                        self._disabled = True
                    return {'ok': False, 'reason': exc.code}
                self._disabled = True
                self._session = uuid.uuid4().hex
                try:
                    if self._store is not None:
                        self._store.invalidate()
                except Exception:
                    pass
                return {'ok': False, 'reason': 'skills_failed', 'outcome_uncertain': True}

    def catalog(self):
        def read(store, session):
            result = {'skills': store.catalog()}
            if callable(getattr(store, 'agent_catalog_status', None)):
                result['agent_catalog'] = store.agent_catalog_status()
            return result
        return self._call(read)

    def plan_public(self, candidate):
        return self._call(lambda store, session: {'plan':store.plan_public(candidate, session)})

    def public_snapshot(self, name, source):
        return self._call(lambda store, session: {'snapshot':{**store.public_snapshot(name, source), '_session':session}})

    def validate_public_snapshot(self, snapshot):
        def check(store, session):
            if snapshot.get('_session') != session:
                return {'ok':False, 'reason':'skills_scope_changed'}
            store.validate_public_snapshot(snapshot)
            return {'valid':True}
        return self._call(check)

    def plan_public_update(self, candidate):
        def plan(store, session):
            snapshot = candidate.get('_update_snapshot', {})
            if snapshot.get('_session') != session:
                return {'ok':False, 'reason':'skills_scope_changed'}
            return {'plan':store.plan_public_update(snapshot, candidate, session)}
        return self._call(plan)

    def read(self, name, source):
        return self._call(lambda store, session: {'skill': store.read(name, source)})

    def plan_change(self, operation, changes):
        return self._call(lambda store, session: {'plan': store.plan_change(operation, changes, session)})

    def plan_import(self):
        def plan(store, session):
            if self._picker is None:
                return {'cancelled': True}
            directory = self._picker()
            authenticated, busy = self._scope()
            if self._session != session or not authenticated or busy:
                return {'cancelled': True, 'reason': 'skills_scope_changed'}
            if directory is None:
                return {'cancelled': True}
            return {'plan': store.plan_import(directory, session)}
        return self._call(plan)

    def confirm_change(self, plan_id, confirmed=False):
        return self._call(lambda store, session: {'result': store.confirm_change(plan_id, session, confirmed)})

    def plan_rollback(self, save_id):
        return self._call(lambda store, session: {'plan': store.plan_rollback(save_id, session)})

    def confirm_rollback(self, rollback_id, confirmed=False):
        return self._call(lambda store, session: {'result': store.confirm_rollback(rollback_id, session, confirmed)})
