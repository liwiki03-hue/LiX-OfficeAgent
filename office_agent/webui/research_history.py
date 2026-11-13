import json

from .professional import Professional
from ..range_history import FORMAT, STATES, RangeHistoryArchive
from ..rag_evidence import EvidenceError
from ..writes import GuardError


class ResearchHistory(Professional):
    def __init__(self, lock, scope, editor, selected, attempt_reader, embedding_reader):
        super().__init__(lock, scope, editor, selected)
        self.attempt_reader = attempt_reader
        self.embedding_reader = embedding_reader
        self._archive_backend = None
        self._history_plans = {}

    def invalidate(self):
        super().invalidate()
        self._history_plans.clear()
        if self._archive_backend:
            self._archive_backend.invalidate()

    def _archive(self, editor):
        if self._archive_backend is None or self._archive_backend.root != editor.space.root:
            self._archive_backend = RangeHistoryArchive(editor.space)
        return self._archive_backend

    def _history_call(self, operation):
        def guarded(editor, path, binding):
            try:
                return operation(editor, path, binding)
            except GuardError:
                raise
            except EvidenceError as error:
                if error.code == 'evidence_uncertain':
                    self._disabled = True
                raise GuardError('professional_history_' + error.code) from None
            except (ValueError, KeyError, OSError, TypeError):
                raise GuardError('professional_history_receipt_invalid') from None
        return self._call(guarded)

    def _attempt(self, name):
        reply = self.attempt_reader(name)
        if not reply.get('ok'):
            raise GuardError('professional_history_attempt_unavailable')
        attempt = reply['ranges'].get('attempt')
        if not attempt:
            raise GuardError('professional_history_no_attempt')
        return json.loads(json.dumps(attempt))

    def plan_attempt(self, name):
        def operation(editor, path, binding):
            document = {'format': FORMAT, 'source_name': name, 'attempt': self._attempt(name)}
            if document['attempt'].get('state') not in STATES:
                raise GuardError('professional_history_not_terminal')
            archive = self._archive(editor)
            for token, entry in list(self._history_plans.items()):
                if entry['operation'] == 'save':
                    archive._plans.pop(token, None)
                    self._history_plans.pop(token)
            plan = archive.plan_attempt(document, self._epoch)
            self._history_plans = {k: v for k, v in self._history_plans.items() if k in archive._plans}
            self._history_plans[plan['plan_id']] = {'operation': 'save', 'binding': binding, 'document': document}
            return {**plan, 'record': document, 'restores_index': False}
        return self._history_call(operation)

    def confirm_attempt(self, plan_id, approved=False):
        def operation(editor, path, binding):
            entry = self._history_plans.pop(plan_id, None) if type(plan_id) is str else None
            if not entry or entry['operation'] != 'save' or entry['binding'] != binding or type(approved) is not bool:
                raise GuardError('professional_plan_invalid')
            archive = self._archive(editor)
            if approved:
                try:
                    if self._attempt(entry['document']['source_name']) != entry['document']['attempt']:
                        raise GuardError('professional_history_attempt_changed')
                except Exception:
                    archive.confirm(plan_id, self._epoch, False)
                    raise
            result = archive.confirm(plan_id, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            return result
        return self._history_call(operation)

    def restore_attempt(self, save_id):
        def operation(editor, path, binding):
            from ..research_pdf import PdfResearchSource
            from ..research_ranges import processing_stamp, range_scope
            archive = self._archive(editor)
            document, saved = archive.inspect(save_id)
            attempt = document['attempt']
            source_state = 'unavailable_or_refused'
            try:
                data, _, fingerprint = PdfResearchSource(editor.space)._snapshot(document['source_name'])
                if data.startswith(b'%PDF-'):
                    source_state = 'matches' if fingerprint == attempt['source_fingerprint'] and range_scope(editor.space, document['source_name']) == attempt['scope'] else 'changed_or_moved'
            except Exception:
                pass
            embedding = None
            try:
                stamp = self.embedding_reader()
                if stamp is not None:
                    embedding = stamp == attempt['embedding_stamp']
            except Exception:
                pass
            archive._saved_check(saved)
            return {'save_id': save_id, 'record': document, 'source_state': source_state,
                    'processing_current': processing_stamp() == attempt['processing_stamp'], 'embedding_current': embedding,
                    'historical_only': True, 'restores_index': False, 'restores_attempt': False,
                    'restores_approval': False, 'restores_rollback_authority': False, 'wrote_nothing': True}
        return self._history_call(operation)

    def plan_remove_attempt(self, save_id):
        def operation(editor, path, binding):
            plan = self._archive(editor).plan_removal(save_id, self._epoch)
            self._history_plans = {k: v for k, v in self._history_plans.items() if k in self._archive_backend._plans}
            self._history_plans[plan['rollback_id']] = {'operation': 'remove', 'binding': binding, 'save_id': save_id}
            return {**plan, 'save_id': save_id, 'scope': 'attempt_artifact_only_backups_and_audits_retained'}
        return self._history_call(operation)

    def confirm_remove_attempt(self, plan_id, approved=False):
        def operation(editor, path, binding):
            entry = self._history_plans.pop(plan_id, None) if type(plan_id) is str else None
            if not entry or entry['operation'] != 'remove' or entry['binding'] != binding or type(approved) is not bool:
                raise GuardError('professional_plan_invalid')
            archive = self._archive(editor)
            if approved:
                try:
                    archive.inspect(entry['save_id'])
                except Exception:
                    archive.confirm_rollback(plan_id, self._epoch, False)
                    raise
            result = archive.confirm_rollback(plan_id, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            return result
        return self._history_call(operation)
