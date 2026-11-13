from copy import deepcopy
import hashlib
import json
import time
import uuid

from .professional import Professional
from ..output_dependencies import OutputDependencies, file_identity
from ..writes import GuardError
from ..review_archive import FORMAT, ReviewSetArchive
from ..rag_evidence import EvidenceError

TTL = 300


class OutputReview(Professional):
    def __init__(self, lock, scope, editor, selected, dependencies, rollback, authority_check=None):
        super().__init__(lock, scope, editor, selected)
        self.dependencies = dependencies
        self.rollback_output = rollback
        self.authority_check = authority_check
        self.reviews = {}
        self.archive_backend = None
        self.archive_plans = {}

    def invalidate(self):
        super().invalidate()
        self.reviews.clear()
        self.archive_plans.clear()
        if self.archive_backend:
            self.archive_backend.invalidate()

    def _snapshot(self, editor, path, target, receipts, seeds=()):
        self.dependencies._bind(editor)
        fork = OutputDependencies()
        fork.root = self.dependencies.root
        fork.known = deepcopy(self.dependencies.known)
        for receipt in seeds:
            fork.known.setdefault(fork._id(receipt), {})
        result = fork.preview(editor, target, receipts)
        rows = fork.known
        public = {r['receipt_id']: r for r in result['records']}
        edges, outside = [], []
        for child, row in sorted(rows.items()):
            if not public[child]['readable']:
                continue
            parent = row.get('parent')
            if parent and parent not in rows:
                outside.append({'parent': parent, 'child': child})
            for candidate, other in sorted(rows.items()):
                if parent == candidate or (other.get('output') and row.get('source') == other['output']):
                    edges.append({'parent': candidate, 'child': child, 'reason': 'parent_receipt' if parent == candidate else 'source_output'})
        active = {key for key, row in public.items() if row['status'] != 'rolled_back'}
        children = {key: set() for key in rows}
        indegree = {key: 0 for key in rows}
        for edge in edges:
            if edge['child'] not in children[edge['parent']]:
                children[edge['parent']].add(edge['child'])
                indegree[edge['child']] += 1
        queue, ordered = sorted(key for key in rows if indegree[key] == 0), []
        while queue:
            current = queue.pop(0)
            ordered.append(current)
            for child in sorted(children[current]):
                indegree[child] -= 1
                if indegree[child] == 0:
                    queue.append(child)
                    queue.sort()
        cycles = sorted(set(rows) - set(ordered))
        descendants, pending = set(), list(children[target])
        while pending:
            child = pending.pop()
            if child not in descendants:
                descendants.add(child)
                pending.extend(children[child] - descendants)
        unresolved = sorted(key for key, row in public.items() if not row['readable'])
        blocking = sorted((descendants & active) - {target})
        target_row = public[target]
        blocked = bool(unresolved or cycles or blocking or not target_row['readable'] or
                       target_row['status'] != 'verified' or target_row.get('output_state') != 'present_not_revalidated')
        view = {**result, 'edges': edges, 'outside_scope_parents': outside, 'unresolved_receipts': unresolved,
                'cycle_or_cycle_dependent_receipts': cycles, 'active_descendants': blocking,
                'suggested_child_first_order': [key for key in reversed(ordered) if key in active] if not cycles and not unresolved else [],
                'review_blocked': blocked, 'file_content_revalidated': False}
        identity = {'rows': rows, 'selection': file_identity(editor, path), 'view': view}
        signature = hashlib.sha256(json.dumps(identity, sort_keys=True, separators=(',', ':')).encode('utf8')).hexdigest()
        return {'signature': signature, 'identity': identity, 'view': view, 'rows': rows}

    def _public(self, identifier, entry, changes):
        authority = False
        try:
            authority = bool(not entry.get('historical_only') and self.authority_check and self.authority_check(entry['target']))
        except GuardError:
            pass
        return {**entry['snapshot']['view'], 'review_id': identifier, 'revision': entry['revision'],
                'snapshot_sha256': entry['snapshot']['signature'], 'changes': changes,
                'review_acknowledged': entry.get('ticket') is not None, 'storage': 'authorization_context_memory_only',
                'target_has_current_undo_authority': authority,
                'historical_only': entry.get('historical_only', False),
                'expires_in_seconds': max(0, int(TTL - (time.monotonic() - entry['created']))),
                'write_approval_granted': False, 'restores_rollback_authority': False, 'wrote_nothing': True}

    def _entry(self, identifier, binding):
        entry = self.reviews.get(identifier) if type(identifier) is str else None
        if not entry or entry['binding'] != binding:
            raise GuardError('professional_review_invalid')
        if time.monotonic() - entry['created'] >= TTL:
            self.reviews.pop(identifier)
            raise GuardError('professional_review_expired')
        return entry

    def start(self, target, receipts):
        def operation(editor, path, binding):
            snapshot = self._snapshot(editor, path, target, receipts)
            self.reviews = {key: value for key, value in self.reviews.items() if value['binding'] != binding and time.monotonic() - value['created'] < TTL}
            if len(self.reviews) >= 4:
                raise GuardError('professional_review_capacity')
            identifier = uuid.uuid4().hex
            self.reviews[identifier] = {'binding': binding, 'target': target, 'receipts': list(receipts),
                                       'created': time.monotonic(), 'snapshot': snapshot, 'revision': 1, 'ticket': None}
            return self._public(identifier, self.reviews[identifier], {'added': [], 'changed': [], 'removed': []})
        return self._call(operation)

    def refresh(self, identifier):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            current = self._snapshot(editor, path, entry['target'], entry['receipts'], entry.get('seeds', ()))
            old, new = entry['snapshot']['rows'], current['rows']
            changes = {'added': sorted(new.keys() - old.keys()), 'removed': sorted(old.keys() - new.keys()),
                       'changed': sorted(key for key in old.keys() & new.keys() if old[key] != new[key])}
            if current['signature'] != entry['snapshot']['signature']:
                entry['revision'] += 1
            entry['ticket'] = None
            entry['snapshot'] = current
            return self._public(identifier, entry, changes)
        return self._call(operation)

    def _fresh(self, editor, path, entry):
        current = self._snapshot(editor, path, entry['target'], entry['receipts'], entry.get('seeds', ()))
        if current['signature'] != entry['snapshot']['signature']:
            entry['ticket'] = None
            raise GuardError('professional_review_changed')
        if current['view']['review_blocked']:
            entry['ticket'] = None
            raise GuardError('professional_review_blocked')
        return current

    def acknowledge(self, identifier, revision, reviewed=False):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            if reviewed is not True or type(revision) is not int or revision != entry['revision']:
                raise GuardError('professional_review_not_acknowledged')
            self._fresh(editor, path, entry)
            entry['ticket'] = uuid.uuid4().hex
            return {**self._public(identifier, entry, {'added': [], 'changed': [], 'removed': []}),
                    'review_token': entry['ticket'], 'review_is_write_approval': False}
        return self._call(operation)

    def cancel(self, identifier):
        def operation(editor, path, binding):
            self._entry(identifier, binding)
            self.reviews.pop(identifier)
            return {'status': 'cancelled', 'wrote_nothing': True}
        return self._call(operation)

    def rollback(self, identifier, ticket, confirmed=False):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            if entry.get('historical_only'):
                raise GuardError('professional_review_read_only')
            expected, entry['ticket'] = entry.get('ticket'), None
            if type(ticket) is not str or not expected or ticket != expected or type(confirmed) is not bool:
                raise GuardError('professional_review_not_acknowledged')
            if not confirmed:
                return {'status': 'cancelled', 'wrote_nothing': True}
            current = self._fresh(editor, path, entry)
            self.dependencies.known.update(deepcopy(current['rows']))
            reply = self.rollback_output(entry['target'], True)
            if not reply.get('ok'):
                raise GuardError(reply.get('reason', 'professional_review_invalid'))
            return reply['result']
        return self._call(operation)

    def _archive(self, editor):
        if self.archive_backend is None or self.archive_backend.root != editor.space.root:
            self.archive_backend = ReviewSetArchive(editor.space)
        return self.archive_backend

    def _archive_call(self, operation):
        def guarded(editor, path, binding):
            try:
                return operation(editor, path, binding)
            except GuardError:
                raise
            except EvidenceError as error:
                if error.code == 'evidence_uncertain':
                    self._disabled = True
                raise GuardError('professional_review_archive_' + error.code) from None
            except (ValueError, KeyError, OSError, TypeError):
                raise GuardError('professional_review_archive_invalid') from None
        return self._call(guarded)

    @staticmethod
    def _hash(value):
        return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(',', ':')).encode('utf8')).hexdigest()

    def _document(self, editor, entry, snapshot):
        public = {row['receipt_id']: row for row in snapshot['view']['records']}
        return {'format': FORMAT, 'scope': self._hash(self.dependencies.root), 'target': entry['target'],
                'receipts': sorted(snapshot['rows']),
                'stamps': {key: self._hash([row, public[key]]) for key, row in snapshot['rows'].items()}}

    def plan_archive(self, identifier, revision):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            if type(revision) is not int or revision != entry['revision']:
                raise GuardError('professional_review_changed')
            snapshot = self._snapshot(editor, path, entry['target'], entry['receipts'], entry.get('seeds', ()))
            if snapshot['signature'] != entry['snapshot']['signature']:
                raise GuardError('professional_review_changed')
            archive = self._archive(editor)
            for key, plan in list(self.archive_plans.items()):
                if plan['operation'] == 'save':
                    archive._plans.pop(key, None)
                    self.archive_plans.pop(key)
            document = self._document(editor, entry, snapshot)
            result = archive.plan_set(document, self._epoch)
            self.archive_plans = {key: value for key, value in self.archive_plans.items() if key in archive._plans}
            self.archive_plans[result['plan_id']] = {'operation': 'save', 'binding': binding,
                'review': identifier, 'signature': snapshot['signature']}
            return {**result, 'target': document['target'], 'receipts': document['receipts'],
                    'restores_approval': False, 'restores_rollback_authority': False}
        return self._archive_call(operation)

    def confirm_archive(self, identifier, approved=False):
        def operation(editor, path, binding):
            plan = self.archive_plans.pop(identifier, None) if type(identifier) is str else None
            if not plan or plan['operation'] != 'save' or plan['binding'] != binding or type(approved) is not bool:
                raise GuardError('professional_plan_invalid')
            archive = self._archive(editor)
            if approved:
                try:
                    entry = self._entry(plan['review'], binding)
                    current = self._snapshot(editor, path, entry['target'], entry['receipts'], entry.get('seeds', ()))
                    if current['signature'] != plan['signature']:
                        raise GuardError('professional_review_changed')
                except Exception:
                    archive.confirm(identifier, self._epoch, False)
                    raise
            result = archive.confirm(identifier, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            return result
        return self._archive_call(operation)

    def restore_archive(self, save_id):
        def operation(editor, path, binding):
            archive = self._archive(editor)
            document, saved = archive.inspect(save_id)
            self.dependencies._bind(editor)
            if document['scope'] != self._hash(self.dependencies.root):
                raise GuardError('professional_review_archive_scope')
            snapshot = self._snapshot(editor, path, document['target'], [], document['receipts'])
            current = self._document(editor, document, snapshot)
            old, new = document['stamps'], current['stamps']
            changes = {'added': sorted(new.keys() - old.keys()), 'removed': sorted(old.keys() - new.keys()),
                       'changed': sorted(key for key in old.keys() & new.keys() if old[key] != new[key])}
            archive._saved_check(saved)
            self.reviews = {key: value for key, value in self.reviews.items() if value['binding'] != binding and time.monotonic() - value['created'] < TTL}
            if len(self.reviews) >= 4:
                raise GuardError('professional_review_capacity')
            identifier = uuid.uuid4().hex
            entry = {'binding': binding, 'target': document['target'], 'receipts': [], 'seeds': document['receipts'],
                     'created': time.monotonic(), 'snapshot': snapshot, 'revision': 1, 'ticket': None, 'historical_only': True}
            self.reviews[identifier] = entry
            return {**self._public(identifier, entry, changes), 'archive_receipt': save_id, 'restores_approval': False}
        return self._archive_call(operation)

    def plan_remove_archive(self, save_id):
        def operation(editor, path, binding):
            archive = self._archive(editor)
            result = archive.plan_removal(save_id, self._epoch)
            self.archive_plans = {key: value for key, value in self.archive_plans.items() if key in archive._plans}
            self.archive_plans[result['rollback_id']] = {'operation': 'remove', 'binding': binding, 'save_id': save_id}
            return {**result, 'save_id': save_id, 'scope': 'review_set_only_backups_and_audits_retained'}
        return self._archive_call(operation)

    def confirm_remove_archive(self, identifier, approved=False):
        def operation(editor, path, binding):
            plan = self.archive_plans.pop(identifier, None) if type(identifier) is str else None
            if not plan or plan['operation'] != 'remove' or plan['binding'] != binding or type(approved) is not bool:
                raise GuardError('professional_plan_invalid')
            archive = self._archive(editor)
            if approved:
                try:
                    archive.inspect(plan['save_id'])
                except Exception:
                    archive.confirm_rollback(identifier, self._epoch, False)
                    raise
            result = archive.confirm_rollback(identifier, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            return result
        return self._archive_call(operation)
