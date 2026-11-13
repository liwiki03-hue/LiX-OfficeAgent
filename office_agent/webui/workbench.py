from copy import deepcopy
import hashlib
import json
import re
import time
import uuid

from .professional import Professional
from ..project_library import ProjectArchive, admitted_record, canonical_csl, clean, fail, filtered, groups, import_rows, parse_csl, record_from_csl, validate, paper_families
from ..literature import csl, payload_hash
from ..rag_evidence import EvidenceError, _dir
from ..rag_sources import SourceExtractor, SourceError
from ..writes import GuardError, digest


class Workbench(Professional):
    def __init__(self, lock, scope, editor, selected, structure, literature, grounded, assist, rag):
        super().__init__(lock, scope, editor, selected)
        self.structure, self.literature, self.grounded, self.assist, self.rag = structure, literature, grounded, assist, rag
        self.project = None
        self.archive = None
        self.receipt = ''
        self.revision = 0
        self.dirty = False
        self.target = None
        self.evidence = {}
        self.captures = {}
        self.capture_windows = {}
        self.plans = {}
        self.history_views = {}
        self.assist_tasks = {}
        self.assist_session = object()
        self.live_drafts = {}
        self.search_rounds = {}
        self.approved_evidence = ''
        self.document_token = None
        self.structure.project_validator = self._validate_write
        self.structure.layout_observer = self.layout_changed
        self.last_numeric = None
        self.numeric_record_id = None
        self.structure.numeric_observer = self.observe_numeric

    def invalidate(self):
        self._retire_assist()
        if self.archive:
            self.archive.invalidate()
        self.project = self.archive = self.target = None
        self.receipt = ''
        self.revision = 0
        self.dirty = False
        self.evidence.clear()
        self.captures.clear()
        self.capture_windows.clear()
        self.plans.clear()
        self.history_views.clear()
        self.live_drafts.clear()
        self.search_rounds.clear()
        self.approved_evidence = ''
        self.document_token = None
        self.last_numeric = None
        self.numeric_record_id = None
        super().invalidate()

    def _discard_task_plans(self):
        for identifier, entry in list(self.plans.items()):
            if entry.get('task'):
                self.plans.pop(identifier, None)
                self.assist._task_plans.pop(identifier, None)
                if self.assist._task_archive:
                    self.assist._task_archive._plans.pop(identifier, None)

    def _retire_assist(self):
        for identifier in self.assist_tasks:
            job = self.assist.jobs.get(identifier)
            if job and job['status'] == 'running':
                job['cancel'].set()
                job.update(status='cancelled', stage='stopped', stop_reason='workbench_context_changed', result=None)
            self.literature.unpin(identifier)
            self.rag._research_dependencies.pop(identifier, None)
        self.assist_tasks.clear()
        self.assist_session = object()
        self._discard_task_plans()

    def _ensure(self, editor, path):
        scope = payload_hash([str(editor.space.root), _dir(editor.space.root)])
        if self.project is None or self.project['scope'] != scope:
            self.invalidate()
            self.project = {'format': 'research-project-v1', 'id': uuid.uuid4().hex, 'name': '研究项目', 'scope': scope,
                'generation': 0, 'parent_receipt': '', 'entries': {}, 'canonical': {}, 'claims': [], 'tasks': [], 'outputs': [], 'layouts': [], 'drafts': [], 'evidence_records': []}
            self.archive = ProjectArchive(editor.space)
        token = (path, id(editor), digest(path))
        if self.document_token != token:
            self._retire_assist()
            self.target = None
            self.last_numeric = None
            self.numeric_record_id = None
            self.evidence.clear()
            self.captures.clear()
            self.capture_windows.clear()
            self.live_drafts.clear()
            self.search_rounds.clear()
            self.approved_evidence = ''
            self.document_token = token

    def _call(self, operation):
        from ..rag import RagError
        from ..vectorstore import VectorStoreError
        def guarded(editor, path, binding):
            self._ensure(editor, path)
            try:
                return operation(editor, path, binding)
            except EvidenceError as error:
                fail('professional_project_' + error.code)
            except SourceError:
                fail('professional_project_source_refused')
            except VectorStoreError:
                fail('professional_project_database_integrity')
            except RagError as error:
                fail('professional_project_' + error.code)
        return super()._call(guarded)

    def _changed(self):
        self._discard_task_plans()
        self.history_views.clear()
        for identifier, entry in list(self.plans.items()):
            if 'document' in entry or 'removal' in entry:
                self.plans.pop(identifier, None)
                if self.archive:
                    self.archive._plans.pop(identifier, None)
        self.revision += 1
        self.dirty = True

    def _entry(self, key):
        if type(key) is not str or key not in self.project['entries']:
            fail()
        return self.project['entries'][key]

    def _snapshot(self, filters=None, offset=0):
        if type(offset) is not int or not 0 <= offset <= 200:
            fail()
        rows = filtered(self.project, filters or {})
        summary = [{key: value for key, value in group.items() if key not in {'variants', 'possible_duplicates'}} |
                   {'variants': [row['id'] for row in group['variants']], 'possible_duplicates': group['possible_duplicates'][:8],
                    'possible_duplicate_count': len(group['possible_duplicates'])}
                   for group in groups(self.project['entries'])]
        drafts = deepcopy(self.project['drafts'])
        for row in drafts:
            if len(row['target']['text']) > 1000:
                row['target']['text'] = row['target']['text'][:1000]
                row['target_display_truncated'] = True
        result = {'id': self.project['id'], 'name': self.project['name'], 'generation': self.project['generation'],
                'revision': self.revision, 'receipt': self.receipt, 'parent_receipt': self.project['parent_receipt'], 'dirty': self.dirty, 'total': len(self.project['entries']),
                'matched': len(rows), 'entries': [], 'next_offset': None,
                'groups': [], 'groups_scope': 'current_page_only', 'canonical': self.project['canonical'], 'drafts': drafts,
                'target': self.target, 'tasks': self.project['tasks'], 'outputs': self.project['outputs'], 'permissions_restored': False,
                'evidence_records': self.project['evidence_records'],
                'coverage': 'explicit_project_records_only_not_systematic_review_completeness'}
        for row in rows[offset:offset + 20]:
            trial = deepcopy(result)
            trial['entries'].append(row)
            visible = {entry['id'] for entry in trial['entries']}
            trial['groups'] = [group for group in summary if visible & set(group['variants'])]
            if len(json.dumps(trial, ensure_ascii=False).encode('utf8')) > 200000:
                if result['entries']:
                    break
                row = deepcopy(row)
                row['metadata_display_truncated'] = True
                row['author_total'] = len(row['item']['authors'])
                row['item']['authors'] = row['item']['authors'][:10]
                if 'author' in row['csl']:
                    row['csl']['author'] = row['csl']['author'][:10]
                trial['entries'] = [row]
                if len(json.dumps(trial, ensure_ascii=False).encode('utf8')) > 200000:
                    fail('professional_project_capacity')
            result = trial
        stop = offset + len(result['entries'])
        result['next_offset'] = stop if stop < len(rows) else None
        return result

    def snapshot(self, filters=None, offset=0):
        return self._call(lambda editor, path, binding: self._snapshot(filters, offset))

    def action(self, action, payload):
        def operation(editor, path, binding):
            if type(payload) is not dict:
                fail()
            dispatch = {'import_csl': self._import, 'add_online': self._online, 'review': self._review,
                        'canonical': self._canonical, 'associate': self._associate, 'target': self._target,
                        'capture': self._capture, 'queue': self._queue, 'queue_review': self._queue_review,
                        'queue_remove': self._queue_remove, 'rename': self._rename, 'export_csl': self._export,
                        'search': self._search, 'impact': self._impact, 'index_status': self._index_status,
                        'adopt_suggestion': self._adopt_suggestion, 'approved_evidence': self._approved_evidence,
                        'queue_rebind': self._queue_rebind, 'queue_edit': self._queue_edit,
                        'numeric': self._numeric, 'numeric_comment': self._numeric_comment,
                        'numeric_source': self._numeric_source, 'mark_read': self._mark_read,
                        'new_project': self._new_project, 'relate': self._relate, 'project_history': self._history}
            dispatch['queue_target'] = self._queue_target
            dispatch['record_page'] = self._record_page
            if type(action) is not str or action not in dispatch:
                fail()
            return dispatch[action](editor, path, binding, payload)
        return self._call(operation)

    def _history(self, editor, path, binding, payload):
        if payload:
            fail()
        self.history_views.clear()
        receipt = self.receipt or self.project['parent_receipt']
        expected = self.project['generation'] - (0 if self.receipt else 1)
        rows, checked, visited = [], [], set()
        while receipt and len(rows) < 8:
            if receipt in visited:
                fail('professional_project_history_lineage')
            visited.add(receipt)
            document, saved = self.archive.inspect(receipt)
            if document['scope'] != self.project['scope'] or document['id'] != self.project['id'] or document['generation'] != expected:
                fail('professional_project_history_lineage')
            if not rows and self.receipt and document['parent_receipt'] != self.project['parent_receipt']:
                fail('professional_project_history_lineage')
            parent = document['parent_receipt']
            if (expected == 1 and parent) or (expected > 1 and not parent) or expected < 1 or parent in visited:
                fail('professional_project_history_lineage')
            rows.append({'receipt': receipt, 'name': document['name'], 'generation': expected,
                'materials': len(document['entries']), 'tasks': len(document['tasks']), 'drafts': len(document['drafts']),
                'parent_receipt': parent, 'current_saved': receipt == self.receipt})
            checked.append(saved)
            receipt, expected = parent, expected - 1
        for saved in checked:
            self.archive._saved_check(saved)
        self.history_views = {row['receipt']: saved for row, saved in zip(rows, checked)}
        return {'snapshots': rows, 'truncated': bool(receipt), 'limit': 8, 'permissions_restored': False,
            'coverage': 'current_project_registered_parent_chain_only', 'project_id': self.project['id'],
            'revision': self.revision, 'generation': self.project['generation'], 'current_receipt': self.receipt}

    def _record_page(self, editor, path, binding, payload):
        if set(payload) != {'id'}:
            fail()
        self._entry(payload['id'])
        rows = filtered(self.project, {})
        index = next(index for index, row in enumerate(rows) if row['id'] == payload['id'])
        return {'snapshot': self._snapshot({}, index), 'offset': index, 'selected_record_id': payload['id'], 'files_read': 0}

    def _import(self, editor, path, binding, payload):
        if set(payload) != {'content'}:
            fail()
        rows = import_rows(payload['content'])
        fresh = {row['id']: record_from_csl(row) for row, source in rows if row['id'] not in self.project['entries']}
        for entry in fresh.values():
            entry['source_records'] = []
        if len(self.project['entries']) + len(fresh) > 200:
            fail('professional_project_capacity')
        trial = deepcopy(self.project)
        trial['entries'].update(fresh)
        for row, source in rows:
            sources = trial['entries'][row['id']]['source_records']
            if source not in sources:
                if len(sources) >= 8:
                    fail('professional_project_capacity')
                sources.append(source)
        trial = validate(trial)
        altered = payload_hash(trial) != payload_hash(self.project)
        self.project = trial
        if altered:
            self._changed()
        return {'added': len(fresh), 'identical_skipped': len(rows) - len(fresh), 'snapshot': self._snapshot(), 'files_read': 0, 'downloads': 0}

    def _online(self, editor, path, binding, payload):
        if set(payload) != {'record_ids'} or type(payload['record_ids']) is not list or not 1 <= len(payload['record_ids']) <= 20:
            fail()
        self.literature._bind()
        rows = []
        for key in payload['record_ids']:
            original = self.literature._record(key)
            entry = record_from_csl(csl(original), source=original['source'])
            entry['source_records'] = [{'source': original['source'], 'sha256': original['source_record_sha256'], 'record_id': key}]
            entry['relations'] = deepcopy(original.get('relations', []))
            entry['item']['retraction_status'] = original['retraction_status']
            rows.append(entry)
        trial = deepcopy(self.project)
        for entry in rows:
            key = entry['item']['id']
            if key not in trial['entries']:
                trial['entries'][key] = entry
            elif entry['source_records'][0] not in trial['entries'][key]['source_records']:
                trial['entries'][key]['source_records'].append(entry['source_records'][0])
            if entry['item']['retraction_status'] == 'reported_retracted':
                trial['entries'][key]['item']['retraction_status'] = 'reported_retracted'
            for relation in entry['relations']:
                if relation not in trial['entries'][key]['relations']:
                    trial['entries'][key]['relations'].append(relation)
        validate(trial)
        self.project = trial
        self._changed()
        return self._snapshot()

    def _review(self, editor, path, binding, payload):
        if set(payload) != {'id', 'decision', 'reason', 'tags'} or payload['decision'] not in {'include', 'exclude', 'undecided'}:
            fail()
        entry = self._entry(payload['id'])
        reason = clean(payload['reason'], 400, empty=payload['decision'] == 'undecided')
        if type(payload['tags']) is not list or len(payload['tags']) > 12:
            fail()
        tags = sorted(set(clean(tag, 40) for tag in payload['tags']))
        if len(entry['decisions']) >= 20:
            fail('professional_project_capacity')
        entry.update(decision=payload['decision'], reason=reason, tags=tags)
        entry['decisions'].append({'decision': payload['decision'], 'reason': reason, 'revision': self.revision + 1})
        if payload['id'] in self.literature.records:
            self.literature.decisions[payload['id']] = {'decision': payload['decision'], 'reason': reason}
        self._changed()
        for row in self.project['drafts']:
            if payload['id'] in row['record_ids']:
                row['status'] = 'recheck'
                self.live_drafts.pop(row['id'], None)
        return self._snapshot()

    def _canonical(self, editor, path, binding, payload):
        if set(payload) != {'id'}:
            fail()
        entry = self._entry(payload['id'])
        self.project['canonical'][entry['item']['doi'] or payload['id']] = payload['id']
        self._changed()
        return self._snapshot()

    def _new_project(self, editor, path, binding, payload):
        if set(payload) != {'name'}:
            fail()
        name = clean(payload['name'], 100)
        self.invalidate()
        self._ensure(editor, path)
        self.project['name'] = name
        self._changed()
        return self._snapshot()

    def _relate(self, editor, path, binding, payload):
        if set(payload) != {'id', 'target_id', 'relation', 'reason'} or payload['id'] == payload['target_id']:
            fail()
        allowed = {'preprint_of', 'version_of', 'correction_of', 'retraction_of', 'possible_duplicate_of'}
        if payload['relation'] not in allowed:
            fail()
        source, target = self._entry(payload['id']), self._entry(payload['target_id'])
        reason = clean(payload['reason'], 400)
        relation = {'type': payload['relation'], 'id': target['item']['doi'] or payload['target_id'], 'reason': reason}
        if relation not in source['relations']:
            if len(source['relations']) >= 20:
                fail('professional_project_capacity')
            source['relations'].append(relation)
            if payload['relation'] == 'retraction_of':
                target['item']['retraction_status'] = 'reported_retracted'
                if payload['target_id'] in self.literature.decisions:
                    self.literature.decisions[payload['target_id']] = {'decision': 'exclude', 'reason': reason}
            self._changed()
        return self._snapshot()

    def _associate(self, editor, path, binding, payload):
        if set(payload) != {'id', 'filename'}:
            fail()
        entry = self._entry(payload['id'])
        name = clean(payload['filename'], 128)
        if any(char in name for char in '/\\:') or not name.lower().endswith('.pdf'):
            fail('professional_project_source_refused')
        data, identity, version = SourceExtractor(editor.space)._snapshot(name)
        if not data.startswith(b'%PDF-'):
            fail('professional_project_source_refused')
        previous = deepcopy(entry['link'])
        if previous:
            if len(entry['link_history']) >= 8:
                fail('professional_project_capacity')
            entry['link_history'].append(previous)
        entry['link'] = {'name': name, 'sha256': version, 'identity': json.loads(json.dumps(identity))}
        entry['read_pages'] = []
        entry['parsed_pages'] = []
        entry['indexed_pages'] = []
        self._changed()
        self.captures.pop(payload['id'], None)
        self.capture_windows = {key: record_id for key, record_id in self.capture_windows.items() if record_id != payload['id']}
        self.evidence = {key: row for key, row in self.evidence.items() if row['record_id'] != payload['id']}
        for row in self.project['drafts']:
            if payload['id'] in row['record_ids']:
                row['status'] = 'recheck'
                self.live_drafts.pop(row['id'], None)
        return {'previous': previous, 'current': entry['link'], 'read': False, 'indexed': False, 'support_assessed': False}

    def _target(self, editor, path, binding, payload):
        if set(payload) != {'locator_id', 'offset'}:
            fail()
        result = self.structure.read_structure(payload['offset'])
        if not result['ok']:
            fail(result['reason'])
        row = next((row for row in result['result']['anchors'] if row['locator_id'] == payload['locator_id']), None)
        if not row:
            fail('professional_stale_locator')
        target = {**row, 'version': result['result']['document']['source_sha256'], 'name': path.name}
        if self.target != target:
            self._retire_assist()
            self.last_numeric = None
            self.numeric_record_id = None
        self.target = target
        return self.target

    def _admit(self, key):
        entry = self._entry(key)
        if entry['decision'] != 'include' or entry['item']['retraction_status'] == 'reported_retracted':
            fail('professional_project_not_included')
        variants = [row for row in self.project['entries'].values() if entry['item']['doi'] and row['item']['doi'] == entry['item']['doi']]
        if len(variants) > 1 and self.project['canonical'].get(entry['item']['doi']) != key:
            fail('professional_project_conflict')
        self.literature._bind()
        if key not in self.literature.records and len(self.literature.records) >= 100:
            fail('professional_project_capacity')
        self.literature.records[key] = admitted_record(entry)
        self.literature.decisions[key] = {'decision': 'include', 'reason': entry['reason']}
        return entry

    def _check_link(self, editor, entry):
        if not entry['link']:
            fail('professional_project_unassociated')
        _, identity, version = SourceExtractor(editor.space)._snapshot(entry['link']['name'])
        if version != entry['link']['sha256'] or json.loads(json.dumps(identity)) != entry['link']['identity']:
            fail('professional_project_source_changed')

    def _capture(self, editor, path, binding, payload):
        if set(payload) != {'id', 'page', 'limit'}:
            fail()
        entry = self._admit(payload['id'])
        self._check_link(editor, entry)
        result = self.grounded.capture(payload['id'], entry['link']['name'], payload['page'], payload['limit'])
        if not result['ok']:
            fail(result['reason'])
        result = result['result']
        key = result['capture_id']
        capture = self.grounded._captures[key]
        self.captures[payload['id']] = key
        self.capture_windows[key] = payload['id']
        for fragment in capture['proof']['fragments']:
            self.evidence[fragment['id']] = {**deepcopy(fragment), 'record_id': payload['id'], 'capture_id': key,
                'title': entry['item']['title'], 'literature': entry['csl']}
        entry['parsed_pages'] = sorted(set(entry['parsed_pages']) | set(capture['proof'].get('coverage', {}).get('complete_pages', [])))
        self._changed()
        return {**result, 'evidence': list(self.evidence.values())[:40], 'read_is_support': False}

    def _mark_read(self, editor, path, binding, payload):
        if set(payload) != {'id', 'reviewed'} or payload['reviewed'] is not True:
            fail()
        entry = self._entry(payload['id'])
        self._check_link(editor, entry)
        capture = self.grounded._captures.get(self.captures.get(payload['id']))
        if not capture or capture['binding'] != self.grounded._binding()[2]:
            fail('professional_evidence_invalid')
        entry['read_pages'] = sorted(set(entry['read_pages']) | set(capture['proof'].get('coverage', {}).get('complete_pages', [])))
        self._changed()
        return self._snapshot()

    def _queue(self, editor, path, binding, payload):
        allowed = {'operation', 'text', 'evidence_ids', 'quote', 'side'}
        if not self.target or not {'operation', 'text', 'evidence_ids'} <= payload.keys() or not payload.keys() <= allowed:
            fail()
        if self.target['version'] != digest(path):
            fail('professional_source_changed')
        if payload['operation'] not in {'replace', 'insert', 'comment'} or len(self.project['drafts']) >= 12:
            fail('professional_project_capacity')
        ids = payload['evidence_ids']
        if type(ids) is not list or len(ids) > 3 or len(set(ids)) != len(ids) or any(key not in self.evidence for key in ids):
            fail('professional_evidence_invalid')
        if payload['operation'] == 'insert' and not ids:
            fail('professional_project_insufficient')
        body = clean(payload['text'], 2000)
        row = {'id': uuid.uuid4().hex, 'operation': payload['operation'], 'text': body,
            'target': deepcopy(self.target), 'evidence_ids': ids, 'record_ids': sorted({self.evidence[key]['record_id'] for key in ids}),
            'quote': clean(payload.get('quote') or self.target['text'], 2000), 'side': payload.get('side', 'after'),
            'status': 'pending', 'relation': 'insufficient', 'applicability': '', 'limitation': '', 'numeric_reviews': []}
        if row['side'] not in {'before', 'after'}:
            fail()
        self.project['drafts'].append(row)
        self.live_drafts[row['id']] = binding
        self._changed()
        return self._snapshot()

    def _queue_review(self, editor, path, binding, payload):
        required = {'id', 'accepted', 'reviewed', 'relation', 'applicability', 'limitation', 'text'}
        if set(payload) != required or type(payload['accepted']) is not bool or type(payload['reviewed']) is not bool:
            fail()
        row = next((row for row in self.project['drafts'] if row['id'] == payload['id']), None)
        if not row or self.live_drafts.get(row['id']) != binding or row['target']['version'] != digest(path):
            fail('professional_project_recheck')
        if payload['accepted'] and not payload['reviewed']:
            fail('professional_assist_review_required')
        if payload['relation'] not in {'supports', 'contradicts', 'insufficient'}:
            fail()
        if row['operation'] == 'insert' and payload['accepted'] and (payload['relation'] != 'supports' or not row['evidence_ids']):
            fail('professional_project_insufficient')
        for key in row['record_ids']:
            self._admit(key)
            self._check_link(editor, self._entry(key))
        new_text = clean(payload['text'], 2000)
        new_applicability = clean(payload['applicability'], 500, empty=row['operation'] != 'insert')
        new_limitation = clean(payload['limitation'], 500, empty=row['operation'] != 'insert')
        additions = []
        if payload['accepted']:
            for key in row['evidence_ids']:
                evidence = self.evidence.get(key)
                if not evidence:
                    fail('professional_evidence_invalid')
                claim = {'id': uuid.uuid4().hex, 'record_id': evidence['record_id'], 'source_sha256': evidence['source_sha256'],
                    'evidence_id': key, 'location': evidence['location'], 'text': new_text, 'target': row['target'], 'relation': payload['relation'],
                    'layout_stamp': evidence.get('layout_stamp', '')}
                claim.update(original_quote=evidence['original_text'], original_header=evidence.get('header', ''),
                             char_indices=evidence.get('char_indices', []))
                if not any(value['evidence_id'] == key and value['text'] == new_text for value in self.project['claims']):
                    additions.append(claim)
        if len(self.project['claims']) + len(additions) > 32:
            fail('professional_project_capacity')
        row.update(text=new_text, relation=payload['relation'], applicability=new_applicability,
                   limitation=new_limitation, status='accepted' if payload['accepted'] else 'rejected')
        self.project['claims'].extend(additions)
        self._changed()
        return self._snapshot()

    def _queue_remove(self, editor, path, binding, payload):
        if set(payload) != {'id'} or not any(row['id'] == payload['id'] for row in self.project['drafts']):
            fail()
        self.project['drafts'] = [row for row in self.project['drafts'] if row['id'] != payload['id']]
        self.live_drafts.pop(payload['id'], None)
        self._changed()
        return self._snapshot()

    def _rename(self, editor, path, binding, payload):
        if set(payload) != {'name'}:
            fail()
        self.project['name'] = clean(payload['name'], 100)
        self._changed()
        return self._snapshot()

    def _queue_rebind(self, editor, path, binding, payload):
        if set(payload) != {'id'}:
            fail()
        row = next((row for row in self.project['drafts'] if row['id'] == payload['id']), None)
        if not row or not self.target or row['target']['locator_id'] != self.target['locator_id'] or row['target']['version'] != digest(path):
            fail('professional_project_recheck')
        if any(key not in self.evidence for key in row['evidence_ids']):
            fail('professional_evidence_invalid')
        for key in row['record_ids']:
            self._admit(key)
            self._check_link(editor, self._entry(key))
        self.live_drafts[row['id']] = binding
        row['status'] = 'pending'
        self._changed()
        return self._snapshot()

    def _queue_target(self, editor, path, binding, payload):
        if set(payload) != {'id'}:
            fail()
        row = next((row for row in self.project['drafts'] if row['id'] == payload['id']), None)
        if not row or row['target']['version'] != digest(path) or row['target']['name'] != path.name:
            fail('professional_project_recheck')
        from ..advanced_writes import _parts
        from ..professional_edits import _inventory
        _, inventory = _inventory(_parts(path), digest(path))
        index = next((index for index, target in enumerate(inventory) if target['locator_id'] == row['target']['locator_id']), None)
        if index is None:
            fail('professional_stale_locator')
        offset = index
        target = self._target(editor, path, binding, {'locator_id': row['target']['locator_id'], 'offset': offset})
        return {'target': target, 'offset': offset, 'write_approval_restored': False}

    def _queue_edit(self, editor, path, binding, payload):
        if set(payload) != {'id', 'text', 'evidence_ids', 'quote'}:
            fail()
        row = next((row for row in self.project['drafts'] if row['id'] == payload['id']), None)
        ids = payload['evidence_ids']
        if not row or row['target']['version'] != digest(path) or self.live_drafts.get(row['id']) != binding:
            fail('professional_project_recheck')
        if type(ids) is not list or len(ids) > 3 or len(ids) != len(set(ids)) or any(key not in self.evidence for key in ids):
            fail('professional_evidence_invalid')
        if row['operation'] == 'insert' and not ids:
            fail('professional_project_insufficient')
        new_text, quote = clean(payload['text'], 2000), clean(payload['quote'], 2000)
        row.update(text=new_text, quote=quote, evidence_ids=ids, record_ids=sorted({self.evidence[key]['record_id'] for key in ids}),
                   status='pending', relation='insufficient', applicability='', limitation='', numeric_reviews=[])
        self._changed()
        return self._snapshot()

    def layout_changed(self, name, fingerprint, stamp, parser_stamp):
        if not self.project:
            return
        keys = {key for key, entry in self.project['entries'].items() if entry['link'] and entry['link']['name'] == name and entry['link']['sha256'] == fingerprint}
        if not keys:
            return
        previous = next((row for row in self.project['layouts'] if row['name'] == name), None)
        if previous and previous['stamp'] == stamp and previous['parser_stamp'] == parser_stamp:
            return
        if not previous and len(self.project['layouts']) >= 8:
            fail('professional_project_capacity')
        self.project['layouts'] = [row for row in self.project['layouts'] if row['name'] != name]
        self.project['layouts'].append({'name': name, 'source_sha256': fingerprint, 'stamp': stamp, 'parser_stamp': parser_stamp})
        for row in self.project['drafts']:
            if row['numeric_reviews'] and set(row['record_ids']) & keys:
                row['status'] = 'recheck'
                self.live_drafts.pop(row['id'], None)
        self._changed()

    def _numeric(self, editor, path, binding, payload):
        if set(payload) != {'id', 'claim'} or not self.target:
            fail()
        entry = self._entry(payload['id'])
        self._check_link(editor, entry)
        result = self.structure.verify_numeric(self.target['locator_id'], payload['claim'])
        if not result['ok']:
            fail(result['reason'])
        result = result['result']
        if result['evidence']['source'] != entry['link']['name'] or result['evidence']['source_sha256'] != entry['link']['sha256']:
            fail('professional_evidence_invalid')
        self.last_numeric = {'result': result, 'record_id': payload['id'], 'binding': binding}
        return result

    def _numeric_source(self, editor, path, binding, payload):
        if set(payload) != {'id'} or not self.target:
            fail()
        entry = self._entry(payload['id'])
        self._check_link(editor, entry)
        self.numeric_record_id = payload['id']
        self.last_numeric = None
        return {'filename': entry['link']['name'], 'target': self.target}

    def observe_numeric(self, locator, result):
        self.last_numeric = None
        if not self.project or not self.numeric_record_id or not self.target or self.target['locator_id'] != locator or not result.get('numeric_review_id'):
            return
        entry = self._entry(self.numeric_record_id)
        source = result.get('evidence')
        if not entry['link'] or not source or source['source'] != entry['link']['name'] or source['source_sha256'] != entry['link']['sha256']:
            return
        self.last_numeric = {'result': deepcopy(result), 'record_id': self.numeric_record_id, 'binding': self._binding()[2]}

    def _numeric_comment(self, editor, path, binding, payload):
        if payload or not self.last_numeric or self.last_numeric['binding'] != binding:
            fail('professional_project_recheck')
        result = self.last_numeric['result']
        review = self.structure._numeric_reviews.get(result.get('numeric_review_id'))
        if not review or time.monotonic() - review['created'] >= 300:
            fail('professional_pdf_expired')
        entry = self._entry(self.last_numeric['record_id'])
        self._check_link(editor, entry)
        item = result['evidence']
        key = 'ev:' + payload_hash([item['source_sha256'], item['location'], item['raw']])
        self.evidence[key] = {'id': key, 'record_id': self.last_numeric['record_id'], 'original_text': item['raw'],
            'source_sha256': item['source_sha256'], 'location': item['location'], 'capture_id': None,
            'title': entry['item']['title'], 'literature': entry['csl'], 'layout_stamp': item['grid_sha256'],
            'header': item['header'], 'char_indices': item['char_indices']}
        queued = self._queue(editor, path, binding, {'operation': 'comment', 'text': result['suggestion'],
            'evidence_ids': [key], 'quote': result['claim']['value']})
        self.project['drafts'][-1]['numeric_reviews'] = [result['numeric_review_id']]
        return self._snapshot()

    def _index_status(self, editor, path, binding, payload):
        if set(payload) != {'id'}:
            fail()
        entry = self._entry(payload['id'])
        self._check_link(editor, entry)
        result = self.rag.pdf_ranges_status(entry['link']['name'])
        if not result['ok']:
            fail('professional_project_index_unavailable')
        report = result['result']['ranges']
        entry['indexed_pages'] = report.get('indexed_pages', [])
        self._changed()
        return {'ranges': report, 'snapshot': self._snapshot(), 'metadata_status_is_not_index_authorization': True}

    def _check_evidence(self, editor):
        from ..grounded_edits import validate_source, seal
        for capture_id in self.capture_windows:
            capture = self.grounded._captures.get(capture_id)
            if not capture or capture['binding'] != self.grounded._binding()[2] or capture['session'] != self.literature.session:
                fail('professional_evidence_invalid')
            validate_source(capture['proof'], seal(capture['proof']), self._selected())
        if self.approved_evidence:
            self.rag._approved_bundle(self.approved_evidence)
        for key in {row['record_id'] for row in self.evidence.values()}:
            self._admit(key)
            self._check_link(editor, self._entry(key))
        if any(row['source_sha256'] != self._entry(row['record_id'])['link']['sha256'] for row in self.evidence.values()):
            fail('professional_project_source_changed')

    def _search(self, editor, path, binding, payload):
        if set(payload) != {'question', 'focus', 'dense', 'read_ids'} or type(payload['dense']) is not bool:
            fail()
        question = clean(payload['question'], 200)
        focus = clean(payload['focus'], 50, empty=True)
        if not self.target or not self.evidence:
            return {'status': 'insufficient', 'evidence': [], 'reason': 'no_current_admitted_materials', 'model_calls': 0,
                    **self._coverage_report(question, {})}
        ids = payload['read_ids']
        if type(ids) is not list or len(ids) > 6 or any(key not in self.evidence for key in ids):
            fail()
        stamp = payload_hash([self.project['id'], question, [(key, row['source_sha256']) for key, row in self.evidence.items()]])
        if self.search_rounds.get(stamp, 0) >= 2:
            fail('professional_project_search_budget')
        self._check_evidence(editor)
        self.search_rounds[stamp] = self.search_rounds.get(stamp, 0) + 1
        query = clean(question + (' ' + focus if focus else ''), 200)
        models, degradation = None, None
        if payload['dense']:
            from ..embeddings import EmbeddingError
            try:
                models = self.assist.retriever_factory() if self.assist.retriever_factory else None
            except EmbeddingError as error:
                if str(error) not in {'embedding_local_not_installed', 'embedding_local_directory', 'embedding_failed'}:
                    raise
                degradation = 'local_embedding_unavailable_lexical_only'
            if models is None and degradation is None:
                degradation = 'local_dense_not_configured_lexical_only'
        from ..research_capture_retrieval import retrieve_captures
        families = paper_families(self.project)
        for row in self.evidence.values():
            entry = self._entry(row['record_id'])
            row['diversity_key'] = families[row['record_id']]
        found, report = retrieve_captures(self.evidence, question, [query], ids, models,
                                         check=lambda: self._check_evidence(editor))
        self._check_evidence(editor)
        return {'status': 'completed', 'evidence': list(found.values()), 'retrieval': report,
            'degradation': degradation, 'rounds_used': self.search_rounds[stamp], 'round_limit': 2, 'model_calls': 0,
            'queries_are_evidence': False, 'effective_query': query, **self._coverage_report(question, found)}

    def _coverage_report(self, question, found):
        claims = self.project['claims']
        from ..retrieval_pipeline import subqueries
        coverage = []
        for part in subqueries(question, True):
            terms = re.findall(r'[A-Za-z0-9]+|[\u4e00-\u9fff]{2,}', part.casefold())
            matches = [key for key, row in found.items() if any(term in row['original_text'].casefold() for term in terms)]
            ignored = {'the', 'and', 'for', 'does', 'this', 'that', 'with', 'not', 'infer', 'beyond', 'report', 'reported'}
            literals = list(dict.fromkeys(term for term in terms if term not in ignored and len(term) > 1))
            missing = [term for term in literals if not any(term in found[key]['original_text'].casefold() for key in matches)]
            coverage.append({'question': part, 'retrieved_ids': matches,
                'reviewed_support_ids': [row['evidence_id'] for row in claims if row['relation'] == 'supports' and row['evidence_id'] in matches],
                'reviewed_contradiction_ids': [row['evidence_id'] for row in claims if row['relation'] == 'contradicts' and row['evidence_id'] in matches],
                'support_assessed_by_search': False, 'pending_numeric_terms': re.findall(r'\d+(?:\.\d+)?', part),
                'condition_review_required': True, 'unmatched_literals': missing[:12],
                'missing_conditions_assessment': 'literal_gaps_not_semantic_condition_verdict',
                'status': 'retrieved_requires_verification' if matches else 'not_retrieved_in_current_materials'})
        materials = []
        for key, entry in list(self.project['entries'].items())[:20]:
            pending = sorted(set(entry['parsed_pages']) - set(entry['read_pages']))
            materials.append({'id': key, 'title': entry['item']['title'], 'decision': entry['decision'],
                'associated': bool(entry['link']), 'parsed_unread_pages': pending[:20], 'parsed_unread_count': len(pending),
                'read_pages': entry['read_pages'][:20], 'indexed_pages': entry['indexed_pages'],
                'coverage': 'registered_metadata_not_current_read_authorization', 'unprocessed_pages_unknown': True})
        return {'coverage': coverage, 'original_question': question, 'materials': materials,
            'materials_truncated': len(self.project['entries']) > 20, 'registered_material_count': len(self.project['entries']),
            'unread_records': [key for key, row in self.project['entries'].items() if not row['read_pages']],
            'absence_of_contradictions_proven': False, 'complete_fulltext_coverage': False}

    def _approved_evidence(self, editor, path, binding, payload):
        if set(payload) != {'receipt', 'id'}:
            fail()
        entry = self._admit(payload['id'])
        self._check_link(editor, entry)
        try:
            bundle, stamp = self.rag._approved_bundle(payload['receipt'])
        except ValueError as error:
            if str(error) != 'research_evidence_unavailable':
                raise
            fail('professional_project_evidence_reapproval')
        if not 1 <= len(bundle['hits']) <= 32:
            fail('professional_project_capacity')
        registration = {'receipt': payload['receipt'], 'record_id': payload['id'], 'source_sha256': entry['link']['sha256'],
                        'evidence_ids': []}
        from ..provenance import identifier
        for hit in bundle['hits']:
            if hit['source'] != entry['link']['name'] or hit['source_fingerprint'] != entry['link']['sha256']:
                fail('professional_evidence_invalid')
            key = 'ev:' + identifier(hit['source_fingerprint'], hit['location'], hit['text'])
            if key not in registration['evidence_ids']:
                registration['evidence_ids'].append(key)
        from ..project_library import EvidenceRecord
        EvidenceRecord.model_validate(registration)
        if registration not in self.project['evidence_records'] and len(self.project['evidence_records']) >= 32:
            fail('professional_project_capacity')
        for hit in bundle['hits']:
            key = 'ev:' + identifier(hit['source_fingerprint'], hit['location'], hit['text'])
            self.evidence[key] = {'id': key, 'original_text': hit['text'], 'source_sha256': hit['source_fingerprint'],
                'location': hit['location'], 'record_id': payload['id'], 'capture_id': None, 'title': entry['item']['title'],
                'literature': entry['csl'], 'page': None, 'context': [], 'saved_evidence_stamp': stamp}
        self.approved_evidence = payload['receipt']
        if registration not in self.project['evidence_records']:
            self.project['evidence_records'].append(registration)
            self._changed()
        return {'evidence': list(self.evidence.values())[:40], 'mode': 'approved_persistent_evidence', 'new_embeddings': 0}

    def _impact(self, editor, path, binding, payload):
        if set(payload) != {'record_ids'} or type(payload['record_ids']) is not list or not 1 <= len(payload['record_ids']) <= 8:
            fail()
        output, reads, text_cache, invalidated = [], 0, {}, set()
        for key in payload['record_ids']:
            entry = self._entry(key)
            if not entry['link']:
                output.append({'kind': 'source', 'record_id': key, 'reason': 'unassociated', 'certainty': 'unknown'})
                continue
            try:
                _, identity, current = SourceExtractor(editor.space)._snapshot(entry['link']['name'])
                new_identity = json.loads(json.dumps(identity))
            except FileNotFoundError:
                current, new_identity = None, None
            except SourceError:
                fail('professional_project_source_refused')
            identity_changed = new_identity != entry['link']['identity']
            layout = next((row['stamp'] for row in self.project['layouts'] if row['name'] == entry['link']['name']), '')
            layout_claims = {claim['evidence_id'] for claim in self.project['claims'] if claim['record_id'] == key and
                             claim.get('layout_stamp') and layout and claim['layout_stamp'] != layout}
            def old_identity(fingerprint):
                identities = [row['identity'] for row in [entry['link'], *entry['link_history']] if row['sha256'] == fingerprint]
                unique = {json.dumps(value, sort_keys=True): value for value in identities}
                return next(iter(unique.values())) if len(unique) == 1 else None
            for registration in self.project['evidence_records']:
                if registration['record_id'] == key and registration['source_sha256'] != current:
                    output.append({'kind': 'evidence', 'record_id': key, 'id': registration['receipt'],
                        'location': 'saved_evidence_receipt', 'reason': 'source_changed_or_reassociated',
                        'old_sha256': registration['source_sha256'], 'new_sha256': current,
                        'old_identity': old_identity(registration['source_sha256']), 'new_identity': new_identity,
                        'certainty': 'registered_receipt_only_body_not_read', 'action': 'requery_current_source_no_permission_restored'})
            for evidence in self.evidence.values():
                if evidence['record_id'] != key:
                    continue
                layout_changed = bool(evidence.get('layout_stamp') and layout and evidence['layout_stamp'] != layout)
                if evidence['source_sha256'] != current or identity_changed or layout_changed:
                    output.append({'kind': 'evidence', 'record_id': key, 'id': evidence['id'], 'location': evidence['location'],
                        'reason': 'layout_interpretation_changed' if layout_changed else 'source_changed_or_missing',
                        'old_sha256': evidence['source_sha256'], 'new_sha256': current,
                        'old_identity': old_identity(evidence['source_sha256']), 'new_identity': new_identity,
                        'certainty': 'registered_evidence_identity_not_fact_verdict', 'action': 'reacquire_and_recheck'})
            for claim in self.project['claims']:
                layout_changed = bool(claim.get('layout_stamp') and layout and claim['layout_stamp'] != layout)
                if claim['record_id'] == key and (claim['source_sha256'] != current or layout_changed or identity_changed):
                    output.append({'kind': 'claim', 'record_id': key, 'id': claim['id'], 'reason': 'source_changed_or_missing',
                        'old_sha256': claim['source_sha256'], 'new_sha256': current, 'old_identity': old_identity(claim['source_sha256']),
                        'new_identity': new_identity, 'location': claim['location'], 'claim': claim['text'],
                        'certainty': 'exact_registered_source_digest', 'action': 'reacquire_and_recheck'})
                    if layout_changed:
                        output[-1].update(reason='layout_interpretation_changed', old_layout=claim['layout_stamp'], new_layout=layout,
                                          certainty='explicit_registered_layout_stamp')
                    comparison = {'status': 'not_rechecked', 'original_quote': claim.get('original_quote', '')}
                    candidate = getattr(self.structure, '_pdf_candidate', None)
                    if layout_changed and candidate and candidate['name'] == entry['link']['name']:
                        cell = next((row for row in candidate.get('records', []) if row['char_indices'] == claim.get('char_indices') and row['char_indices']), None)
                        if cell:
                            comparison.update(status='same_character_scope_reinterpreted', original_header=claim.get('original_header', ''),
                                              current_header=cell['header'], current_raw=cell['raw'], current_unit=cell['unit'])
                    match = re.fullmatch(r'pdf-page:(\d+)/char:(\d+)', claim['location'])
                    if current and match and claim.get('original_quote'):
                        cache_key = (entry['link']['name'], current, claim['location'])
                        if cache_key not in text_cache and reads < 5:
                            from ..pdf_reader import read_pdf
                            reads += 1
                            try:
                                report = read_pdf(editor.space.file(str(editor.space.root / entry['link']['name'])),
                                    page=int(match[1]) - 1, limit=1, char_offset=int(match[2]), max_chars=1000, layout=False)
                                if report['source_sha256'] != current or SourceExtractor(editor.space)._snapshot(entry['link']['name'])[2] != current:
                                    fail('professional_project_source_changed')
                                text_cache[cache_key] = report['pages'][0]['text']
                            except SourceError as error:
                                if error.code not in {'unsupported', 'capacity', 'pdf_text_missing'}:
                                    fail('professional_project_source_refused')
                                comparison.update(status='parse_failed', parser_status=error.code)
                        if cache_key in text_cache:
                            value = text_cache[cache_key]
                            old = claim['original_quote']
                            comparison.update(status='quote_unchanged_at_registered_location' if value.startswith(old) else
                                                     'quote_moved_within_bounded_window' if old in value else 'quote_changed_or_moved_requires_review',
                                              current_excerpt=value, source_sha256=current)
                            if len(old) > 1000:
                                comparison['status'] = 'quote_prefix_unchanged_full_quote_not_checked' if old.startswith(value) and value else 'quote_exceeds_window_requires_review'
                        elif reads >= 5:
                            comparison['status'] = 'bounded_read_budget_exhausted'
                    output[-1]['content_comparison'] = comparison
            changed = current != entry['link']['sha256'] or new_identity != entry['link']['identity']
            if changed:
                output.append({'kind': 'source', 'record_id': key, 'reason': 'associated_version_changed',
                    'old_sha256': entry['link']['sha256'], 'new_sha256': current, 'certainty': 'file_identity_and_digest', 'action': 'explicit_reassociate'})
            for row in self.project['drafts']:
                historical = any(claim['record_id'] == key and claim['source_sha256'] != current and claim['evidence_id'] in row['evidence_ids'] for claim in self.project['claims'])
                layout_dependency = bool(set(row['evidence_ids']) & layout_claims) or any(
                    self.evidence.get(identifier, {}).get('layout_stamp') and layout and
                    self.evidence[identifier]['layout_stamp'] != layout for identifier in row['evidence_ids'])
                if key in row['record_ids'] and (changed or historical or layout_dependency):
                    invalidated.add(row['id'])
                    output.append({'kind': 'draft', 'id': row['id'], 'record_id': key, 'location': row['target']['location'],
                        'reason': 'layout_interpretation_changed' if layout_dependency else 'source_changed_or_reassociated',
                        'old_sha256': next((claim['source_sha256'] for claim in self.project['claims'] if claim['evidence_id'] in row['evidence_ids'] and claim['record_id'] == key), None),
                        'new_sha256': current, 'new_identity': new_identity,
                        'certainty': 'registered_evidence_link', 'action': 'review_again_before_new_plan'})
            for item in self.project['tasks'] + self.project['outputs']:
                if key in item['record_ids'] and (item['source_stamps'].get(key) != current or layout_claims):
                    output.append({'kind': item['kind'], 'id': item['receipt'], 'record_id': key,
                        'old_sha256': item['source_stamps'].get(key), 'new_sha256': current,
                        'old_identity': old_identity(item['source_stamps'].get(key)), 'new_identity': new_identity,
                        'location': item['target_name'], 'target_version': item['target_version'],
                        'reason': 'source_changed_or_reassociated' if item['source_stamps'].get(key) != current else 'possible_layout_dependency',
                        'registration_state': 'historical_receipt_current_output_or_rollback_not_checked',
                        'certainty': 'possible_registered_layout_dependency' if item['source_stamps'].get(key) == current else 'registered_source_link_not_full_document_proof',
                        'action': 'review_receipt_and_copy_no_automatic_rewrite'})
        if len(output) > 96:
            fail('professional_project_capacity')
        if any(row['id'] in invalidated and row['status'] != 'recheck' for row in self.project['drafts']):
            for row in self.project['drafts']:
                if row['id'] in invalidated:
                    row['status'] = 'recheck'
                    self.live_drafts.pop(row['id'], None)
            self._changed()
        return {'impacts': output, 'registered_records_checked': len(payload['record_ids']),
            'registered_claims': len(self.project['claims']), 'registered_outputs': len(self.project['outputs']),
            'content_windows_read': reads, 'content_window_limit': 5, 'characters_per_window_limit': 1000,
            'coverage': 'explicit_registered_project_only', 'complete_dependency_graph': False,
            'history_deleted': False, 'documents_rewritten': False, 'invalidated_drafts': sorted(invalidated)}

    def start_assist(self, payload):
        def operation(editor, path, binding):
            if not self.target or type(payload) is not dict or set(payload) != {'question', 'mode', 'outbound_approved'}:
                fail()
            self._check_evidence(editor)
            project_id = self.project['id']
            session, target = self.assist_session, deepcopy(self.target)
            sources = {key: deepcopy(self._entry(key)['link']) for key in {row['record_id'] for row in self.evidence.values()}}
            def check_project():
                if not self.project or self.project['id'] != project_id or self.assist_session is not session or self.target != target:
                    fail('professional_project_changed')
                for key, link in sources.items():
                    entry = self._entry(key)
                    if entry['link'] != link:
                        fail('professional_project_source_changed')
                    self._admit(key)
                    self._check_link(editor, entry)
            families = paper_families(self.project)
            paper_groups = {(self._entry(key)['item'].get('doi') or link['sha256']): families[key] for key, link in sources.items()}
            result = self.assist.start({'question': payload['question'], 'mode': payload['mode'], 'outbound_approved': payload['outbound_approved'],
                'source_sha256': digest(path), 'locator_ids': [self.target['locator_id']], 'capture_ids': list(self.capture_windows),
                'evidence_save_id': self.approved_evidence, 'protected_terms': [], 'max_calls': 3, 'compute_budget': 60}, preserve_queries=True, project_guard=check_project, paper_groups=paper_groups)
            if result.get('ok') and result.get('result', {}).get('id') in self.assist.jobs:
                self._discard_task_plans()
                self.assist_tasks = {result['result']['id']: {'target': target, 'record_ids': sorted(sources),
                    'source_stamps': {key: link['sha256'] for key, link in sources.items()},
                    'target_name': path.name, 'target_version': digest(path)}}
            return result
        result = self._call(operation)
        return result.get('result') if result.get('ok') else result

    def _assist_task(self, identifier):
        entry = self.assist_tasks.get(identifier) if type(identifier) is str else None
        job = self.assist.jobs.get(identifier) if entry else None
        if not job or job['binding'] != self.assist._binding()[2] or job['status'] != 'completed':
            fail('professional_assist_task_invalid')
        if entry['target'] != self.target:
            fail('professional_project_changed')
        self.assist._fresh(job)
        return job, entry

    def _adopt_suggestion(self, editor, path, binding, payload):
        if set(payload) != {'task_id', 'operation'}:
            fail()
        job, entry = self._assist_task(payload['task_id'])
        suggestion = next((row for row in job['result']['suggestions'] if self.target and row['locator_id'] == self.target['locator_id']), None)
        if not suggestion or any(key not in self.evidence for key in suggestion['evidence_ids']):
            fail('professional_evidence_invalid')
        return self._queue(editor, path, binding, {'operation': payload['operation'], 'text': suggestion['new_text'],
            'evidence_ids': suggestion['evidence_ids']})

    def confirm_queue(self, identifier, approved=False):
        def operation(editor, path, binding):
            selected = [row for row in self.project['drafts'] if row['status'] == 'accepted']
            if len(self.project['outputs']) >= 32:
                fail('professional_project_capacity')
            result = self.structure.confirm(identifier, approved)
            if result.get('result', {}).get('status') == 'saved':
                keys = sorted({key for row in selected for key in row['record_ids']})
                self.project['outputs'].append({'receipt': result['result']['save_id'], 'record_ids': keys,
                    'source_stamps': {key: self._entry(key)['link']['sha256'] for key in keys},
                    'target_name': path.name, 'target_version': digest(path), 'kind': 'document'})
                for row in selected:
                    row['status'] = 'delivered'
                    self.live_drafts.pop(row['id'], None)
                self._changed()
            return result
        result = self._call(operation)
        return result.get('result') if result.get('ok') else result

    def _export(self, editor, path, binding, payload):
        if not payload.keys() <= {'offset'}:
            fail()
        offset = payload.get('offset', 0)
        if type(offset) is not int or not 0 <= offset <= len(self.project['entries']):
            fail()
        values, selected = list(self.project['entries'].values()), []
        for entry in values[offset:offset + 100]:
            trial = selected + [entry['csl']]
            if len(json.dumps(json.dumps(trial, ensure_ascii=False, indent=2), ensure_ascii=False).encode('utf8')) > 150000:
                break
            selected = trial
        if offset < len(values) and not selected:
            fail('professional_project_capacity')
        return {'csl_json': json.dumps(selected, ensure_ascii=False, indent=2), 'count': len(selected),
                'next_offset': offset + len(selected) if offset + len(selected) < len(values) else None,
                'metadata_only': True, 'permissions_exported': False}

    def plan_save(self):
        def operation(editor, path, binding):
            if not self.dirty:
                fail('professional_project_unchanged')
            if self.receipt:
                self.archive.inspect(self.receipt)
            document = deepcopy(self.project)
            document['generation'] += 1
            document['parent_receipt'] = self.receipt
            document = validate(document)
            result = self.archive.plan_project(document, self._epoch)
            self.plans[result['plan_id']] = {'revision': self.revision, 'document': document, 'parent': self.receipt}
            return result
        return self._call(operation)

    def confirm_save(self, identifier, approved=False):
        def operation(editor, path, binding):
            entry = self.plans.pop(identifier, None) if type(identifier) is str else None
            if not entry or entry['revision'] != self.revision or type(approved) is not bool:
                fail('professional_project_changed')
            if approved and entry['parent']:
                self.archive.inspect(entry['parent'])
            result = self.archive.confirm(identifier, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            if result['status'] == 'saved':
                self.project = entry['document']
                self.receipt = result['save_id']
                self.dirty = False
            return result
        return self._call(operation)

    def restore(self, receipt):
        def operation(editor, path, binding):
            if type(receipt) is str and receipt in self.history_views:
                self.archive._saved_check(self.history_views[receipt])
            document, saved = self.archive.inspect(receipt)
            if document['scope'] != self.project['scope']:
                fail('professional_project_scope')
            self.archive._saved_check(saved)
            self._retire_assist()
            self.archive.invalidate()
            self.archive = ProjectArchive(editor.space)
            self.project = document
            self.receipt = receipt
            self.revision += 1
            self.dirty = False
            self.evidence.clear()
            self.captures.clear()
            self.capture_windows.clear()
            self.live_drafts.clear()
            self.plans.clear()
            self.history_views.clear()
            for row in self.project['drafts']:
                row['status'] = 'recheck'
            return {**self._snapshot(), 'source_status': 'registered_sources_require_explicit_revalidation', 'write_approval_restored': False}
        return self._call(operation)

    def _validate_write(self, proof):
        editor, path, binding = self._binding()
        self._ensure(editor, path)
        if proof != {'project': self.project['id'], 'revision': self.revision, 'digest': payload_hash(self.project)}:
            fail('professional_project_changed')
        for row in self.project['drafts']:
            if row['status'] == 'accepted':
                if self.live_drafts.get(row['id']) != binding or row['target']['version'] != digest(path):
                    fail('professional_project_recheck')
                for key in row['record_ids']:
                    self._admit(key)
                    self._check_link(editor, self._entry(key))

    def plan_queue(self, options=None):
        def operation(editor, path, binding):
            values = options or {}
            if type(values) is not dict or not values.keys() <= {'bibliography_anchor', 'authorized', 'bibliography', 'number_records'}:
                fail()
            selected = [row for row in self.project['drafts'] if row['status'] == 'accepted']
            if not selected or any(self.live_drafts.get(row['id']) != binding for row in selected):
                fail('professional_project_recheck')
            proof = {'project': self.project['id'], 'revision': self.revision, 'digest': payload_hash(self.project)}
            self._validate_write(proof)
            counts = {kind: sum(row['operation'] == kind for row in selected) for kind in ('replace', 'insert', 'comment')}
            if counts['insert'] > 4 or counts['comment'] > 8 or len({key for row in selected for key in row['record_ids']}) > 8:
                fail('professional_project_transaction_capacity')
            request = {'source_sha256': digest(path), 'insertions': [], 'comments': [], 'replacements': [],
                       'numeric_reviews': list(dict.fromkeys(key for row in selected for key in row['numeric_reviews']))}
            request.update(values)
            for key in values.get('number_records', {}).values():
                self._admit(key)
            for row in selected:
                target = row['target']
                if row['operation'] == 'replace':
                    request['replacements'].append({'locator_id': target['locator_id'], 'expected_text_sha256': target['text_sha256'], 'new_text': row['text'], 'reason': '项目工作台明确审阅'})
                elif row['operation'] == 'comment':
                    start = target['text'].find(row['quote'])
                    if start < 0 or target['text'].count(row['quote']) != 1:
                        fail('professional_comment_range')
                    request['comments'].append({'anchor': target['locator_id'], 'start': start, 'end': start + len(row['quote']), 'quote': row['quote'], 'text': row['text']})
                else:
                    attrs = {}
                    for key in row['evidence_ids']:
                        evidence = self.evidence.get(key)
                        if not evidence:
                            fail('professional_evidence_invalid')
                        if not evidence.get('capture_id'):
                            fail('professional_project_capture_required')
                        attr = attrs.setdefault(evidence['capture_id'], {'capture_id': evidence['capture_id'], 'evidence_ids': [],
                            'relation': row['relation'], 'applicability': row['applicability'], 'limitation': row['limitation'],
                            'assessment': 'user_reviewed', 'identity_reviewed': True})
                        attr['evidence_ids'].append(key)
                    request['insertions'].append({'anchor': target['locator_id'], 'side': row['side'], 'text': row['text'], 'attributions': list(attrs.values())})
            return self.structure.plan_structure(request, proof)
        result = self._call(operation)
        return result.get('result') if result.get('ok') else result

    def plan_remove(self, receipt):
        def operation(editor, path, binding):
            result = self.archive.plan_removal(receipt, self._epoch)
            self.plans[result['rollback_id']] = {'revision': self.revision, 'removal': receipt}
            return result
        return self._call(operation)

    def confirm_remove(self, identifier, approved=False):
        def operation(editor, path, binding):
            entry = self.plans.pop(identifier, None) if type(identifier) is str else None
            if not entry or 'removal' not in entry or entry['revision'] != self.revision or type(approved) is not bool:
                fail('professional_project_changed')
            result = self.archive.confirm_rollback(identifier, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            elif result['status'] == 'rolled_back' and self.receipt == entry['removal']:
                self.receipt = ''
                self.dirty = True
            return result
        return self._call(operation)

    def plan_task(self, task_id):
        def operation(editor, path, binding):
            job, context = self._assist_task(task_id)
            if len(self.project['tasks']) >= 32:
                fail('professional_project_capacity')
            keys = context['record_ids']
            if len(keys) > 8:
                fail('professional_project_transaction_capacity')
            result = self.assist.save_task({'task_id': task_id, 'drafts': []})
            if result.get('ok'):
                self.plans[result['result']['plan_id']] = {**deepcopy(context), 'task': True,
                    'task_id': task_id, 'revision': self.revision}
            return result
        result = self._call(operation)
        return result.get('result') if result.get('ok') else result

    def confirm_task(self, identifier, approved=False):
        def operation(editor, path, binding):
            entry = self.plans.pop(identifier, None) if type(identifier) is str else None
            if not entry or not entry.get('task') or entry['revision'] != self.revision:
                fail('professional_project_changed')
            if approved:
                self._assist_task(entry['task_id'])
                if len(self.project['tasks']) >= 32:
                    fail('professional_project_capacity')
            result = self.assist.confirm_task(identifier, approved)
            if result.get('result', {}).get('status') == 'saved':
                self.project['tasks'].append({key: entry[key] for key in ('record_ids', 'source_stamps', 'target_name', 'target_version')} |
                                            {'receipt': result['result']['save_id'], 'kind': 'task'})
                self._changed()
            return result
        result = self._call(operation)
        return result.get('result') if result.get('ok') else result
