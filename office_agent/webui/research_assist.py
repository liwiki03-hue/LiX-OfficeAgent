from copy import deepcopy
import json
import re
import threading
import time
import uuid

from pydantic import Field

from .professional import Professional
from ..assisted_edits import AssistedDocx
from ..grounded_edits import seal, validate_source
from ..hybrid_ranking import rankings, fuse
from ..professional_edits import LocalChange, _protect, snapshot
from ..provenance import identifier
from ..rag import SECRET
from ..research_model import ResearchModelError
from ..research_contracts import EvidenceCard, Suggestion, Suggestions, ResearchPlan
from ..writes import Strict, GuardError, digest


class ResearchRequest(Strict):
    question: str = Field(min_length=1, max_length=200)
    mode: str = Field(pattern='^(polish|supplement|critique)$')
    source_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    locator_ids: list[str] = Field(min_length=1, max_length=12)
    capture_ids: list[str] = Field(default_factory=list, max_length=8)
    evidence_save_id: str = Field(default='', pattern=r'^(|[0-9a-f]{32})$')
    protected_terms: list[str] = Field(default_factory=list, max_length=30)
    outbound_approved: bool
    max_calls: int = Field(default=3, ge=1, le=3)
    compute_budget: int = Field(default=60, ge=5, le=60)


class ResearchAssist(Professional):
    def __init__(self, lock, scope, editor, selected, grounded, rag, model_factory, retriever_factory=None):
        super().__init__(lock, scope, editor, selected)
        self.grounded, self.rag, self.model_factory = grounded, rag, model_factory
        self.retriever_factory = retriever_factory
        self._retrieval_disabled = False
        self._threads = set()
        self.jobs = {}
        self._task_archive = None
        self._task_plans = {}

    def invalidate(self):
        for job in self.jobs.values():
            job['cancel'].set()
            self.grounded.literature.unpin(job['id'])
            self.rag._research_dependencies.pop(job['id'], None)
        for key in self._pending:
            self.grounded.literature.unpin(key)
            self.rag._research_dependencies.pop(key, None)
        self.jobs.clear()
        self._task_plans.clear()
        if self._task_archive:
            self._task_archive.invalidate()
        super().invalidate()

    def _archive(self, editor):
        from ..research_archive import ResearchArchive
        if self._task_archive is None or self._task_archive.root != editor.space.root:
            self._task_archive = ResearchArchive(editor.space)
        return self._task_archive

    def save_task(self, payload):
        def operation(editor, path, binding):
            if type(payload) is not dict or set(payload) != {'task_id', 'drafts'}:
                raise GuardError('professional_invalid_request')
            job = self.jobs.get(payload['task_id'])
            if not job or job['binding'] != binding or job['status'] != 'completed':
                raise GuardError('professional_assist_task_invalid')
            self._fresh(job)
            drafts = payload['drafts']
            known = {row['locator_id'] for row in job['targets']}
            if type(drafts) is not list or len(drafts) > 12:
                raise GuardError('professional_invalid_request')
            for row in drafts:
                if (type(row) is not dict or set(row) != {'locator_id', 'new_text', 'accepted', 'reviewed'} or row['locator_id'] not in known or
                        type(row['new_text']) is not str or not 1 <= len(row['new_text']) <= 2000 or
                        type(row['accepted']) is not bool or type(row['reviewed']) is not bool):
                    raise GuardError('professional_invalid_request')
            task = self._public(job)
            task['request']['outbound_approved'] = False
            from ..professional_edits import parser_identity
            data = {'format': 'research-task-v1', 'task': task, 'drafts': drafts, 'source_parser': parser_identity()}
            encoded = json.dumps(data, ensure_ascii=False, allow_nan=False)
            if SECRET.search(encoded):
                raise GuardError('professional_sensitive_text')
            plan = self._archive(editor).plan_payload(encoded.encode('utf8'), self._epoch, {'kind': 'research-task-v1'})
            self._task_plans[plan['plan_id']] = {'binding': binding, 'task_id': job['id']}
            return plan
        return self._call(operation)

    def confirm_task(self, plan_id, approved=False):
        def operation(editor, path, binding):
            entry = self._task_plans.pop(plan_id, None) if type(plan_id) is str else None
            if not entry or entry['binding'] != binding or type(approved) is not bool:
                raise GuardError('professional_plan_invalid')
            job = self.jobs.get(entry['task_id'])
            if approved:
                if not job:
                    raise GuardError('professional_assist_task_invalid')
                self._fresh(job)
            result = self._archive(editor).confirm(plan_id, self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            if result['status'] == 'saved':
                archives = {item['proof']['archive']['save_id'] for item in job['proofs'] if item['proof'].get('archive')}
                self.grounded.literature.pin(plan_id, archives, state='output')
                self._pin_rag(plan_id, job['request']['evidence_save_id'], output=True)
            return result
        return self._call(operation)

    def restore_task(self, save_id):
        def operation(editor, path, binding):
            data = self._archive(editor).restore(save_id)
            task = data['task']
            if len(json.dumps(data, ensure_ascii=False).encode('utf8')) > 180000:
                raise GuardError('professional_response_capacity')
            request = ResearchRequest.model_validate(task['request'])
            changes = []
            if task['source_name'] != path.name or request.source_sha256 != digest(path):
                changes.append('source_changed')
            from ..professional_edits import parser_identity
            if data['source_parser'] != parser_identity():
                changes.append('source_parser_changed')
            for evidence in task['evidence'].values():
                try:
                    if evidence.get('capture_id'):
                        found = next((value for value in self.grounded._captures.values() if any(row['id'] == evidence['id'] for row in value['proof']['fragments'])), None)
                        if not found:
                            raise ValueError()
                        validate_source(found['proof'], seal(found['proof']), path)
                        if (evidence.get('coverage') or {}).get('parser_identity') != found['proof'].get('coverage', {}).get('parser_identity'):
                            changes.append('evidence_parser_changed')
                    else:
                        raise ValueError()
                except Exception:
                    changes.append('evidence_or_index_requires_revalidation')
            try:
                if self.model_factory().identity != task['model']:
                    changes.append('model_configuration_changed')
            except Exception:
                changes.append('model_configuration_unavailable')
            retrieval = task.get('retrieval') or {}
            if retrieval.get('embedding_identity'):
                try:
                    models = self.retriever_factory() if self.retriever_factory else None
                    if not models or models[0].identity != retrieval['embedding_identity'] or retrieval.get('reranker_identity') and (not models[1] or models[1].identity != retrieval['reranker_identity']):
                        changes.append('retrieval_model_artifacts_changed')
                except Exception:
                    changes.append('retrieval_model_artifacts_unavailable')
            task['status'] = 'restored'
            task['stage'] = 'comparison_only_requires_fresh_admission'
            task['request']['outbound_approved'] = False
            task['stop_reason'] = 'restored_without_authority'
            for row in (task.get('result') or {}).get('suggestions', []):
                row['assessment'] = 'unverified'
            resumed = None
            if not changes and task.get('result'):
                captured = {}
                for key, value in self.grounded._captures.items():
                    if value['binding'] == self.grounded._binding()[2] and value['session'] == self.grounded.literature.session:
                        for row in value['proof']['fragments']:
                            if row['id'] in task['evidence']:
                                captured[key] = value
                request.capture_ids = list(captured)
                request.evidence_save_id = ''
                from ..advanced_writes import _parts
                from ..professional_edits import _inventory
                _, inventory = _inventory(_parts(path), request.source_sha256)
                current_targets = {row['locator_id']: row for row in inventory}
                if any(row['locator_id'] not in current_targets or row['text_sha256'] != current_targets[row['locator_id']]['text_sha256'] for row in task['targets']):
                    changes.append('target_changed')
                else:
                    token = uuid.uuid4().hex
                    checked_targets = [{key: current_targets[row['locator_id']][key] for key in ('locator_id', 'location', 'text', 'text_sha256')} | {'index': row['index']} for row in task['targets']]
                    job = {**task, 'id': token, 'targets': checked_targets, 'status': 'completed', 'stage': 'restored_requires_new_human_review', 'binding': binding, 'path': path,
                           'request': request.model_dump(), 'proofs': [{'proof': deepcopy(value['proof']), 'seal': seal(value['proof']), 'capture_id': key} for key, value in captured.items()],
                           'rag_stamp': None, 'cancel': threading.Event()}
                    self._fresh(job)
                    if len(self.jobs) >= 8:
                        raise GuardError('professional_capacity')
                    self.jobs[token] = job
                    resumed = self._public(job)
            return {**data, 'task': task, 'resumable_task': resumed, 'stale_reasons': list(dict.fromkeys(changes)),
                    'restores_approval': False, 'restores_rollback_authority': False,
                    'validation': 'source_and_parser_checked_evidence_requires_current_capture_model_configuration_only_not_weight_attestation'}
        return self._call(operation)

    def rollback_task(self, save_id, approved=False):
        def operation(editor, path, binding):
            archive = self._archive(editor)
            plan = archive.plan_rollback(save_id, self._epoch)
            result = archive.confirm_rollback(plan['rollback_id'], self._epoch, approved)
            if result['status'] == 'uncertain':
                self._disabled = True
            if result['status'] == 'rolled_back':
                self.grounded.literature.unpin(save_id)
                self.rag._research_dependencies.pop(save_id, None)
            return result
        return self._call(operation)

    def _current(self, job):
        authenticated, busy, token = self._scope()
        return authenticated and not busy and token == job['binding'][0] and self._epoch == job['binding'][2] and not job['cancel'].is_set()

    def _pin_rag(self, token, save_id, output=False):
        if save_id:
            self.rag._research_dependencies[token] = {'save_id': save_id, 'expires_at': None if output else time.monotonic() + 300}

    def _fresh(self, job):
        if callable(job.get('project_guard')):
            job['project_guard']()
        if job['status'] == 'running' and time.monotonic() > job.get('deadline', float('inf')):
            raise GuardError('professional_assist_budget_expired')
        if not self._current(job) or digest(job['path']) != job['request']['source_sha256']:
            raise GuardError('professional_source_changed')
        for item in job['proofs']:
            validate_source(item['proof'], item['seal'], job['path'])
            if not self.grounded._included(item['proof']):
                raise GuardError('professional_evidence_invalid')
        if job['request']['evidence_save_id']:
            _, stamp = self.rag._approved_bundle(job['request']['evidence_save_id'])
            if stamp != job['rag_stamp']:
                raise GuardError('professional_evidence_changed')

    def _public(self, job):
        return deepcopy({key: job[key] for key in ('id', 'status', 'stage', 'request', 'source_name', 'model', 'calls',
                                                  'rounds', 'stop_reason', 'degradation', 'evidence', 'result', 'targets', 'retrieval', 'failure_stage')})

    def _discard_pending(self):
        for key, value in self._pending.items():
            value['editor'].plans.pop(key, None)
            self.grounded.literature.unpin(key)
            self.rag._research_dependencies.pop(key, None)
        self._pending.clear()

    def _validate_confirmation(self, plan):
        from ..assisted_edits import validate_review
        job = self.jobs.get(plan.request.review['task_id'])
        if not job or job['status'] != 'completed':
            raise GuardError('professional_assist_task_invalid')
        self._fresh(job)
        validate_review(plan.request)

    def read(self, offset=0):
        def operation(editor, path, binding):
            source = snapshot(editor, path, offset=offset, limit=40)
            captures = [{'id': key, 'title': value['proof']['csl']['title'], 'coverage': value['proof'].get('coverage'),
                         'ocr_reviewed': value['proof'].get('ocr', {}).get('reviewed'), 'parser': value['proof']['document']['parser_version']}
                        for key, value in self.grounded._captures.items()
                        if value['binding'] == self.grounded._binding()[2] and value['session'] == self.grounded.literature.session]
            return {**source, 'captures': captures, 'model_calls_on_read': 0}
        return self._call(operation)

    def start(self, payload, preserve_queries=False, project_guard=None, paper_groups=None):
        def operation(editor, path, binding):
            if self._retrieval_disabled:
                raise GuardError('professional_assist_retrieval_disabled')
            self._threads = {thread for thread in self._threads if thread.is_alive()}
            if self._threads:
                raise GuardError('busy')
            request = ResearchRequest.model_validate_json(json.dumps(payload))
            if request.outbound_approved is not True or SECRET.search(json.dumps(payload, ensure_ascii=False)):
                raise GuardError('professional_assist_approval_required')
            if any(job['status'] == 'running' for job in self.jobs.values()):
                raise GuardError('busy')
            from ..advanced_writes import _parts
            from ..professional_edits import _inventory
            if digest(path) != request.source_sha256 or path.suffix.lower() != '.docx':
                raise GuardError('professional_source_changed')
            _, rows = _inventory(_parts(path), request.source_sha256)
            by_id = {row['locator_id']: dict(row, index=index) for index, row in enumerate(rows)}
            targets = []
            if len(set(request.locator_ids)) != len(request.locator_ids):
                raise GuardError('professional_invalid_request')
            for key in request.locator_ids:
                row = by_id.get(key)
                if not row or not row['writable'] or row['style'] or not re.fullmatch(r'body:\d+/paragraph', row['location']) or SECRET.search(row['text']):
                    raise GuardError('professional_unsupported_target')
                targets.append({key: row[key] for key in ('locator_id', 'location', 'text', 'text_sha256', 'index')})
            if sum(len(row['text']) for row in targets) > 8000:
                raise GuardError('professional_capacity')
            proofs, evidence, rag_stamp = [], {}, None
            for key in request.capture_ids:
                capture = self.grounded._captures.get(key)
                if (not capture or capture['binding'] != self.grounded._binding()[2] or
                        capture['session'] != self.grounded.literature.session or not self.grounded._included(capture['proof'])):
                    raise GuardError('professional_evidence_invalid')
                proof = deepcopy(capture['proof'])
                fragments = validate_source(proof, seal(proof), path)
                proofs.append({'proof': proof, 'seal': seal(proof), 'capture_id': key})
                for fragment in fragments:
                    item = fragment.model_dump(mode='json')
                    evidence[item['id']] = {**item, 'capture_id': key, 'title': proof['csl']['title'],
                                            'literature': proof['csl'], 'coverage': proof.get('coverage')}
            if request.evidence_save_id:
                bundle, rag_stamp = self.rag._approved_bundle(request.evidence_save_id)
                for hit in bundle['hits']:
                    key = 'ev:' + identifier(hit['source_fingerprint'], hit['location'], hit['text'])
                    evidence[key] = {'id': key, 'original_text': hit['text'], 'location': hit['location'],
                                     'source_sha256': hit['source_fingerprint'], 'title': hit['source'], 'capture_id': None,
                                     'context': next((row['sections'] for row in bundle.get('contexts', []) if row['hit_id'] == hit['id']), []),
                                     'page': None, 'literature': None, 'coverage': bundle.get('retrieval')}
            if preserve_queries:
                for row in evidence.values():
                    metadata = row.get('literature') or {}
                    identity = metadata.get('DOI') or row.get('source_sha256')
                    row['diversity_key'] = (paper_groups or {}).get(identity, identity)
            from ..research_capture_retrieval import material_pool
            balanced = material_pool(evidence, [request.question], 16)
            ranked = [{'id': index, 'text': row['original_text'], 'distance': 0} for index, row in enumerate(balanced.values(), 1)]
            orders = rankings(ranked, request.question, 30)[:2] if ranked else []
            ordered_ids = fuse(orders, 16)
            order = list(dict.fromkeys(ordered_ids + [row['id'] for row in ranked]))[:16]
            values = list(balanced.values())
            picked, used = {}, 0
            for index in order:
                row = values[index - 1]
                cost = len(row['original_text'])
                if used + cost <= 10000 and not SECRET.search(json.dumps(row, ensure_ascii=False)):
                    picked[row['id']] = row
                    used += cost
            task_id = uuid.uuid4().hex
            self._discard_pending()
            job = {'id': task_id, 'status': 'running', 'stage': 'materials_validated', 'binding': binding, 'path': path,
                   'request': request.model_dump(), 'targets': targets, 'source_name': path.name, 'proofs': proofs,
                   'evidence': picked, 'evidence_pool': {key: row for key, row in evidence.items() if not SECRET.search(json.dumps(row, ensure_ascii=False))},
                   'rag_stamp': rag_stamp, 'cancel': threading.Event(), 'model': None, 'retrieval': None, 'compute_seconds': 0., 'deadline': time.monotonic() + request.compute_budget,
                   'calls': 0, 'rounds': 0, 'stop_reason': '', 'result': None, 'failure_stage': None,
                   'degradation': 'approved_saved_retrieval' if request.evidence_save_id else 'bounded_capture_reading_with_lexical_priority_not_vector_retrieval'}
            for key, value in self.jobs.items():
                if value['status'] != 'running':
                    self.grounded.literature.unpin(key)
                    self.rag._research_dependencies.pop(key, None)
            self.jobs = {key: value for key, value in self.jobs.items() if value['status'] == 'running'}
            archives = {item['proof']['archive']['save_id'] for item in proofs if item['proof'].get('archive')}
            self.grounded.literature.pin(task_id, archives)
            if request.evidence_save_id:
                self._pin_rag(task_id, request.evidence_save_id)
            job['preserve_query_context'] = preserve_queries is True
            job['project_guard'] = project_guard
            self.jobs[task_id] = job
            self._fresh(job)
            thread = threading.Thread(target=self._run, args=(job,), name='OfficeResearch-' + task_id[:8], daemon=True)
            job['thread'] = thread
            self._threads.add(thread)
            thread.start()
            return self._public(job)
        return self._call(operation)

    def _run(self, job):
        started = time.monotonic()
        try:
            if not job['evidence'] and job['request']['mode'] != 'polish':
                with self._lock:
                    self._fresh(job)
                    job.update(status='completed', stage='stopped', stop_reason='insufficient_materials',
                               result={'cards': [], 'suggestions': [], 'missing_evidence': ['未取得可核对证据；请先检索或读取相关页段。'], 'stop_reason': 'insufficient_materials'})
                return
            model = self.model_factory()
            if job['evidence'] and job['request'].get('max_calls', 3) >= 2 and not job.get('regenerate') and not job['request'].get('evidence_save_id') and callable(getattr(model, 'plan', None)):
                with self._lock:
                    self._fresh(job)
                    job.update(stage='question_and_query_planning', model=model.identity, calls=job['calls'] + 1, rounds=job['rounds'] + 1)
                    planning_input = {'question': job['request']['question'], 'mode': job['request']['mode'],
                                      'targets': job['targets'], 'catalog': [{'id': row['id'], 'title': row['title'], 'location': row['location'], 'preview': row['original_text'][:80]} for row in list(job['evidence'].values())[:16]],
                                      'partial_sources_only': True}
                planned = ResearchPlan.model_validate_json(json.dumps(model.plan(planning_input, lambda: not self._current(job) or time.monotonic() > job['deadline'])))
                if job.get('preserve_query_context'):
                    original = job['request']['question']
                    kept = [original] + [query for query in planned.queries if original.casefold() in query.casefold() and query != original]
                    planned = ResearchPlan(queries=list(dict.fromkeys(kept))[:3], read_ids=planned.read_ids)
                if not set(planned.read_ids) <= set(job['evidence']) or SECRET.search(json.dumps(planned.model_dump(), ensure_ascii=False)):
                    raise GuardError('professional_assist_response')
                with self._lock:
                    self._fresh(job)
                    job['stage'] = 'retrieving_selected_sources'
                models = None
                if self.retriever_factory:
                    try:
                        models = self.retriever_factory()
                    except Exception as error:
                        if job.get('preserve_query_context'):
                            from ..embeddings import EmbeddingError
                            if not isinstance(error, EmbeddingError) or str(error) not in {'embedding_local_not_installed', 'embedding_local_directory', 'embedding_failed'}:
                                raise
                from ..research_capture_retrieval import retrieve_captures
                def check():
                    with self._lock:
                        self._fresh(job)
                try:
                    selected, retrieval = retrieve_captures(job.get('evidence_pool', job['evidence']), job['request']['question'], planned.queries, planned.read_ids, models, check=check)
                except ValueError as error:
                    if str(error) == 'research_retrieval_cleanup_failed':
                        raise ResearchModelError('research_retrieval_cleanup_failed') from None
                    raise
                with self._lock:
                    self._fresh(job)
                    job.update(stage='original_contexts_read', evidence=selected, retrieval=retrieval, degradation=retrieval['quality'])
            with self._lock:
                self._fresh(job)
                if job['calls'] >= job['request'].get('max_calls', 3):
                    raise GuardError('professional_assist_budget_exhausted')
                job.update(stage='model_review', model=model.identity, calls=job['calls'] + 1, rounds=job['rounds'] + 1)
                targets = [row for row in job['targets'] if not job.get('regenerate') or row['locator_id'] == job['regenerate']['locator_id']]
                model_input = {'question': job['request']['question'], 'mode': job['request']['mode'],
                               'targets': targets, 'protected_terms': job['request']['protected_terms'],
                               'evidence': [{key: row.get(key) for key in ('id', 'original_text', 'context', 'location', 'page', 'title', 'source_sha256', 'extraction_quality')} |
                                            {'coverage_partial': not bool((row.get('coverage') or {}).get('complete'))} for row in job['evidence'].values()], 'coverage_is_partial': True,
                               'feedback': job.get('regenerate', {}).get('feedback', ''),
                               'budget': {'maximum_rounds': job['request'].get('max_calls', 3), 'remaining_calls': job['request'].get('max_calls', 3) - job['calls'], 'selected_fragments': len(job['evidence'])}}
            result = Suggestions.model_validate_json(json.dumps(model.suggest(model_input, lambda: not self._current(job) or time.monotonic() > job['deadline'])))
            encoded = result.model_dump(mode='json')
            if SECRET.search(json.dumps(encoded, ensure_ascii=False)):
                raise GuardError('professional_sensitive_text')
            known = set(job['evidence'])
            targets = {row['locator_id']: row for row in targets}
            if len({row.locator_id for row in result.suggestions}) != len(result.suggestions):
                raise GuardError('professional_assist_response')
            for card in result.cards:
                if not set(card.evidence_ids) <= known or (card.relation != 'insufficient' and not card.evidence_ids):
                    raise GuardError('professional_assist_response')
            encoded['cards'] = [dict(row, assessment='unverified') for row in encoded['cards']]
            checked = []
            for item in result.suggestions:
                if item.locator_id not in targets or not set(item.evidence_ids) <= known:
                    raise GuardError('professional_assist_response')
                row = item.model_dump(mode='json')
                row.update(assessment='unverified', protection_status='not_checked')
                try:
                    _protect(targets[item.locator_id]['text'], item.new_text, type('Protection', (), {
                        'preserve_numbers': True, 'preserve_formulas': True, 'preserve_citations': True,
                        'protected_literals': job['request']['protected_terms']})())
                    row['protection_status'] = 'passed'
                except GuardError as error:
                    row['protection_status'] = str(error).split(':')[0]
                if job['request']['mode'] != 'polish' and not item.evidence_ids:
                    row['needs_evidence'] = True
                checked.append(row)
            with self._lock:
                self._fresh(job)
                encoded['suggestions'] = checked
                if job.get('regenerate'):
                    old = job['regenerate']['previous']
                    locator = job['regenerate']['locator_id']
                    if len(checked) != 1 or checked[0]['locator_id'] != locator:
                        raise GuardError('professional_assist_response')
                    old['suggestions'] = [checked[0] if row['locator_id'] == locator else row for row in old['suggestions']]
                    old['cards'].extend(encoded['cards'])
                    old['missing_evidence'] = list(dict.fromkeys(old['missing_evidence'] + encoded['missing_evidence']))[:12]
                    old['stop_reason'] = encoded['stop_reason']
                    encoded = old
                    job.pop('regenerate')
                job.update(result=encoded, status='completed', stage='awaiting_human_review', stop_reason=encoded['stop_reason'])
        except Exception as error:
            with self._lock:
                code = getattr(error, 'code', None)
                safe = code if type(code) is str and code.startswith('research_') else str(error) if isinstance(error, GuardError) else 'research_model_response'
                if safe == 'research_retrieval_cleanup_failed':
                    self._retrieval_disabled = True
                if self.jobs.get(job['id']) is job and not job['cancel'].is_set() and job['binding'][2] == self._epoch:
                    previous = job.pop('regenerate', {}).get('previous')
                    job.update(status='completed' if previous else 'failed', failure_stage=job['stage'], stage='stopped', stop_reason=safe.split(':')[0], result=previous)
        finally:
            with self._lock:
                job['compute_seconds'] = job.get('compute_seconds', 0.) + time.monotonic() - started
                self._threads.discard(threading.current_thread())

    def regenerate(self, payload):
        def operation(editor, path, binding):
            if self._retrieval_disabled:
                raise GuardError('professional_assist_retrieval_disabled')
            if (type(payload) is not dict or set(payload) != {'task_id', 'locator_id', 'feedback', 'outbound_approved'} or
                    payload['outbound_approved'] is not True or type(payload['feedback']) is not str or
                    len(payload['feedback']) > 400 or SECRET.search(payload['feedback'])):
                raise GuardError('professional_invalid_request')
            job = self.jobs.get(payload['task_id'])
            if not job or job['binding'] != binding or job['status'] != 'completed' or not job['result']:
                raise GuardError('professional_assist_task_invalid')
            self._fresh(job)
            if job['calls'] >= job['request'].get('max_calls', 3) or job['rounds'] >= job['request'].get('max_calls', 3):
                raise GuardError('professional_assist_budget_exhausted')
            if job.get('compute_seconds', 0) >= job['request'].get('compute_budget', 60):
                raise GuardError('professional_assist_budget_expired')
            if payload['locator_id'] not in {row['locator_id'] for row in job['result']['suggestions']}:
                raise GuardError('professional_invalid_request')
            for key, value in list(self._pending.items()):
                value['editor'].plans.pop(key, None)
                self.grounded.literature.unpin(key)
                self.rag._research_dependencies.pop(key, None)
            self._pending.clear()
            job['regenerate'] = {'locator_id': payload['locator_id'], 'feedback': payload['feedback'], 'previous': deepcopy(job['result'])}
            job.update(status='running', stage='regenerating_one_suggestion', result=None, deadline=time.monotonic() + job['request'].get('compute_budget', 60) - job.get('compute_seconds', 0))
            thread = threading.Thread(target=self._run, args=(job,), name='OfficeResearch-' + job['id'][:8], daemon=True)
            job['thread'] = thread
            self._threads.add(thread)
            thread.start()
            return self._public(job)
        return self._call(operation)

    def status(self, task_id):
        def operation(editor, path, binding):
            job = self.jobs.get(task_id) if type(task_id) is str else None
            if not job or job['binding'] != binding:
                raise GuardError('professional_assist_task_invalid')
            if job['status'] != 'cancelled':
                self._fresh(job)
            return self._public(job)
        return self._call(operation)

    def cancel(self, task_id):
        def operation(editor, path, binding):
            job = self.jobs.get(task_id) if type(task_id) is str else None
            if not job or job['binding'] != binding:
                raise GuardError('professional_assist_task_invalid')
            job['cancel'].set()
            self.grounded.literature.unpin(task_id)
            self.rag._research_dependencies.pop(task_id, None)
            job.update(status='cancelled', stage='stopped', stop_reason='user_cancelled', result=None)
            for key in list(self._pending):
                self._pending[key]['editor'].plans.pop(key, None)
                self._pending.pop(key, None)
                self.grounded.literature.unpin(key)
                self.rag._research_dependencies.pop(key, None)
            return self._public(job)
        return self._call(operation)

    def plan(self, payload):
        def operation(editor, path, binding):
            if type(payload) is not dict or not {'task_id', 'edits'} <= set(payload) or set(payload) - {'task_id', 'edits', 'tracked', 'granular'} or type(payload.get('tracked', False)) is not bool or type(payload.get('granular', False)) is not bool or payload.get('granular') and not payload.get('tracked'):
                raise GuardError('professional_invalid_request')
            job = self.jobs.get(payload['task_id'])
            if not job or job['binding'] != binding or job['status'] != 'completed':
                raise GuardError('professional_assist_task_invalid')
            self._fresh(job)
            if job['request']['mode'] != 'polish':
                raise GuardError('professional_assist_citation_review_required')
            edits = payload['edits']
            if type(edits) is not list or not 1 <= len(edits) <= 12:
                raise GuardError('professional_invalid_request')
            targets = {row['locator_id']: row for row in job['targets']}
            changes = []
            for edit in edits:
                if (type(edit) is not dict or set(edit) != {'locator_id', 'new_text', 'reason', 'reviewed'} or
                        edit['reviewed'] is not True or edit['locator_id'] not in targets):
                    raise GuardError('professional_assist_review_required')
                row = targets[edit['locator_id']]
                changes.append(LocalChange(locator_id=edit['locator_id'], expected_text_sha256=row['text_sha256'],
                                           new_text=edit['new_text'], reason=edit['reason']))
            known = set(job['evidence'])
            selected_proofs = []
            for item in job['proofs']:
                proof = {**item['proof'], 'fragments': [row for row in item['proof']['fragments'] if row['id'] in known]}
                if proof['fragments']:
                    selected_proofs.append({'proof': proof, 'seal': seal(proof)})
            review = {'task_id': job['id'], 'question': job['request']['question'], 'mode': 'polish',
                      'source_sha256': job['request']['source_sha256'], 'edits': deepcopy(edits), 'model': job['model'],
                      'proofs': selected_proofs, 'selected_evidence': job['evidence'], 'protected_terms': job['request']['protected_terms'],
                      'rag_receipt': {'save_id': job['request']['evidence_save_id'], 'stamp': json.loads(json.dumps(job['rag_stamp']))} if job['rag_stamp'] else None,
                      'semantic_check': 'user_review_required_not_machine_proof'}
            request_type, extra = AssistedDocx, {}
            if payload.get('tracked'):
                from datetime import datetime, timezone
                from ..tracked_edits import TrackedAssistedDocx
                request_type = TrackedAssistedDocx
                review['revision_mode'] = 'native_plain_paragraph'
                extra['revision_date'] = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
                extra['revision_scope'] = 'word' if payload.get('granular') else 'paragraph'
            request = request_type(path=str(path), source_sha256=job['request']['source_sha256'], changes=changes,
                                   protected_literals=job['request']['protected_terms'], review=review, review_seal=seal(review), **extra)
            for key, value in self._pending.items():
                value['editor'].plans.pop(key, None)
                self.grounded.literature.unpin(key)
                self.rag._research_dependencies.pop(key, None)
            self._pending.clear()
            plan = editor.plan(request)
            self._pending[plan.id] = {'editor': editor, 'binding': binding, 'created': time.monotonic()}
            archives = {item['proof']['archive']['save_id'] for item in job['proofs'] if item['proof'].get('archive')}
            self.grounded.literature.pin(plan.id, archives)
            if job['request']['evidence_save_id']:
                self._pin_rag(plan.id, job['request']['evidence_save_id'])
            return {**plan.public(), 'items': json.loads(plan.after)['items'], 'wrote_nothing': True, 'expires_in_seconds': 300}
        return self._call(operation)

    def confirm(self, plan_id, confirmed=False):
        with self._lock:
            reply = super().confirm(plan_id, confirmed)
            status = reply.get('result', {}).get('status')
            if status == 'saved':
                if plan_id in self.grounded.literature._dependencies:
                    self.grounded.literature._dependencies[plan_id]['state'] = 'output'
                lease = self.rag._research_dependencies.get(plan_id)
                if lease:
                    lease['expires_at'] = None
            elif status != 'uncertain' and plan_id not in self._pending:
                self.grounded.literature.unpin(plan_id)
                self.rag._research_dependencies.pop(plan_id, None)
            return reply

    def rollback(self, save_id, confirmed=False):
        with self._lock:
            reply = super().rollback(save_id, confirmed)
            if reply.get('result', {}).get('status') == 'rolled_back':
                self.grounded.literature.unpin(save_id)
                self.rag._research_dependencies.pop(save_id, None)
            return reply
