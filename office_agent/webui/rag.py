"""Authenticated host-only RAG management. No client roots, rows or session IDs.

Synchronous admitted IO serializes with logout; not an interruptible transaction.
Epoch callback reaches core fresh-source boundaries, including after embedding.
"""
import json
import re
import time
import uuid

from ..rag import RagError
from ..rag_sources import SourceError
from ..rag_evidence import EvidenceError, CODES as EVIDENCE_CODES
from ..embeddings import EmbeddingError
from ..settings_store import SettingsError
from ..vectorstore import VectorStoreError

_CODES = {
    RagError: frozenset({'rag_invalid','rag_sensitive_content','rag_sensitive_source','rag_capacity','rag_uncertain','rag_source_changed','rag_model_mismatch','rag_embedding_failed','rag_query_failed','rag_id_collision'}),
    SourceError: frozenset({'source_changed','invalid','capacity','unsupported','sensitive','pdf_text_missing','pdf_text_truncated'}),
    EvidenceError: EVIDENCE_CODES,
    EmbeddingError: frozenset({'embedding_invalid','embedding_not_configured','embedding_api_configuration','embedding_local_directory','embedding_local_changed','embedding_local_not_installed','embedding_local_capacity','embedding_local_modules','embedding_local_weights','embedding_input','embedding_response','embedding_timeout','embedding_failed'}),
    SettingsError: frozenset({'settings_invalid','settings_path_changed','settings_file_invalid','settings_source_changed','settings_io_failed','settings_uncertain'}),
    VectorStoreError: frozenset({'invalid_input','unsafe_path','invalid_dimension','stale_root','unsafe_sidecar','stale_database','database_limit','database_error','model_mismatch','invalid_database','capacity_limit','verification_failed','invalid_vector','uncertain_state','authority_limit','document_limit','document_missing','invalid_authority','expired_authority','session_mismatch','confirmation_required','stale_owned_file','audit_limit','evidence_missing','stale_source'}),
}


_CODES[VectorStoreError] |= frozenset({'index_profile_mismatch'})
_CODES[RagError] |= frozenset({'rag_index_current'})
_CODES[RagError] |= frozenset({'rag_corpus_stale'})
_CODES[RagError] |= frozenset({'rag_evidence_in_use'})


def _pdf_coverage(value):
    if (type(value) is not dict or set(value) != {'start_page','end_page','total_pages','complete','text_order'} or
            any(type(value[key]) is not int for key in ('start_page','end_page','total_pages')) or
            not 1 <= value['start_page'] <= value['end_page'] <= value['total_pages'] <= 2000 or
            value['end_page'] - value['start_page'] >= 20 or type(value['complete']) is not bool or
            value['complete'] != (value['start_page'] == 1 and value['end_page'] == value['total_pages']) or
            value['text_order'] != 'content_order_unreviewed'):
        raise ValueError()


class Rag:
    def __init__(self, lock, scope, factory, selected_name):
        self._lock, self._scope, self._factory, self._selected_name = lock, scope, factory, selected_name
        self._session = uuid.uuid4().hex
        self._store = self._store_scope = None
        self._disabled = False
        self._research_dependencies = {}

    def invalidate(self):
        with self._lock:
            self._session = uuid.uuid4().hex
            previous, self._store, self._store_scope = self._store, None, None
            if previous is not None:
                try:
                    previous.invalidate()
                except Exception:
                    self._disabled = True

    def _lease(self, session, token):
        authenticated, busy, current = self._scope()
        return authenticated is True and not busy and current == token and self._session == session

    def _result(self, result):
        if type(result) is not dict or len(result) != 1 or not result.keys() <= {'plan','result','snapshot','evidence','research','context','corpus','ranges'}:
            raise TypeError()
        if 'ranges' in result:
            value = result['ranges']
            if type(value) is not dict or value.get('wrote_nothing') is not True:
                raise ValueError()
            for key in ('pages', 'selected_pages', 'parsed_pages', 'indexed_pages', 'failed_pages', 'native_pages', 'ocr_pages', 'gaps'):
                if type(value.get(key)) is not list or len(value[key]) > 2000 or any(type(n) is not int or not 1 <= n <= 2000 for n in value[key]):
                    raise ValueError()
        if 'corpus' in result:
            value = result['corpus']
            if (type(value) is not dict or set(value) != {'status','documents','changed','removed','unchanged','inventory','corpus_sha256','wrote_nothing'} or
                    value['wrote_nothing'] is not True or value['status'] not in {'current','needs_sync'} or
                    type(value['inventory']) is not list or len(value['inventory']) > 256):
                raise ValueError()
        if 'research' in result:
            value = result['research']
            if type(value) is dict and 'pdf_coverage' in value:
                _pdf_coverage(value['pdf_coverage'])
                value = {key: item for key, item in value.items() if key != 'pdf_coverage'}
            if (type(value) is not dict or set(value) != {'status','chunk_count','revision','source','source_fingerprint','wrote_nothing'} or
                    value['status'] not in {'current','missing','stale','profile_mismatch'} or value['wrote_nothing'] is not True or
                    type(value['chunk_count']) is not int or not 0 <= value['chunk_count'] <= 500 or
                    type(value['revision']) is not int or value['revision'] < 0 or
                    type(value['source']) is not str or len(value['source']) > 128 or
                    type(value['source_fingerprint']) is not str or not re.fullmatch('[a-f0-9]{64}', value['source_fingerprint'])):
                raise ValueError()
        if 'context' in result:
            value = result['context']
            if (type(value) is not dict or set(value) != {'hit_id','source','source_fingerprint','retrieved_location','context_version',
                    'tokenizer','token_budget','tokens','sections','writable'} or value.get('writable') is not False or
                    type(value['hit_id']) is not int or value['hit_id'] < 1 or type(value['source']) is not str or len(value['source']) > 128 or
                    type(value['source_fingerprint']) is not str or not re.fullmatch('[a-f0-9]{64}', value['source_fingerprint']) or
                    type(value.get('sections')) is not list or
                    not 1 <= len(value['sections']) <= 4 or type(value.get('tokens')) is not int or not 1 <= value['tokens'] <= 768):
                raise ValueError()
            for section in value['sections']:
                if (type(section) is not dict or set(section) != {'role','location','char_start','char_end','text','tokens','truncated'} or
                        section['role'] not in {'parent','heading','table_first_row','adjacent'} or
                        type(section['text']) is not str or len(section['text']) > 8000 or
                        type(section['location']) is not str or len(section['location']) > 128 or
                        type(section['truncated']) is not bool):
                    raise ValueError()
        if 'plan' in result:
            plan = result['plan']
            if type(plan) is dict and 'pdf_coverage' in plan:
                _pdf_coverage(plan['pdf_coverage'])
                plan = {key: item for key, item in plan.items() if key != 'pdf_coverage'}
            if type(plan) is dict and 'retrieval' in plan:
                info = plan['retrieval']
                if (type(info) is not dict or set(info) != {'queries','candidates','rerank','query_expansion','query_text_is_evidence',
                        'ranking','elapsed_ms','support_assessed','scope','documents','corpus_sha256','coverage_complete','lexical_version','duplicates_removed'} or
                        info['query_text_is_evidence'] is not False or info['support_assessed'] is not False or
                        info['scope'] != 'directory' or info['ranking'] not in {'hybrid','dense','lexical'} or
                        info['query_expansion'] not in {'literal_clauses','original'} or
                        info['rerank'] not in {'not_requested','not_configured','completed','timeout','unavailable','invalid_response','capacity','configuration_invalid'} or
                        any(type(info[key]) is not int for key in ('queries','candidates','elapsed_ms','documents')) or
                        not 1 <= info['queries'] <= 3 or not 0 <= info['candidates'] <= 50 or not 0 <= info['documents'] <= 20 or
                        type(info['corpus_sha256']) is not str or not re.fullmatch('[a-f0-9]{64}', info['corpus_sha256'])):
                    raise ValueError()
                plan = dict(plan)
                plan.pop('retrieval')
            schemas = (
                {'plan_id','operation','selected_pages','reused_pages','new_pages','removed_pages','removed_page_numbers','chunk_count','total_pages','generation','requires_confirmation','wrote_nothing'},
                {'plan_id','operation','confirmation_required','changed','removed','unchanged','chunk_count','inventory','embedding','expires_in_seconds','wrote_nothing'},
                {'plan_id','doc_id','chunk_count','source_fingerprint','embedding','provider_call_after_confirmation','requires_confirmation'},
                {'plan_id','operation','hit_count','byte_count','confirmation_required','wrote_nothing','payload_sha256','expires_in_seconds'},
                {'plan_id','operation','hit_count','byte_count','confirmation_required','wrote_nothing','payload_sha256','expires_in_seconds','directory'},
                {'rollback_id','operation','confirmation_required','wrote_nothing','expires_in_seconds'},
                {'rollback_id','operation','confirmation_required','revision','chunk_count','first_absent_restored_empty'},
            )
            if type(plan) is not dict or not any(plan.keys() == keys for keys in schemas):
                raise ValueError()
            identifier = plan.get('plan_id',plan.get('rollback_id'))
            if type(identifier) is not str or not re.fullmatch('[a-f0-9]{32}',identifier):
                raise ValueError()
            if plan.get('requires_confirmation',plan.get('confirmation_required')) is not True:
                raise ValueError()
            if 'embedding' in plan:
                identity = plan['embedding']
                if type(identity) is not dict or identity.keys() != {'channel','model','dimension'}:
                    raise ValueError()
        if 'evidence' in result:
            evidence = result['evidence']
            if (type(evidence) is not dict or evidence.keys() != {'text','offset','has_more'} or
                    type(evidence['text']) is not str or len(evidence['text']) > 2000 or
                    type(evidence['offset']) is not int or type(evidence['has_more']) is not bool):
                raise ValueError()
        # Transport budget and primitive-only finite JSON; backend is host-owned.
        if len(json.dumps(result,allow_nan=False).encode('utf-8')) > 65536:
            raise ValueError()
        outcome = result.get('result')
        if outcome is not None:
            if type(outcome) is not dict:
                raise TypeError()
            status = outcome.get('status')
            if status == 'uncertain' or outcome.get('applied','missing') is None or outcome.get('restored','missing') is None:
                self._disabled = True
                return {'result': {'status':'uncertain','applied':None,'restored':None,'wrote_nothing':False}}
            if status == 'saved':
                if (outcome.get('applied') is not True or outcome.get('wrote_nothing') is not False or
                        type(outcome.get('save_id')) is not str or not re.fullmatch('[a-f0-9]{32}',outcome['save_id'])):
                    raise ValueError()
            elif status == 'rolled_back':
                if outcome.get('restored') is not True or outcome.get('wrote_nothing') is not False:
                    raise ValueError()
            elif status == 'rejected':
                if outcome.get('wrote_nothing') is not True or outcome.get('applied',False) is not False or outcome.get('restored',False) is not False:
                    raise ValueError()
            else:
                raise ValueError()
            if 'hits' in outcome:
                if type(outcome['hits']) is not list or len(outcome['hits']) > 10:
                    raise ValueError()
                for hit in outcome['hits']:
                    if type(hit) is not dict or 'text' in hit or type(hit.get('excerpt')) is not str or len(hit['excerpt']) > 200:
                        raise ValueError()
        if result.get('snapshot',{}).get('uncertain') is True:
            self._disabled = True
        return result

    def _call(self, operation):
        with self._lock:
            try:
                authenticated, busy, token = self._scope()
                if authenticated is not True:
                    self.invalidate(); return {'ok':False,'reason':'not_authenticated'}
                if busy:
                    self.invalidate(); return {'ok':False,'reason':'busy'}
                if token is None:
                    self.invalidate(); return {'ok':False,'reason':'no_file'}
                if self._disabled:
                    return {'ok':False,'reason':'rag_service_disabled'}
                if self._store is not None and self._store_scope != token:
                    self.invalidate()
                session = self._session
                if self._store is None:
                    store = self._factory(lambda:self._lease(session,token))
                    if not self._lease(session,token):
                        store.invalidate()
                        return {'ok':False,'reason':'rag_scope_changed'}
                    self._store, self._store_scope = store, token
                result = operation(self._store,session)
                if not self._lease(session,token):
                    self._disabled = True
                    self.invalidate()
                    return {'ok':False,'reason':'rag_scope_changed','outcome_uncertain':True}
                return {'ok':True,**self._result(result)}
            except Exception as exc:
                codes = _CODES.get(type(exc),frozenset())
                code = getattr(exc,'code',None)
                if type(code) is str and code in codes:
                    if code in {'rag_uncertain','evidence_uncertain','uncertain_state','settings_uncertain'}:
                        self._disabled = True
                    return {'ok':False,'reason':code}
                self._disabled = True
                self.invalidate()
                return {'ok':False,'reason':'rag_failed','outcome_uncertain':True}

    def snapshot(self):
        return self._call(lambda store,session:{'snapshot':store.snapshot()})

    def plan_index(self):
        return self._call(lambda store,session:{'plan':store.plan_index(self._selected_name(),session)})

    def pdf_ranges_status(self, name=None):
        return self._call(lambda store, session: {'ranges': store.pdf_ranges_status(name if name is not None else self._selected_name(), session)})

    def plan_pdf_ranges(self, ranges, name=None):
        return self._call(lambda store, session: {'plan': store.plan_pdf_ranges(name if name is not None else self._selected_name(), ranges, session)})

    def query_pdf_ranges(self, question, top_k=5, name=None):
        return self._call(lambda store, session: {'plan': store.query_pdf_ranges(name if name is not None else self._selected_name(), question, session, top_k)})

    def plan_research_index(self, start_page=None, end_page=None):
        return self._call(lambda store,session:{'plan':store.plan_research_index(self._selected_name(),session, **self._range(start_page, end_page))})

    def _range(self, start_page, end_page):
        if start_page is None and end_page is None:
            return {}
        if type(start_page) is not int or type(end_page) is not int or not 1 <= start_page <= end_page <= 2000 or end_page - start_page >= 20:
            raise RagError('rag_invalid')
        return {'start_page': start_page, 'end_page': end_page}

    def plan_research_query(self, question, top_k=5, start_page=None, end_page=None):
        return self._call(lambda store,session:{'plan':store.plan_research_query(self._selected_name(),question,session,top_k=top_k, **self._range(start_page, end_page))})

    def research_status(self, start_page=None, end_page=None):
        return self._call(lambda store,session:{'research':store.research_status(self._selected_name(),session, **self._range(start_page, end_page))})

    def corpus_status(self):
        return self._call(lambda store,session:{'corpus':store.corpus_status(self._selected_name(),session)})

    def plan_corpus(self):
        return self._call(lambda store,session:{'plan':store.plan_corpus(self._selected_name(),session)})

    def confirm_corpus(self, token, confirmed=False):
        return self._call(lambda store,session:{'result':store.confirm_corpus(token,session,confirmed)})

    def plan_corpus_query(self, question, top_k=5, rerank=False, expand=False):
        return self._call(lambda store,session:{'plan':store.plan_corpus_query(self._selected_name(),question,session,top_k=top_k,rerank=rerank,expand=expand)})

    def read_research_context(self, save_id, hit_id):
        return self._call(lambda store,session:{'context':store.read_research_context(save_id,hit_id,session)})

    def confirm_index(self, plan_id, confirmed=False):
        return self._call(lambda store,session:{'result':store.confirm_index(plan_id,session,confirmed)})

    def plan_query(self, question, top_k=5):
        def plan(store,session):
            store.bind_source(self._selected_name(),session)
            return {'plan':store.plan_query(question,session,top_k=top_k)}
        return self._call(plan)

    def plan_directory_query(self, question, top_k=5):
        return self._call(lambda store, session: {'plan': store.plan_directory_query(
            self._selected_name(), question, session, top_k=top_k)})

    def confirm_query(self, plan_id, confirmed=False):
        return self._call(lambda store,session:{'result':store.confirm_query(plan_id,session,confirmed)})

    def read_evidence(self, save_id, offset=0, limit=2000):
        return self._call(lambda store,session:{'evidence':store.read_evidence(save_id,session,offset=offset,limit=limit)})

    def _approved_bundle(self, save_id):
        from ..rag_evidence import _read
        with self._lock:
            authenticated, busy, token = self._scope()
            if not authenticated or busy or self._store is None or token != self._store_scope:
                raise ValueError('research_evidence_unavailable')
            evidence = self._store._evidence
            saved = evidence._saved_for(save_id, self._session)
            evidence._fresh(saved['hits'])
            data, stamp = _read(saved['directory'] / evidence.artifact_name)
            return json.loads(data), stamp

    def plan_rollback(self, kind, save_id):
        def plan(store,session):
            if kind == 'evidence' and self._research_in_use(save_id):
                raise RagError('rag_evidence_in_use')
            if type(kind) is not str or kind not in {'index','evidence'}:
                raise RagError()
            method = store.plan_index_rollback if kind == 'index' else store.plan_evidence_rollback
            return {'plan':method(save_id,session)}
        return self._call(plan)

    def confirm_rollback(self, kind, rollback_id, confirmed=False):
        def confirm(store,session):
            if type(kind) is not str or kind not in {'index','evidence'}:
                raise RagError()
            if kind == 'evidence':
                pending = store._evidence._plans.get(rollback_id) if type(rollback_id) is str else None
                if confirmed is True and pending and self._research_in_use(pending.get('save_id')):
                    store._evidence._plans.pop(rollback_id, None)
                    raise RagError('rag_evidence_in_use')
                result = store.confirm_evidence_rollback(rollback_id,session,confirmed)
            else:
                result = store.confirm_index_rollback(rollback_id,session,confirmed)
                if type(result) is not dict or result.get('uncertain') is not False or type(result.get('revision')) is not int:
                    raise ValueError()
                result = {**result,'status':'rolled_back','restored':True,'wrote_nothing':False}
            return {'result':result}
        return self._call(confirm)

    def _research_in_use(self, save_id):
        now = time.monotonic()
        self._research_dependencies = {key: value for key, value in self._research_dependencies.items()
                                       if value['expires_at'] is None or value['expires_at'] > now}
        return any(value['save_id'] == save_id for value in self._research_dependencies.values())
