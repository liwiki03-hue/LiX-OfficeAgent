import json
import re
import uuid
import time
from copy import deepcopy

from pydantic import ValidationError

from ..literature import LiteratureError, LiteratureQuery, VERSION, group_records, csl, doi, text
from ..literature_client import LiteratureClient, LiteratureConfig
from ..literature_archive import LiteratureArchive
from ..rag_evidence import EvidenceError, _dir
from ..settings_store import SettingsStore
from ..provenance import acquired_at, SourceDocument, identifier


class Literature:
    def __init__(self, lock, scope, editor, settings_root):
        self.lock, self.scope, self.editor, self.settings_root = lock, scope, editor, settings_root
        self.factory = self._client
        self.binding = None
        self.root_identity = None
        self.session = uuid.uuid4().hex
        self.client = None
        self.records, self.reports, self.decisions, self.plans, self.saves = {}, [], {}, {}, {}
        self.archives = {}
        self.running = False
        self.disabled = False
        self._dependencies = {}

    def _client(self):
        settings = SettingsStore(self.settings_root)
        values = settings._values(settings._read(settings.root / '.env'))
        return LiteratureClient(LiteratureConfig.from_values(values))

    def invalidate(self):
        with self.lock:
            for archive in self.archives.values():
                archive.invalidate()
            if self.client:
                self.client.invalidate()
            self.binding = self.client = None
            self.root_identity = None
            self.session = uuid.uuid4().hex
            self.records, self.reports, self.decisions, self.plans, self.saves = {}, [], {}, {}, {}
            self.archives = {}

    def _bind(self):
        authenticated, busy, token = self.scope()
        if not authenticated:
            self.invalidate()
            raise LiteratureError('not_authenticated')
        if busy:
            raise LiteratureError('busy')
        editor = self.editor()
        if token is None or editor is None:
            self.invalidate()
            raise LiteratureError('no_file')
        current = (token, id(editor))
        if self.binding != current:
            self.invalidate()
            self.binding = current
            self.root_identity = _dir(editor.space.root)
        if _dir(editor.space.root) != self.root_identity:
            raise LiteratureError('literature_scope_changed')
        try:
            if editor.space.file(str(token[0])) != token[0]:
                raise ValueError()
        except Exception:
            raise LiteratureError('literature_path_refused') from None
        if self.disabled:
            raise LiteratureError('literature_uncertain')
        return editor

    def _check(self, lease):
        with self.lock:
            authenticated, busy, token = self.scope()
            if (not authenticated or busy or (token, id(self.editor())) != self.binding or self.session != lease):
                raise LiteratureError('literature_scope_changed')
            if _dir(self.editor().space.root) != self.root_identity:
                raise LiteratureError('literature_scope_changed')

    def _error(self, error):
        if isinstance(error, LiteratureError):
            return {'ok': False, 'reason': error.code}
        if isinstance(error, EvidenceError):
            return {'ok': False, 'reason': error.code}
        if isinstance(error, ValidationError):
            return {'ok': False, 'reason': 'literature_invalid'}
        return {'ok': False, 'reason': 'literature_failed'}

    def _call(self, operation):
        with self.lock:
            try:
                editor = self._bind()
                if self.running:
                    raise LiteratureError('busy')
                return {'ok': True, 'result': operation(editor)}
            except Exception as error:
                return self._error(error)

    def _network(self, operation, accept, expected=None):
        with self.lock:
            try:
                editor = self._bind()
                if expected is not None and self.session != expected:
                    raise LiteratureError('literature_scope_changed')
                if self.running:
                    raise LiteratureError('busy')
                if self.client is None:
                    self.client = self.factory()
                lease, client = self.session, self.client
                self.running = True
            except Exception as error:
                return self._error(error)
        try:
            result = operation(client, lambda: self._check(lease))
            with self.lock:
                self._check(lease)
                return {'ok': True, 'result': accept(editor, result)}
        except Exception as error:
            return self._error(error)
        finally:
            with self.lock:
                self.running = False

    def _record(self, record_id):
        if type(record_id) is not str or record_id not in self.records:
            raise LiteratureError('literature_selection_invalid')
        return self.records[record_id]

    def _accept(self, records, reports, query=None):
        if len(self.records) + len(records) > 100 or len(self.reports) + len(reports) > 30:
            raise LiteratureError('literature_capacity')
        for row in records:
            previous = self.records.get(row['id'], {})
            for key in ('fulltext_state', 'fulltext_source', 'local_binding'):
                if key in previous:
                    row[key] = deepcopy(previous[key])
            self.records[row['id']] = row
        self.reports.extend(reports)
        return {'groups': group_records(records), 'reports': reports, 'query': query, 'decisions': self.decisions,
                'requests_used': self.client.calls, 'request_limit': self.client.config.request_limit,
                'scope': 'workspace_session', 'support_assessed': False}

    def search(self, payload):
        try:
            if type(payload) is not dict or type(payload.get('sources', [])) is not list:
                raise LiteratureError()
            query = LiteratureQuery.model_validate({**payload, 'sources': tuple(payload.get('sources', ['crossref', 'openalex']))})
        except Exception as error:
            return self._error(error)
        return self._network(lambda client, check: client.search(query, check=check),
                             lambda editor, result: self._accept(*result, query.model_dump(mode='json')))

    def review(self, record_id, decision, reason):
        def operation(editor):
            self._record(record_id)
            if decision not in {'include', 'exclude', 'undecided'}:
                raise LiteratureError()
            reason_value = text(reason, 400, empty=decision == 'undecided')
            self.decisions[record_id] = {'decision': decision, 'reason': reason_value, 'at': acquired_at()}
            return {'record_id': record_id, **self.decisions[record_id], 'persisted': False}
        return self._call(operation)

    def verify(self, record_id):
        try:
            with self.lock:
                self._bind()
                record = self._record(record_id)
                expected = self.session
                if not record['doi']:
                    raise LiteratureError('literature_doi_invalid')
                query = LiteratureQuery(query=record['doi'], mode='doi')
        except Exception as error:
            return self._error(error)
        def accept(editor, result):
            records, reports = result
            response = self._accept(records, reports)
            successful = {row['source'] for row in records if row['doi'] == record['doi']}
            response['verification'] = {'doi': record['doi'], 'sources_found': sorted(successful),
                                        'both_found': len(successful) == 2, 'claim_supported': False,
                                        'retraction_clearance': False}
            return response
        return self._network(lambda client, check: client.search(query, check=check), accept, expected)

    def citing(self, record_id):
        try:
            with self.lock:
                self._bind()
                record = self._record(record_id)
                expected = self.session
        except Exception as error:
            return self._error(error)
        return self._network(lambda client, check: client.citing(record, check=check),
                             lambda editor, result: self._accept(result[0], [result[1]]), expected)

    def _archive(self, editor, kind):
        if kind not in self.archives:
            self.archives[kind] = LiteratureArchive(editor.space, kind)
        return self.archives[kind]

    def _plan(self, editor, kind, payload, extra=None):
        if len(self.plans) >= 8:
            raise LiteratureError('literature_capacity')
        plan = self._archive(editor, kind).plan_payload(payload, self.session, extra)
        self.plans[plan['plan_id']] = (kind, extra)
        return {'plan': {**plan, **(extra or {})}}

    def export(self, kind):
        def operation(editor):
            if kind not in {'csl', 'record'}:
                raise LiteratureError()
            selected = [row for key, row in self.records.items() if self.decisions.get(key, {}).get('decision') == 'include']
            if (kind == 'csl' and not selected) or (kind == 'record' and not self.reports):
                raise LiteratureError('literature_selection_invalid')
            seen, unique = set(), []
            for row in selected:
                key = row['doi'] or row['id']
                if key in seen:
                    raise LiteratureError('literature_duplicate_selection')
                seen.add(key)
                unique.append(row)
            value = [csl(row) for row in unique] if kind == 'csl' else {
                'version': VERSION, 'exported_at': acquired_at(), 'records': list(self.records.values()),
                'decisions': self.decisions, 'searches': self.reports, 'bibliography': [csl(row) for row in unique],
                'support_assessed': False, 'systematic_review': False}
            payload = json.dumps(value, ensure_ascii=False, allow_nan=False).encode('utf-8')
            return self._plan(editor, kind, payload, {'selected': len(unique), 'citation_style': 'CSL-JSON data only'})
        return self._call(operation)

    def download(self, record_id):
        try:
            with self.lock:
                self._bind()
                record = self._record(record_id)
                expected = self.session
                if any(kind == 'pdf' and extra['record_id'] == record_id for kind, extra in self.saves.values()):
                    raise LiteratureError('literature_fulltext_exists')
        except Exception as error:
            return self._error(error)
        return self._network(lambda client, check: client.fulltext(record, check=check),
                             lambda editor, body: self._plan(editor, 'pdf', body,
                                 {'record_id': record_id, 'source_url': record['public_fulltext']['url'],
                                  'source_record_sha256': record['source_record_sha256'], 'retrieved_at': acquired_at(),
                                  'license': record['public_fulltext']['license'], 'state': 'downloaded_in_memory',
                                  'csl': csl(record), 'retraction_status': record['retraction_status']}), expected)

    def confirm(self, plan_id, approved=False):
        def operation(editor):
            entry = self.plans.pop(plan_id, None) if type(plan_id) is str else None
            if entry is None or type(approved) is not bool:
                raise LiteratureError('literature_plan_invalid')
            kind, extra = entry
            if approved and kind == 'pdf' and any(saved_kind == 'pdf' and saved_extra['record_id'] == extra['record_id'] for saved_kind, saved_extra in self.saves.values()):
                self._archive(editor, kind).confirm(plan_id, self.session, False)
                raise LiteratureError('literature_fulltext_exists')
            outcome = self._archive(editor, kind).confirm(plan_id, self.session, approved)
            if outcome['status'] == 'uncertain':
                self.disabled = True
            if outcome['status'] == 'saved':
                self.saves[plan_id] = (kind, extra)
                if kind == 'pdf':
                    self.records[extra['record_id']]['fulltext_state'] = 'downloaded_unparsed'
            return outcome
        return self._call(operation)

    def read_fulltext(self, save_id):
        def operation(editor):
            if type(save_id) is not str or save_id not in self.saves or self.saves[save_id][0] != 'pdf':
                raise LiteratureError('literature_selection_invalid')
            from ..pdf_reader import read_pdf
            archive = self._archive(editor, 'pdf')
            saved = archive._saved_for(save_id, self.session)
            report = read_pdf(saved['directory'] / archive.artifact_name, limit=5, max_chars=20000, layout=True)
            archive._saved_check(saved)
            record_id = self.saves[save_id][1]['record_id']
            origin = self.saves[save_id][1]
            source_document = self._source_document(editor, record_id, archive.artifact_name, report['source_sha256'],
                                                    'pdfium-content-order-v1', source_url=origin['source_url'], when=origin['retrieved_at'])
            has_text = any(page['text'].strip() for page in report['pages'])
            self.records[record_id]['fulltext_state'] = ('parsed_partial' if report['truncated'] else 'parsed_text_layer') if has_text else 'text_layer_empty_in_preview'
            self.records[record_id]['fulltext_source'] = source_document
            return {'record_id': record_id, 'source_sha256': report['source_sha256'], 'pages': report['pages'],
                    'source_document': source_document,
                    'text_layer_available': has_text,
                    'page_count': report['page_count'], 'truncated': report['truncated'], 'ocr_performed': False,
                    'support_assessed': False, 'coordinate_system': report['coordinate_system']}
        return self._call(operation)

    def _source_document(self, editor, record_id, name, fingerprint, parser, *, source_url=None, when=None):
        return SourceDocument(id='src:' + identifier(record_id, fingerprint), source_name=name, source_sha256=fingerprint,
                              parser_version=parser, scope_id=identifier(str(editor.space.root), str(self.root_identity)),
                              retrieved_at=when or acquired_at(), access='public_fulltext' if source_url else 'user_local',
                              source_url=source_url, literature_id=record_id).model_dump(mode='json')

    def local_fulltext(self, record_id):
        def operation(editor):
            record = self._record(record_id)
            selected = self.scope()[2][0]
            path = editor.space.file(str(selected))
            if path.suffix.lower() != '.pdf':
                raise LiteratureError('literature_pdf_required')
            from ..research_pdf import PdfResearchSource
            source = PdfResearchSource(editor.space).extract(path.name)
            binding = {'record_id': record_id, 'source': path.name, 'source_sha256': source['fingerprint'],
                       'parser_version': source['parser_version'], 'blocks': source['blocks'],
                       'source_document': self._source_document(editor, record_id, path.name, source['fingerprint'], source['parser_version']),
                       'doi_match_verified': False, 'support_assessed': False, 'access': 'user_local'}
            record['fulltext_state'] = 'user_local_text_layer'
            record['local_binding'] = {key: value for key, value in binding.items() if key != 'blocks'}
            return binding
        return self._call(operation)

    def check_citations(self, value):
        def operation(editor):
            content = text(value, 8000)
            selected = [row for key, row in self.records.items() if self.decisions.get(key, {}).get('decision') == 'include']
            citations = [int(match) for match in re.findall(r'\[(\d{1,5})\]', content)]
            return {'style': 'plain-numbered-v1', 'recognized': sorted(set(citations)),
                    'unknown': sorted({index for index in citations if not 1 <= index <= len(selected)}),
                    'uncited': [index for index in range(1, len(selected) + 1) if index not in citations],
                    'missing_fields': [{'id': row['id'], 'fields': row['missing_fields']} for row in selected if row['missing_fields']],
                    'reported_retracted': [row['id'] for row in selected if row['retraction_status'] == 'reported_retracted'],
                    'support_assessed': False, 'doi_resolved': False,
                    'duplicate_dois': sorted({row['doi'] for row in selected if row['doi'] and sum(other['doi'] == row['doi'] for other in selected) > 1})}
        return self._call(operation)

    def rollback(self, save_id, approved=False):
        def operation(editor):
            if type(save_id) is not str or save_id not in self.saves or type(approved) is not bool:
                raise LiteratureError('literature_plan_invalid')
            kind, extra = self.saves[save_id]
            self._dependencies = {key: value for key, value in self._dependencies.items() if value['state'] != 'plan' or time.monotonic() - value['created'] < 300}
            if approved and kind == 'pdf' and any(save_id in value['archives'] for value in self._dependencies.values()):
                raise LiteratureError('literature_fulltext_in_use')
            archive = self._archive(editor, kind)
            plan = archive.plan_rollback(save_id, self.session)
            result = archive.confirm_rollback(plan['rollback_id'], self.session, approved)
            if result['status'] == 'uncertain':
                self.disabled = True
            if result['status'] == 'rolled_back':
                self.saves.pop(save_id)
                if kind == 'pdf':
                    row = self.records[extra['record_id']]
                    row['fulltext_state'] = 'user_local_text_layer' if row.get('local_binding') else 'not_acquired'
                    row.pop('fulltext_source', None)
            return result
        return self._call(operation)

    def pin(self, token, archives, state='plan'):
        if archives:
            self._dependencies[token] = {'archives': list(archives), 'state': state, 'created': time.monotonic()}

    def unpin(self, token):
        self._dependencies.pop(token, None)

    def archive_receipt(self, save_id):
        editor = self._bind()
        archive = self._archive(editor, 'pdf')
        receipt, origin = archive.pdf_receipt(save_id)
        metadata = origin.get('csl')
        if not metadata:
            metadata = csl(self._record(origin['record_id']))
        if metadata['id'] != origin['record_id'] or not metadata.get('DOI') or not metadata.get('author') or not metadata.get('issued'):
            raise LiteratureError('literature_receipt_changed')
        return archive.check_pdf_receipt(receipt), receipt, origin, metadata
