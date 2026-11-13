import json
import re
import time

from pydantic import Field

from .professional import Professional
from ..citation_lifecycle import CitationDocx, VERSION, reconcile, scan
from ..grounded_edits import seal, validate_proof
from ..literature import csl
from ..professional_edits import LocalChange, MAX_TARGETS, _inventory, snapshot
from ..provenance import ClaimSupport
from ..rag_evidence import _read
from ..rag import SECRET
from ..writes import Strict, GuardError, digest


class Attribution(Strict):
    capture_id: str = Field(pattern=r'^[0-9a-f]{64}$')
    evidence_ids: list[str] = Field(min_length=1, max_length=3)
    relation: str
    applicability: str
    limitation: str
    assessment: str
    identity_reviewed: bool


class Claim(Strict):
    locator_id: str = Field(pattern=r'^loc:[0-9a-f]{64}$')
    text: str = Field(min_length=1, max_length=2000)
    attributions: list[Attribution] = Field(min_length=1, max_length=3)


class ChangeSet(Strict):
    source_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    authorized: list[str] = Field(min_length=1, max_length=12)
    bibliography: list[str] = Field(min_length=1, max_length=8)
    number_records: dict[str, str]
    claims: list[Claim] = Field(default_factory=list, max_length=4)
    removals: dict[str, list[int]] = Field(default_factory=dict)
    mapping_reviewed: bool


class Citations(Professional):
    def __init__(self, lock, scope, editor, selected, grounded):
        super().__init__(lock, scope, editor, selected)
        self.grounded = grounded
        self.literature = grounded.literature
        self._restored = {}

    def invalidate(self):
        for token in self._pending:
            self.literature.unpin(token)
        super().invalidate()
        self._restored.clear()

    def _rows(self, editor, path, version):
        from ..advanced_writes import _parts
        if path.suffix.lower() != '.docx':
            raise GuardError('professional_docx_required')
        if digest(path) != version:
            raise GuardError('professional_source_changed')
        _, rows = _inventory(_parts(path), version)
        if editor.space.file(str(path)) != path or digest(path) != version:
            raise GuardError('professional_source_changed')
        return rows

    def _eligibility(self, row):
        if row.get('refusal') == 'sensitive_or_oversize' or len(row['text']) > 8000 or SECRET.search(row['text']):
            return 'sensitive_or_oversize'
        if not re.fullmatch(r'body:\d+/paragraph', row['location']):
            return 'not_body_paragraph'
        if not row['writable']:
            return row.get('refusal') or 'unsupported_structure'
        if row['style']:
            return 'explicit_paragraph_style'
        return None

    def locations(self, number, source_sha256, offset=0):
        def operation(editor, path, binding):
            if (type(number) is not int or not 1 <= number <= 999 or type(source_sha256) is not str or
                    not re.fullmatch('[a-f0-9]{64}', source_sha256) or type(offset) is not int or not 0 <= offset <= MAX_TARGETS):
                raise GuardError('professional_invalid_request')
            rows = self._rows(editor, path, source_sha256)
            found, omitted = [], 0
            for index, row in enumerate(rows):
                reason = self._eligibility(row)
                if reason == 'sensitive_or_oversize':
                    omitted += 1
                    continue
                matches = [match for match, ns in scan(row['text']) if number in ns]
                if not matches:
                    continue
                start = max(0, matches[0].start() - 60)
                excerpt = row['text'][start:start + 180]
                found.append({'index': index, 'location': row['location'], 'locator_id': row['locator_id'],
                              'occurrences': len(matches), 'excerpt': excerpt, 'excerpt_truncated': start > 0 or start + 180 < len(row['text']),
                              'citation_refusal': reason, 'eligible_not_authorized': reason is None,
                              'kind': 'bibliography_candidate' if re.match(r'^\[\d+\]', row['text']) and re.search(r'10\.\d{4,9}/', row['text']) else 'citation_occurrence'})
            if offset and offset >= len(found):
                raise GuardError('professional_invalid_page')
            page = found[offset:offset + 40]
            if editor.space.file(str(path)) != path or digest(path) != source_sha256:
                raise GuardError('professional_source_changed')
            return {'source_sha256': source_sha256, 'number': number, 'items': page, 'offset': offset,
                    'total_locations': len(found), 'total_occurrences': sum(row['occurrences'] for row in found),
                    'next_offset': offset + len(page) if offset + len(page) < len(found) else None,
                    'omitted_targets': omitted, 'scope': 'body_inventory', 'read_only': True,
                    'occurrence_unit': 'citation_cluster_or_bibliography_label',
                    'write_approval_granted': False, 'impact_analysis': 'occurrences_only_not_complete_renumbering_impact'}
        return self._call(operation)

    def read(self, receipt_id='', offset=0, source_sha256=''):
        def operation(editor, path, binding):
            if (type(receipt_id) is not str or type(offset) is not int or
                    type(source_sha256) is not str or
                    (source_sha256 and not re.fullmatch('[a-f0-9]{64}', source_sha256)) or
                    (offset != 0 and not source_sha256)):
                raise GuardError('professional_invalid_page')
            source = snapshot(editor, path, offset=offset, limit=40)
            version = source['document']['source_sha256']
            if source_sha256 and source_sha256 != version:
                raise GuardError('professional_source_changed')
            rows = self._rows(editor, path, version)
            if offset and offset >= len(rows):
                raise GuardError('professional_invalid_page')
            self.literature._bind()
            manifest = None
            restored = {}
            if receipt_id:
                self._restored.clear()
                if type(receipt_id) is not str or not re.fullmatch('[a-f0-9]{32}', receipt_id):
                    raise GuardError('professional_citation_receipt')
                try:
                    data, stamp = _read(editor.space.file(str(editor.audit / receipt_id / 'audit.json')))
                    audit = json.loads(data)
                    if audit['id'] != receipt_id or audit['status'] != 'verified' or editor.space.file(audit['output']) != path or audit['output_sha256'] != digest(path):
                        raise ValueError()
                    after = json.loads(audit['after'])
                    manifest = next(row['citation_manifest'] for row in after['items'] if 'citation_manifest' in row)
                    from ..structured_edits import RECEIPT_VERSION, restore_citation_receipt
                    if manifest['version'] == RECEIPT_VERSION:
                        manifest = restore_citation_receipt(after, rows)
                    else:
                        if manifest['version'] != VERSION:
                            raise ValueError()
                        by_location = {row['location']: row for row in rows}
                        for row in after['items']:
                            if by_location[row['location']]['text'] != row['text']:
                                raise ValueError()
                except Exception:
                    raise GuardError('professional_citation_receipt') from None
                restored = {'binding': binding, 'version': version, 'manifest': manifest, 'receipt_id': receipt_id,
                            'stamp': json.loads(json.dumps(stamp))}
            references = [csl(row) for key, row in self.literature.records.items() if self.literature.decisions.get(key, {}).get('decision') == 'include']
            references.extend(value['proof']['csl'] for value in self.grounded._captures.values()
                              if value['binding'] == self.grounded._binding()[2] and value['session'] == self.literature.session and self.grounded._included(value['proof']))
            if manifest:
                references.extend(manifest['catalog'].values())
            public_manifest = {key: value for key, value in manifest.items() if key in {'version', 'style', 'catalog', 'number_to_key', 'slots', 'parent_receipt'}} if manifest else None
            if public_manifest and SECRET.search(json.dumps(public_manifest, ensure_ascii=False)):
                raise GuardError('professional_sensitive_text')
            if editor.space.file(str(path)) != path or digest(path) != version:
                raise GuardError('professional_source_changed')
            citation_numbers = sorted({n for row in rows for _, ns in scan(row['text']) for n in ns})
            self._restored = restored
            source['items'] = [dict(row, citation_refusal=self._eligibility(row)) for row in source['items']]
            return {**source, 'offset': offset, 'citation_numbers_scope': 'body_inventory', 'citation_numbers': citation_numbers,
                    'references': list({row['id']: row for row in references}.values()),
                    'captures': [{'capture_id': key, 'title': value['proof']['csl']['title'], 'doi': value['proof']['csl']['DOI']}
                                 for key, value in self.grounded._captures.items() if value['binding'] == self.grounded._binding()[2] and value['session'] == self.literature.session],
                    'manifest': public_manifest, 'receipt_restores_write_approval': False}
        return self._call(operation)

    def capture(self, capture_id):
        def operation(editor, path, binding):
            value = self.grounded._captures.get(capture_id) if type(capture_id) is str else None
            if not value or value['binding'] != self.grounded._binding()[2] or value['session'] != self.literature.session:
                raise GuardError('professional_evidence_invalid')
            return {'document': value['proof']['document'], 'fragments': value['proof']['fragments'],
                    'coverage': value['proof'].get('coverage'),
                    'origin': value['proof'].get('origin'), 'support_assessed': False}
        return self._call(operation)

    def _catalog_item(self, record_id, binding):
        for value in self.grounded._captures.values():
            if (value['proof']['csl']['id'] == record_id and value['binding'] == self.grounded._binding()[2] and
                    value['session'] == self.literature.session and self.grounded._included(value['proof'])):
                return value['proof']['csl']
        if self._restored.get('binding') == binding:
            for value in self._restored['manifest']['catalog'].values():
                if value['id'] == record_id:
                    return value
        record = self.literature._record(record_id)
        if self.literature.decisions.get(record_id, {}).get('decision') != 'include' or record['retraction_status'] == 'reported_retracted':
            raise GuardError('professional_metadata_incomplete')
        return csl(record)

    def plan(self, payload):
        def operation(editor, path, binding):
            from ..advanced_writes import _parts
            if type(payload) is not dict:
                raise GuardError('professional_invalid_request')
            draft = ChangeSet.model_validate_json(json.dumps(payload))
            self.literature._bind()
            if not draft.mapping_reviewed:
                raise GuardError('professional_citation_review')
            rows = self._rows(editor, path, draft.source_sha256)
            if self._restored and (self._restored['version'] != draft.source_sha256 or self._restored['binding'] != binding):
                raise GuardError('professional_citation_receipt')
            targets = {row['locator_id']: row for row in rows}
            def location(key):
                if key not in targets:
                    raise GuardError('professional_stale_locator')
                return targets[key]['location']
            catalog, bindings = {}, {}
            def add(metadata):
                if SECRET.search(json.dumps(metadata, ensure_ascii=False)):
                    raise GuardError('professional_sensitive_text')
                if not metadata.get('DOI') or not metadata.get('author') or not metadata.get('issued') or not metadata.get('title'):
                    raise GuardError('professional_metadata_incomplete')
                key = 'doi:' + metadata['DOI'].casefold()
                normalized = {k: v for k, v in metadata.items() if k not in {'id', 'URL'}}
                if key in catalog and normalized != {k: v for k, v in catalog[key].items() if k not in {'id', 'URL'}}:
                    raise GuardError('professional_duplicate_metadata')
                catalog.setdefault(key, metadata)
                return key
            for number, record_id in draft.number_records.items():
                if not re.fullmatch(r'[1-9]\d{0,2}', number):
                    raise GuardError('professional_citation_unmapped')
                bindings[number] = add(self._catalog_item(record_id, binding))
            claims = []
            for change in draft.claims:
                attributions = []
                for item in change.attributions:
                    capture = self.grounded._captures.get(item.capture_id)
                    if (not capture or capture['binding'] != self.grounded._binding()[2] or capture['session'] != self.literature.session or
                            not self.grounded._included(capture['proof'])):
                        raise GuardError('professional_evidence_invalid')
                    if not item.identity_reviewed:
                        raise GuardError('professional_identity_required')
                    proof = {**capture['proof'], 'fragments': [row for row in capture['proof']['fragments'] if row['id'] in item.evidence_ids]}
                    support = ClaimSupport.model_validate_json(json.dumps({'claim': change.text, **item.model_dump(exclude={'capture_id', 'identity_reviewed'})}))
                    add(proof['csl'])
                    attributions.append({'proof': proof, 'seal': seal(proof), 'support': support.model_dump(mode='json')})
                claims.append({'location': location(change.locator_id), 'text': change.text, 'attributions': attributions})
            intent = {'authorized': [location(key) for key in draft.authorized], 'slots': [location(key) for key in draft.bibliography],
                      'bindings': bindings, 'catalog': catalog, 'claims': claims,
                      'removals': [{'location': location(key), 'numbers': value} for key, value in draft.removals.items()],
                      'deadline': time.monotonic() + 300}
            archives = {item['proof']['archive']['save_id'] for claim in claims for item in claim['attributions'] if item['proof'].get('archive')}
            if self._restored:
                inherited = self._restored['manifest'].get('archive_dependencies', [])
                if any(type(key) is not str or not re.fullmatch('[a-f0-9]{32}', key) for key in inherited) or len(inherited) > 8:
                    raise GuardError('professional_citation_receipt')
                archives.update(inherited)
                intent['parent_receipt'] = self._restored['receipt_id']
                intent['parent_audit'] = {'receipt_id': self._restored['receipt_id'], 'stamp': self._restored['stamp']}
            intent['archive_dependencies'] = sorted(archives)
            target = targets[draft.authorized[0]]
            placeholder = LocalChange(locator_id=target['locator_id'], expected_text_sha256=target['text_sha256'], new_text='pending', reason='planning')
            request = CitationDocx(path=str(path), source_sha256=draft.source_sha256, changes=[placeholder], intent=intent, intent_seal=seal(intent))
            changes, manifest = reconcile(_parts(path), request)
            request = request.model_copy(update={'changes': changes})
            for key, value in self._pending.items():
                value['editor'].plans.pop(key, None)
                self.literature.unpin(key)
            self._pending.clear()
            plan = editor.plan(request)
            self._pending[plan.id] = {'editor': editor, 'binding': binding, 'created': time.monotonic()}
            self.literature.pin(plan.id, archives)
            return {**plan.public(), 'items': json.loads(plan.after)['items'], 'manifest': manifest,
                    'wrote_nothing': True, 'expires_in_seconds': 300, 'paragraphs_inserted_or_deleted': False}
        return self._call(operation)

    def _validate_confirmation(self, plan):
        from ..citation_lifecycle import reconcile
        from ..advanced_writes import _parts
        self.literature._bind()
        reconcile(_parts(self._editor().space.file(plan.request.path)), plan.request)
        for claim in plan.request.intent['claims']:
            for item in claim['attributions']:
                proof = item['proof']
                if not self.grounded._included(proof):
                    raise GuardError('professional_evidence_invalid')
                support = ClaimSupport.model_validate_json(json.dumps(item['support']))
                check = type('ProofRequest', (), {'proof': proof, 'proof_seal': item['seal'], 'path': plan.request.path, 'support': support})()
                validate_proof(check)

    def confirm(self, plan_id, confirmed=False):
        with self._lock:
            reply = super().confirm(plan_id, confirmed)
            state = reply.get('result', {}).get('status')
            if state == 'saved' and plan_id in self.literature._dependencies:
                self.literature._dependencies[plan_id]['state'] = 'output'
            elif state != 'uncertain' and plan_id not in self._pending:
                self.literature.unpin(plan_id)
            return reply

    def rollback(self, save_id, confirmed=False):
        with self._lock:
            reply = super().rollback(save_id, confirmed)
            if reply.get('result', {}).get('status') == 'rolled_back':
                self.literature.unpin(save_id)
            return reply
