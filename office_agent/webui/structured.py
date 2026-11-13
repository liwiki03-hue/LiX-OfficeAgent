from datetime import datetime, timezone
import hashlib
import json
import re
import time
import uuid

from pydantic import Field

from .professional import Professional
from .citations import Attribution
from ..structured_edits import StructuredDocx, ParagraphInsertion, ReviewComment, anchor_record
from ..professional_edits import _inventory, LocalChange
from ..grounded_edits import seal
from ..provenance import ClaimSupport
from ..rag import SECRET
from ..writes import Strict, GuardError, digest


class DraftInsertion(Strict):
    anchor: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    side: str
    text: str = Field(min_length=1, max_length=2000)
    style: str | None = None
    attributions: list[Attribution] = Field(default_factory=list, max_length=3)


class StructureDraft(Strict):
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    insertions: list[DraftInsertion] = Field(default_factory=list, max_length=4)
    comments: list[ReviewComment] = Field(default_factory=list, max_length=8)
    review_notes: list[str] = Field(default_factory=list, max_length=8)
    replacements: list[LocalChange] = Field(default_factory=list, max_length=12)
    authorized: list[str] = Field(default_factory=list, max_length=12)
    bibliography: list[str] = Field(default_factory=list, max_length=8)
    bibliography_anchor: str | None = None
    number_records: dict[str, str] = Field(default_factory=dict)
    numeric_reviews: list[str] = Field(default_factory=list, max_length=8)


class Structured(Professional):
    def invalidate(self):
        if hasattr(self, 'grounded'):
            for key in self._pending:
                self.grounded.literature.unpin(key)
        self._observations = {}
        self._pdf_candidate = None
        self._numeric_reviews = {}
        super().invalidate()

    def review_claim(self, locator_id, selections):
        def operation(editor, path, binding):
            from ..advanced_writes import _parts
            from ..grounded_edits import validate_source
            from ..research_contracts import Suggestions, ClaimVerification
            if type(selections) is not list or not 1 <= len(selections) <= 3:
                raise GuardError('professional_evidence_invalid')
            version = digest(path)
            _, rows = _inventory(_parts(path), version)
            target = next((r for r in rows if r['locator_id'] == locator_id), None)
            if not target or SECRET.search(target['text']) or len(target['text']) > 2000:
                raise GuardError('professional_stale_locator')
            evidence, proofs = {}, []
            for selected in selections:
                if type(selected) is not dict or set(selected) != {'capture_id', 'evidence_ids'} or type(selected['evidence_ids']) is not list:
                    raise GuardError('professional_evidence_invalid')
                capture = self.grounded._captures.get(selected['capture_id'])
                if not capture or capture['binding'] != self.grounded._binding()[2] or capture['session'] != self.grounded.literature.session or not self.grounded._included(capture['proof']):
                    raise GuardError('professional_evidence_invalid')
                fragments = validate_source(capture['proof'], seal(capture['proof']), path)
                ids = selected['evidence_ids']
                if not 1 <= len(ids) <= 3 or not set(ids) <= {r.id for r in fragments}:
                    raise GuardError('professional_evidence_invalid')
                proofs.append(capture['proof'])
                for row in fragments:
                    if row.id in ids:
                        evidence[row.id] = {'id': row.id, 'text': row.original_text, 'location': row.location,
                                            'source_sha256': row.source_sha256, 'coverage': capture['proof'].get('coverage')}
            if sum(len(r['text']) for r in evidence.values()) > 10000:
                raise GuardError('professional_capacity')
            model = self.model_factory()
            payload = {'question': '核对所选段落的原子主张。对齐对象、样本、时间、条件、否定、范围与相关/因果；保留支持和反对来源。信息不足时明确不足；不要强行判定矛盾。',
                       'targets': [{'locator_id': locator_id, 'text': target['text']}], 'evidence': list(evidence.values()),
                       'coverage_is_partial': True, 'mode': 'critique', 'protected_terms': []}
            response = model.verify_claim(payload) if callable(getattr(model, 'verify_claim', None)) else model.suggest(payload)
            parsed = ClaimVerification.model_validate_json(json.dumps(response)) if 'claims' in response else Suggestions.model_validate_json(json.dumps(response))
            judgments = parsed.claims if hasattr(parsed, 'claims') else parsed.cards
            if any(getattr(row, 'verdict', getattr(row, 'relation', 'insufficient')) in {'supports', 'contradicts'} and not row.evidence_ids for row in judgments):
                raise GuardError('professional_assist_invalid_response')
            if SECRET.search(json.dumps(response, ensure_ascii=False)) or any(key not in evidence for row in [*judgments, *parsed.suggestions] for key in row.evidence_ids):
                raise GuardError('professional_assist_invalid_response')
            if any(row.locator_id != locator_id for row in parsed.suggestions):
                raise GuardError('professional_assist_invalid_response')
            for proof in proofs:
                validate_source(proof, seal(proof), path)
            if digest(path) != version or self._binding()[2] != binding:
                raise GuardError('professional_source_changed')
            return {'paragraph': target['text'], 'source_sha256': version, 'evidence': list(evidence.values()),
                    'review': parsed.model_dump(), 'user_reviewed': False, 'model_calls': 1,
                    'support_score_is_probability': False, 'scope': 'selected_evidence_only', 'wrote_nothing': True}
        return self._call(operation)

    def read_numeric_source(self, name):
        def operation(editor, path, binding):
            self._pdf_candidate = None
            from ..evidence_checks import table_observations
            records = table_observations(editor.space, name)
            self._observations = {'binding': binding, 'records': records, 'source': name}
            return {'observations': records, 'origin': 'direct_source', 'user_reviewed': False, 'wrote_nothing': True}
        return self._call(operation)

    def _pdf_call(self, operation):
        from ..rag_sources import SourceError
        def guarded(editor, path, binding):
            try:
                return operation(editor, path, binding)
            except SourceError as error:
                raise GuardError('professional_pdf_' + error.code) from None
        return self._call(guarded)

    def _pdf_entry(self, editor, identifier, binding):
        from ..rag_sources import SourceExtractor
        entry = getattr(self, '_pdf_candidate', None)
        if not entry or type(identifier) is not str or entry['id'] != identifier or entry['binding'] != binding:
            raise GuardError('professional_pdf_stale')
        if time.monotonic() - entry['created'] >= 300:
            self._pdf_candidate = None
            raise GuardError('professional_pdf_expired')
        _, identity, fingerprint = SourceExtractor(editor.space)._snapshot(entry['name'])
        if identity != entry['identity'] or fingerprint != entry['source_sha256']:
            self._pdf_candidate = None
            raise GuardError('professional_pdf_source_changed')
        return entry

    def preview_pdf_numeric(self, name, page, proposal_box=None):
        def operation(editor, path, binding):
            from ..pdf_table import read_table
            self._observations = {}
            self._pdf_candidate = None
            report = read_table(editor.space, name, page, proposal_box=proposal_box)
            entry = {'id': uuid.uuid4().hex, 'binding': binding, 'created': time.monotonic(), 'name': name,
                     'identity': report['identity'], 'source_sha256': report['source_sha256'], 'page': page}
            self._pdf_candidate = entry
            return {**{key: value for key, value in report.items() if key != 'identity'}, 'pdf_table_id': entry['id'],
                    'layout_confirmed': False, 'wrote_nothing': True, 'expires_in_seconds': 300}
        return self._pdf_call(operation)

    def read_pdf_numeric(self, identifier, grid):
        def operation(editor, path, binding):
            from ..pdf_table import read_table, validate_grid
            from ..evidence_checks import observations_from_tables, UNITS
            entry = self._pdf_entry(editor, identifier, binding)
            self._observations = {}
            self._pdf_candidate = None
            validate_grid(grid)
            if grid['page'] != entry['page']:
                raise GuardError('professional_pdf_invalid')
            report = read_table(editor.space, entry['name'], entry['page'], grid)
            if report['identity'] != entry['identity'] or report['source_sha256'] != entry['source_sha256']:
                raise GuardError('professional_pdf_source_changed')
            stamp = hashlib.sha256(json.dumps([grid, report['parser_stamp']], sort_keys=True).encode()).hexdigest()[:16]
            location = f"pdf-page:{entry['page']}/grid:{stamp}"
            following = [(f"pdf-page:{entry['page']}/note-box", report['note_text'])] if report.get('note_text') else []
            header_rows = grid.get('header_rows', 1)
            records = observations_from_tables(entry['name'], report['source_sha256'], [(location, report['rows'], False, following, False)], header_rows=header_rows)
            for record, cell in zip(records, report['cells']):
                record.update(extraction='pdf_explicit_grid', uncertain_structure=True, layout_confirmed=False,
                              pdf_page=entry['page'], cell_box=cell['box'], char_indices=cell['char_indices'],
                              grid_sha256=stamp, parser_stamp=report['parser_stamp'], context_outside_selection_included=False,
                              pdf_note_box=report.get('note_box'), pdf_note_text=report.get('note_text', ''),
                              pdf_note_char_indices=report.get('note_char_indices', []))
                if report.get('note_box'):
                    from ..table_context import NOTE
                    if not report.get('note_text') or not NOTE.fullmatch(report['note_text']):
                        record['context_flags'].append('pdf_note_region_requires_manual_review')
                for column, unit_column in grid['units']:
                    if cell['column'] == column and cell['row'] >= header_rows:
                        unit = report['rows'][cell['row']][unit_column]
                        if unit not in UNITS or (record['unit'] and record['unit'] != unit):
                            raise GuardError('professional_pdf_invalid')
                        record.update(unit=unit, unit_source=f"{location}/row:{cell['row']}/column:{unit_column}")
            entry.update(id=uuid.uuid4().hex, grid=grid, records=records)
            if callable(getattr(self, 'layout_observer', None)):
                self.layout_observer(entry['name'], report['source_sha256'], stamp, report['parser_stamp'])
            self._pdf_candidate = entry
            return {**{key: value for key, value in report.items() if key != 'identity'}, 'observations': records,
                    'pdf_table_id': entry['id'], 'layout_confirmed': False, 'wrote_nothing': True}
        return self._pdf_call(operation)

    def acknowledge_pdf_numeric(self, identifier, reviewed=False):
        def operation(editor, path, binding):
            entry = self._pdf_entry(editor, identifier, binding)
            self._pdf_candidate = None
            if type(reviewed) is not bool:
                raise GuardError('professional_pdf_invalid')
            if not reviewed:
                return {'status': 'cancelled', 'wrote_nothing': True}
            if 'records' not in entry:
                raise GuardError('professional_pdf_stale')
            for record in entry['records']:
                record.update(uncertain_structure=False, layout_confirmed=True)
            self._observations = {**entry, 'mode': 'pdf_grid'}
            return {'observations': entry['records'], 'layout_confirmed': True, 'user_reviewed': False,
                    'write_approval_granted': False, 'wrote_nothing': True}
        return self._pdf_call(operation)

    def verify_numeric(self, locator_id, payload):
        def operation(editor, path, binding):
            from ..evidence_checks import table_observations, check_numeric
            from ..advanced_writes import _parts
            saved = getattr(self, '_observations', {})
            if saved.get('binding') != binding:
                raise GuardError('professional_evidence_invalid')
            if saved.get('mode') == 'pdf_grid':
                from ..rag_sources import SourceExtractor, SourceError
                try:
                    _, identity, fingerprint = SourceExtractor(editor.space)._snapshot(saved['name'])
                except SourceError:
                    self._observations = {}
                    raise GuardError('professional_pdf_source_changed') from None
                if identity != saved['identity'] or fingerprint != saved['source_sha256']:
                    self._observations = {}
                    raise GuardError('professional_pdf_source_changed')
                records = saved['records']
            else:
                records = table_observations(editor.space, saved['source'])
            if records != saved['records']:
                raise GuardError('professional_evidence_changed')
            _, rows = _inventory(_parts(path), digest(path))
            row = next((r for r in rows if r['locator_id'] == locator_id), None)
            if not row:
                raise GuardError('professional_stale_locator')
            from ..evidence_checks import NUMBER, _decimal
            if type(payload) is not dict or not any(_decimal(n) == _decimal(payload.get('value')) for n in re.findall(NUMBER, row['text'])):
                raise GuardError('professional_numeric_claim_range')
            checked = check_numeric(records, payload)
            checked['paragraph'] = row['text']
            checked['paragraph_sha256'] = row['text_sha256']
            checked['source_sha256'] = digest(path)
            checked['claim_scope'] = 'selected_numeric_value_only_not_entire_paragraph'
            checked['semantic_flags'] = [label for pattern, label in [(r'因果|导致|caus', 'causality_requires_review'), (r'所有|普遍|always|all ', 'scope_generalization_requires_review'), (r'不|未|没有|not |no ', 'negation_requires_review')] if re.search(pattern, row['text'], re.I)]
            if saved.get('mode') == 'pdf_grid':
                self._numeric_reviews = {key: value for key, value in getattr(self, '_numeric_reviews', {}).items() if time.monotonic() - value['created'] < 300}
                if len(self._numeric_reviews) >= 8:
                    raise GuardError('professional_pdf_review_capacity')
                identifier = uuid.uuid4().hex
                observation = next(record for record in records if record['id'] == payload['observation_id'])
                self._numeric_reviews[identifier] = {'binding': binding, 'created': time.monotonic(), 'proof': {
                    'name': saved['name'], 'source_sha256': saved['source_sha256'], 'identity': json.loads(json.dumps(saved['identity'])),
                    'page': saved['page'], 'grid': saved['grid'], 'observation_id': observation['id'],
                    'location': observation['location'], 'parser_stamp': observation['parser_stamp'],
                    'target_locator': locator_id, 'claim': checked['claim'], 'status': checked['status'], 'user_reviewed': False}}
                checked['numeric_review_id'] = identifier
            if callable(getattr(self, 'numeric_observer', None)):
                self.numeric_observer(locator_id, checked)
            return checked
        return self._call(operation)

    def read_structure(self, offset=0):
        def operation(editor, path, binding):
            from ..professional_edits import snapshot
            from ..advanced_writes import _parts
            data = snapshot(editor, path, offset=offset)
            _, rows = _inventory(_parts(path), data['document']['source_sha256'])
            data['anchors'] = [{'locator_id': r['locator_id'], 'location': r['location'], 'text': r['text'], 'text_sha256': r['text_sha256'], 'style': r['style']}
                               for r in rows[offset:data['next_offset'] or len(rows)] if not SECRET.search(r['text']) and r['writable'] and re.fullmatch(r'body:\d+/paragraph', r['location'])]
            return data
        return self._call(operation)

    def plan_structure(self, payload, project_proof=None):
        def operation(editor, path, binding):
            from ..advanced_writes import _parts
            draft = StructureDraft.model_validate(payload)
            if digest(path) != draft.source_sha256:
                raise GuardError('professional_source_changed')
            restored = getattr(self.citations, '_restored', {}) if hasattr(self, 'citations') else {}
            if restored and (restored['version'] != draft.source_sha256 or restored['binding'] != self.citations._binding()[2]):
                raise GuardError('professional_citation_receipt')
            _, rows = _inventory(_parts(path), draft.source_sha256)
            targets = {r['locator_id']: r for r in rows}
            catalog, bindings, insertions = {}, {}, []
            def add(metadata):
                if not metadata.get('DOI'):
                    raise GuardError('professional_metadata_incomplete')
                key = 'doi:' + metadata['DOI'].casefold()
                if key in catalog and {k: v for k, v in metadata.items() if k not in {'id', 'URL'}} != {k: v for k, v in catalog[key].items() if k not in {'id', 'URL'}}:
                    raise GuardError('professional_duplicate_metadata')
                catalog[key] = metadata
                return key
            for number, record in draft.number_records.items():
                bindings[number] = add(self.citations._catalog_item(record, self.citations._binding()[2]))
            archives = set()
            for item in draft.insertions:
                attrs = []
                for attribution in item.attributions:
                    capture = self.grounded._captures.get(attribution.capture_id)
                    if not capture or capture['binding'] != self.grounded._binding()[2] or capture['session'] != self.grounded.literature.session or not self.grounded._included(capture['proof']):
                        raise GuardError('professional_evidence_invalid')
                    if not attribution.identity_reviewed:
                        raise GuardError('professional_identity_required')
                    proof = {**capture['proof'], 'fragments': [r for r in capture['proof']['fragments'] if r['id'] in attribution.evidence_ids]}
                    support = ClaimSupport.model_validate_json(json.dumps({'claim': item.text, **attribution.model_dump(exclude={'capture_id', 'identity_reviewed'})}))
                    attrs.append({'proof': proof, 'seal': seal(proof), 'support': support.model_dump(mode='json')})
                    add(proof['csl'])
                    if proof.get('archive'):
                        archives.add(proof['archive']['save_id'])
                insertions.append(ParagraphInsertion(**item.model_dump(exclude={'attributions'}), attributions=attrs))
            keys = set(draft.authorized + draft.bibliography + [r.anchor for r in draft.insertions] + [r.anchor for r in draft.comments])
            if draft.bibliography_anchor:
                keys.add(draft.bibliography_anchor)
            if not keys <= targets.keys():
                raise GuardError('professional_stale_locator')
            intent = {'source_sha256': draft.source_sha256, 'anchors': {k: anchor_record(targets[k], rows, draft.source_sha256) for k in keys},
                      'insertions': [r.model_dump() for r in insertions], 'comments': [r.model_dump() for r in draft.comments], 'replacements': [r.model_dump() for r in draft.replacements],
                      'review_notes': draft.review_notes, 'date': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'), 'catalog': catalog, 'bindings': bindings,
                      'authorized': draft.authorized, 'bibliography': draft.bibliography, 'bibliography_anchor': draft.bibliography_anchor}
            proofs = []
            if len(set(draft.numeric_reviews)) != len(draft.numeric_reviews):
                raise GuardError('professional_pdf_stale')
            for identifier in draft.numeric_reviews:
                review = getattr(self, '_numeric_reviews', {}).get(identifier)
                if not review or review['binding'] != binding or time.monotonic() - review['created'] >= 300:
                    raise GuardError('professional_pdf_stale')
                if review['proof']['target_locator'] not in keys:
                    raise GuardError('professional_pdf_stale')
                proofs.append(review['proof'])
            if proofs:
                intent['pdf_numeric_reviews'] = proofs
            if restored:
                inherited = restored['manifest'].get('archive_dependencies', [])
                if len(inherited) > 8 or any(type(k) is not str or not re.fullmatch('[a-f0-9]{32}', k) for k in inherited):
                    raise GuardError('professional_citation_receipt')
                archives.update(inherited)
                intent['parent_receipt'] = restored['receipt_id']
                intent['parent_audit'] = {'receipt_id': restored['receipt_id'], 'stamp': restored['stamp']}
                intent['prior_citations'] = restored['manifest']
            intent['archive_dependencies'] = sorted(archives)
            if project_proof is not None:
                if not callable(getattr(self, 'project_validator', None)):
                    raise GuardError('professional_project_invalid')
                self.project_validator(project_proof)
                intent['project_workbench'] = project_proof
            request = StructuredDocx(path=str(path), source_sha256=draft.source_sha256, insertions=insertions, comments=draft.comments, replacements=draft.replacements, intent=intent, intent_seal=seal(intent))
            for key, lease in self._pending.items():
                lease['editor'].plans.pop(key, None)
                self.grounded.literature.unpin(key)
            self._pending.clear()
            plan = editor.plan(request)
            self._pending[plan.id] = {'editor': editor, 'binding': binding, 'created': time.monotonic()}
            self.grounded.literature.pin(plan.id, archives)
            return {**plan.public(), 'items': json.loads(plan.after)['items'], 'wrote_nothing': True, 'expires_in_seconds': 300}
        return self._call(operation)

    def _validate_confirmation(self, plan):
        if plan.request.kind != 'structured_docx':
            return
        from ..structured_edits import validate_parent
        validate_parent(plan.request)
        if 'project_workbench' in plan.request.intent:
            if not callable(getattr(self, 'project_validator', None)):
                raise GuardError('professional_project_invalid')
            self.project_validator(plan.request.intent['project_workbench'])
        for insertion in plan.request.insertions:
            for item in insertion.attributions:
                from ..grounded_edits import validate_source
                validate_source(item['proof'], item['seal'], plan.request.path)
                if not self.grounded._included(item['proof']):
                    raise GuardError('professional_evidence_invalid')

    def confirm(self, plan_id, confirmed=False):
        reply = super().confirm(plan_id, confirmed)
        if hasattr(self, 'grounded'):
            if reply.get('result', {}).get('status') == 'saved':
                dependency = self.grounded.literature._dependencies.get(plan_id)
                if dependency:
                    dependency['state'] = 'output'
            elif reply.get('result', {}).get('status') != 'uncertain' and plan_id not in self._pending:
                self.grounded.literature.unpin(plan_id)
        return reply

    def rollback(self, save_id, confirmed=False):
        reply = super().rollback(save_id, confirmed)
        if reply.get('result', {}).get('status') == 'rolled_back':
            self.grounded.literature.unpin(save_id)
        return reply
