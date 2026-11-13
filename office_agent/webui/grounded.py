import json
import time
from copy import deepcopy

from pydantic import Field

from .professional import Professional
from ..grounded_edits import GroundedDocx, format_reference, seal
from ..literature import csl
from ..professional_edits import LocalChange, snapshot
from ..provenance import ClaimSupport, EvidenceFragment, identifier
from ..rag import SECRET
from ..writes import Strict, GuardError


class Draft(Strict):
    source_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    body_locator: str = Field(pattern=r'^loc:[0-9a-f]{64}$')
    reference_locator: str = Field(pattern=r'^loc:[0-9a-f]{64}$')
    capture_id: str = Field(pattern=r'^[0-9a-f]{64}$')
    support: ClaimSupport
    identity_reviewed: bool


class Grounded(Professional):
    def __init__(self, lock, scope, editor, selected, literature):
        super().__init__(lock, scope, editor, selected)
        self.literature = literature
        self._captures = {}
        from ..research_reading import PdfCoverage
        self._coverage = PdfCoverage()

    def invalidate(self):
        for token in self._pending:
            self.literature.unpin(token)
        super().invalidate()
        self._captures.clear()
        self._coverage.clear()

    def sources(self):
        def operation(editor, path, binding):
            self.literature._bind()
            return {'records': [{'id': key, 'title': row['title'], 'doi': row['doi']} for key, row in self.literature.records.items()
                                if self.literature.decisions.get(key, {}).get('decision') == 'include'],
                    'archives': [{'save_id': key, 'title': value[1].get('csl', {}).get('title', '已保存的公开 PDF')}
                                 for key, value in self.literature.saves.items() if value[0] == 'pdf'],
                    'scope': '当前文献会话中明确纳入的条目；本地 PDF 须人工核对身份。'}
        return self._call(operation)

    def _validate_confirmation(self, plan):
        from ..grounded_edits import validate_proof
        self.literature._bind()
        capture = next((value for value in self._captures.values() if
                        all(value['proof'][key] == plan.request.proof[key] for key in ('document', 'filename', 'csl'))), None)
        if (not capture or capture['session'] != self.literature.session or
                not self._included(plan.request.proof)):
            raise GuardError('professional_evidence_invalid')
        validate_proof(plan.request)

    def _included(self, proof):
        decision = self.literature.decisions.get(proof['document']['literature_id'], {}).get('decision')
        return decision == 'include' or bool(proof.get('archive') and decision is None)

    def _keep_capture(self, capture_id, proof, binding):
        self._captures = {key: value for key, value in self._captures.items()
                          if value['binding'] == binding and value['session'] == self.literature.session}
        if len(self._captures) >= 8 and capture_id not in self._captures:
            raise GuardError('professional_capacity')
        self._captures[capture_id] = {'proof': proof, 'binding': binding, 'session': self.literature.session}

    def capture_archive(self, save_id, page=0, limit=5, char_offset=0, allow_empty=False):
        start_page = page
        def operation(editor, path, binding):
            from ..pdf_reader import read_pdf
            try:
                source, receipt, origin, metadata = self.literature.archive_receipt(save_id)
                try:
                    report = read_pdf(source, page=start_page, limit=limit, char_offset=char_offset, max_chars=20000, layout=True)
                except Exception:
                    from ..pdf_reader import parser_identity
                    self._coverage.failure(save_id, receipt['files']['fulltext.pdf'][-1], start_page, limit, parser_identity())
                    raise
                self.literature._archive(editor, 'pdf').check_pdf_receipt(receipt)
            except Exception:
                raise GuardError('professional_archive_changed') from None
            if origin.get('retraction_status') == 'reported_retracted' or self.literature.records.get(origin['record_id'], {}).get('retraction_status') == 'reported_retracted':
                raise GuardError('professional_metadata_incomplete')
            from ..research_pdf import PDF_PARSER_VERSION
            document = self.literature._source_document(editor, origin['record_id'], 'fulltext.pdf', report['source_sha256'],
                                                       PDF_PARSER_VERSION, source_url=origin['source_url'], when=origin['retrieved_at'])
            fragments = []
            for page in report['pages']:
                for run in page['runs']:
                    text = run['text']
                    if not text.strip() or SECRET.search(text):
                        continue
                    location = f"pdf-page:{page['page'] + 1}/char:{run['char_index']}"
                    fragments.append(EvidenceFragment(id='ev:' + identifier(report['source_sha256'], location, text), document_id=document['id'],
                        source_sha256=report['source_sha256'], location=location, page=page['page'] + 1,
                        bbox=tuple(float(value) for value in run['bbox']), original_text=text,
                        extraction_quality='pdf_content_order', chunk_version='pdf-80char-v1').model_dump(mode='json'))
            if not fragments and not allow_empty:
                raise GuardError('professional_evidence_empty')
            proof = {'filename': 'fulltext.pdf', 'document': document, 'fragments': fragments, 'csl': metadata,
                     'archive': receipt, 'origin': origin}
            if SECRET.search(json.dumps(proof, ensure_ascii=False)):
                raise GuardError('professional_sensitive_text')
            coverage = self._coverage.record(save_id, report)
            proof['coverage'] = coverage
            capture_id = identifier(self._epoch, save_id, report['source_sha256'], coverage['range_id'])
            self._keep_capture(capture_id, proof, binding)
            if origin['record_id'] in self.literature.records:
                record = self.literature.records[origin['record_id']]
                record['fulltext_state'] = 'parsed_partial' if report['truncated'] else 'parsed_text_layer'
                record['fulltext_source'] = document
            return {'capture_id': capture_id, 'document': document, 'fragments': fragments, 'truncated': report['truncated'],
                    'coverage': coverage,
                    'title': metadata['title'], 'archive_save_id': save_id, 'license': origin['license'],
                    'inline': format_reference(metadata)[0], 'bibliography': format_reference(metadata)[1],
                    'doi_match_verified': False, 'coordinate_system': report['coordinate_system'],
                    'limitations': ['PDF content order; multicolumn order is not verified', 'Only displayed pages/characters were read']}
        return self._call(operation)

    def capture(self, record_id, filename, page=0, limit=5, char_offset=0, allow_empty=False):
        start_page = page
        def operation(editor, path, binding):
            from ..pdf_reader import read_pdf
            from ..rag_sources import SourceExtractor
            self.literature._bind()
            record = self.literature._record(record_id)
            if (self.literature.decisions.get(record_id, {}).get('decision') != 'include' or
                    record['retraction_status'] == 'reported_retracted' or not record['doi'] or
                    not record['title'] or not record['author_names'] or not record['year']):
                raise GuardError('professional_metadata_incomplete')
            if (type(filename) is not str or not filename or '/' in filename or '\\' in filename or ':' in filename or
                    not filename.lower().endswith('.pdf') or filename != filename.rstrip(' .') or path.parent != editor.space.root):
                raise GuardError('professional_evidence_path')
            extractor = SourceExtractor(editor.space)
            _, identity, fingerprint = extractor._snapshot(filename)
            try:
                report = read_pdf(editor.space.file(str(path.parent / filename)), page=start_page, limit=limit, char_offset=char_offset, max_chars=20000, layout=True)
            except Exception:
                from ..pdf_reader import parser_identity
                self._coverage.failure(filename, fingerprint, start_page, limit, parser_identity())
                raise
            _, after_identity, after = extractor._snapshot(filename)
            if identity != after_identity or fingerprint != after or report['source_sha256'] != fingerprint:
                raise GuardError('professional_evidence_changed')
            from ..research_pdf import PDF_PARSER_VERSION
            document = self.literature._source_document(editor, record_id, filename, fingerprint, PDF_PARSER_VERSION)
            fragments = []
            for page in report['pages']:
                for run in page['runs']:
                    text = run['text']
                    if not text.strip() or SECRET.search(text):
                        continue
                    location = f"pdf-page:{page['page'] + 1}/char:{run['char_index']}"
                    fragments.append(EvidenceFragment(id='ev:' + identifier(fingerprint, location, text), document_id=document['id'],
                                                     source_sha256=fingerprint, location=location, page=page['page'] + 1,
                                                     bbox=tuple(float(value) for value in run['bbox']), original_text=text,
                                                     extraction_quality='pdf_content_order', chunk_version='pdf-80char-v1').model_dump(mode='json'))
            if not fragments and not allow_empty:
                raise GuardError('professional_evidence_empty')
            proof = {'filename': filename, 'document': document, 'fragments': fragments, 'csl': csl(record)}
            if SECRET.search(json.dumps(proof, ensure_ascii=False)):
                raise GuardError('professional_sensitive_text')
            coverage = self._coverage.record(filename, report)
            proof['coverage'] = coverage
            capture_id = identifier(self._epoch, record_id, fingerprint, coverage['range_id'])
            self._keep_capture(capture_id, proof, binding)
            inline, bibliography = format_reference(proof['csl'])
            return {'capture_id': capture_id, 'document': document, 'fragments': fragments, 'truncated': report['truncated'],
                    'coverage': coverage,
                    'inline': inline, 'bibliography': bibliography, 'doi_match_verified': False,
                    'semantic_verification': 'not_assessed', 'coordinate_system': report['coordinate_system']}
        return self._call(operation)

    def reading_status(self):
        def operation(editor, path, binding):
            from ..rag_sources import SourceExtractor
            import re
            rows = self._coverage.statuses()
            for row in rows:
                try:
                    if re.fullmatch('[a-f0-9]{32}', row['source']):
                        _, receipt, _, _ = self.literature.archive_receipt(row['source'])
                        current = receipt['files']['fulltext.pdf'][-1]
                    else:
                        _, _, current = SourceExtractor(editor.space)._snapshot(row['source'])
                    row['stale'] = current != row['source_sha256']
                except Exception:
                    row['stale'] = True
                row['complete_current_version'] = row['complete'] and not row['stale']
            return {'sources': rows, 'read_only': True}
        return self._call(operation)

    def plan(self, payload):
        def operation(editor, path, binding):
            if type(payload) is not dict:
                raise GuardError('professional_invalid_request')
            draft = Draft.model_validate_json(json.dumps(payload))
            self.literature._bind()
            capture = self._captures.get(draft.capture_id)
            if (not capture or capture['binding'] != binding or capture['session'] != self.literature.session or
                    not self._included(capture['proof'])):
                raise GuardError('professional_evidence_invalid')
            if not draft.identity_reviewed:
                raise GuardError('professional_identity_required')
            source = snapshot(editor, path, limit=40)
            if source['document']['source_sha256'] != draft.source_sha256:
                raise GuardError('professional_source_changed')
            targets = {row['locator_id']: row for row in source['items']}
            inline, bibliography = format_reference(capture['proof']['csl'])
            changes = []
            for locator, text, reason in ((draft.body_locator, draft.support.claim + ' ' + inline, '依据人工核对的原文与适用限制修改主张'),
                                          (draft.reference_locator, bibliography, '插入固定版本 IEEE CSL 普通文本书目')):
                target = targets.get(locator)
                if not target:
                    raise GuardError('professional_stale_locator')
                changes.append(LocalChange(locator_id=locator, expected_text_sha256=target['text_sha256'], new_text=text, reason=reason))
            proof = {**capture['proof'], 'fragments': [row for row in capture['proof']['fragments'] if row['id'] in draft.support.evidence_ids]}
            request = GroundedDocx(path=str(path), source_sha256=draft.source_sha256, changes=changes,
                                   proof=proof, proof_seal=seal(proof), support=draft.support)
            for key, value in self._pending.items():
                self.literature.unpin(key)
                value['editor'].plans.pop(key, None)
            self._pending.clear()
            plan = editor.plan(request)
            self._pending[plan.id] = {'editor': editor, 'binding': binding, 'created': time.monotonic()}
            self.literature.pin(plan.id, [proof['archive']['save_id']] if proof.get('archive') else [])
            return {**plan.public(), 'items': json.loads(plan.after)['items'], 'wrote_nothing': True, 'expires_in_seconds': 300,
                    'style_version': request.style_version, 'semantic_verification': 'user_reviewed_not_machine_proven'}
        return self._call(operation)

    def page(self, capture_id, evidence_id):
        def operation(editor, path, binding):
            from ..grounded_edits import validate_source
            from ..research_page import render_page
            capture = self._captures.get(capture_id) if type(capture_id) is str else None
            if not capture or capture['binding'] != binding or capture['session'] != self.literature.session or not self._included(capture['proof']):
                raise GuardError('professional_evidence_invalid')
            proof = capture['proof']
            fragments = validate_source(proof, seal(proof), path)
            fragment = next((row for row in fragments if row.id == evidence_id), None)
            if fragment is None or fragment.page is None or fragment.bbox is None:
                raise GuardError('professional_evidence_invalid')
            source = self.literature._archive(editor, 'pdf').check_pdf_receipt(proof['archive']) if proof.get('archive') else editor.space.file(str(path.parent / proof['filename']))
            result = render_page(source, proof['document']['source_sha256'], fragment.page - 1, list(fragment.bbox))
            validate_source(proof, seal(proof), path)
            return result
        return self._call(operation)

    def ocr(self, capture_id, page):
        def operation(editor, path, binding):
            from ..grounded_edits import validate_source
            from ..research_page import render_page
            capture = self._captures.get(capture_id) if type(capture_id) is str else None
            if not capture or capture['binding'] != binding or capture['session'] != self.literature.session or not self._included(capture['proof']):
                raise GuardError('professional_evidence_invalid')
            proof = deepcopy(capture['proof'])
            validate_source(proof, seal(proof), path)
            if type(page) is not int or page not in [row['page'] for row in proof['coverage']['window']]:
                raise GuardError('professional_invalid_page')
            source = self.literature._archive(editor, 'pdf').check_pdf_receipt(proof['archive']) if proof.get('archive') else editor.space.file(str(path.parent / proof['filename']))
            try:
                result = render_page(source, proof['document']['source_sha256'], page - 1, [], ocr=True)
            except Exception:
                raise GuardError('professional_ocr_unavailable') from None
            validate_source(proof, seal(proof), path)
            if not result['lines'] or SECRET.search(json.dumps(result, ensure_ascii=False)):
                raise GuardError('professional_evidence_empty')
            fragments = []
            fingerprint = proof['document']['source_sha256']
            for index, line in enumerate(result['lines']):
                location = f'pdf-page:{page}/ocr-raw:{index}'
                fragments.append(EvidenceFragment(id='ev:' + identifier(fingerprint, location, line['text']), document_id=proof['document']['id'],
                    source_sha256=fingerprint, location=location, page=page, bbox=tuple(line['box_pdf']), original_text=line['text'],
                    extraction_quality='ocr_unverified', chunk_version='windows-ocr-v1').model_dump(mode='json'))
            proof.update(fragments=fragments, ocr={'raw': result, 'raw_fragments': deepcopy(fragments), 'corrections': [], 'reviewed': False})
            proof['document']['parser_version'] = 'windows-ocr-unverified-v1'
            key = identifier(self._epoch, fingerprint, str(page), json.dumps(fragments, ensure_ascii=False), 'ocr')
            self._keep_capture(key, proof, binding)
            return {'capture_id': key, 'document': proof['document'], 'fragments': fragments, 'coverage': proof['coverage'], 'truncated': True,
                    'ocr': proof['ocr'], 'inline': format_reference(proof['csl'])[0], 'bibliography': format_reference(proof['csl'])[1]}
        return self._call(operation)

    def review_ocr(self, payload):
        def operation(editor, path, binding):
            from ..grounded_edits import validate_source
            if type(payload) is not dict or set(payload) != {'capture_id', 'corrections', 'reviewed'} or payload['reviewed'] is not True:
                raise GuardError('professional_ocr_review_required')
            capture = self._captures.get(payload['capture_id'])
            if not capture or capture['binding'] != binding or capture['session'] != self.literature.session or not capture['proof'].get('ocr'):
                raise GuardError('professional_evidence_invalid')
            proof = deepcopy(capture['proof'])
            validate_source(proof, seal(proof), path)
            raw = proof['ocr']['raw_fragments']
            values = payload['corrections']
            if type(values) is not list or len(values) != len(raw):
                raise GuardError('professional_invalid_request')
            fragments = []
            for correction, original in zip(values, raw):
                if (type(correction) is not dict or set(correction) != {'evidence_id', 'text'} or correction['evidence_id'] != original['id'] or
                        type(correction['text']) is not str or not correction['text'].strip() or len(correction['text']) > 2000 or SECRET.search(correction['text'])):
                    raise GuardError('professional_invalid_request')
                location = original['location'] + '/human-corrected'
                fragments.append({**original, 'id': 'ev:' + identifier(original['source_sha256'], location, correction['text']), 'location': location,
                                  'original_text': correction['text'], 'context': '用户对原始OCR的独立校正；数值可靠性仍需人工核对。'})
            proof['fragments'] = fragments
            proof['ocr'].update(corrections=deepcopy(values), reviewed=True)
            key = identifier(self._epoch, json.dumps(proof, ensure_ascii=False), 'ocr-reviewed')
            self._keep_capture(key, proof, binding)
            return {'capture_id': key, 'document': proof['document'], 'fragments': fragments, 'coverage': proof['coverage'], 'truncated': True,
                    'ocr': proof['ocr'], 'inline': format_reference(proof['csl'])[0], 'bibliography': format_reference(proof['csl'])[1]}
        return self._call(operation)

    def confirm(self, plan_id, confirmed=False):
        with self._lock:
            reply = super().confirm(plan_id, confirmed)
            result = reply.get('result', {})
            if result.get('status') == 'saved' and plan_id in self.literature._dependencies:
                self.literature._dependencies[plan_id]['state'] = 'output'
            elif result.get('status') != 'uncertain' and plan_id not in self._pending:
                self.literature.unpin(plan_id)
            return reply

    def rollback(self, save_id, confirmed=False):
        with self._lock:
            reply = super().rollback(save_id, confirmed)
            if reply.get('result', {}).get('status') == 'rolled_back':
                self.literature.unpin(save_id)
            return reply
