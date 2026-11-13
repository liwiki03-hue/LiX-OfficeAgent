import json
import os
import re

from .literature_archive import LiteratureArchive
from .rag_evidence import EvidenceError, MAX_PENDING, _dir, _hash, _read
from .rag import SECRET

FORMAT = 'pdf-range-attempt-v1'
STATES = {'failed', 'rejected', 'expired', 'cancelled'}
REASONS = {'rag_capacity', 'rag_source_changed', 'pdf_text_missing', 'pdf_text_truncated', 'rag_invalid', 'rag_model_mismatch', 'processing_failed'}


def entries(directory):
    names = set()
    with os.scandir(directory) as iterator:
        for item in iterator:
            names.add(item.name)
            if len(names) >= 5:
                break
    return names


def validate(document):
    if type(document) is not dict or set(document) != {'format', 'source_name', 'attempt'} or document['format'] != FORMAT:
        raise EvidenceError()
    name, attempt = document['source_name'], document['attempt']
    if (type(name) is not str or not 1 <= len(name) <= 128 or not name.lower().endswith('.pdf') or name.lower().startswith('.env') or
            any(c in name for c in '/\\:') or any(ord(c) < 32 for c in name)):
        raise EvidenceError()
    required = {'id', 'scope', 'source_fingerprint', 'embedding_stamp', 'processing_stamp', 'selected_pages', 'parsed_pages',
                'reused_pages', 'failed_pages', 'unprocessed_pages', 'state', 'published', 'source_is_current'}
    if type(attempt) is not dict or not required <= attempt.keys() or not attempt.keys() <= required | {'reason'}:
        raise EvidenceError()
    for key, length in [('id', 32), ('scope', 24), ('source_fingerprint', 64), ('embedding_stamp', 64), ('processing_stamp', 64)]:
        if type(attempt[key]) is not str or not re.fullmatch('[a-f0-9]{' + str(length) + '}', attempt[key]):
            raise EvidenceError()
    pages = {}
    for key in ('selected_pages', 'parsed_pages', 'reused_pages', 'failed_pages', 'unprocessed_pages'):
        values = attempt[key]
        if (type(values) is not list or len(values) > 20 or any(type(n) is not int or not 1 <= n <= 2000 for n in values) or
                values != sorted(set(values))):
            raise EvidenceError()
        pages[key] = set(values)
    if (any(not values <= pages['selected_pages'] for values in pages.values()) or
            not pages['reused_pages'] <= pages['parsed_pages'] or pages['parsed_pages'] & pages['failed_pages'] or
            pages['unprocessed_pages'] != pages['selected_pages'] - pages['parsed_pages'] - pages['failed_pages'] or
            type(attempt['state']) is not str or attempt['state'] not in STATES or attempt['published'] is not False or
            type(attempt['source_is_current']) is not bool or
            ('reason' in attempt and (type(attempt['reason']) is not str or attempt['reason'] not in REASONS))):
        raise EvidenceError()
    encoded = json.dumps(document, ensure_ascii=False, allow_nan=False).encode('utf8')
    if len(encoded) > 16384 or SECRET.search(encoded.decode('utf8')):
        raise EvidenceError('evidence_sensitive')
    return document


class RangeHistoryArchive(LiteratureArchive):
    format = FORMAT
    validate = staticmethod(validate)

    def __init__(self, workspace):
        super().__init__(workspace, 'record')
        self.artifact_name = 'pdf-range-attempt.json'

    def plan_attempt(self, document, session):
        validate(document)
        return self.plan_payload(json.dumps(document, ensure_ascii=False).encode('utf8'), session, {'kind': FORMAT})

    def inspect(self, save_id):
        with self._lock:
            self._guard()
            if type(save_id) is not str or not re.fullmatch('[a-f0-9]{32}', save_id):
                raise EvidenceError('evidence_invalid')
            directory = self.root / ('.office-rag-evidence-' + save_id)
            identity = _dir(directory)
            names = {'backup', 'audit-intent.json', 'audit-verified.json', self.artifact_name}
            if entries(directory) != names:
                raise EvidenceError('evidence_path_changed')
            values, files = {}, {}
            for name in names:
                values[name], files[name] = _read(directory / name)
            audit = json.loads(values['audit-verified.json'])
            if (values['backup'] != b'' or values['audit-intent.json'] != values['audit-verified.json'] or
                    audit.get('operation') != 'save_literature' or audit.get('artifact') != self.artifact_name or
                    audit.get('payload_sha256') != _hash(values[self.artifact_name]) or audit.get('provenance') != {'kind': self.format}):
                raise EvidenceError('evidence_verification_failed')
            saved = {'directory': directory, 'identity': identity, 'files': files, 'hits': []}
            document = self.validate(json.loads(values[self.artifact_name]))
            self._saved_check(saved)
            if entries(directory) != names:
                raise EvidenceError('evidence_path_changed')
            return document, saved

    def plan_removal(self, save_id, session):
        with self._lock:
            self._session(session)
            self._guard()
            _, saved = self.inspect(save_id)
            if len(self._saved) >= MAX_PENDING and save_id not in self._saved:
                raise EvidenceError('evidence_capacity')
            self._saved[save_id] = {**saved, 'session': session}
            return self.plan_rollback(save_id, session)
