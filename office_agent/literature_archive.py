import json
import re

from .rag_evidence import RagEvidence, MAX_BYTES, TTL, _hash, _json, _dir, _read
from .literature import LiteratureError


class LiteratureArchive(RagEvidence):
    def __init__(self, workspace, kind):
        super().__init__(workspace, lambda key: None)
        self.artifact_name = {'csl': 'references.csl.json', 'record': 'research.json', 'pdf': 'fulltext.pdf'}[kind]

    def pdf_receipt(self, save_id):
        with self._lock:
            self._guard()
            if self.artifact_name != 'fulltext.pdf' or type(save_id) is not str or not re.fullmatch('[a-f0-9]{32}', save_id):
                raise LiteratureError('literature_plan_invalid')
            directory = self.root / ('.office-rag-evidence-' + save_id)
            identity = _dir(directory)
            if {item.name for item in directory.iterdir()} != {'backup', 'audit-intent.json', 'audit-verified.json', 'fulltext.pdf'}:
                raise LiteratureError('literature_receipt_changed')
            values, files = {}, {}
            for name in ('backup', 'audit-intent.json', 'audit-verified.json', 'fulltext.pdf'):
                values[name], files[name] = _read(directory / name)
            audit = json.loads(values['audit-verified.json'])
            if (values['backup'] != b'' or values['audit-intent.json'] != values['audit-verified.json'] or
                    audit.get('operation') != 'save_literature' or audit.get('artifact') != 'fulltext.pdf' or
                    audit.get('payload_sha256') != _hash(values['fulltext.pdf'])):
                raise LiteratureError('literature_receipt_changed')
            saved = {'directory': directory, 'identity': identity, 'files': files}
            self._saved_check(saved)
            return {'save_id': save_id, 'identity': identity, 'files': files}, audit['provenance']

    def check_pdf_receipt(self, receipt):
        def tuples(value):
            return tuple(tuples(item) for item in value) if isinstance(value, (tuple, list)) else value
        if (type(receipt) is not dict or set(receipt) != {'save_id', 'identity', 'files'} or
                type(receipt['save_id']) is not str or not re.fullmatch('[a-f0-9]{32}', receipt['save_id']) or
                set(receipt['files']) != {'backup', 'audit-intent.json', 'audit-verified.json', 'fulltext.pdf'}):
            raise LiteratureError('literature_receipt_changed')
        saved = {'directory': self.root / ('.office-rag-evidence-' + receipt['save_id']),
                 'identity': tuples(receipt['identity']), 'files': {key: tuples(value) for key, value in receipt['files'].items()}}
        if {item.name for item in saved['directory'].iterdir()} != set(saved['files']):
            raise LiteratureError('literature_receipt_changed')
        self._saved_check(saved)
        return saved['directory'] / 'fulltext.pdf'

    def _save_audit(self, plan):
        return _json({**json.loads(super()._save_audit(plan)), 'operation': 'save_literature',
                      'artifact': self.artifact_name, 'provenance': plan['provenance']})

    def plan_payload(self, payload, session, provenance=None):
        with self._lock:
            self.ensure_available(session)
            if type(payload) is not bytes or not payload or len(payload) > MAX_BYTES:
                raise LiteratureError('literature_capacity')
            if self.artifact_name.endswith('.json'):
                json.loads(payload)
            elif not payload.startswith(b'%PDF-'):
                raise LiteratureError('literature_response')
            token = self._id()
            self._plans[token] = {'session': session, 'time': self._clock(), 'operation': 'save', 'payload': payload,
                                 'hits': [], 'provenance': provenance or {}}
            return {'plan_id': token, 'operation': 'save_literature', 'filename': self.artifact_name,
                    'byte_count': len(payload), 'payload_sha256': _hash(payload),
                    'confirmation_required': True, 'wrote_nothing': True, 'expires_in_seconds': TTL}
