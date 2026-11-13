import json
import re

from .literature_archive import LiteratureArchive
from .rag_evidence import _dir, _read, _hash
from .rag import SECRET
from .writes import GuardError


class ResearchArchive(LiteratureArchive):
    def __init__(self, workspace):
        super().__init__(workspace, 'record')
        self.artifact_name = 'research-task.json'

    def restore(self, save_id):
        self._guard()
        if type(save_id) is not str or not re.fullmatch('[a-f0-9]{32}', save_id):
            raise GuardError('professional_assist_receipt_invalid')
        directory = self.root / ('.office-rag-evidence-' + save_id)
        identity = _dir(directory)
        if {item.name for item in directory.iterdir()} != {'backup', 'audit-intent.json', 'audit-verified.json', self.artifact_name}:
            raise GuardError('professional_assist_receipt_invalid')
        values, files = {}, {}
        for name in ('backup', 'audit-intent.json', 'audit-verified.json', self.artifact_name):
            values[name], files[name] = _read(directory / name)
        audit = json.loads(values['audit-verified.json'])
        if (values['backup'] != b'' or values['audit-intent.json'] != values['audit-verified.json'] or
                audit.get('operation') != 'save_literature' or audit.get('artifact') != self.artifact_name or
                audit.get('payload_sha256') != _hash(values[self.artifact_name]) or
                audit.get('provenance') != {'kind': 'research-task-v1'} or SECRET.search(values[self.artifact_name].decode('utf8'))):
            raise GuardError('professional_assist_receipt_invalid')
        self._saved_check({'directory': directory, 'identity': identity, 'files': files})
        data = json.loads(values[self.artifact_name])
        if (set(data) != {'format', 'task', 'drafts', 'source_parser'} or data['format'] != 'research-task-v1' or
                data['task']['request'].get('outbound_approved') is not False):
            raise GuardError('professional_assist_receipt_invalid')
        return data
