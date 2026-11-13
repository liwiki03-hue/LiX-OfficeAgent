import json
import re

from .document_tasks import TaskSpec, MAX_ITEMS
from .range_edits import RangeOperation
from .block_edits import BlockOperation
from .range_history import RangeHistoryArchive
from .rag_evidence import EvidenceError
from .rag import SECRET

FORMAT = 'document-task-v1'
KEYS = {'id', 'spec', 'source_sha256', 'current_sha256', 'source_name', 'current_name', 'items', 'groups',
        'group_index', 'coverage', 'receipts', 'generation', 'expected', 'metrics', 'uncertain', 'checks', 'anchor_map'}


def validate(document):
    try:
        if type(document) is not dict or set(document) != {'format', 'scope', 'task'} or document['format'] != FORMAT:
            raise ValueError()
        task = document['task']
        if type(task) is not dict or set(task) != KEYS or not re.fullmatch('[a-f0-9]{64}', document['scope']):
            raise ValueError()
        TaskSpec.model_validate(task['spec'])
        for key in ('source_sha256', 'current_sha256'):
            if not re.fullmatch('[a-f0-9]{64}', task[key]):
                raise ValueError()
        for key in ('source_name', 'current_name'):
            name = task[key]
            if type(name) is not str or not name.lower().endswith('.docx') or len(name) > 220 or any(c in name for c in '/\\:') or any(ord(c) < 32 for c in name):
                raise ValueError()
        if not re.fullmatch('[a-f0-9]{32}', task['id']) or type(task['items']) is not list or not 1 <= len(task['items']) <= MAX_ITEMS:
            raise ValueError()
        ids = []
        for item in task['items']:
            if type(item) is not dict or not {'id', 'requirement', 'location', 'operation', 'state', 'reason', 'receipt', 'position'} <= item.keys() or not set(item) <= {'id', 'requirement', 'location', 'operation', 'state', 'reason', 'receipt', 'position', 'occurrences', 'source_location'}:
                raise ValueError()
            if not re.fullmatch('[a-f0-9]{32}', item['id']) or item['state'] not in {'pending', 'completed', 'no_change', 'excluded', 'restricted', 'failed', 'uncertain'}:
                raise ValueError()
            if item['operation'] is not None:
                if item['operation'].get('action') in {'paragraphs_replace', 'paragraphs_merge', 'chapter_move', 'bookmark_text'}:
                    BlockOperation.model_validate(item['operation'])
                else:
                    RangeOperation.model_validate(item['operation'])
            if type(item['reason']) is not str or len(item['reason']) > 400 or type(item['position']) is not int:
                raise ValueError()
            ids.append(item['id'])
        if len(set(ids)) != len(ids) or type(task['groups']) is not list or len(task['groups']) > 16:
            raise ValueError()
        grouped = []
        for group in task['groups']:
            if type(group) is not list or not 1 <= len(group) <= 12 or any(key not in ids for key in group):
                raise ValueError()
            grouped += group
        if len(set(grouped)) != len(grouped) or type(task['group_index']) is not int or not 0 <= task['group_index'] <= len(task['groups']):
            raise ValueError()
        if type(task['generation']) is not int or task['generation'] < 1 or type(task['uncertain']) is not bool or type(task['expected']) is not dict or len(task['expected']) > 4000:
            raise ValueError()
        for key, value in task['expected'].items():
            if type(key) is not str or len(key) > 200 or type(value) is not dict or set(value) != {'text', 'node'} or any(not re.fullmatch('[a-f0-9]{64}', v) for v in value.values()):
                raise ValueError()
        if type(task['anchor_map']) is not dict or len(task['anchor_map']) > 4000 or any(not re.fullmatch('loc:[a-f0-9]{64}', key) or (value is not None and (type(value) is not str or len(value) > 200)) for key, value in task['anchor_map'].items()):
            raise ValueError()
        if len(task['receipts']) > 16 or any(not re.fullmatch('[a-f0-9]{32}', key) for key in task['receipts']):
            raise ValueError()
        if SECRET.search(json.dumps(document, ensure_ascii=False)):
            raise EvidenceError('evidence_sensitive')
        return document
    except EvidenceError:
        raise
    except Exception:
        raise EvidenceError('evidence_invalid') from None


class TaskArchive(RangeHistoryArchive):
    format = FORMAT
    validate = staticmethod(validate)

    def __init__(self, workspace):
        super().__init__(workspace)
        self.artifact_name = 'document-task.json'

    def plan_task(self, task, scope, session):
        document = {'format': FORMAT, 'scope': scope, 'task': {key: task[key] for key in KEYS}}
        validate(document)
        return self.plan_payload(json.dumps(document, ensure_ascii=False).encode('utf8'), session, {'kind': FORMAT})
