import json
import re

from .range_history import RangeHistoryArchive
from .rag_evidence import EvidenceError

FORMAT = 'output-review-set-v1'


def validate(document):
    if type(document) is not dict or set(document) != {'format', 'scope', 'target', 'receipts', 'stamps'} or document['format'] != FORMAT:
        raise EvidenceError()
    receipts, stamps = document['receipts'], document['stamps']
    if (type(receipts) is not list or not 1 <= len(receipts) <= 32 or
            any(type(value) is not str or not re.fullmatch('[a-f0-9]{32}', value) for value in receipts) or
            receipts != sorted(set(receipts)) or type(document['target']) is not str or document['target'] not in receipts or
            type(stamps) is not dict or set(stamps) != set(receipts) or
            any(type(value) is not str or not re.fullmatch('[a-f0-9]{64}', value) for value in [document['scope'], *stamps.values()])):
        raise EvidenceError()
    return document


class ReviewSetArchive(RangeHistoryArchive):
    format = FORMAT
    validate = staticmethod(validate)

    def __init__(self, workspace):
        super().__init__(workspace)
        self.artifact_name = 'output-review-set.json'

    def plan_set(self, document, session):
        self.validate(document)
        return self.plan_payload(json.dumps(document, sort_keys=True).encode('utf8'), session, {'kind': self.format})
