import re

from .review_archive import ReviewSetArchive, validate as validate_set
from .rag_evidence import EvidenceError
from .rag import SECRET

FORMAT = 'range-review-set-v2'


def validate(document):
    if type(document) is not dict or set(document) != {'format', 'scope', 'target', 'receipts', 'stamps', 'review_items', 'pair'} or document['format'] != FORMAT:
        raise EvidenceError()
    validate_set({**{key: document[key] for key in ('scope', 'target', 'receipts', 'stamps')}, 'format': 'output-review-set-v1'})
    pair = document['pair']
    if type(pair) is not dict or set(pair) != {'source', 'output', 'analysis'} or any(type(value) is not str or not re.fullmatch('[a-f0-9]{64}', value) for value in pair.values()):
        raise EvidenceError()
    items = document['review_items']
    if type(items) is not list or not 1 <= len(items) <= 12:
        raise EvidenceError()
    for index, row in enumerate(items):
        if (type(row) is not dict or set(row) != {'index', 'state', 'reason'} or type(row['index']) is not int or row['index'] != index or
                type(row['state']) is not str or row['state'] not in {'unreviewed', 'viewed', 'needs_review', 'ignored'} or
                type(row['reason']) is not str or len(row['reason']) > 400 or SECRET.search(row['reason']) or
                any(ord(c) < 32 for c in row['reason']) or (row['state'] == 'ignored' and not row['reason'].strip())):
            raise EvidenceError()
    return document


class RangeReviewArchive(ReviewSetArchive):
    format = FORMAT
    validate = staticmethod(validate)

