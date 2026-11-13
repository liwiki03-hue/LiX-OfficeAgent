from copy import deepcopy
from datetime import date
import hashlib
import json
import re
from typing import Literal

from pydantic import Field, model_validator

from .literature import doi, text, safe_url, payload_hash
from .provenance import Record, LiteratureItem, acquired_at
from .rag import SECRET
from .rag_evidence import EvidenceError, MAX_BYTES
from .range_history import RangeHistoryArchive

FORMAT = 'research-project-v1'
CSL_FIELDS = {'id', 'type', 'title', 'author', 'issued', 'DOI', 'URL', 'container-title', 'publisher', 'volume', 'issue', 'page'}
KINDS = {'article-journal': 'journal-article', 'article': 'preprint', 'paper-conference': 'proceedings-article',
         'book': 'book', 'document': 'document', 'report': 'report', 'thesis': 'thesis', 'dataset': 'dataset'}


def fail(code='professional_project_invalid'):
    from .writes import GuardError
    raise GuardError(code)


def clean(value, maximum=500, empty=False):
    try:
        return text(value, maximum, empty=empty)
    except Exception:
        fail('professional_project_sensitive' if SECRET.search(str(value)) else 'professional_project_invalid')


def pairs(values):
    result = {}
    for key, value in values:
        if key in result:
            fail()
        result[key] = value
    return result


def canonical_csl(row):
    if type(row) is not dict or not {'title', 'type'} <= row.keys() or not row.keys() <= CSL_FIELDS or row['type'] not in KINDS:
        fail()
    result = {'title': clean(row['title'], 1000), 'type': row['type']}
    if 'id' in row:
        clean(row['id'], 300)
    for key in ('container-title', 'publisher', 'volume', 'issue', 'page'):
        if key in row:
            result[key] = clean(row[key], 500)
    if 'DOI' in row:
        try:
            result['DOI'] = doi(row['DOI'])
        except Exception:
            fail()
    if 'URL' in row:
        if not safe_url(row['URL']):
            fail('professional_project_url')
        result['URL'] = row['URL']
    if 'author' in row:
        if type(row['author']) is not list or not 1 <= len(row['author']) <= 100:
            fail()
        result['author'] = []
        for author in row['author']:
            if (type(author) is not dict or not author or not author.keys() <= {'given', 'family', 'literal'} or
                    ('literal' in author and len(author) != 1) or ('literal' not in author and not author.get('family'))):
                fail()
            result['author'].append({key: clean(value, 200) for key, value in author.items()})
    if 'issued' in row:
        issued = row['issued']
        if type(issued) is not dict or set(issued) != {'date-parts'} or type(issued['date-parts']) is not list or len(issued['date-parts']) != 1:
            fail()
        parts = issued['date-parts'][0]
        if type(parts) is not list or not 1 <= len(parts) <= 3 or any(type(value) is not int for value in parts) or not 1500 <= parts[0] <= 2100:
            fail()
        try:
            date(parts[0], parts[1] if len(parts) > 1 else 1, parts[2] if len(parts) > 2 else 1)
        except ValueError:
            fail()
        result['issued'] = {'date-parts': [parts]}
    identifier = 'lit:' + payload_hash(result)
    return {'id': identifier, **result}


def parse_csl(content):
    if type(content) is not str:
        fail()
    if SECRET.search(content):
        fail('professional_project_sensitive')
    if len(content.encode('utf8')) > 256000:
        fail('professional_project_capacity')
    try:
        rows = json.loads(content, object_pairs_hook=pairs, parse_constant=lambda value: fail())
    except Exception:
        fail()
    if type(rows) is not list or not 1 <= len(rows) <= 100:
        fail('professional_project_capacity')
    return [canonical_csl(row) for row in rows]


def import_rows(content):
    normalized = parse_csl(content)
    original = json.loads(content, object_pairs_hook=pairs)
    return [(row, {'source': 'csl_import', 'sha256': payload_hash(raw), 'record_id': raw.get('id', '')})
            for row, raw in zip(normalized, original)]


def record_from_csl(value, source='csl_import'):
    row = canonical_csl(value)
    authors = [' '.join(author[key] for key in ('given', 'family') if key in author) or author.get('literal', '') for author in row.get('author', [])]
    item = LiteratureItem(id=row['id'], title=row['title'], authors=tuple(authors),
        year=row.get('issued', {}).get('date-parts', [[None]])[0][0], doi=row.get('DOI'), source_url=row.get('URL'), retrieved_at=acquired_at())
    return {'item': item.model_dump(mode='json'), 'csl': row, 'source_records': [{'source': source, 'sha256': payload_hash(row)}],
            'tags': [], 'decision': 'undecided', 'reason': '', 'decisions': [], 'relations': [], 'link': None, 'link_history': [], 'read_pages': [], 'parsed_pages': [], 'indexed_pages': []}


def admitted_record(entry):
    row, item = entry['csl'], entry['item']
    return {**deepcopy(item), 'source': 'project_library', 'source_record_sha256': payload_hash(row),
            'kind': KINDS[row['type']], 'author_names': row.get('author', []), 'abstract': '', 'relations': entry['relations'],
            'container_title': row.get('container-title', ''), 'publisher': row.get('publisher', ''),
            'bibliographic': {key: row.get(key, '') for key in ('volume', 'issue', 'page')}, 'references': [],
            'references_truncated': False, 'missing_fields': [key for key in ('author', 'issued', 'DOI') if not row.get(key)],
            'fulltext_state': 'unassociated', 'pdf_url': None, 'license': None}


class SourceLink(Record):
    name: str = Field(min_length=1, max_length=128)
    sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    identity: list

    @model_validator(mode='after')
    def bounded(self):
        if any(char in self.name for char in '/\\:') or not self.name.lower().endswith('.pdf') or self.name != self.name.rstrip(' .'):
            fail()
        if len(json.dumps(self.identity)) > 2000:
            fail()
        return self


class SourceRecord(Record):
    source: str = Field(min_length=1, max_length=80)
    sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    record_id: str = Field(default='', max_length=300)


class Decision(Record):
    decision: Literal['include', 'exclude', 'undecided']
    reason: str = Field(max_length=400)
    revision: int = Field(ge=0)


class Relation(Record):
    type: str = Field(min_length=1, max_length=80)
    id: str = Field(min_length=1, max_length=256)
    reason: str = Field(default='', max_length=400)


class Target(Record):
    locator_id: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    location: str = Field(max_length=200)
    text: str = Field(max_length=8000)
    text_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    style: str | None
    version: str = Field(pattern=r'^[a-f0-9]{64}$')
    name: str = Field(min_length=1, max_length=128)


class QueueItem(Record):
    id: str = Field(pattern=r'^[a-f0-9]{32}$')
    operation: Literal['replace', 'insert', 'comment']
    text: str = Field(min_length=1, max_length=2000)
    target: Target
    evidence_ids: list[str] = Field(max_length=3)
    record_ids: list[str] = Field(max_length=3)
    quote: str = Field(min_length=1, max_length=2000)
    side: Literal['before', 'after']
    status: Literal['pending', 'accepted', 'rejected', 'recheck', 'delivered']
    relation: Literal['supports', 'contradicts', 'insufficient']
    applicability: str = Field(max_length=500)
    limitation: str = Field(max_length=500)
    numeric_reviews: list[str] = Field(default_factory=list, max_length=8)


class RegisteredClaim(Record):
    id: str = Field(pattern=r'^[a-f0-9]{32}$')
    record_id: str
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    evidence_id: str = Field(pattern=r'^ev:[a-f0-9]{64}$')
    location: str = Field(max_length=512)
    text: str = Field(min_length=1, max_length=2000)
    target: Target
    relation: Literal['supports', 'contradicts', 'insufficient']
    layout_stamp: str = Field(default='', max_length=64)
    original_quote: str = Field(default='', max_length=8000)
    original_header: str = Field(default='', max_length=1000)
    char_indices: list[int] = Field(default_factory=list, max_length=2000)


class LayoutRecord(Record):
    name: str = Field(max_length=128)
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    stamp: str = Field(pattern=r'^[a-f0-9]{16}$')
    parser_stamp: str = Field(pattern=r'^[a-f0-9]{64}$')


class OutputRecord(Record):
    receipt: str = Field(pattern=r'^[a-f0-9]{32}$')
    record_ids: list[str] = Field(max_length=8)
    source_stamps: dict[str, str] = Field(max_length=8)
    target_name: str = Field(max_length=128)
    target_version: str = Field(pattern=r'^[a-f0-9]{64}$')
    kind: Literal['task', 'document']


class EvidenceRecord(Record):
    receipt: str = Field(pattern=r'^[a-f0-9]{32}$')
    record_id: str = Field(max_length=100)
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    evidence_ids: list[str] = Field(min_length=1, max_length=32)

    @model_validator(mode='after')
    def bounded(self):
        if len(set(self.evidence_ids)) != len(self.evidence_ids) or any(not re.fullmatch(r'ev:[a-f0-9]{64}', value) for value in self.evidence_ids):
            fail()
        return self


class Entry(Record):
    item: LiteratureItem
    csl: dict
    source_records: list[SourceRecord] = Field(max_length=8)
    tags: list[str] = Field(max_length=12)
    decision: Literal['include', 'exclude', 'undecided']
    reason: str = Field(max_length=400)
    decisions: list[Decision] = Field(max_length=20)
    relations: list[Relation] = Field(max_length=20)
    link: SourceLink | None
    link_history: list[SourceLink] = Field(max_length=8)
    read_pages: list[int] = Field(max_length=2000)
    parsed_pages: list[int] = Field(max_length=2000)
    indexed_pages: list[int] = Field(max_length=20)

    @model_validator(mode='after')
    def consistent(self):
        row = canonical_csl(self.csl)
        if row != self.csl or row['id'] != self.item.id or row['title'] != self.item.title or row.get('DOI') != self.item.doi:
            fail()
        expected = record_from_csl(row)['item']
        if list(self.item.authors) != expected['authors'] or self.item.year != expected['year']:
            fail()
        if self.decision != 'undecided' and not self.reason.strip():
            fail()
        if any(not 1 <= page <= 2000 for page in self.read_pages) or self.read_pages != sorted(set(self.read_pages)):
            fail()
        for tag in self.tags:
            clean(tag, 40)
        return self


class Project(Record):
    format: Literal['research-project-v1'] = FORMAT
    id: str = Field(pattern=r'^[a-f0-9]{32}$')
    name: str = Field(min_length=1, max_length=100)
    scope: str = Field(pattern=r'^[a-f0-9]{64}$')
    generation: int = Field(ge=0, le=1000000)
    parent_receipt: str = Field(pattern=r'^(|[a-f0-9]{32})$')
    entries: dict[str, Entry] = Field(max_length=200)
    canonical: dict[str, str] = Field(max_length=200)
    claims: list[RegisteredClaim] = Field(max_length=32)
    tasks: list[OutputRecord] = Field(max_length=32)
    outputs: list[OutputRecord] = Field(max_length=32)
    evidence_records: list[EvidenceRecord] = Field(default_factory=list, max_length=32)
    layouts: list[LayoutRecord] = Field(max_length=8)
    drafts: list[QueueItem] = Field(max_length=12)

    @model_validator(mode='after')
    def bounded(self):
        if any(key != value.item.id for key, value in self.entries.items()):
            fail()
        if any(row.record_id not in self.entries for row in self.claims):
            fail()
        if any(row.record_id not in self.entries for row in self.evidence_records):
            fail()
        for entry in self.entries.values():
            for pages in (entry.read_pages, entry.parsed_pages, entry.indexed_pages):
                if pages != sorted(set(pages)) or any(type(page) is not int or not 1 <= page <= 2000 for page in pages):
                    fail()
        for row in self.tasks + self.outputs:
            if set(row.record_ids) != set(row.source_stamps) or not set(row.record_ids) <= self.entries.keys():
                fail()
            if any(not re.fullmatch(r'[a-f0-9]{64}', value) for value in row.source_stamps.values()):
                fail()
        if any(not set(row.record_ids) <= self.entries.keys() for row in self.drafts):
            fail()
        for key, value in self.canonical.items():
            if value not in self.entries or key != (self.entries[value].item.doi or value):
                fail()
        encoded = self.model_dump_json()
        if len(encoded.encode('utf8')) > MAX_BYTES:
            fail('professional_project_capacity')
        if SECRET.search(encoded):
            fail('professional_project_sensitive')
        return self


def validate(document):
    try:
        return Project.model_validate_json(json.dumps(document, allow_nan=False)).model_dump(mode='json')
    except Exception:
        raise EvidenceError('evidence_invalid') from None


class ProjectArchive(RangeHistoryArchive):
    format = FORMAT
    validate = staticmethod(validate)

    def __init__(self, workspace):
        super().__init__(workspace)
        self.artifact_name = 'research-project.json'

    def plan_project(self, document, session):
        document = validate(document)
        return self.plan_payload(json.dumps(document, ensure_ascii=False).encode('utf8'), session, {'kind': FORMAT})


def groups(entries):
    from .literature import group_records
    return group_records([admitted_record(entry) for entry in entries.values()])


def paper_families(project):
    entries = project['entries']
    parents = {key: key for key in entries}
    identities = {}
    def root(key):
        while parents[key] != key:
            parents[key] = parents[parents[key]]
            key = parents[key]
        return key
    def join(left, right):
        left, right = root(left), root(right)
        parents[max(left, right)] = min(left, right)
    for key, entry in entries.items():
        for identity in {key, entry['item'].get('doi')} - {None, ''}:
            if identity in identities:
                join(key, identities[identity])
            identities[identity] = key
    for key, entry in entries.items():
        for relation in entry['relations']:
            if relation['type'] in {'preprint_of', 'version_of'} and relation['id'] in identities:
                join(key, identities[relation['id']])
    return {key: 'registered-paper:' + payload_hash(root(key)) for key in entries}


def filtered(project, filters):
    allowed = {'query', 'author', 'year', 'type', 'tag', 'decision', 'fulltext', 'read', 'indexed'}
    if type(filters) is not dict or not filters.keys() <= allowed or any(type(v) is not str or len(v) > 200 for v in filters.values()):
        fail()
    result = []
    query = filters.get('query', '').casefold().strip()
    for key, entry in project['entries'].items():
        item, row = entry['item'], entry['csl']
        checks = [not filters.get('author') or filters['author'].casefold() in ' '.join(item['authors']).casefold(),
                  not filters.get('year') or filters['year'] == str(item['year']),
                  not filters.get('type') or filters['type'] == row['type'],
                  not filters.get('tag') or filters['tag'] in entry['tags'],
                  not filters.get('decision') or filters['decision'] == entry['decision'],
                  not filters.get('fulltext') or filters['fulltext'] == ('associated_unverified' if entry['link'] else 'none'),
                  not filters.get('read') or filters['read'] == ('partial' if entry['read_pages'] else 'unread')]
        checks.append(not filters.get('indexed') or filters['indexed'] == ('recorded_requires_revalidation' if entry['indexed_pages'] else 'unknown'))
        haystack = (item['title'] + ' ' + ' '.join(item['authors']) + ' ' + (item['doi'] or '')).casefold()
        terms = query.split()
        score = sum(haystack.count(term) for term in terms)
        if all(checks) and (not terms or score):
            result.append({'id': key, **deepcopy(entry), 'ranking': {'kind': 'metadata_literal_term_count', 'score': score,
                'fact_probability': False}, 'indexed_status': 'unknown_until_explicit_range_check'})
    return sorted(result, key=lambda row: (-row['ranking']['score'], row['item']['title'], row['id']))
