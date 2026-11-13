import hashlib
import html
import json
import re
import unicodedata
from urllib.parse import unquote, urlsplit

from pydantic import Field, model_validator

from .provenance import Record, LiteratureItem, acquired_at, identifier
from .rag import SECRET

VERSION = 'literature-metadata-v1'
KINDS = {'article': ('article', 'journal-article'), 'preprint': ('preprint', 'posted-content'),
         'book': ('book', 'book'), 'proceedings': ('proceedings-article', 'proceedings-article')}


class LiteratureError(ValueError):
    def __init__(self, code='literature_invalid'):
        self.code = code
        super().__init__(code)


def text(value, maximum=300, *, empty=False):
    if (type(value) is not str or len(value) > maximum or (not empty and not value.strip()) or
            any(ord(char) < 32 or ord(char) == 127 or 0xD800 <= ord(char) <= 0xDFFF for char in value)):
        raise LiteratureError()
    if SECRET.search(value):
        raise LiteratureError('literature_sensitive')
    return value.strip()


def doi(value):
    value = text(value, 256)
    value = re.sub(r'^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)', '', value, flags=re.I)
    value = unquote(value).strip().casefold()
    text(value, 256)
    if not re.fullmatch(r'10\.\d{4,9}/[^\s<>"{}\\]+', value) or any(ord(char) < 32 for char in value):
        raise LiteratureError('literature_doi_invalid')
    return value


def safe_url(value):
    if type(value) is not str or len(value) > 1000 or SECRET.search(value):
        return None
    try:
        parsed = urlsplit(value)
        if (parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password or
                parsed.port not in (None, 443) or parsed.query or parsed.fragment or
                any(ord(char) < 33 or char == '\\' for char in value)):
            return None
        return value
    except ValueError:
        return None


def clean(value, maximum=2000):
    if type(value) is not str:
        return ''
    value = html.unescape(re.sub(r'<[^>]{0,1000}>', '', value))
    value = re.sub(r'\s+', ' ', value).strip()
    if SECRET.search(value):
        raise LiteratureError('literature_sensitive')
    if len(value) > maximum:
        raise LiteratureError('literature_capacity')
    return text(value, maximum, empty=True)


class LiteratureQuery(Record):
    query: str = Field(min_length=1, max_length=300)
    mode: str = 'keywords'
    sources: tuple[str, ...] = ('crossref', 'openalex')
    author: str = Field(default='', max_length=100)
    year_from: int | None = Field(default=None, ge=1500, le=2100)
    year_to: int | None = Field(default=None, ge=1500, le=2100)
    kind: str = ''
    page: int = Field(default=1, ge=1, le=3)

    @model_validator(mode='after')
    def valid(self):
        text(self.query)
        text(self.author, 100, empty=True)
        if (self.mode not in {'keywords', 'title', 'doi'} or not self.sources or len(set(self.sources)) != len(self.sources) or
                not set(self.sources) <= {'crossref', 'openalex'} or self.kind not in {'', *KINDS} or
                any(char in self.author for char in ',:|!<>/\\') or
                (self.year_from and self.year_to and self.year_from > self.year_to)):
            raise LiteratureError()
        if self.mode == 'doi':
            doi(self.query)
            if self.page != 1 or self.author or self.year_from or self.year_to or self.kind:
                raise LiteratureError()
        return self


def abstract_from_index(value):
    if value is None:
        return ''
    if type(value) is not dict or len(value) > 1500:
        raise LiteratureError('literature_capacity')
    words = {}
    for word, positions in value.items():
        if type(word) is not str or type(positions) is not list or len(positions) > 1500:
            raise LiteratureError('literature_response')
        for position in positions:
            if type(position) is not int or not 0 <= position < 1500 or position in words:
                raise LiteratureError('literature_response')
            words[position] = word
    if words and sorted(words) != list(range(len(words))):
        raise LiteratureError('literature_response')
    return clean(' '.join(words[index] for index in sorted(words)), 12000)


def normalize(raw, source, when=None):
    if type(raw) is not dict:
        raise LiteratureError('literature_response')
    when = when or acquired_at()
    raw_doi = raw.get('DOI') if source == 'crossref' else raw.get('doi')
    canonical = doi(raw_doi) if raw_doi else None
    title = clean((raw.get('title') or [''])[0] if source == 'crossref' else raw.get('display_name'))
    authors, names, relations, references = [], [], [], []
    author_rows = raw.get('author') or [] if source == 'crossref' else raw.get('authorships') or []
    if type(author_rows) is not list or len(author_rows) > 500:
        raise LiteratureError('literature_capacity')
    for author in author_rows[:100]:
        if source == 'crossref':
            name = {key: clean(author.get(key), 200) for key in ('family', 'given') if author.get(key)}
            if not name and author.get('name'):
                name = {'literal': clean(author['name'], 200)}
            display = ' '.join(name[key] for key in ('given', 'family') if key in name) or name.get('literal', '')
        else:
            display = clean(author.get('author', {}).get('display_name'), 200)
            name = {'literal': display}
        if display:
            names.append(display)
            authors.append(name)
    year = raw.get('publication_year')
    if source == 'crossref':
        dates = (raw.get('issued') or {}).get('date-parts') or [[]]
        year = dates[0][0] if dates and dates[0] else None
    year = year if type(year) is int and 1500 <= year <= 2100 else None
    kind = clean(raw.get('type'), 80)
    if source == 'crossref':
        container_title = clean((raw.get('container-title') or [''])[0], 500)
        bibliographic = {key: clean(str(raw.get(key) or ''), 100) for key in ('volume', 'issue', 'page')}
        publisher = clean(raw.get('publisher'), 500)
    else:
        primary = raw.get('primary_location') or {}
        container_title = clean((primary.get('source') or {}).get('display_name'), 500)
        biblio = raw.get('biblio') or {}
        bibliographic = {key: clean(str(biblio.get(key) or ''), 100) for key in ('volume', 'issue')}
        bibliographic['page'] = clean('-'.join(str(biblio[key]) for key in ('first_page', 'last_page') if biblio.get(key)), 100)
        publisher = ''
    abstract = clean(raw.get('abstract'), 12000) if source == 'crossref' else abstract_from_index(raw.get('abstract_inverted_index'))
    public_id = raw.get('id') if source == 'openalex' else ('https://doi.org/' + canonical if canonical else None)
    source_url = safe_url(public_id)
    if source == 'openalex' and (not source_url or not re.fullmatch(r'https://openalex.org/W\d+', source_url)):
        raise LiteratureError('literature_response')
    if source == 'crossref':
        for relation, values in (raw.get('relation') or {}).items():
            for value in values[:20]:
                relations.append({'type': clean(relation, 80), 'id': clean(value.get('id'), 256)})
        for update in (raw.get('update-to') or [])[:20]:
            relations.append({'type': clean(update.get('type'), 80), 'id': clean(update.get('DOI'), 256)})
        references = [doi(row['DOI']) for row in (raw.get('reference') or [])[:20] if row.get('DOI')]
    else:
        references = [value for value in (raw.get('referenced_works') or [])[:20] if safe_url(value) and re.fullmatch(r'https://openalex.org/W\d+', value)]
    pdf_url, license_id = None, None
    location = raw.get('best_oa_location') or {} if source == 'openalex' else {}
    if location.get('is_oa') is True and location.get('pdf_url'):
        pdf_url = safe_url(location['pdf_url'])
        license_id = clean(location.get('license'), 100) or None
    retracted = source == 'openalex' and raw.get('is_retracted') is True
    source_version = payload_hash(raw)
    item = LiteratureItem(id='lit:' + identifier(source, source_url or canonical or title, source_version), title=title,
                          authors=tuple(names), year=year, doi=canonical, source_url=source_url, retrieved_at=when,
                          access='abstract' if abstract else 'metadata', retraction_status='reported_retracted' if retracted else 'unknown')
    result = item.model_dump(mode='json')
    result.update(source=source, source_record_sha256=source_version, kind=kind, abstract=abstract, author_names=authors, relations=relations,
                  container_title=container_title, publisher=publisher, bibliographic=bibliographic,
                  references=references, references_truncated=len(raw.get('reference') or raw.get('referenced_works') or []) > 20,
                  public_fulltext={'url': pdf_url, 'license': license_id, 'status': 'available_unverified' if pdf_url else 'not_available'},
                  fulltext_state='not_acquired', support_assessed=False, metadata_verified=False,
                  authors_truncated=len(author_rows) > 100, missing_fields=[field for field, value in [('title', title), ('authors', names), ('year', year), ('doi', canonical)] if not value])
    return result


def group_records(records):
    by_doi, groups, fingerprints = {}, [], set()
    for record in records:
        fingerprint = (record['source'], record['id'])
        if fingerprint in fingerprints:
            continue
        fingerprints.add(fingerprint)
        canonical = record['doi']
        if canonical and canonical in by_doi:
            group = by_doi[canonical]
            group['variants'].append(record)
        else:
            group = {'id': record['id'], 'doi': canonical, 'variants': [record], 'conflicts': [], 'possible_duplicates': []}
            groups.append(group)
            if canonical:
                by_doi[canonical] = group
    for group in groups:
        group['conflicts'] = [field for field in ('title', 'authors', 'year', 'kind', 'container_title', 'bibliographic')
                              if len({json.dumps(row[field], ensure_ascii=False) for row in group['variants']}) > 1]
        row = group['variants'][0]
        if not row['doi']:
            signature = (unicodedata.normalize('NFKC', row['title']).casefold(), tuple(row['authors']), row['year'])
            group['possible_duplicates'] = [other['id'] for other in groups if other is not group and
                                           (unicodedata.normalize('NFKC', other['variants'][0]['title']).casefold(),
                                            tuple(other['variants'][0]['authors']), other['variants'][0]['year']) == signature]
    return groups


def csl(record):
    mapping = {'journal-article': 'article-journal', 'article': 'article-journal', 'preprint': 'article',
               'posted-content': 'article', 'proceedings-article': 'paper-conference', 'book': 'book'}
    value = {'id': record['id'], 'type': mapping.get(record['kind'], 'document'), 'title': record['title']}
    if record['author_names']:
        value['author'] = record['author_names']
    if record['year']:
        value['issued'] = {'date-parts': [[record['year']]]}
    if record['doi']:
        value['DOI'] = record['doi']
    if record['source_url']:
        value['URL'] = record['source_url']
    if record['container_title']:
        value['container-title'] = record['container_title']
    if record['publisher']:
        value['publisher'] = record['publisher']
    value.update({key: field for key, field in record['bibliographic'].items() if field})
    return value


def payload_hash(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()
