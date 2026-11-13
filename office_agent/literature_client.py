from copy import deepcopy
from dataclasses import dataclass, field
import json
import math
import re
import time
from urllib.parse import quote, urlsplit

import httpx2

from .literature import LiteratureError, LiteratureQuery, KINDS, doi, normalize, payload_hash
from .provenance import acquired_at

BASES = {'crossref': 'https://api.crossref.org', 'openalex': 'https://api.openalex.org'}


@dataclass(frozen=True, repr=False)
class LiteratureConfig:
    openalex_key: str = field(default='', repr=False)
    key_requests: int = 0
    timeout: float = 8
    request_limit: int = 20

    def __post_init__(self):
        if (type(self.openalex_key) is not str or len(self.openalex_key) > 512 or
                any(char.isspace() or ord(char) < 32 for char in self.openalex_key) or
                type(self.key_requests) is not int or not 0 <= self.key_requests <= 20 or
                type(self.request_limit) is not int or not 1 <= self.request_limit <= 50 or
                type(self.timeout) not in (float, int) or not math.isfinite(self.timeout) or not 1 <= self.timeout <= 10):
            raise LiteratureError('literature_configuration')

    @classmethod
    def from_values(cls, values):
        try:
            return cls(values.get('OPENALEX_API_KEY') or '', int(values.get('OFFICE_LITERATURE_KEY_REQUESTS') or 0),
                       float(values.get('OFFICE_LITERATURE_TIMEOUT') or 8), int(values.get('OFFICE_LITERATURE_REQUEST_LIMIT') or 20))
        except Exception:
            raise LiteratureError('literature_configuration') from None


def _unique(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise LiteratureError('literature_response')
        result[key] = value
    return result


class LiteratureClient:
    def __init__(self, config=None, *, transport_factory=None, clock=time.monotonic):
        self.config = config or LiteratureConfig()
        self.transport = transport_factory or (lambda: httpx2.HTTPTransport(retries=0, trust_env=False))
        self.clock = clock
        self.calls = 0
        self.key_calls = 0
        self.next_at = {}
        self.cache = {}

    def __repr__(self):
        return '<LiteratureClient redacted>'

    def invalidate(self):
        self.cache.clear()

    def _get(self, source, path, params=None, *, binary=False, check=lambda: None):
        check()
        if self.calls >= self.config.request_limit:
            raise LiteratureError('literature_budget')
        now = self.clock()
        if now < self.next_at.get(source, 0):
            raise LiteratureError('literature_rate_limited')
        headers = {'Accept': 'application/pdf' if binary else 'application/json', 'Accept-Encoding': 'identity',
                   'User-Agent': 'OfficeAgent-Literature/1.0'}
        if source == 'openalex' and self.config.openalex_key and self.config.key_requests:
            if self.key_calls >= self.config.key_requests:
                raise LiteratureError('literature_budget')
            headers['Authorization'] = 'Bearer ' + self.config.openalex_key
            self.key_calls += 1
        url = path if binary else BASES[source] + path
        if binary and not re.fullmatch(r'https://arxiv\.org/pdf/(?:\d{4}\.\d{4,5}(?:v\d+)?|[a-z-]+/\d{7}(?:v\d+)?)(?:\.pdf)?', url):
            raise LiteratureError('literature_fulltext_refused')
        maximum = 512 * 1024 if binary else 2_000_000
        self.calls += 1
        self.next_at[source] = now + 1
        deadline = now + self.config.timeout
        try:
            with httpx2.Client(transport=self.transport(), trust_env=False, follow_redirects=False) as client:
                with client.stream('GET', url, params=params, headers=headers, timeout=self.config.timeout) as response:
                    if response.status_code == 429:
                        wait = response.headers.get('retry-after', '')
                        self.next_at[source] = self.clock() + (min(300, max(1, int(wait))) if wait.isdecimal() else 60)
                        raise LiteratureError('literature_rate_limited')
                    if response.status_code in (401, 403):
                        raise LiteratureError('literature_auth_required')
                    if response.status_code == 404:
                        raise LiteratureError('literature_not_found')
                    if 300 <= response.status_code < 400:
                        raise LiteratureError('literature_redirect_refused')
                    if response.status_code != 200:
                        raise LiteratureError('literature_service_failed')
                    mime = response.headers.get('content-type', '').split(';')[0].strip().lower()
                    if (response.headers.get('content-encoding', 'identity').lower() != 'identity' or
                            (binary and mime != 'application/pdf') or
                            (not binary and not (mime == 'application/json' or mime.endswith('+json')))):
                        raise LiteratureError('literature_response')
                    length = response.headers.get('content-length')
                    if length is not None and (not length.isdecimal() or int(length) > maximum):
                        raise LiteratureError('literature_capacity')
                    body = bytearray()
                    for part in response.iter_raw():
                        check()
                        if self.clock() > deadline:
                            raise LiteratureError('literature_timeout')
                        if len(body) + len(part) > maximum:
                            raise LiteratureError('literature_capacity')
                        body.extend(part)
            check()
            if self.clock() > deadline:
                raise LiteratureError('literature_timeout')
            if binary:
                if not body.startswith(b'%PDF-'):
                    raise LiteratureError('literature_response')
                return bytes(body)
            try:
                result = json.loads(body, object_pairs_hook=_unique, parse_constant=lambda value: (_ for _ in ()).throw(ValueError()))
            except Exception:
                raise LiteratureError('literature_response') from None
            if type(result) is not dict:
                raise LiteratureError('literature_response')
            return result
        except LiteratureError:
            raise
        except httpx2.TimeoutException:
            raise LiteratureError('literature_timeout') from None
        except Exception:
            raise LiteratureError('literature_service_failed') from None

    def _parameters(self, source, query):
        filters = []
        if source == 'crossref':
            if query.mode == 'doi':
                return '/works/' + quote(doi(query.query), safe=''), {}
            params = {'rows': 10, 'offset': (query.page - 1) * 10,
                      'query.bibliographic' if query.mode == 'title' else 'query': query.query}
            if query.author:
                params['query.author'] = query.author
            if query.year_from:
                filters.append(f'from-pub-date:{query.year_from}-01-01')
            if query.year_to:
                filters.append(f'until-pub-date:{query.year_to}-12-31')
            if query.kind:
                filters.append('type:' + KINDS[query.kind][1])
        else:
            if query.mode == 'doi':
                return '/works/https://doi.org/' + quote(doi(query.query), safe=''), {}
            params = {'per_page': 10, 'page': query.page}
            if query.mode == 'title':
                if any(char in query.query for char in ',:|!<>'):
                    raise LiteratureError('literature_invalid')
                filters.append('title.search:' + query.query)
            else:
                params['search'] = query.query
            if query.author:
                filters.append('raw_author_name.search:' + query.author)
            if query.year_from or query.year_to:
                filters.append(f'publication_year:{query.year_from or 1500}-{query.year_to or 2100}')
            if query.kind:
                filters.append('type:' + KINDS[query.kind][0])
        if filters:
            params['filter'] = ','.join(filters)
        return '/works', params

    def search(self, query, *, check=lambda: None):
        if type(query) is not LiteratureQuery:
            raise LiteratureError()
        records, reports = [], []
        for source in query.sources:
            check()
            expression = query.model_dump(mode='json')
            key = payload_hash([source, expression])
            cached = self.cache.get(key)
            if cached and self.clock() - cached[0] < 120:
                records.extend(deepcopy(cached[1]))
                reports.append({**deepcopy(cached[2]), 'cached': True})
                continue
            when = acquired_at()
            try:
                path, params = self._parameters(source, query)
                raw = self._get(source, path, params, check=check)
                if query.mode == 'doi':
                    rows = [raw['message'] if source == 'crossref' else raw]
                    total = 1
                elif source == 'crossref':
                    rows, total = raw['message']['items'], raw['message']['total-results']
                else:
                    rows, total = raw['results'], raw['meta']['count']
                if type(rows) is not list or len(rows) > 10 or type(total) is not int or total < 0:
                    raise LiteratureError('literature_response')
                normalized = [normalize(row, source, when) for row in rows]
                if len(json.dumps(normalized, ensure_ascii=False).encode('utf-8')) > 250000:
                    raise LiteratureError('literature_capacity')
                if query.mode == 'doi' and any(row['doi'] != doi(query.query) for row in normalized):
                    raise LiteratureError('literature_response')
                report = {'source': source, 'status': 'completed', 'expression': expression, 'retrieved_at': when,
                          'page': query.page, 'total_reported': total, 'returned': len(rows),
                          'truncated': query.mode != 'doi' and total > query.page * 10,
                          'cached': False, 'response_sha256': payload_hash(raw),
                          'auth_mode': 'configured_key' if source == 'openalex' and self.config.openalex_key and self.config.key_requests else 'anonymous'}
                check()
                if len(self.cache) >= 8:
                    self.cache.pop(next(iter(self.cache)))
                self.cache[key] = (self.clock(), deepcopy(normalized), deepcopy(report))
                records.extend(normalized)
                reports.append(report)
            except LiteratureError as error:
                reports.append({'source': source, 'status': error.code, 'expression': expression, 'retrieved_at': when,
                                'page': query.page, 'cached': False})
            except Exception:
                reports.append({'source': source, 'status': 'literature_response', 'expression': expression, 'retrieved_at': when,
                                'page': query.page, 'cached': False})
        check()
        return records, reports

    def fulltext(self, record, *, check=lambda: None):
        public = record['public_fulltext']
        if public['license'] not in {'cc0', 'cc-by', 'cc-by-sa', 'cc-by-nc', 'cc-by-nc-sa'} or not public['url']:
            raise LiteratureError('literature_fulltext_refused')
        return self._get('arxiv', public['url'], binary=True, check=check)

    def citing(self, record, *, check=lambda: None):
        parsed = urlsplit(record['source_url'] or '')
        if record['source'] != 'openalex' or not re.fullmatch(r'/W\d+', parsed.path):
            raise LiteratureError('literature_expansion_refused')
        raw = self._get('openalex', '/works', {'filter': 'cites:' + parsed.path[1:], 'per_page': 10}, check=check)
        rows, total = raw['results'], raw['meta']['count']
        if type(rows) is not list or len(rows) > 10 or type(total) is not int or total < 0:
            raise LiteratureError('literature_response')
        when = acquired_at()
        normalized = [normalize(row, 'openalex', when) for row in rows]
        if len(json.dumps(normalized, ensure_ascii=False).encode('utf-8')) > 250000:
            raise LiteratureError('literature_capacity')
        return normalized, {'source': 'openalex', 'status': 'completed',
                'expression': {'cites': record['source_url']}, 'retrieved_at': when, 'returned': len(rows),
                'truncated': total > 10, 'page': 1, 'response_sha256': payload_hash(raw), 'cached': False,
                'auth_mode': 'configured_key' if self.config.openalex_key and self.config.key_requests else 'anonymous'}
