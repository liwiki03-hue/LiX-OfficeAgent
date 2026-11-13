import hashlib

from .pdf_reader import read_pdf, parser_identity
import json
import time
from pathlib import Path
from .rag import structure_chunks
from .rag_sources import SourceExtractor, SourceError, MAX_IDENTITIES

_PARSER_IDENTITY = parser_identity()
_ADAPTER_VERSION = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
PDF_PARSER_VERSION = 'pdfium-content-order-v2:' + hashlib.sha256(json.dumps(_PARSER_IDENTITY, sort_keys=True).encode()).hexdigest()[:24]


class PdfResearchSource(SourceExtractor):
    def __init__(self, workspace):
        super().__init__(workspace)
        self._pdf_cached = None

    def clear_cache(self):
        super().clear_cache()
        self._pdf_cached = None

    def extract(self, name, *, start_page=None, end_page=None):
        from copy import deepcopy

        if type(name) is not str or not name.lower().endswith('.pdf') or '/' in name or '\\' in name or ':' in name:
            raise SourceError('invalid')
        explicit = start_page is not None or end_page is not None
        if explicit and (type(start_page) is not int or type(end_page) is not int or not 1 <= start_page <= end_page <= 2000 or end_page - start_page >= 20):
            raise SourceError('invalid')
        data, identity, fingerprint = self._snapshot(name)
        if not data.startswith(b'%PDF-'):
            raise SourceError('invalid')
        doc_id = hashlib.sha256(name.casefold().encode('utf-8')).hexdigest()
        if doc_id not in self._bindings and len(self._bindings) >= MAX_IDENTITIES:
            raise SourceError('capacity')
        key = (name, identity, fingerprint, PDF_PARSER_VERSION, _ADAPTER_VERSION, start_page, end_page)
        if self._pdf_cached is not None and self._pdf_cached[0] == key:
            result = deepcopy(self._pdf_cached[1])
        else:
            if explicit:
                pages, total, characters, started = [], None, 0, time.monotonic()
                for number in range(start_page, end_page + 1):
                    if time.monotonic() - started > 60:
                        raise SourceError('capacity')
                    report = read_pdf(self._path(name), page=number - 1, limit=1, max_chars=20000, layout=True)
                    if report['parser_identity'] != _PARSER_IDENTITY:
                        raise SourceError('source_changed')
                    if total is not None and total != report['page_count'] or report['source_sha256'] != fingerprint:
                        raise SourceError('source_changed')
                    total = report['page_count']
                    if end_page > total or len(report['pages']) != 1 or report['pages'][0]['page'] != number - 1:
                        raise SourceError('invalid')
                    current = report['pages'][0]
                    if current['text_truncated']:
                        raise SourceError('pdf_text_truncated')
                    if not current['text'].strip():
                        raise SourceError('pdf_text_missing')
                    characters += len(current['text'])
                    if characters > 100000:
                        raise SourceError('capacity')
                    pages.append(current)
                report = {'page_count': total, 'pages': pages}
            else:
                report = read_pdf(self._path(name), limit=5, max_chars=20000, layout=True)
                if report['truncated'] or report['page_count'] > 5:
                    raise SourceError('capacity')
            blocks = []
            for page in report['pages']:
                runs = page['runs']
                if explicit:
                    merged = []
                    for offset in range(0, len(runs), 10):
                        group = runs[offset:offset + 10]
                        merged.append({'char_index': group[0]['char_index'], 'text': ''.join(run['text'] for run in group),
                                       'bbox': [min(run['bbox'][0] for run in group), min(run['bbox'][1] for run in group),
                                                max(run['bbox'][2] for run in group), max(run['bbox'][3] for run in group)]})
                    runs = merged
                for run in runs:
                    if run['text'].strip():
                        box = ','.join(str(round(value, 2)) for value in run['bbox'])
                        location = f"pdf-page:{page['page'] + 1}/char:{run['char_index']}/box:{box}"
                        if len(location) > 128:
                            raise SourceError('capacity')
                        blocks.append({'source': name, 'type': 'paragraph', 'group': f"pdf-page:{page['page'] + 1}",
                                       'location': location, 'text': run['text']})
            if not blocks:
                raise SourceError('unsupported')
            structure_chunks(blocks)
            result = {'doc_id': doc_id, 'fingerprint': fingerprint, 'blocks': blocks, 'parser_version': PDF_PARSER_VERSION}
            if explicit:
                result['pdf_coverage'] = {'start_page': start_page, 'end_page': end_page, 'total_pages': total,
                                          'complete': start_page == 1 and end_page == total,
                                          'text_order': 'content_order_unreviewed'}
                version = hashlib.sha256((PDF_PARSER_VERSION + _ADAPTER_VERSION).encode()).hexdigest()[:24]
                result['parser_version'] = f'pdf-range-v1:{start_page}-{end_page}/{total}:' + version
        _, after_identity, after_fingerprint = self._snapshot(name)
        if identity != after_identity or fingerprint != after_fingerprint:
            raise SourceError('source_changed')
        self._bindings[doc_id] = (name, identity)
        self._pdf_cached = (key, deepcopy(result))
        return result
