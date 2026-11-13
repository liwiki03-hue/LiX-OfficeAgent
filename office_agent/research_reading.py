from .provenance import identifier
import json


class PdfCoverage:
    def __init__(self):
        self.documents = {}

    def clear(self):
        self.documents.clear()

    def failure(self, key, fingerprint, page, limit, parser=None):
        if type(page) is not int or type(limit) is not int or not 0 <= page <= 1999 or not 1 <= limit <= 5:
            return
        value = self.documents.get(key)
        if value is None or value['source_sha256'] != fingerprint:
            value = {'source_sha256': fingerprint, 'parser': parser['implementation'] if parser else 'pdfium-content-order-v1',
                     'parser_identity': parser, 'page_count': None, 'intervals': {}, 'totals': {}, 'empty_pages': [], 'failed_pages': [], 'indexed_ranges': []}
        value['failed_ranges'] = (value.get('failed_ranges', []) + [{'start_page': page + 1, 'limit': limit, 'code': 'page_read_failed'}])[-16:]
        if value['page_count'] is not None:
            value['failed_pages'] = sorted(set(value['failed_pages']) | set(range(page + 1, min(page + limit, value['page_count']) + 1)))
        if len(self.documents) >= 8 and key not in self.documents:
            self.documents.pop(next(iter(self.documents)))
        self.documents[key] = value

    def statuses(self):
        values = []
        for key, value in self.documents.items():
            complete = sorted(number for number, intervals in value['intervals'].items() if intervals == [[0, value['totals'][number]]])
            values.append({'source': key, **value, 'complete_pages': complete, 'read_pages': sorted(value['intervals']),
                           'complete': value['page_count'] is not None and len(complete) == value['page_count'], 'index_coverage': 'not_published_by_reading'})
        return json.loads(json.dumps(values))

    def record(self, key, report):
        fingerprint = report['source_sha256']
        previous = self.documents.get(key)
        parser = report.get('parser_identity')
        if previous is None or previous['source_sha256'] != fingerprint or previous.get('parser_identity') != parser:
            previous = {'source_sha256': fingerprint, 'parser': parser['implementation'] if parser else 'pdfium-content-order-v1', 'page_count': report['page_count'],
                        'parser_identity': parser, 'intervals': {}, 'totals': {}, 'empty_pages': [], 'failed_pages': [], 'indexed_ranges': []}
        for page in report['pages']:
            number = page['page'] + 1
            end = max((run['char_index'] + run['char_count'] for run in page['runs']), default=page['char_offset'])
            if not page['text_truncated']:
                end = page['total_chars']
            previous['intervals'].setdefault(number, []).append([page['char_offset'], end])
            merged = []
            for start, stop in sorted(previous['intervals'][number]):
                if merged and start <= merged[-1][1]:
                    merged[-1][1] = max(stop, merged[-1][1])
                else:
                    merged.append([start, stop])
            previous['intervals'][number] = merged
            previous['totals'][number] = page['total_chars']
            if page['text_layer_empty'] and number not in previous['empty_pages']:
                previous['empty_pages'].append(number)
        if len(self.documents) >= 8 and key not in self.documents:
            self.documents.pop(next(iter(self.documents)))
        self.documents[key] = previous
        complete = sorted(number for number, intervals in previous['intervals'].items()
                          if intervals == [[0, previous['totals'][number]]])
        previous['page_count'] = report['page_count']
        reread = {row['page'] + 1 for row in report['pages'] if row['char_offset'] == 0 and not row['text_truncated']}
        previous['failed_pages'] = [number for number in previous['failed_pages'] if number not in reread]
        window = [{'page': row['page'] + 1, 'char_offset': row['char_offset'], 'total_chars': row['total_chars'],
                   'partial': row['text_truncated']} for row in report['pages']]
        return json.loads(json.dumps({**previous, 'complete_pages': complete, 'read_pages': sorted(previous['intervals']),
                'complete': len(complete) == report['page_count'], 'window': window,
                'range_id': identifier(fingerprint, str(parser), str(window)),
                'next_page': report['next_page'], 'next_char_offset': report['next_char_offset'],
                'index_coverage': 'not_published_by_reading', 'text_quality': 'content_order_requires_review',
                'numeric_reliability': 'not_independently_verified', 'multicolumn_reading_order': 'not_inferred'}))
