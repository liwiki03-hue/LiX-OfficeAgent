from __future__ import annotations

from contextlib import closing
import hashlib
import json
import math
from pathlib import Path
import sys
import tempfile

from .fidelity import _run, _temporary_root
from .rag_sources import MAX_SOURCE_BYTES, SourceError


def parser_identity():
    import pypdfium2 as pdfium
    return {'implementation': 'pdfium-content-order-v2', 'wrapper': str(pdfium.PYPDFIUM_INFO.version),
            'engine': str(pdfium.PDFIUM_INFO.version), 'source_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
            'run_characters': 80, 'page_budget': 5, 'character_budget': 20000}


def read_pdf_payload(path, *, page=0, limit=3, char_offset=0, max_chars=4000, layout=True):
    import pypdfium2 as pdfium

    for value, low, high in ((page, 0, 1999), (limit, 1, 5), (char_offset, 0, 2_000_000), (max_chars, 200, 20000)):
        if type(value) is not int or not low <= value <= high:
            raise SourceError('invalid')
    if type(layout) is not bool:
        raise SourceError('invalid')
    try:
        with pdfium.PdfDocument(str(path)) as document:
            total = len(document)
            if not 1 <= total <= 2000 or page >= total:
                raise SourceError('capacity' if total > 2000 else 'invalid')
            results, budget, cursor, next_char = [], max_chars, page, 0
            while cursor < total and len(results) < limit and budget:
                with closing(document[cursor]) as current:
                    size = current.get_size()
                    if not all(math.isfinite(value) and 0 < value < 1_000_000 for value in size):
                        raise SourceError('invalid')
                    with closing(current.get_textpage()) as textpage:
                        count = textpage.count_chars()
                        start = char_offset if cursor == page else 0
                        if count > 2_000_000 or start > count:
                            raise SourceError('capacity' if count > 2_000_000 else 'invalid')
                        take = min(count - start, budget)
                        text = textpage.get_text_range(start, take, errors='strict') if take else ''
                        if len(text) > budget:
                            raise SourceError('capacity')
                        runs = []
                        if layout:
                            for chunk_start in range(start, start + take, 80):
                                chunk_count = min(80, start + take - chunk_start)
                                boxes = [textpage.get_charbox(index) for index in range(chunk_start, chunk_start + chunk_count)]
                                if any(not all(math.isfinite(value) for value in box) for box in boxes):
                                    raise SourceError('invalid')
                                runs.append({'char_index': chunk_start, 'char_count': chunk_count,
                                             'text': textpage.get_text_range(chunk_start, chunk_count, errors='strict'),
                                             'bbox': [min(box[0] for box in boxes), min(box[1] for box in boxes),
                                                      max(box[2] for box in boxes), max(box[3] for box in boxes)]})
                        partial = start + take < count
                        results.append({'page': cursor, 'size_points': list(size), 'char_offset': start,
                                        'total_chars': count, 'text': text, 'runs': runs,
                                        'text_truncated': partial, 'text_layer_empty': count == 0})
                        budget -= take
                        if partial:
                            next_char = start + take
                            break
                cursor += 1
            return {'page_count': total, 'pages': results, 'truncated': cursor < total,
                    'parser_identity': parser_identity(),
                    'next_page': cursor if cursor < total else None,
                    'next_char_offset': next_char if cursor < total else None,
                    'coordinate_system': 'PDF points; origin bottom-left; 80-character run bounds',
                    'read_only': True, 'ocr_performed': False,
                    'limitations': ['Text extraction order is PDF content order, not inferred reading order',
                                    'Empty text layer may contain scanned content; OCR is not performed']}
    except SourceError:
        raise
    except Exception:
        raise SourceError('unsupported') from None


def read_pdf(path, *, page=0, limit=3, char_offset=0, max_chars=4000, layout=True):
    path = Path(path)
    if path.suffix.lower() != '.pdf' or path.is_symlink() or path.stat().st_size > MAX_SOURCE_BYTES:
        raise SourceError('unsupported')
    with path.open('rb') as stream:
        data = stream.read(MAX_SOURCE_BYTES + 1)
    if len(data) > MAX_SOURCE_BYTES or not data.startswith(b'%PDF-'):
        raise SourceError('invalid')
    options = dict(page=page, limit=limit, char_offset=char_offset, max_chars=max_chars, layout=layout)
    for value, low, high in ((page, 0, 1999), (limit, 1, 5), (char_offset, 0, 2_000_000), (max_chars, 200, 20000)):
        if type(value) is not int or not low <= value <= high:
            raise SourceError('invalid')
    if type(layout) is not bool:
        raise SourceError('invalid')
    fingerprint = hashlib.sha256(data).hexdigest()
    processes = []
    with tempfile.TemporaryDirectory(prefix='orp-', dir=_temporary_root()) as raw:
        scratch = Path(raw)
        copy, config, target = scratch / 'input.pdf', scratch / 'options.json', scratch / 'result.json'
        copy.write_bytes(data)
        config.write_text(json.dumps(options), encoding='utf-8')
        _run((sys.executable, str(Path(__file__).with_name('read_worker.py')), 'pdf', str(copy), str(config), str(target)), 30000, processes)
        if not target.is_file() or target.stat().st_size > 500000:
            raise SourceError('capacity')
        report = json.loads(target.read_text(encoding='utf-8'))
    if report.get('error'):
        raise SourceError(report['error'])
    with path.open('rb') as stream:
        current = stream.read(MAX_SOURCE_BYTES + 1)
    if hashlib.sha256(current).hexdigest() != fingerprint:
        raise SourceError('source_changed')
    return {**report, 'source_sha256': fingerprint, 'processes': processes}
