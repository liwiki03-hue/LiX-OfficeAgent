import base64
import hashlib
from io import BytesIO
import json
import math
from pathlib import Path
import sys
import tempfile

from .fidelity import _run, _temporary_root
from .rag_sources import SourceError, SourceExtractor


def validate_proofs(request):
    from .writes import GuardError, Workspace
    proofs = request.intent.get('pdf_numeric_reviews', [])
    if type(proofs) is not list or len(proofs) > 8:
        raise GuardError('professional_pdf_stale')
    if not proofs:
        return
    source = SourceExtractor(Workspace(Path(request.path).parent))
    for proof in proofs:
        try:
            _, identity, fingerprint = source._snapshot(proof['name'])
            if fingerprint != proof['source_sha256'] or json.loads(json.dumps(identity)) != proof['identity']:
                raise ValueError()
        except Exception:
            raise GuardError('professional_pdf_source_changed') from None


def validate_grid(grid):
    if type(grid) is not dict or not {'page', 'rows', 'columns', 'units'} <= set(grid) or not set(grid) <= {'page', 'rows', 'columns', 'units', 'note_box', 'header_rows'}:
        raise SourceError('invalid')
    if type(grid['page']) is not int or not 1 <= grid['page'] <= 2000:
        raise SourceError('invalid')
    for key, maximum in [('rows', 22), ('columns', 13)]:
        values = grid[key]
        if (type(values) is not list or not 3 <= len(values) <= maximum or
                any(type(value) not in (int, float) or not math.isfinite(value) or not 0 <= value <= 1 for value in values) or
                values != sorted(set(values))):
            raise SourceError('invalid')
    if (len(grid['rows']) - 1) * (len(grid['columns']) - 1) > 100:
        raise SourceError('capacity')
    header_rows = grid.get('header_rows', 1)
    if type(header_rows) is not int or header_rows not in (1, 2) or len(grid['rows']) < header_rows + 2:
        raise SourceError('invalid')
    note = grid.get('note_box')
    if note is not None:
        if (type(note) is not list or len(note) != 4 or any(type(value) not in (int, float) or not math.isfinite(value) or not 0 <= value <= 1 for value in note) or
                not note[0] < note[2] or not note[1] < note[3] or
                (note[2] > grid['columns'][0] and note[0] < grid['columns'][-1] and note[3] > grid['rows'][0] and note[1] < grid['rows'][-1])):
            raise SourceError('invalid')
    pairs = grid['units']
    if (type(pairs) is not list or len(pairs) > 12 or any(type(pair) is not list or len(pair) != 2 or
            any(type(value) is not int or not 0 <= value < len(grid['columns']) - 1 for value in pair) for pair in pairs) or
            len({pair[0] for pair in pairs}) != len(pairs)):
        raise SourceError('invalid')
    return grid


def cell_text(chars, layer):
    ordered = sorted(chars, key=lambda item: item[1])
    if ordered and max(item[0] for item in ordered) - min(item[0] for item in ordered) > 2:
        raise SourceError('unsupported')
    parts = []
    for offset, item in enumerate(ordered):
        if offset:
            previous = ordered[offset - 1][4]
            box = item[4]
            overlap = min(previous[2], box[2]) - max(previous[0], box[0])
            if overlap > .7 * min(previous[2] - previous[0], box[2] - box[0]):
                raise SourceError('unsupported')
        if offset and 1 < item[3] - ordered[offset - 1][3] < 100:
            gap = layer.get_text_range(ordered[offset - 1][3] + 1, item[3] - ordered[offset - 1][3] - 1, errors='strict')
            if gap.isspace() and not any(char in gap for char in '\r\n'):
                parts.append(gap)
        parts.append(item[2])
    return ''.join(parts), [item[3] for item in ordered]


def payload(path, page, grid=None):
    import pypdfium2 as pdfium
    if type(page) is not int or not 1 <= page <= 2000:
        raise SourceError('invalid')
    if grid is not None and validate_grid(grid)['page'] != page:
        raise SourceError('invalid')
    with pdfium.PdfDocument(path) as document:
        if not 1 <= len(document) <= 2000 or page > len(document):
            raise SourceError('invalid')
        current = document[page - 1]
        try:
            width, height = current.get_size()
            if any(not math.isfinite(value) or not 0 < value < 100000 for value in (width, height)):
                raise SourceError('invalid')
            if current.get_rotation() != 0:
                raise SourceError('unsupported')
            if any(abs(actual - expected) > .01 for actual, expected in zip(current.get_bbox(), (0, 0, width, height))):
                raise SourceError('unsupported')
            bitmap = current.render(scale=min(600 / width, 900 / height))
            try:
                stream = BytesIO()
                bitmap.to_pil().convert('RGB').save(stream, format='JPEG', quality=55)
                preview = stream.getvalue()
                if len(preview) > 65000:
                    raise SourceError('capacity')
            finally:
                bitmap.close()
            report = {'page': page, 'total_pages': len(document), 'size_points': [width, height],
                      'parser_identity': {'wrapper': str(pdfium.PYPDFIUM_INFO.version), 'engine': str(pdfium.PDFIUM_INFO.version),
                                          'adapter_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest()},
                      'image': 'data:image/jpeg;base64,' + base64.b64encode(preview).decode(),
                      'rows': [], 'cells': [], 'coordinate_system': 'normalized top-left; boundaries 0..1',
                      'ocr_performed': False, 'automatic_table_detection': False}
            if grid is None:
                return report
            layer = current.get_textpage()
            try:
                count = layer.count_chars()
                if count == 0:
                    raise SourceError('pdf_text_missing')
                if count > 20000:
                    raise SourceError('capacity')
                rows, columns = grid['rows'], grid['columns']
                note_box, note_chars = grid.get('note_box'), []
                cells = [[[] for _ in columns[1:]] for _ in rows[1:]]
                for index in range(count):
                    text = layer.get_text_range(index, 1, errors='strict')
                    if not text.strip():
                        continue
                    left, bottom, right, top = layer.get_charbox(index)
                    if not all(math.isfinite(value) for value in (left, bottom, right, top)):
                        raise SourceError('invalid')
                    box = [left / width, 1 - top / height, right / width, 1 - bottom / height]
                    x, y = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
                    if note_box and box[2] > note_box[0] and box[0] < note_box[2] and box[3] > note_box[1] and box[1] < note_box[3]:
                        if box[0] < note_box[0] or box[2] > note_box[2] or box[1] < note_box[1] or box[3] > note_box[3]:
                            raise SourceError('unsupported')
                        note_chars.append((round(y * height / 3), x, text, index, box))
                    intersects = box[2] > columns[0] and box[0] < columns[-1] and box[3] > rows[0] and box[1] < rows[-1]
                    ri = next((i for i in range(len(rows) - 1) if rows[i] <= y < rows[i + 1]), None)
                    ci = next((i for i in range(len(columns) - 1) if columns[i] <= x < columns[i + 1]), None)
                    if ri is None or ci is None:
                        if intersects:
                            raise SourceError('unsupported')
                        continue
                    if box[0] < columns[ci] or box[2] > columns[ci + 1] or box[1] < rows[ri] or box[3] > rows[ri + 1]:
                        raise SourceError('unsupported')
                    cells[ri][ci].append((round(y * height / 3), x, text, index, box))
                for ri, row in enumerate(cells):
                    values = []
                    for ci, chars in enumerate(row):
                        text, indices = cell_text(chars, layer)
                        values.append(text)
                        report['cells'].append({'row': ri, 'column': ci, 'raw': text,
                            'box': [columns[ci], rows[ri], columns[ci + 1], rows[ri + 1]],
                            'char_indices': indices})
                    report['rows'].append(values)
                header_rows = grid.get('header_rows', 1)
                if sum(len(value) for row in report['rows'] for value in row) > 8000 or not all(value for row in report['rows'][:header_rows] for value in row):
                    raise SourceError('unsupported')
                report['header_rows'] = header_rows
                report['note_text'], report['note_char_indices'] = cell_text(note_chars, layer)
                report['note_box'] = note_box
                report['context_outside_selection_included'] = False
                if len(report['note_text']) > 2000:
                    raise SourceError('capacity')
                return report
            finally:
                layer.close()
        finally:
            current.close()


def read_table(workspace, name, page, grid=None, proposal_box=None):
    source = SourceExtractor(workspace)
    if type(page) is not int or not 1 <= page <= 2000:
        raise SourceError('invalid')
    if type(name) is not str or not name.lower().endswith('.pdf') or any(char in name for char in '/\\:'):
        raise SourceError('invalid')
    data, identity, fingerprint = source._snapshot(name)
    if not data.startswith(b'%PDF-'):
        raise SourceError('invalid')
    if grid is not None:
        validate_grid(grid)
    if proposal_box is not None:
        from .pdf_layout import validate_box
        validate_box(proposal_box)
        if grid is not None:
            raise SourceError('invalid')
    jobs = []
    with tempfile.TemporaryDirectory(prefix='opt-', dir=_temporary_root()) as raw:
        folder = Path(raw)
        (folder / 'input.pdf').write_bytes(data)
        options = {'page': page, 'grid': grid} if proposal_box is None else {'page': page, 'proposal_box': proposal_box}
        (folder / 'options.json').write_text(json.dumps(options), encoding='utf8')
        target = folder / 'result.json'
        _run((sys.executable, str(Path(__file__).with_name('pdf_table_worker.py')), str(folder)), 30000, jobs)
        if not target.is_file() or target.stat().st_size > 180000:
            raise SourceError('capacity')
        result = json.loads(target.read_text(encoding='utf8'))
    if result.get('error'):
        raise SourceError(result['error'])
    if source._snapshot(name)[1:] != (identity, fingerprint):
        raise SourceError('source_changed')
    return {**result, 'source_sha256': fingerprint, 'identity': identity, 'processes': jobs,
            'parser_stamp': hashlib.sha256(json.dumps(result['parser_identity'], sort_keys=True).encode()).hexdigest()}
