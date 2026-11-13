import hashlib
import math
from pathlib import Path
from statistics import median
import re

from .pdf_table import payload, validate_grid
from .rag_sources import SourceError


def validate_box(box):
    if (type(box) is not list or len(box) != 4 or any(type(v) not in (float, int) or not math.isfinite(v) or not 0 <= v <= 1 for v in box)
            or box[0] >= box[2] or box[1] >= box[3]):
        raise SourceError('invalid')
    return box


def unit_review(parsed):
    from .evidence_checks import UNITS, NUMBER
    rows = parsed['rows']
    sources = [column for column, header in enumerate(rows[0]) if header.casefold() in {'unit', 'units', '单位'}]
    result, candidates = [], []
    for column, header in enumerate(rows[0]):
        if column in sources or not all(re.fullmatch(NUMBER, row[column]) for row in rows[1:]):
            continue
        explicit = {value for pair in re.findall(r'\(([^()]+)\)|（([^（）]+)）', header) for value in pair if value in UNITS}
        options = []
        for source in sources:
            values = list(dict.fromkeys(row[source] for row in rows[1:]))
            valid = all(value in UNITS for value in values)
            status = 'unsupported_or_missing_unit' if not valid else 'ambiguous_unit_scope'
            if valid and explicit:
                status = 'header_and_row_unit_agreement' if len(explicit) == 1 and set(values) == explicit else 'header_unit_conflict'
            options.append({'unit_column': source + 1, 'unit_values': values, 'status': status})
        matching = [value for value in options if value['status'] == 'header_and_row_unit_agreement']
        if len(matching) == 1:
            candidates.append({'value_column': column + 1, 'unit_column': matching[0]['unit_column'],
                'confirmed': False, 'kind': 'header_and_row_unit_agreement'})
        result.append({'value_column': column + 1, 'header': header, 'explicit_header_units': sorted(explicit),
            'header_cell': next(cell for cell in parsed['cells'] if cell['row'] == 0 and cell['column'] == column),
            'options': options, 'status': 'multiple_unit_sources_require_review' if len(matching) > 1 else
                'header_and_row_unit_agreement' if len(matching) == 1 else 'no_confirmed_unit_scope', 'confirmed': False})
    return candidates, result


def propose(path, page, box):
    import pypdfium2 as pdfium
    validate_box(box)
    report = payload(path, page)
    report['parser_identity']['layout_candidate'] = {'kind': 'native_glyph_alignment-v2-unit-scope',
        'source_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}
    report.update(proposal_box=box, proposal=None, candidate_only=True, layout_confirmed=False,
                  support_assessed=False, ocr_performed=False, automatic_table_detection=True,
                  analysis_scope='explicit_rectangle_only')
    uncertain = ['header_depth_requires_review', 'unit_and_footnote_association_requires_review', 'outside_rectangle_not_read']
    with pdfium.PdfDocument(path) as document:
        current = document[page - 1]
        try:
            width, height = current.get_size()
            layer = current.get_textpage()
            try:
                if layer.count_chars() > 20000:
                    raise SourceError('capacity')
                chars = []
                for index in range(layer.count_chars()):
                    left, bottom, right, top = layer.get_charbox(index)
                    bounds = [left / width, 1 - top / height, right / width, 1 - bottom / height]
                    if not all(math.isfinite(v) for v in bounds):
                        raise SourceError('invalid')
                    intersects = bounds[2] > box[0] and bounds[0] < box[2] and bounds[3] > box[1] and bounds[1] < box[3]
                    if not intersects:
                        continue
                    value = layer.get_text_range(index, 1, errors='strict')
                    if not value.strip():
                        continue
                    if bounds[0] < box[0] or bounds[1] < box[1] or bounds[2] > box[2] or bounds[3] > box[3]:
                        report['uncertainties'] = uncertain + ['rectangle_clips_glyphs']
                        return report
                    chars.append({'index': index, 'raw': value, 'box': bounds, 'x': bounds[0], 'y': (bounds[1] + bounds[3]) / 2})
                if not chars:
                    report['uncertainties'] = uncertain + ['no_native_text_in_rectangle']
                    return report
                line_tolerance = median(item['box'][3] - item['box'][1] for item in chars) * .45
                lines = []
                for item in sorted(chars, key=lambda item: (item['y'], item['x'])):
                    if not lines or abs(median(value['y'] for value in lines[-1]) - item['y']) > line_tolerance:
                        lines.append([])
                    lines[-1].append(item)
                if not 2 <= len(lines) <= 21:
                    report['uncertainties'] = uncertain + ['row_alignment_not_simple']
                    return report
                rows = []
                for line in lines:
                    ordered = sorted(line, key=lambda item: item['x'])
                    gap = max(median(item['box'][2] - item['box'][0] for item in ordered) * 2.5, (box[2] - box[0]) * .015)
                    cells = []
                    for item in ordered:
                        if not cells or item['box'][0] - cells[-1][-1]['box'][2] > gap:
                            cells.append([])
                        cells[-1].append(item)
                    rows.append(cells)
                columns = len(rows[0])
                if not 2 <= columns <= 12 or any(len(row) != columns for row in rows) or columns * len(rows) > 100:
                    report['uncertainties'] = uncertain + ['inconsistent_columns_or_possible_merged_cells']
                    return report
                intervals = [(min(item['box'][0] for row in rows for item in row[column]),
                              max(item['box'][2] for row in rows for item in row[column])) for column in range(columns)]
                if any(intervals[i][1] >= intervals[i + 1][0] for i in range(columns - 1)):
                    report['uncertainties'] = uncertain + ['overlapping_column_alignment']
                    return report
                vertical = [box[0]] + [(intervals[i][1] + intervals[i + 1][0]) / 2 for i in range(columns - 1)] + [box[2]]
                line_boxes = [(min(item['box'][1] for item in line), max(item['box'][3] for item in line)) for line in lines]
                if any(line_boxes[i][1] >= line_boxes[i + 1][0] for i in range(len(lines) - 1)):
                    report['uncertainties'] = uncertain + ['overlapping_row_alignment']
                    return report
                horizontal = [box[1]] + [(line_boxes[i][1] + line_boxes[i + 1][0]) / 2 for i in range(len(lines) - 1)] + [box[3]]
                grid = validate_grid({'page': page, 'rows': horizontal, 'columns': vertical, 'units': [], 'header_rows': 1})
            finally:
                layer.close()
        finally:
            current.close()
    parsed = payload(path, page, grid)
    unit_candidates, review = unit_review(parsed)
    footnote_candidates = []
    for cell in parsed['cells']:
        for marker in re.findall(r'\[([A-Za-z0-9]{1,8})\]', cell['raw']):
            footnote_candidates.append({'marker': marker, 'cell_box': cell['box'], 'association': 'unresolved_no_outside_read'})
    report.update(proposal=grid, candidate_cells=parsed['cells'], candidate_rows=parsed['rows'],
                  header_candidates=[1, 2] if len(lines) >= 3 else [1], merge_candidates=[],
                  unit_candidates=unit_candidates, unit_review=review, footnote_candidates=footnote_candidates, uncertainties=uncertain)
    return report
