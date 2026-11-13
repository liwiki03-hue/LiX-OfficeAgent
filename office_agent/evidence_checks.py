from decimal import Decimal, InvalidOperation, localcontext, ROUND_HALF_EVEN
from io import BytesIO
import hashlib
import re

from pydantic import Field

from .writes import Strict, GuardError
from .rag_sources import SourceExtractor
from .rag import SECRET
from .table_context import notes, context_for, compare_context, REFERENCE

UNITS = {'ms': ('time', Decimal('.001')), 's': ('time', Decimal(1)), '秒': ('time', Decimal(1)),
         '毫秒': ('time', Decimal('.001')), 'kg': ('mass', Decimal(1)), 'g': ('mass', Decimal('.001')),
         'mg': ('mass', Decimal('.000001')), 'm': ('length', Decimal(1)), 'cm': ('length', Decimal('.01')),
         'mm': ('length', Decimal('.001')), '%': ('ratio', Decimal('.01')), '百分比': ('ratio', Decimal('.01')),
         '百分点': ('percentage_point', Decimal('.01')), 'pp': ('percentage_point', Decimal('.01')),
         '1': ('dimensionless', Decimal(1))}
NUMBER = r'[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?'
VALUE = re.compile(r'^\s*(' + NUMBER + r')\s*([^\d\s]+)?\s*$')


def _decimal(value):
    if type(value) is not str or len(value) > 50 or re.fullmatch(NUMBER, value) is None:
        raise GuardError('professional_numeric_invalid')
    try:
        result = Decimal(value)
        if not result.is_finite() or abs(result.adjusted()) > 100:
            raise ValueError()
        return result
    except (InvalidOperation, ValueError):
        raise GuardError('professional_numeric_invalid') from None


class NumericClaim(Strict):
    observation_id: str = Field(pattern=r'^obs:[a-f0-9]{64}$')
    value: str = Field(max_length=50)
    unit: str = Field(min_length=1, max_length=30)
    baseline_id: str | None = None
    operation: str = Field(default='value', max_length=30)
    sample: str = Field(default='', max_length=200)
    conditions: str = Field(default='', max_length=300)
    time: str = Field(default='', max_length=100)
    statistic: str = Field(default='value', max_length=20)


def table_observations(workspace, name):
    if type(name) is not str or any(c in name for c in '/\\:'):
        raise GuardError('professional_evidence_invalid')
    extractor = SourceExtractor(workspace)
    source = extractor.extract(name)
    data, _, version = extractor._snapshot(name)
    if version != source['fingerprint']:
        raise GuardError('professional_source_changed')
    tables = []
    if name.lower().endswith('.docx'):
        from docx import Document
        from .professional_edits import W
        document = Document(BytesIO(data))
        for ti, table in enumerate(document.tables):
            merged = any(n.tag in {W + 'gridSpan', W + 'vMerge'} for n in table._tbl.iter())
            following = []
            for node in table._tbl.itersiblings():
                if node.tag != W + 'p':
                    break
                text = ''.join(n.text or '' for n in node.iter(W + 't'))
                if not text.strip():
                    continue
                from .table_context import NOTE
                if not NOTE.fullmatch(text) or len(following) >= 8:
                    break
                following.append((f'table:{ti}/following-note:{len(following)}', text))
            unsupported = any(n.tag == W + 'footnoteReference' for n in table._tbl.iter())
            tables.append((f'table:{ti}', [[cell.text for cell in row.cells] for row in table.rows], merged, following, unsupported))
    elif name.lower().endswith('.xlsx'):
        from openpyxl import load_workbook
        book = load_workbook(BytesIO(data), data_only=False, read_only=False)
        try:
            for si, sheet in enumerate(book.worksheets):
                if sheet.max_row > 50 or sheet.max_column > 20:
                    raise GuardError('professional_numeric_capacity')
                tables.append((f'sheet:{si}', [[cell.value for cell in row] for row in sheet.iter_rows()], bool(sheet.merged_cells.ranges), [], False))
        finally:
            book.close()
    else:
        raise GuardError('professional_numeric_table_format')
    records = observations_from_tables(name, version, tables)
    if extractor._snapshot(name)[2] != version:
        raise GuardError('professional_source_changed')
    return records


def observations_from_tables(name, version, tables, *, header_rows=1):
    if type(header_rows) is not int or header_rows not in (1, 2):
        raise GuardError('professional_numeric_invalid')
    if len(tables) > 8 or sum(sum(len(row) for row in table[1]) for table in tables) > 400:
        raise GuardError('professional_numeric_capacity')
    records = []
    for location, rows, merged, following, unsupported in tables:
        if header_rows == 2 and (len(rows) < 3 or any(len(row) != len(rows[0]) for row in rows) or not all(str(v or '').strip() for row in rows[:2] for v in row)):
            raise GuardError('professional_numeric_invalid')
        definitions = notes(rows, following, location)
        if any(len(note['raw']) > 2000 or SECRET.search(note['raw']) for values in definitions.values() for note in values):
            raise GuardError('professional_sensitive_text')
        headers = [' / '.join(str(rows[level][column]) if rows[level][column] is not None else '' for level in range(header_rows)) for column in range(len(rows[0]))] if rows else []
        for ri, row in enumerate(rows):
            for ci, value in enumerate(row):
                raw = '' if value is None else str(value)
                if len(raw) > 2000 or SECRET.search(raw):
                    raise GuardError('professional_sensitive_text')
                where = f'{location}/row:{ri}/column:{ci}'
                context, numeric_text = context_for(rows, ri, ci, definitions, location, unsupported, header_rows)
                uncertainty = re.fullmatch(r'\s*(' + NUMBER + r')\s*±\s*(' + NUMBER + r')\s*([^\d\s]+)?\s*', numeric_text)
                interval = re.fullmatch(r'\s*(' + NUMBER + r')\s*(?:to|至|–|~)\s*(' + NUMBER + r')\s*([^\d\s]+)?\s*', numeric_text)
                match = VALUE.fullmatch(numeric_text) if not uncertainty and not interval else uncertainty
                header = headers[ci] if ci < len(headers) else ''
                header_units = [unit for pair in re.findall(r'\(([^()]+)\)|（([^（）]+)）', REFERENCE.sub('', header)) for unit in pair if unit in UNITS]
                unit = uncertainty[3] if uncertainty else interval[3] if interval else match[2] if match else None
                if (match or interval) and not unit:
                    units = [unit for pair in re.findall(r'\(([^()]+)\)|（([^（）]+)）', REFERENCE.sub('', header)) for unit in pair if unit in UNITS]
                    if header_rows == 2:
                        units = list(dict.fromkeys(units))
                    unit = units[0] if len(units) == 1 else None
                numeric = match[1] if match else None
                if header_rows == 2:
                    if len(set(header_units)) > 1 or (unit and header_units and unit != header_units[0]):
                        context['context_flags'].append('ambiguous_header_unit')
                    if re.search(r'均值|平均值|\bmean\b', header, re.I) and re.search(r'不确定度|uncertainty', header, re.I):
                        context['context_flags'].append('ambiguous_header_statistic')
                    if ri < header_rows:
                        numeric = None
                if numeric:
                    _decimal(numeric)
                if uncertainty and _decimal(uncertainty[2]) < 0:
                    context['context_flags'].append('invalid_uncertainty')
                if interval and _decimal(interval[1]) > _decimal(interval[2]):
                    context['context_flags'].append('invalid_range')
                records.append({'id': 'obs:' + hashlib.sha256((version + where + raw).encode()).hexdigest(),
                                'source': name, 'source_sha256': version, 'location': where, 'raw': raw,
                                'header': header, 'row_label': str(row[0] or '') if row else '',
                                'value': numeric, 'unit': unit, 'range': [interval[1], interval[2]] if interval else None,
                                'uncertainty': uncertainty[2] if uncertainty else None,
                                'baseline': None, 'sample': None, 'time': None, 'conditions': None,
                                'missing': value is None or raw.strip() in {'', 'NA', 'N/A', '—', '-'},
                                'origin': 'direct_source', 'extraction': 'native_table_cell',
                                'uncertain_structure': merged, **context,
                                'statistic': 'uncertainty' if re.search(r'不确定度|uncertainty', header, re.I) else
                                             'mean' if re.search(r'均值|平均值|\bmean\b', header, re.I) else 'value'})
                if header_rows == 2:
                    records[-1].update(header_rows=2, is_header=ri < 2,
                        header_sources=[{'raw': str(rows[level][ci]), 'location': f'{location}/row:{level}/column:{ci}'} for level in range(2)])
    return records


def check_numeric(observations, claim):
    request = NumericClaim.model_validate(claim)
    if SECRET.search(str(claim)):
        raise GuardError('professional_sensitive_text')
    by_id = {r['id']: r for r in observations}
    row = by_id.get(request.observation_id)
    base = by_id.get(request.baseline_id) if request.baseline_id else None
    result = {'claim': request.model_dump(), 'evidence': row, 'baseline': base,
              'status': 'insufficient', 'numeric_status': 'insufficient', 'user_reviewed': False,
              'formula': None, 'inputs': [], 'rounding': 'Decimal precision=28 ROUND_HALF_EVEN; no display rounding before comparison',
              'suggestion': '补齐单位、样本、时间和条件，再人工核对原表；数值吻合不证明整段论点成立。'}
    if not row or row['missing'] or not row['value'] or row['uncertain_structure'] or row['unit'] not in UNITS:
        return result
    if request.statistic not in {'value', 'mean', 'uncertainty'}:
        raise GuardError('professional_numeric_invalid')
    result['context_issues'] = compare_context(row, base, request)
    result['footnote_records'] = row.get('footnote_records', [])
    if row.get('context_flags') or (base and base.get('context_flags')):
        return {**result, 'suggestion': '脚注缺失、重复、条件冲突或含未结构化限制；请核对所列原文，不能自动比较。'}
    if result['context_issues']:
        return {**result, 'status': 'condition_mismatch' if any(issue['status'] == 'condition_mismatch' for issue in result['context_issues']) else 'insufficient',
                'suggestion': '正文或基线的样本、时间、条件与所选证据不一致或缺失；请核对来源位置。'}
    if request.statistic != row.get('statistic', 'value') or (base and row.get('statistic', 'value') != base.get('statistic', 'value')):
        return {**result, 'status': 'contradict', 'numeric_status': 'statistic_mismatch', 'suggestion': '数值的统计口径不一致；均值、不确定度与普通测量值不能互换。'}
    if request.operation not in {'value', 'relative_increase', 'relative_decrease', 'percentage_points'}:
        raise GuardError('professional_numeric_invalid')
    claimed = _decimal(request.value)
    if request.unit not in UNITS:
        result['suggestion'] = '单位未知或有歧义；请人工核对，不自动转换。'
        return result
    dimension, factor = UNITS[row['unit']]
    wanted, target_factor = UNITS[request.unit]
    with localcontext() as context:
        context.prec = 28
        context.rounding = ROUND_HALF_EVEN
        expected = _decimal(row['value']) * factor
        result['inputs'] = [row['id']]
        if request.operation != 'value':
            if not base or base['missing'] or not base['value'] or base['uncertain_structure'] or base['unit'] not in UNITS:
                return result
            if UNITS[base['unit']][0] != dimension:
                return {**result, 'status': 'condition_mismatch', 'numeric_status': 'unit_mismatch'}
            if base['source_sha256'] != row['source_sha256']:
                return {**result, 'suggestion': '跨来源的样本、时间和条件尚未对齐；证据不足，不能自动比较。'}
            if base['row_label'] != row['row_label']:
                return {**result, 'status': 'condition_mismatch', 'suggestion': '两项记录的对象或来源不同；先人工对齐样本、时间及实验条件。'}
            previous = _decimal(base['value']) * UNITS[base['unit']][1]
            result['inputs'].append(base['id'])
            if request.operation == 'percentage_points':
                if dimension != 'ratio':
                    return result
                expected -= previous
                dimension = 'percentage_point'
                result['formula'] = '(current_fraction - baseline_fraction)'
            else:
                if previous == 0:
                    return {**result, 'suggestion': '基线为0，不能计算相对变化。'}
                expected = (expected - previous) / previous
                if request.operation == 'relative_decrease':
                    expected = -expected
                dimension = 'ratio'
                result['formula'] = '(current - baseline) / baseline' if request.operation == 'relative_increase' else '(baseline - current) / baseline'
        if wanted != dimension:
            return {**result, 'status': 'contradict', 'numeric_status': 'unit_mismatch', 'suggestion': '核对量纲；百分比变化与百分点不能互换。'}
        actual = claimed * target_factor
        result.update(numeric_status='match' if actual == expected else 'mismatch', expected_value=str(expected / target_factor))
    result['status'] = 'needs_manual_judgment' if result['numeric_status'] == 'match' else 'contradict'
    result['suggestion'] = '数值与单位吻合；样本、时间、适用范围及因果论断仍需人工核对。' if result['numeric_status'] == 'match' else '正文数值或变化方向与所选证据不符；请核对基线、原值及单位。'
    return result
