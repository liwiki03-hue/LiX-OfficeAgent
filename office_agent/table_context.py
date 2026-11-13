import re

ALIASES = {'sample': {'sample', 'specimen', '样本', '对象'}, 'time': {'time', '时间', '日期'},
           'conditions': {'conditions', 'condition', '条件', '实验条件'}}
NOTE = re.compile(r'^\s*(?:注|note)?\s*\[([A-Za-z0-9]{1,8})\]\s*[:：]\s*(.+)$', re.I)
REFERENCE = re.compile(r'\[([A-Za-z0-9]{1,8})\]')


def normalized(value):
    return ' '.join(str(value).split()).casefold()


def field_name(value):
    return next((key for key, names in ALIASES.items() if normalized(value) in names), None)


def notes(rows, following, location):
    result = {}
    for index, row in enumerate(rows):
        cells = [str(value) for value in row if value is not None and str(value).strip()]
        if len(cells) == 1:
            match = NOTE.fullmatch(cells[0])
            if match:
                result.setdefault(match[1], []).append({'raw': cells[0], 'text': match[2], 'location': f'{location}/row:{index}'})
    for index, text in following:
        match = NOTE.fullmatch(text)
        if match:
            result.setdefault(match[1], []).append({'raw': text, 'text': match[2], 'location': index})
    return result


def context_for(rows, ri, ci, definitions, location, native_unsupported=False, header_rows=1):
    headers = [' / '.join(str(rows[level][column] or '') for level in range(header_rows)) for column in range(len(rows[0]))] if rows else []
    raw = '' if rows[ri][ci] is None else str(rows[ri][ci])
    references = sorted(set(REFERENCE.findall(raw) + REFERENCE.findall(headers[ci] if ci < len(headers) else '')))
    fields, sources, flags = {}, [], []
    def add(key, value, where):
        if not value.strip():
            return
        if key in fields and normalized(fields[key]) != normalized(value):
            flags.append('conflicting_' + key)
        else:
            fields[key] = value.strip()
        sources.append({'field': key, 'raw': value, 'location': where})
    if ri >= header_rows:
        for column, header in enumerate(headers):
            keys = {field_name(rows[level][column]) for level in range(header_rows)} - {None}
            if len(keys) > 1:
                flags.append('ambiguous_context_header')
            elif keys and column < len(rows[ri]):
                add(next(iter(keys)), str(rows[ri][column] or ''), f'{location}/row:{ri}/column:{column}')
    resolved = []
    for reference in references:
        candidates = definitions.get(reference, [])
        if len(candidates) != 1:
            flags.append('missing_footnote' if not candidates else 'ambiguous_footnote')
            continue
        note = {'marker': reference, **candidates[0]}
        resolved.append(note)
        for part in re.split('[;；]', note['text']):
            pair = re.fullmatch(r'\s*([^=:：]+)\s*[=:：]\s*(.+?)\s*', part)
            key = field_name(pair[1]) if pair else None
            if key:
                add(key, pair[2], note['location'])
            else:
                flags.append('unstructured_footnote_requires_review')
    if native_unsupported:
        flags.append('native_footnote_requires_review')
    return {'sample': fields.get('sample'), 'time': fields.get('time'), 'conditions': fields.get('conditions'),
            'context_sources': sources, 'context_flags': sorted(set(flags)),
            'footnote_references': references, 'footnote_records': resolved,
            'footnotes': 'explicit_notes_resolved_requires_review' if resolved and not flags else
                         'unresolved_requires_review' if flags else 'no_explicit_marker'}, REFERENCE.sub('', raw).strip()


def compare_context(row, base, request):
    issues = []
    for key in ALIASES:
        actual, claimed = row.get(key), getattr(request, key)
        if claimed and not actual:
            issues.append({'field': key, 'status': 'missing_source_context'})
        elif claimed and normalized(claimed) != normalized(actual):
            issues.append({'field': key, 'status': 'condition_mismatch'})
        if base:
            previous = base.get(key)
            if actual and previous and normalized(actual) != normalized(previous):
                issues.append({'field': key, 'status': 'condition_mismatch'})
            elif bool(actual) != bool(previous):
                issues.append({'field': key, 'status': 'missing_source_context'})
    return issues
