import re
import unicodedata


VERSION = 'context-page-span-v3'


def _row_neighbors(node, text_for):
    from .professional_edits import W
    cell = node.getparent()
    if cell is None or cell.tag != W + 'tc':
        return None
    row = cell.getparent()
    if row is None or row.tag != W + 'tr':
        return None
    cells = list(row.findall(W + 'tc'))
    if not 2 <= len(cells) <= 20:
        return None
    if any(len(item.findall(W + 'p')) != 1 or item.find(W + 'tbl') is not None
           or item.find('.//' + W + 'gridSpan') is not None or item.find('.//' + W + 'vMerge') is not None
           or item.find('.//' + W + 'hMerge') is not None for item in cells):
        return None
    index = cells.index(cell)
    values = [text_for(item.find(W + 'p')) for item in cells]
    return ' '.join(values[:index])[-64:], ' '.join(values[index + 1:])[:64]


def normalized(value):
    value = re.sub(r'\s+', ' ', unicodedata.normalize('NFC', value)).strip()
    return re.sub(r'(?<=[\u3400-\u9fff\u3000-\u303f，。；：？！（）]) | (?=[\u3400-\u9fff\u3000-\u303f，。；：？！（）])', '', value)


def _locate(pages, context):
    text = normalized(context.get('text', ''))
    prefix = normalized(context.get('prefix', ''))
    suffix = normalized(context.get('suffix', ''))
    needle = normalized(context.get('prefix', '') + context.get('text', '') + context.get('suffix', ''))
    if len(needle) < 4 or (not text and not (prefix and suffix)):
        return {'pages': [], 'status': 'not_located', 'matches': 0}
    chunks = [normalized(page['text']) for page in pages]
    document = ''
    offsets, cursor = [], 0
    for chunk in chunks:
        separator = '' if not document or not chunk or normalized(document[-1] + ' ' + chunk[0]) == document[-1] + chunk[0] else ' '
        document += separator + chunk
        cursor += len(separator)
        offsets.append((cursor, cursor + len(chunk)))
        cursor += len(chunk)
    matches, position = [], 0
    while len(matches) < 3:
        start = document.find(needle, position)
        if start < 0:
            break
        end = start + len(needle)
        span = [i + 1 for i, (a, b) in enumerate(offsets) if a < end and b > start]
        if span and len(span) <= 2 and span[-1] - span[0] + 1 == len(span):
            matches.append(span)
        position = start + 1
    return {'pages': matches[0] if len(matches) == 1 else [],
            'status': 'unique' if len(matches) == 1 else ('ambiguous' if matches else 'not_located'),
            'matches': len(matches)}


def locate(pages, context):
    match = _locate(pages, context)
    local = context.get('paragraph')
    if match['status'] == 'not_located' and type(local) is dict:
        candidate = _locate(pages, {'text': context.get('text', ''), 'prefix': local.get('prefix', ''), 'suffix': local.get('suffix', '')})
        if candidate['status'] == 'unique':
            return {**candidate, 'method': 'host-paragraph-context-fallback'}
    return match


def align(anchors, snapshots):
    result = []
    for anchor in anchors:
        row = {key: anchor[key] for key in ('location', 'reason', 'action')}
        for side, snapshot in zip(('before', 'after'), snapshots):
            context = anchor.get('contexts', {}).get(side)
            if context is not None:
                match = locate(snapshot['pages'], context)
                row[side + '_pages'] = match['pages']
                row[side + '_status'] = match['status']
                row[side + '_matches'] = match['matches']
                if 'method' in match:
                    row[side + '_method'] = match['method']
            else:
                text = normalized(anchor[side])
                pages = [i + 1 for i, page in enumerate(snapshot['pages']) if text and text in normalized(page['text'])]
                row[side + '_pages'] = pages
                row[side + '_status'] = 'unique' if len(pages) == 1 else ('ambiguous' if pages else 'not_located')
        row['policy'] = VERSION if 'contexts' in anchor else 'legacy-page-anchor'
        if 'contexts' in anchor:
            row['anchor_scope'] = 'bounded operation context; structural insertions use the first 200 characters; not the complete affected-page set'
        result.append(row)
    return result


def anchors_for(parts, request, approved_parts=None, approved_starts=None):
    from .range_edits import _inventory
    from .professional_edits import W, _text
    root, entries = _inventory(parts, request.source_sha256)
    paragraphs = list(root.iter(W + 'p'))
    by_id = {entry['locator_id']: entry for entry in entries}
    by_node = {entry['node']: entry for entry in entries}
    def projected_text(node):
        entry = by_node.get(node)
        value = _text(node)
        if entry:
            changes = [item for item in request.operations if item.locator_id == entry['locator_id']
                       and item.action in {'replace', 'insert', 'delete'}]
            for item in sorted(changes, key=lambda item: item.start, reverse=True):
                value = value[:item.start] + item.text + value[item.end:]
        return value
    if approved_parts is not None:
        approved_root, approved_entries = _inventory(approved_parts, request.source_sha256)
        approved_paragraphs = list(approved_root.iter(W + 'p'))
        approved_locations = {entry['location']: entry for entry in approved_entries}
    anchors = []
    for op in request.operations:
        entry = by_id[op.locator_id]
        value = entry['text']
        index = paragraphs.index(entry['node'])
        neighbors = [paragraphs[index - 1] if index else None, paragraphs[index + 1] if index + 1 < len(paragraphs) else None]
        previous, following = [_text(node) if node is not None and node.getparent() is entry['node'].getparent() else '' for node in neighbors]
        row_before = _row_neighbors(entry['node'], _text)
        row_after = _row_neighbors(entry['node'], projected_text)
        def context(text, start, end, row=None):
            prefix, suffix = text[max(0, start - 64):start], text[end:end + 64]
            if row is not None:
                if start < 64:
                    prefix = row[0][-(64 - start):] + ' ' + prefix
                if len(text) - end < 64:
                    suffix += ' ' + row[1][:64 - len(suffix)]
                return {'text': text[start:end], 'prefix': prefix, 'suffix': suffix}
            if start < 64 and not following:
                prefix = previous[-(64 - start):] + ' ' + prefix
            if len(text) - end < 64:
                suffix += ' ' + following[:64 - len(suffix)]
            return {'text': text[start:end], 'prefix': prefix, 'suffix': suffix,
                    'paragraph': {'prefix': text[max(0, start - 64):start], 'suffix': text[end:end + 64]}}
        before = context(value, op.start, op.end, row_before)
        after = None
        if op.action in {'replace', 'insert', 'delete', 'format', 'comment'}:
            siblings = [other for other in request.operations if other.locator_id == op.locator_id and other.action in {'replace', 'insert', 'delete'}]
            projected = value
            for other in sorted(siblings, key=lambda item: item.start, reverse=True):
                projected = projected[:other.start] + other.text + projected[other.end:]
            shift = sum(len(other.text) - (other.end - other.start) for other in siblings if other.start < op.start)
            start = op.start + shift
            text = op.text if op.action in {'replace', 'insert', 'delete'} else op.quote
            after = context(projected, start, start + len(text), row_after)
            if approved_parts is not None:
                if approved_starts is not None:
                    start = approved_starts[(op.locator_id, op.start, op.end)]
                approved = approved_locations[entry['location']]
                approved_value = approved['text']
                approved_index = approved_paragraphs.index(approved['node'])
                adjacent = [approved_paragraphs[approved_index - 1] if approved_index else None,
                            approved_paragraphs[approved_index + 1] if approved_index + 1 < len(approved_paragraphs) else None]
                previous_after, following_after = [_text(node) if node is not None and node.getparent() is approved['node'].getparent() else '' for node in adjacent]
                prefix, suffix = approved_value[max(0, start - 64):start], approved_value[start + len(text):start + len(text) + 64]
                if start < 64 and not following_after:
                    prefix = previous_after[-(64 - start):] + ' ' + prefix
                if len(approved_value) - start - len(text) < 64:
                    suffix += ' ' + following_after[:64 - len(suffix)]
                after = {'text': approved_value[start:start + len(text)], 'prefix': prefix, 'suffix': suffix}
                approved_row = _row_neighbors(approved['node'], _text)
                if approved_row is None:
                    after['paragraph'] = {'prefix': approved_value[max(0, start - 64):start],
                                          'suffix': approved_value[start + len(text):start + len(text) + 64]}
                if approved_row is not None:
                    after = {'text': approved_value[start:start + len(text)],
                             'prefix': (approved_row[0][-(64 - start):] + ' ' if start < 64 else '') + approved_value[max(0, start - 64):start],
                             'suffix': approved_value[start + len(text):start + len(text) + 64]}
                    if len(approved_value) - start - len(text) < 64:
                        after['suffix'] += ' ' + approved_row[1][:64 - len(after['suffix'])]
        elif op.action == 'paragraph_format' or op.action.startswith('list_'):
            before = after = {'text': value[:200], 'prefix': '', 'suffix': ''}
        elif op.action in {'insert_before', 'insert_after'}:
            before = {'text': value[:200], 'prefix': '', 'suffix': ''}
            after = {'text': op.text[:200], 'prefix': '', 'suffix': ''}
        anchors.append({'location': op.locator_id, 'reason': op.reason, 'action': op.action,
                        'before': op.quote, 'after': op.text,
                        'contexts': {'before': before, 'after': after or {'text': '', 'prefix': '', 'suffix': ''}}})
    return anchors
