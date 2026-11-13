from __future__ import annotations

import hashlib
from pathlib import Path
import posixpath

from .rag_sources import MAX_SOURCE_BYTES, SourceError, _package


P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
MAX_SLIDES = 1000
MAX_ITEMS = 20000


def _number(value, low, high):
    if type(value) is not int or not low <= value <= high:
        raise SourceError('invalid')


def _load(path):
    path = Path(path)
    if path.suffix.lower() != '.pptx' or path.is_symlink() or path.stat().st_size > MAX_SOURCE_BYTES:
        raise SourceError('unsupported')
    with path.open('rb') as stream:
        data = stream.read(MAX_SOURCE_BYTES + 1)
    if len(data) > MAX_SOURCE_BYTES:
        raise SourceError('capacity')
    try:
        parts = _package(data, '.pptx')
        presentation = parts['ppt/presentation.xml']
        if presentation.tag != P + 'presentation':
            raise SourceError('invalid')
        relations = _rels(parts, 'ppt/presentation.xml')
        slides = []
        for node in presentation.findall('./' + P + 'sldIdLst/' + P + 'sldId'):
            target, kind = relations[node.attrib[R + 'id']]
            if not kind.endswith('/slide') or parts[target].tag != P + 'sld':
                raise SourceError('invalid')
            slides.append(target)
        if len(slides) > MAX_SLIDES or len(slides) != len(set(slides)):
            raise SourceError('capacity')
        return parts, slides, hashlib.sha256(data).hexdigest()
    except SourceError:
        raise
    except Exception:
        raise SourceError('invalid') from None


def _rels(parts, part):
    name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    root = parts.get(name)
    if root is None:
        return {}
    return {node.attrib['Id']: (posixpath.normpath(posixpath.join(posixpath.dirname(part), node.attrib['Target']))
                                if not node.attrib['Target'].startswith('/') else node.attrib['Target'][1:],
                                node.attrib.get('Type', '')) for node in root}


def _text(root):
    return '\n'.join(''.join(node.itertext()) for node in root.iter(A + 'p'))


def _anchor(shape):
    transform = shape.find('./' + P + 'spPr/' + A + 'xfrm')
    if transform is None:
        transform = shape.find('./' + P + 'xfrm')
    if transform is None:
        return None
    off, ext = transform.find(A + 'off'), transform.find(A + 'ext')
    if off is None or ext is None:
        return None
    try:
        values = [int(off.attrib['x']), int(off.attrib['y']), int(ext.attrib['cx']), int(ext.attrib['cy'])]
        if any(abs(value) > 10**12 for value in values):
            raise ValueError()
        return dict(zip(('x', 'y', 'cx', 'cy'), values), unit='EMU', coordinate_space='parent')
    except (KeyError, ValueError):
        raise SourceError('invalid') from None


def _items(parts, part):
    root, relations = parts[part], _rels(parts, part)
    items, titles, unsupported = [], [], set()
    tree = root.find('./' + P + 'cSld/' + P + 'spTree')
    if tree is None:
        raise SourceError('invalid')
    for shape in tree.iter():
        if shape.tag not in {P + 'sp', P + 'pic', P + 'graphicFrame', P + 'cxnSp'}:
            continue
        identifier = shape.find('.//' + P + 'cNvPr')
        base = {'shape_id': identifier.get('id') if identifier is not None else None,
                'anchor': _anchor(shape)}
        if shape.tag == P + 'sp':
            text = _text(shape)
            placeholder = shape.find('.//' + P + 'ph')
            is_title = placeholder is not None and placeholder.get('type') in {'title', 'ctrTitle'}
            if is_title:
                titles.append(text)
            items.append({**base, 'type': 'text_box', 'text': text, 'is_title': is_title})
        elif shape.tag == P + 'pic':
            blip = shape.find('.//' + A + 'blip')
            relation = relations.get(blip.get(R + 'embed')) if blip is not None else None
            if relation is None or not relation[1].endswith('/image'):
                raise SourceError('invalid')
            items.append({**base, 'type': 'image', 'text': '', 'part': relation[0],
                          'relationship': blip.get(R + 'embed')})
        elif shape.tag == P + 'graphicFrame':
            table = shape.find('.//' + A + 'tbl')
            chart = shape.find('.//' + C + 'chart')
            if table is not None:
                for row_index, row in enumerate(table.findall(A + 'tr')):
                    for column, cell in enumerate(row.findall(A + 'tc')):
                        items.append({**base, 'type': 'table_cell', 'row': row_index, 'column': column,
                                      'text': _text(cell), 'merge': {key: cell.get(key) for key in
                                        ('gridSpan', 'rowSpan', 'hMerge', 'vMerge') if cell.get(key) is not None}})
            elif chart is not None:
                relation = relations.get(chart.get(R + 'id'))
                if relation is None or not relation[1].endswith('/chart'):
                    raise SourceError('invalid')
                chart_root = parts[relation[0]]
                if getattr(chart_root, 'tag', None) != C + 'chartSpace':
                    raise SourceError('invalid')
                title = chart_root.find('.//' + C + 'title')
                title_text = '' if title is None else ' '.join(node.text or '' for node in title.iter()
                                                              if node.tag in {A + 't', C + 'v'})
                items.append({**base, 'type': 'chart', 'text': title_text, 'part': relation[0]})
            else:
                unsupported.add('graphic_frame')
        else:
            unsupported.add('connector')
        if len(items) > MAX_ITEMS:
            raise SourceError('capacity')
    for target, kind in relations.values():
        if kind.endswith('/notesSlide'):
            notes = parts[target]
            if notes.tag != P + 'notes':
                raise SourceError('invalid')
            for shape in notes.iter(P + 'sp'):
                placeholder = shape.find('.//' + P + 'ph')
                if placeholder is not None and placeholder.get('type') in {'sldNum', 'dt', 'hdr', 'ftr', 'sldImg'}:
                    continue
                text = _text(shape)
                if text:
                    items.append({'type': 'notes', 'text': text, 'part': target})
    if len(items) > MAX_ITEMS:
        raise SourceError('capacity')
    return items, '\n'.join(titles), sorted(unsupported)


def _unchanged(path, expected):
    with Path(path).open('rb') as stream:
        data = stream.read(MAX_SOURCE_BYTES + 1)
    if hashlib.sha256(data).hexdigest() != expected:
        raise SourceError('source_changed')


def read_pptx_structure(path, *, offset=0, limit=20):
    _number(offset, 0, MAX_SLIDES)
    _number(limit, 1, 20)
    parts, slides, fingerprint = _load(path)
    results = []
    for index in range(offset, min(offset + limit, len(slides))):
        items, title, unsupported = _items(parts, slides[index])
        results.append({'slide_index': index, 'part': slides[index], 'title': title[:256],
                        'title_truncated': len(title) > 256, 'item_count': len(items),
                        'images': sum(item['type'] == 'image' for item in items),
                        'charts': sum(item['type'] == 'chart' for item in items),
                        'table_cells': sum(item['type'] == 'table_cell' for item in items),
                        'has_notes': any(item['type'] == 'notes' for item in items),
                        'unsupported': unsupported})
    _unchanged(path, fingerprint)
    more = offset + len(results) < len(slides)
    return {'slide_count': len(slides), 'slides': results, 'offset': offset, 'truncated': more,
            'next_offset': offset + len(results) if more else None, 'source_sha256': fingerprint,
            'read_only': True, 'limitations': ['Text and local shape anchors only; no slide rendering or embedded object execution']}


def read_pptx_slide(path, *, slide_index=0, offset=0, limit=20, text_offset=0, max_chars=4000):
    _number(slide_index, 0, MAX_SLIDES - 1)
    _number(offset, 0, MAX_ITEMS)
    _number(limit, 1, 50)
    _number(text_offset, 0, MAX_SOURCE_BYTES)
    _number(max_chars, 200, 20000)
    parts, slides, fingerprint = _load(path)
    if slide_index >= len(slides):
        raise SourceError('invalid')
    items, title, unsupported = _items(parts, slides[slide_index])
    if offset > len(items) or (text_offset and (offset == len(items) or text_offset > len(items[offset]['text']))):
        raise SourceError('invalid')
    results, budget, cursor, next_text = [], max_chars, offset, 0
    while cursor < len(items) and len(results) < limit and budget:
        item = items[cursor]
        start = text_offset if cursor == offset else 0
        text = item['text'][start:start + min(1024, budget)]
        partial = start + len(text) < len(item['text'])
        results.append({**item, 'text': text, 'item_index': cursor, 'text_offset': start, 'text_truncated': partial})
        budget -= len(text)
        if partial:
            next_text = start + len(text)
            break
        cursor += 1
    _unchanged(path, fingerprint)
    more = cursor < len(items)
    return {'slide_index': slide_index, 'title': title[:256], 'title_truncated': len(title) > 256,
            'items': results, 'item_count': len(items), 'truncated': more,
            'next_offset': cursor if more else None, 'next_text_offset': next_text if more else None,
            'unsupported': unsupported, 'source_sha256': fingerprint, 'read_only': True}
