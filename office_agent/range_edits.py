from copy import deepcopy
from datetime import datetime, timezone
import json
from io import BytesIO
import zipfile
import regex
from typing import Annotated, Literal

from lxml import etree as ET
from pydantic import Field

from .professional_edits import ConstrainedDocx, W, XML, _canonical, _hash_text, _inventory, _protect, _text
from .provenance import identifier
from .rag import SECRET
from .writes import GuardError, Strict, digest


def validate_external_relationships(parts):
    from urllib.parse import urlsplit
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    for name, data in parts.items():
        if not name.endswith('.rels'):
            continue
        root = ET.fromstring(data)
        for node in root.findall(rel + 'Relationship'):
            if node.get('TargetMode') != 'External':
                continue
            target = urlsplit(node.get('Target', ''))
            if name != 'word/_rels/document.xml.rels' or not node.get('Type', '').endswith('/hyperlink') or target.scheme not in {'https', 'http', 'mailto'} or target.username or target.password:
                raise GuardError('professional_unsafe_document')


def static_render_parts(parts):
    from .range_fields import analyze
    analyze(parts)
    validate_external_relationships(parts)
    result = dict(parts)
    for name, data in parts.items():
        lower = name.lower()
        if any(token in lower for token in ('vbaproject', 'activex', 'embeddings/', 'externallinks/', 'connections', 'customui')):
            raise GuardError('professional_unsafe_document')
        if not name.endswith(('.xml', '.rels')):
            continue
        root = ET.fromstring(data)
        for node in list(root.iter()):
            tag = ET.QName(node).localname
            if tag in {'altChunk', 'oleObject', 'OLEObject'}:
                raise GuardError('professional_unsafe_document')
            if tag in {'instrText', 'fldSimple'}:
                if tag == 'fldSimple':
                    parent = node.getparent()
                    index = parent.index(node)
                    for child in list(node):
                        parent.insert(index, child)
                        index += 1
                    parent.remove(node)
                else:
                    node.getparent().remove(node)
            elif tag == 'fldChar':
                node.getparent().remove(node)
            elif tag == 'hyperlink':
                node.attrib.pop('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', None)
            elif tag == 'Relationship' and node.get('TargetMode') == 'External':
                node.getparent().remove(node)
        result[name] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    from .rag_sources import _package, SourceError
    stream = BytesIO()
    with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in result.items():
            archive.writestr(name, data)
    try:
        _package(stream.getvalue(), '.docx')
    except SourceError:
        raise GuardError('professional_unsafe_document') from None
    return result


def _order(properties):
    sequences = {
        'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
        'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
    }
    order = sequences[ET.QName(properties).localname].split()
    rank = {W + value: index for index, value in enumerate(order)}
    original = list(properties)
    if any(child.tag not in rank for child in original):
        raise GuardError('professional_format_scope')
    properties[:] = sorted(original, key=lambda child: rank[child.tag])


class CharacterFormat(Strict):
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font: str | None = Field(default=None, min_length=1, max_length=80)
    size_pt: float | None = Field(default=None, ge=6, le=72, allow_inf_nan=False)


class ParagraphFormat(Strict):
    style: str | None = Field(default=None, min_length=1, max_length=100)
    alignment: Literal['left', 'center', 'right', 'both'] | None = None
    before_pt: float | None = Field(default=None, ge=0, le=144, allow_inf_nan=False)
    after_pt: float | None = Field(default=None, ge=0, le=144, allow_inf_nan=False)
    line_multiple: float | None = Field(default=None, ge=0.5, le=3, allow_inf_nan=False)
    left_pt: float | None = Field(default=None, ge=0, le=360, allow_inf_nan=False)
    first_line_pt: float | None = Field(default=None, ge=-72, le=144, allow_inf_nan=False)


class RangeOperation(Strict):
    locator_id: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    expected_text_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    action: Literal['replace', 'insert', 'delete', 'format', 'comment', 'paragraph_format', 'insert_before', 'insert_after',
                    'paragraph_delete', 'split', 'merge_next', 'list_continue', 'list_restart', 'row_insert', 'row_delete',
                    'revision_accept', 'revision_reject']
    island_id: str | None = Field(default=None, pattern=r'^island:[a-f0-9]{64}$')
    start: int = Field(default=0, ge=0, le=8000)
    end: int = Field(default=0, ge=0, le=8000)
    quote: str = Field(default='', max_length=8000)
    text: str = Field(default='', max_length=2000)
    reason: str = Field(min_length=1, max_length=400)
    inheritance: Literal['reject_mixed', 'start', 'distributed'] = 'reject_mixed'
    character_format: CharacterFormat | None = None
    paragraph_format: ParagraphFormat | None = None
    revision_id: str | None = Field(default=None, pattern=r'^rev:[a-f0-9]{64}$')
    revision_ids: list[Annotated[str, Field(pattern=r'^rev:[a-f0-9]{64}$')]] = Field(default_factory=list, max_length=8)
    numbering_locator_id: str | None = Field(default=None, pattern=r'^loc:[a-f0-9]{64}$')


class RangeIntent(Strict):
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    operations: list[RangeOperation] = Field(min_length=1, max_length=12)
    mode: Literal['clean', 'tracked'] = 'clean'
    preserve_numbers: bool = True
    preserve_citations: bool = True
    preserve_formulas: bool = True
    protected_literals: list[str] = Field(default_factory=list, max_length=30)
    editing_profile: Literal['general', 'technical_report_v1'] = 'general'
    profile_rule_sha256: str | None = Field(default=None, pattern=r'^[a-f0-9]{64}$')


class RangeViewArgs(Strict):
    offset: int = Field(default=0, ge=0, le=4000)


class IntentAssessment(Strict):
    verdict: Literal['consistent', 'concerns', 'insufficient']
    reasons: list[str] = Field(min_length=1, max_length=5)
    requires_human_review: Literal[True] = True


class RangeDocx(ConstrainedDocx):
    kind: Literal['range_docx'] = 'range_docx'
    changes: list = Field(default_factory=list, max_length=0)
    operations: list[RangeOperation] = Field(min_length=1, max_length=12)
    mode: Literal['clean', 'tracked'] = 'clean'
    editing_profile: Literal['general', 'technical_report_v1'] = 'general'
    profile_rule_sha256: str | None = Field(default=None, pattern=r'^[a-f0-9]{64}$')
    revision_date: str = Field(pattern=r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$')


def _boundary(text, offset):
    return offset == 0 or any(match.end() == offset for match in regex.finditer(r'\X', text))


def _islands(entry):
    paragraph = entry['node']
    cursor, field_depth, result, current = 0, 0, [], None

    def barrier():
        nonlocal current
        current = None

    def run(node, container):
        nonlocal cursor, field_depth, current
        value = _text(node)
        marks = node.findall(W + 'fldChar')
        supported = not field_depth and all(child.tag in {W + 'rPr', W + 't'} for child in node)
        supported = supported and not any(child.tag == W + 'rPrChange' for child in node.iter())
        if supported and value:
            if current is None or current['container'] is not container:
                current = {'start': cursor, 'end': cursor, 'container': container, 'runs': []}
                result.append(current)
            current['runs'].append((cursor, cursor + len(value), node))
            current['end'] = cursor + len(value)
        else:
            barrier()
        cursor += len(value)
        for mark in marks:
            if mark.get(W + 'fldCharType') == 'begin':
                field_depth += 1
            elif mark.get(W + 'fldCharType') == 'end':
                field_depth = max(0, field_depth - 1)

    for node in paragraph:
        if node.tag == W + 'pPr':
            continue
        if node.tag == W + 'r':
            run(node, paragraph)
        elif node.tag == W + 'hyperlink':
            barrier()
            for child in node:
                if child.tag == W + 'r':
                    run(child, node)
                else:
                    barrier()
                    cursor += len(_text(child))
            barrier()
        else:
            barrier()
            cursor += len(_text(node))
    for island in result:
        island['id'] = 'island:' + identifier(entry['locator_id'], str(island['start']), str(island['end']),
                                            *(_canonical(run).hex() for _, _, run in island['runs']))
    return result


def _revisions(entry):
    rows = []
    for node in entry['node'].iter():
        if node.tag not in {W + 'ins', W + 'del', W + 'moveFrom', W + 'moveTo', W + 'rPrChange', W + 'pPrChange'}:
            continue
        supported = node.tag in {W + 'ins', W + 'del'} and node.getparent() is entry['node']
        supported = supported and all(child.tag == W + 'r' and all(t.tag in {W + 'rPr', W + 't', W + 'delText'} for t in child) for child in node)
        supported = supported and not any(n is not node and n.tag in {W + name for name in ('ins', 'del', 'rPrChange', 'pPrChange', 'moveFrom', 'moveTo')} for n in node.iter())
        content = ''.join(t.text or '' for t in node.iter() if t.tag in {W + 't', W + 'delText'})[:2000]
        author = node.get(W + 'author', '')[:100]
        date = node.get(W + 'date', '')[:40]
        if SECRET.search(content) or SECRET.search(author) or SECRET.search(date):
            content = author = '[redacted]'
            date = '[redacted]'
            supported = False
        rows.append({'id': 'rev:' + identifier(entry['locator_id'], str(len(rows)), _canonical(node).hex()),
                     'type': ET.QName(node).localname, 'author': author,
                     'date': date, 'text': content,
                     'supported': supported, 'node': node})
    return rows


def view(editor, path, offset=0):
    from .range_fields import analyze, public_fields, impact
    from .advanced_writes import _parts
    from .professional_edits import snapshot
    result = snapshot(editor, path, offset=offset, limit=20, _range_edit=True)
    version = result['document']['source_sha256']
    parts = _parts(editor.space.file(str(path)), range_edit=True)
    root, entries = _inventory(parts, version)
    _, fields = analyze(parts)
    paragraphs = list(root.iter(W + 'p'))
    result['reference_impact'] = impact(root, entries, fields, [], version)
    mapping = {row['locator_id']: row for row in entries}
    styles = ET.fromstring(parts['word/styles.xml']) if 'word/styles.xml' in parts else None
    result['styles'] = [node.get(W + 'styleId') for node in styles.findall(W + 'style')
                        if node.get(W + 'type') == 'paragraph' and node.get(W + 'styleId') and
                        len(node.get(W + 'styleId')) <= 100 and not SECRET.search(node.get(W + 'styleId'))][:150] if styles is not None else []
    result['coordinate_system'] = 'Unicode scalar offsets in paragraph w:t text; half-open; protected structures are barriers'
    result['mode'] = 'clean by default; tracked only for ordinary inline text edits'
    for item in result['items']:
        if item.get('style') and SECRET.search(item['style']):
            item['style'] = None
        if item.get('refusal') == 'sensitive_or_oversize':
            continue
        entry = mapping[item['locator_id']]
        paragraph_index = paragraphs.index(entry['node'])
        item['fields'] = public_fields(fields, paragraph_index)
        item['reference_target_locked'] = any(row.get('target_paragraph') == paragraph_index for row in fields)
        item['action_support'] = {mode: {action: action_reason(entry, action, parts, item['reference_target_locked'], mode)
                                       for action in RangeOperation.model_fields['action'].annotation.__args__}
                                  for mode in ('clean', 'tracked')}
        item['islands'] = [{'id': row['id'], 'start': row['start'], 'end': row['end'],
                            'hyperlink': row['container'].tag == W + 'hyperlink',
                            'runs': [{'start': a, 'end': b, 'direct_properties_sha256': _hash_text(_canonical(run.find(W + 'rPr')).hex()) if run.find(W + 'rPr') is not None else None}
                                     for a, b, run in row['runs']]} for row in _islands(entry)] if entry['refusal'] != 'merged_cell_continuation' else []
        item['revisions'] = [{k: v for k, v in row.items() if k != 'node'} for row in _revisions(entry)]
        item['protected_structures'] = sorted({ET.QName(node).localname for node in entry['node'].iter()
                                               if node.tag not in {W + t for t in ('p', 'pPr', 'r', 'rPr', 't')}})
    if digest(editor.space.file(str(path))) != version:
        raise GuardError('professional_source_changed')
    return result


def _text_run(original, value):
    copied = deepcopy(original)
    for child in list(copied):
        if child.tag != W + 'rPr':
            copied.remove(child)
    text = ET.SubElement(copied, W + 't')
    text.text = value
    text.set(XML + 'space', 'preserve')
    return copied


def _properties(node, name):
    value = node.find(W + name)
    if value is None:
        value = ET.Element(W + name)
        node.insert(0, value)
    return value


def _set(properties, tag, **values):
    node = properties.find(W + tag)
    if node is None:
        node = ET.SubElement(properties, W + tag)
    for key, value in values.items():
        node.set(W + key, str(value))


def _character(run, spec):
    properties = _properties(run, 'rPr')
    for attr, tag in [('bold', 'b'), ('italic', 'i'), ('underline', 'u')]:
        value = getattr(spec, attr)
        if value is not None:
            _set(properties, tag, val=('single' if value else 'none') if attr == 'underline' else ('1' if value else '0'))
            if attr != 'underline':
                _set(properties, tag + 'Cs', val='1' if value else '0')
    if spec.font is not None:
        if not spec.font.strip() or SECRET.search(spec.font) or any(ord(c) < 32 for c in spec.font):
            raise GuardError('professional_invalid_request')
        _set(properties, 'rFonts', ascii=spec.font, hAnsi=spec.font, eastAsia=spec.font, cs=spec.font)
        fonts = properties.find(W + 'rFonts')
        for name in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme', 'csTheme'):
            fonts.attrib.pop(W + name, None)
    if spec.size_pt is not None:
        if spec.size_pt * 2 != round(spec.size_pt * 2):
            raise GuardError('professional_format_precision')
        for name in ('sz', 'szCs'):
            _set(properties, name, val=round(spec.size_pt * 2))
    _order(properties)


def _paragraph(node, spec, parts):
    if spec is None or not spec.model_dump(exclude_none=True):
        raise GuardError('professional_no_changes')
    properties = _properties(node, 'pPr')
    if spec.style is not None:
        styles = ET.fromstring(parts.get('word/styles.xml', b'<styles/>'))
        if not any(s.get(W + 'styleId') == spec.style and s.get(W + 'type') == 'paragraph' for s in styles):
            raise GuardError('professional_style_unavailable')
        _set(properties, 'pStyle', val=spec.style)
    if spec.alignment is not None:
        _set(properties, 'jc', val=spec.alignment)
    for attr, name in [('before_pt', 'before'), ('after_pt', 'after')]:
        if getattr(spec, attr) is not None:
            _set(properties, 'spacing', **{name: round(getattr(spec, attr) * 20)})
            properties.find(W + 'spacing').attrib.pop(W + name + 'Autospacing', None)
    if spec.line_multiple is not None:
        _set(properties, 'spacing', line=round(spec.line_multiple * 240), lineRule='auto')
    if spec.left_pt is not None:
        _set(properties, 'ind', left=round(spec.left_pt * 20))
    if spec.first_line_pt is not None:
        ind = properties.find(W + 'ind')
        if ind is not None:
            for name in ('firstLine', 'hanging', 'firstLineChars', 'hangingChars'):
                ind.attrib.pop(W + name, None)
        _set(properties, 'ind', **{('firstLine' if spec.first_line_pt >= 0 else 'hanging'): round(abs(spec.first_line_pt) * 20)})
    _order(properties)


def _check_range(entry, op, islands):
    if not _boundary(entry['text'], op.start) or not _boundary(entry['text'], op.end) or op.start > op.end:
        raise GuardError('professional_invalid_range')
    if entry['text'][op.start:op.end] != op.quote:
        raise GuardError('professional_stale_locator')
    island = next((row for row in islands if row['id'] == op.island_id), None)
    if island is None or not island['start'] <= op.start <= op.end <= island['end']:
        raise GuardError('professional_protected_boundary')
    if op.action == 'insert' and (op.start != op.end or not op.text):
        raise GuardError('professional_invalid_range')
    if op.action != 'insert' and op.start == op.end:
        raise GuardError('professional_invalid_range')
    if op.action == 'delete' and op.text:
        raise GuardError('professional_invalid_request')
    return island


def _inline(entry, op, island, request, next_revision):
    selected = [(a, b, run) for a, b, run in island['runs'] if (a < op.end and b > op.start) or
                (op.start == op.end and a <= op.start <= b)]
    if op.start == op.end:
        selected = [next((row for row in selected if row[0] <= op.start < row[1]), selected[-1])]
    signatures = {_canonical(run.find(W + 'rPr')) if run.find(W + 'rPr') is not None else b'' for _, _, run in selected}
    if op.action != 'format' and len(signatures) > 1 and op.inheritance == 'reject_mixed':
        raise GuardError('professional_mixed_format_policy')
    if op.inheritance == 'distributed' and len(op.text) != op.end - op.start:
        raise GuardError('professional_distributed_length')
    if request.mode == 'tracked' and (op.action == 'format' or island['container'].tag != W + 'p'):
        raise GuardError('professional_revision_scope')
    first = selected[0][2]
    parent = first.getparent()
    index = parent.index(first)
    output, old = [], []
    for a, b, run in selected:
        value = _text(run)
        left, right = max(op.start - a, 0), min(op.end - a, len(value))
        if left > 0:
            output.append(_text_run(run, value[:left]))
        middle = _text_run(run, value[left:right])
        if right > left:
            old.append(middle)
        if op.action == 'format':
            if op.character_format is None or not op.character_format.model_dump(exclude_none=True):
                raise GuardError('professional_no_changes')
            _character(middle, op.character_format)
            output.append(middle)
        elif op.inheritance == 'distributed':
            output.append(_text_run(run, op.text[max(a, op.start) - op.start:min(b, op.end) - op.start]))
        elif run is first and request.mode == 'clean' and op.text:
            output.append(_text_run(run, op.text))
        if right < len(value):
            output.append(_text_run(run, value[right:]))
    if request.mode == 'tracked':
        prefix = [_text_run(first, _text(first)[:max(0, op.start - selected[0][0])])] if op.start > selected[0][0] else []
        last = selected[-1]
        suffix = [_text_run(last[2], _text(last[2])[op.end - last[0]:])] if op.end < last[1] else []
        output = prefix
        for kind, runs in [('del', old), ('ins', [_text_run(first, op.text)] if op.text else [])]:
            if not runs:
                continue
            wrapper = ET.Element(W + kind, {W + 'id': str(next_revision[0]), W + 'author': 'Office Agent', W + 'date': request.revision_date})
            next_revision[0] += 1
            for run in runs:
                if kind == 'del':
                    for text in run.findall(W + 't'):
                        text.tag = W + 'delText'
                wrapper.append(run)
            output.append(wrapper)
        output += suffix
    for _, _, run in selected:
        parent.remove(run)
    for run in output:
        parent.insert(index, run)
        index += 1


def _ordinary(entry, parts=None):
    node = entry['node']
    if SECRET.search(entry['text']):
        raise GuardError('professional_sensitive_or_control_text')
    properties = node.find(W + 'pPr')
    if '/table/' in entry['location'] or entry['refusal'] or (properties is not None and any(
            n.tag in {W + 'sectPr', W + 'numPr', W + 'pPrChange'} for n in properties.iter())):
        raise GuardError('professional_structure_scope')
    if parts and entry['style']:
        styles = ET.fromstring(parts.get('word/styles.xml', b'<styles/>'))
        name, seen = entry['style'], set()
        for _ in range(12):
            if not name or name in seen:
                break
            seen.add(name)
            style = next((s for s in styles if s.get(W + 'styleId') == name), None)
            if style is None:
                break
            if style.find('./' + W + 'pPr/' + W + 'numPr') is not None:
                raise GuardError('professional_structure_scope')
            base = style.find(W + 'basedOn')
            name = base.get(W + 'val') if base is not None else None
    return node


def action_reason(entry, action, parts, target_locked=False, mode='clean'):
    if target_locked and action not in {'comment', 'format', 'paragraph_format', 'insert_before', 'insert_after'}:
        return 'professional_reference_target_locked'
    if mode == 'tracked' and action not in {'replace', 'insert', 'delete'}:
        return 'professional_revision_scope'
    if action in {'insert_before', 'insert_after', 'paragraph_delete', 'split', 'merge_next'}:
        try:
            _ordinary(entry, parts)
        except GuardError as error:
            return str(error)
    if action in {'replace', 'insert', 'delete', 'format', 'comment'}:
        islands = _islands(entry)
        if not islands or (mode == 'tracked' and not any(row['container'] is entry['node'] for row in islands)):
            return 'professional_protected_boundary'
    if action.startswith('revision_') and not any(row['supported'] for row in _revisions(entry)):
        return 'professional_revision_scope'
    if action == 'comment' and any('comment' in name.casefold() for name in parts):
        return 'professional_existing_comments'
    if action.startswith('row_'):
        try:
            _row_scope(entry['node'], action)
        except GuardError as error:
            return str(error)
    if action == 'paragraph_format' and entry['node'].find('./' + W + 'pPr/' + W + 'pPrChange') is not None:
        return 'professional_structure_scope'
    if action.startswith('list_') and (entry['refusal'] or '/table/' in entry['location']):
        return 'professional_numbering_scope'
    return None


def _new_paragraph_ids(node, source, seed):
    namespace = '{http://schemas.microsoft.com/office/word/2010/wordml}'
    occupied = {(n.get(namespace + 'paraId') or '').upper() for n in source.iter(W + 'p')}
    for index, paragraph in enumerate(node.iter(W + 'p')):
        if paragraph.get('{http://www.w3.org/XML/1998/namespace}id') is not None:
            raise GuardError('professional_structure_scope')
        for name in ('paraId', 'textId'):
            if namespace + name not in paragraph.attrib:
                continue
            candidate = int(identifier(seed, str(index), name)[:8], 16) & 0x7fffffff or 1
            while f'{candidate:08X}' in occupied:
                candidate = (candidate + 1) & 0x7fffffff or 1
            value = f'{candidate:08X}'
            paragraph.set(namespace + name, value)
            occupied.add(value)


def _comment(entry, op, island, index):
    if island['container'] is not entry['node'] or not op.text.strip():
        raise GuardError('professional_comment_range')
    selected = [(a, b, run) for a, b, run in island['runs'] if a < op.end and b > op.start]
    parent = entry['node']
    position = parent.index(selected[0][2])
    first, last = selected[0], selected[-1]
    nodes = []
    if op.start > first[0]:
        nodes.append(_text_run(first[2], _text(first[2])[:op.start - first[0]]))
    nodes.append(ET.Element(W + 'commentRangeStart', {W + 'id': str(index)}))
    for a, b, run in selected:
        nodes.append(_text_run(run, _text(run)[max(0, op.start - a):min(b, op.end) - a]))
    nodes.append(ET.Element(W + 'commentRangeEnd', {W + 'id': str(index)}))
    reference = ET.Element(W + 'r')
    ET.SubElement(reference, W + 'commentReference', {W + 'id': str(index)})
    nodes.append(reference)
    if op.end < last[1]:
        nodes.append(_text_run(last[2], _text(last[2])[op.end - last[0]:]))
    for _, _, run in selected:
        parent.remove(run)
    for node in nodes:
        parent.insert(position, node)
        position += 1


def _comment_parts(parts, comments, date):
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
    root = ET.Element(W + 'comments', nsmap={'w': W[1:-1]})
    for index, op in enumerate(comments):
        node = ET.SubElement(root, W + 'comment', {W + 'id': str(index), W + 'author': 'Office Agent', W + 'date': date})
        ET.SubElement(ET.SubElement(ET.SubElement(node, W + 'p'), W + 'r'), W + 't').text = op.text
    parts['word/comments.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    name = 'word/_rels/document.xml.rels'
    relationships = ET.fromstring(parts[name])
    ids = {node.get('Id') for node in relationships}
    rid = next('rId' + str(i) for i in range(1, len(ids) + 2) if 'rId' + str(i) not in ids)
    ET.SubElement(relationships, rel + 'Relationship', {'Id': rid, 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'Target': 'comments.xml'})
    parts[name] = ET.tostring(relationships, xml_declaration=True, encoding='UTF-8', standalone=True)
    types = ET.fromstring(parts['[Content_Types].xml'])
    ET.SubElement(types, ct + 'Override', {'PartName': '/word/comments.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'})
    parts['[Content_Types].xml'] = ET.tostring(types, xml_declaration=True, encoding='UTF-8', standalone=True)


def _row_scope(node, action):
    from .writes import _REVISION_ALL_MARKS
    cell = node.getparent()
    row = cell.getparent() if cell.tag == W + 'tc' else None
    table = row.getparent() if row is not None and row.tag == W + 'tr' else None
    if table is None or table.tag != W + 'tbl' or len(table.findall(W + 'tr')) > 100:
        raise GuardError('professional_table_scope')
    cells = len(row.findall(W + 'tc'))
    if row.find('./' + W + 'trPr/' + W + 'tblHeader') is not None or any(n.tag == W + mark for n in table.iter() for mark in _REVISION_ALL_MARKS):
        raise GuardError('professional_table_scope')
    if any(n.tag in {W + t for t in ('gridSpan', 'vMerge', 'ins', 'del', 'bookmarkStart', 'commentRangeStart', 'sectPr')} for n in table.iter()) or len(list(table.iter(W + 'tbl'))) != 1:
        raise GuardError('professional_table_scope')
    if any(len(r.findall(W + 'tc')) != cells for r in table.findall(W + 'tr')) or any(
            any(c.tag not in {W + 'tcPr', W + 'p'} for c in tc) or any(
                any(ch.tag not in {W + 'pPr', W + 'r'} for ch in p) or any(
                    ch.tag not in {W + 'rPr', W + 't'} for run in p.findall(W + 'r') for ch in run)
                for p in tc.findall(W + 'p')) for tc in table.iter(W + 'tc')):
        raise GuardError('professional_table_scope')
    if action == 'row_delete' and len(table.findall(W + 'tr')) == 1:
        raise GuardError('professional_table_scope')
    return row, table


def _structure(entry, op, entries, parts):
    node = entry['node']
    if op.action == 'paragraph_format':
        if node.find('./' + W + 'pPr/' + W + 'pPrChange') is not None:
            raise GuardError('professional_structure_scope')
        _paragraph(node, op.paragraph_format, parts)
        return
    if op.action in {'revision_accept', 'revision_reject'}:
        requested = ([op.revision_id] if op.revision_id else []) + op.revision_ids
        inventory = {row['id']: row for row in _revisions(entry)}
        if not requested or len(set(requested)) != len(requested) or any(key not in inventory or not inventory[key]['supported'] for key in requested):
            raise GuardError('professional_revision_scope')
        for key in requested:
            row = inventory[key]
            wrapper = row['node']
            position = node.index(wrapper)
            keep = (row['type'] == 'ins') == (op.action == 'revision_accept')
            if keep:
                for run in list(wrapper):
                    for text in run.findall(W + 'delText'):
                        text.tag = W + 't'
                    node.insert(position, run)
                    position += 1
            node.remove(wrapper)
        return
    if op.action in {'row_insert', 'row_delete'}:
        row, table = _row_scope(node, op.action)
        if op.action == 'row_delete':
            if len(table.findall(W + 'tr')) == 1:
                raise GuardError('professional_table_scope')
            table.remove(row)
        else:
            copied = deepcopy(row)
            _new_paragraph_ids(copied, table.getroottree().getroot(), entry['locator_id'] + ':row-insert')
            for text in copied.iter(W + 't'):
                text.text = ''
            table.insert(table.index(row) + 1, copied)
        return
    node = _ordinary(entry, parts)
    parent = node.getparent()
    if op.action in {'insert_before', 'insert_after'}:
        if not op.text.strip():
            raise GuardError('professional_invalid_request')
        added = ET.Element(W + 'p')
        properties = node.find(W + 'pPr')
        if properties is not None:
            added.append(deepcopy(properties))
        run = ET.SubElement(added, W + 'r')
        text = ET.SubElement(run, W + 't')
        text.text = op.text
        text.set(XML + 'space', 'preserve')
        if op.paragraph_format:
            _paragraph(added, op.paragraph_format, parts)
        parent.insert(parent.index(node) + (op.action == 'insert_after'), added)
    elif op.action == 'paragraph_delete':
        if len(parent.findall(W + 'p')) <= 1:
            raise GuardError('professional_structure_scope')
        parent.remove(node)
    elif op.action == 'merge_next':
        following = node.getnext()
        other = next((row for row in entries if row['node'] is following), None)
        if other is None:
            raise GuardError('professional_structure_scope')
        _ordinary(other, parts)
        prop = lambda p: _canonical(p.find(W + 'pPr')) if p.find(W + 'pPr') is not None else b''
        if prop(node) != prop(following):
            raise GuardError('professional_structure_scope')
        for run in list(following.findall(W + 'r')):
            node.append(run)
        parent.remove(following)
    elif op.action == 'split':
        if not 0 < op.start < len(entry['text']) or not _boundary(entry['text'], op.start):
            raise GuardError('professional_invalid_range')
        from .tracked_edits import _slices
        runs = node.findall(W + 'r')
        left, right = _slices(runs, 0, op.start), _slices(runs, op.start, len(entry['text']))
        added = deepcopy(node)
        _new_paragraph_ids(added, parent.getroottree().getroot(), entry['locator_id'] + ':split')
        for target, values in ((node, left), (added, right)):
            for run in list(target.findall(W + 'r')):
                target.remove(run)
            target.extend(values)
        parent.insert(parent.index(node) + 1, added)
    elif op.action in {'list_continue', 'list_restart'}:
        raise GuardError('professional_numbering_scope')


def _numbering(entry, op, targets, parts):
    if entry['refusal'] or '/table/' in entry['location']:
        raise GuardError('professional_numbering_scope')
    numbering = ET.fromstring(parts.get('word/numbering.xml', b'<numbering/>'))
    styles = ET.fromstring(parts.get('word/styles.xml', b'<styles/>'))
    def effective(row):
        properties = row['node'].find(W + 'pPr')
        direct = properties.find(W + 'numPr') if properties is not None else None
        layers = [direct] if direct is not None else []
        name = row['style']
        visited = set()
        for _ in range(12):
            if not name:
                break
            if name in visited:
                raise GuardError('professional_numbering_scope')
            visited.add(name)
            style = next((s for s in styles if s.get(W + 'styleId') == name), None)
            if style is None:
                break
            num = style.find('./' + W + 'pPr/' + W + 'numPr')
            if num is not None:
                layers.append(num)
            base = style.find(W + 'basedOn')
            name = base.get(W + 'val') if base is not None else None
        else:
            raise GuardError('professional_numbering_scope')
        result = ET.Element(W + 'numPr')
        found = set()
        for layer in layers:
            for child in layer:
                if child.tag not in {W + 'numId', W + 'ilvl'}:
                    raise GuardError('professional_numbering_scope')
                if child.tag not in found:
                    result.append(deepcopy(child))
                    found.add(child.tag)
        if W + 'numId' not in found:
            raise GuardError('professional_numbering_scope')
        order = {W + 'ilvl': 0, W + 'numId': 1}
        result[:] = sorted(result, key=lambda node: order[node.tag])
        return result
    reference = targets.get(op.numbering_locator_id) if op.numbering_locator_id else entry
    if reference is None:
        raise GuardError('professional_numbering_scope')
    properties = effective(reference)
    num_id = properties.find(W + 'numId')
    level = properties.find(W + 'ilvl')
    level_id = level.get(W + 'val') if level is not None else '0'
    definition = next((n for n in numbering.findall(W + 'num') if num_id is not None and n.get(W + 'numId') == num_id.get(W + 'val')), None)
    if definition is None or not level_id.isdigit() or int(level_id) > 8:
        raise GuardError('professional_numbering_scope')
    abstract_id = definition.find(W + 'abstractNumId')
    abstract = next((n for n in numbering.findall(W + 'abstractNum') if abstract_id is not None and n.get(W + 'abstractNumId') == abstract_id.get(W + 'val')), None)
    if abstract is None or abstract.find(W + 'numStyleLink') is not None or abstract.find(W + 'styleLink') is not None:
        raise GuardError('professional_numbering_scope')
    if not any(n.get(W + 'ilvl') == level_id for n in abstract.findall(W + 'lvl')):
        raise GuardError('professional_numbering_scope')
    original_num = num_id.get(W + 'val')
    if op.action == 'list_restart':
        copied = deepcopy(definition)
        new_id = max([int(n.get(W + 'numId')) for n in numbering.findall(W + 'num') if (n.get(W + 'numId') or '').isdigit()] + [0]) + 1
        copied.set(W + 'numId', str(new_id))
        override = next((n for n in copied.findall(W + 'lvlOverride') if n.get(W + 'ilvl') == level_id), None)
        if override is None:
            override = ET.SubElement(copied, W + 'lvlOverride', {W + 'ilvl': level_id})
        start = override.find(W + 'startOverride')
        if start is None:
            start = ET.Element(W + 'startOverride')
            override.insert(0, start)
        start.set(W + 'val', '1')
        numbering.append(copied)
        num_id.set(W + 'val', str(new_id))
        parts['word/numbering.xml'] = ET.tostring(numbering, xml_declaration=True, encoding='UTF-8', standalone=True)
    node = _properties(entry['node'], 'pPr')
    previous = node.find(W + 'numPr')
    if previous is not None:
        node.remove(previous)
    node.append(properties)
    _order(node)
    candidates = []
    for candidate in targets.values():
        if candidate['node'] is entry['node']:
            continue
        try:
            effective_num = effective(candidate).find(W + 'numId')
            if effective_num is not None and effective_num.get(W + 'val') == original_num:
                candidates.append(candidate['location'])
        except GuardError:
            continue
    return {'reference_num_id': original_num, 'new_num_id': num_id.get(W + 'val'), 'level': level_id,
            'other_registered_paragraphs_sharing_numbering': candidates[:40], 'other_count': len(candidates),
            'truncated': len(candidates) > 40, 'display_numbers': 'same numbering stream may renumber; Word list-label rendering remains renderer/human review, not a complete numbering emulator'}


def transform(parts, request):
    from .range_fields import analyze, guard_targets, impact
    if request.editing_profile != 'general' and any(op.action not in {'replace', 'insert', 'delete', 'format', 'paragraph_format', 'comment'} for op in request.operations):
        raise GuardError('professional_profile_operation')
    original_parts = parts
    parts = dict(parts)
    root, entries = _inventory(parts, request.source_sha256)
    _, fields = analyze(parts)
    targets = {row['locator_id']: row for row in entries}
    prepared, before, after, occupied, structural, insertions, comments = [], [], [], {}, set(), {}, []
    wants_comments = any(op.action == 'comment' for op in request.operations)
    if wants_comments and (request.mode != 'clean' or any('comment' in name.casefold() for name in parts) or any(
            node.tag in {W + 'commentRangeStart', W + 'commentReference'} for node in root.iter())):
        raise GuardError('professional_existing_comments')
    for op in request.operations:
        entry = targets.get(op.locator_id)
        if entry is None or entry['text_sha256'] != op.expected_text_sha256:
            raise GuardError('professional_stale_locator')
        if SECRET.search(entry['text']):
            raise GuardError('professional_sensitive_or_control_text')
        if entry['refusal'] == 'merged_cell_continuation' or len(entry['text']) > 8000:
            raise GuardError('professional_structure_scope')
        if any(SECRET.search(value) or any(ord(c) < 32 or 0xD800 <= ord(c) <= 0xDFFF for c in value)
               for value in (op.text, op.reason, op.quote)):
            raise GuardError('professional_sensitive_or_control_text')
        if not op.reason.strip():
            raise GuardError('professional_invalid_reason')
        if ((op.action not in {'replace', 'insert', 'comment', 'insert_before', 'insert_after'} and op.text)
                or (op.character_format is not None and op.action != 'format')
                or (op.paragraph_format is not None and op.action not in {'paragraph_format', 'insert_before', 'insert_after'})
                or (op.revision_id is not None and not op.action.startswith('revision_'))
                or (op.revision_ids and not op.action.startswith('revision_'))
                or (op.numbering_locator_id is not None and not op.action.startswith('list_'))):
            raise GuardError('professional_invalid_request')
        inline = op.action in {'replace', 'insert', 'delete', 'format', 'comment'}
        if inline:
            island = _check_range(entry, op, _islands(entry))
            if op.locator_id in structural or any(max(a, op.start) <= min(b, op.end) for a, b in occupied.get(op.locator_id, [])):
                raise GuardError('professional_overlapping_changes')
            occupied.setdefault(op.locator_id, []).append((op.start, op.end))
            preview = entry['text'][:op.start] + (op.quote if op.action in {'format', 'comment'} else op.text) + entry['text'][op.end:]
            _protect(entry['text'], preview, request)
        else:
            same_anchor_insert = op.action in {'insert_before', 'insert_after'} and all(
                previous.action in {'insert_before', 'insert_after'} and previous.text != op.text
                for previous in insertions.get(op.locator_id, []))
            if (op.locator_id in structural and not same_anchor_insert) or op.locator_id in occupied or request.mode != 'clean':
                raise GuardError('professional_overlapping_changes')
            structural.add(op.locator_id)
            insertions.setdefault(op.locator_id, []).append(op)
            island = None
            if op.action == 'paragraph_delete':
                _protect(entry['text'], '', request)
            if op.action in {'insert_before', 'insert_after'}:
                _protect('', op.text, request)
            if op.action == 'row_delete':
                _protect(_text(entry['node'].getparent().getparent()), '', request)
        prepared.append((entry, op, island))
        before.append({'locator_id': op.locator_id, 'text': op.quote if inline else entry['text'], 'location': entry['location']})
        if op.action == 'row_delete':
            before[-1]['removed_row_text'] = _text(entry['node'].getparent().getparent())
        if op.action == 'merge_next':
            before[-1]['next_paragraph_text'] = _text(entry['node'].getnext()) if entry['node'].getnext() is not None else None
        after.append({'locator_id': op.locator_id, 'action': op.action, 'text': op.text, 'reason': op.reason,
                      'inheritance': op.inheritance, 'character_format': op.character_format.model_dump(exclude_none=True) if op.character_format else None,
                      'paragraph_format': op.paragraph_format.model_dump(exclude_none=True) if op.paragraph_format else None})
        if op.action.startswith('revision_'):
            after[-1]['reviewed_revisions'] = [row['id'] for row in _revisions(entry) if row['id'] in ([op.revision_id] + op.revision_ids)]
    guard_targets(root, fields, prepared)
    reference_impact = impact(root, entries, fields, request.operations, request.source_sha256)
    for entry, op, _ in prepared:
        if op.action == 'merge_next':
            following = entry['node'].getnext()
            other = next((row for row in entries if row['node'] is following), None)
            if other and (other['locator_id'] in occupied or other['locator_id'] in structural):
                raise GuardError('professional_overlapping_changes')
        if op.action in {'row_insert', 'row_delete'}:
            table = entry['node'].getparent().getparent().getparent()
            if any(other is not entry and table in list(other['node'].iterancestors()) for other, _, _ in prepared):
                raise GuardError('professional_overlapping_changes')
    from .writes import _REVISION_ALL_MARKS
    revision_tags = {W + mark for mark in _REVISION_ALL_MARKS}
    next_revision = [max([int(n.get(W + 'id')) for n in root.iter() if n.tag in revision_tags and (n.get(W + 'id') or '').isdigit()] + [-1]) + 1]
    order = sorted(enumerate(prepared), key=lambda pair: (pair[1][0]['location'], pair[1][1].start,
                   pair[0] if pair[1][1].action == 'insert_after' else -pair[0]), reverse=True)
    for _, (entry, op, island) in order:
        if island is not None:
            island = next((row for row in _islands(entry) if row['container'] is island['container'] and
                           row['start'] <= op.start <= op.end <= row['end']), None)
            if island is None:
                raise GuardError('professional_protected_boundary')
            if op.action == 'comment':
                _comment(entry, op, island, len(comments))
                comments.append(op)
            else:
                _inline(entry, op, island, request, next_revision)
        else:
            if op.action in {'list_continue', 'list_restart'}:
                effect = _numbering(entry, op, targets, parts)
                next(row for row in after if row['locator_id'] == op.locator_id)['numbering_impact'] = effect
            else:
                _structure(entry, op, entries, parts)
    if request.editing_profile != 'general':
        from .technical_profile import verify as verify_profile
        for entry in {entry['location']: entry for entry, _, _ in prepared}.values():
            verify_profile(entry['text'], _text(entry['node']), request.editing_profile, request.profile_rule_sha256)
    changed = dict(parts)
    if comments:
        _comment_parts(changed, comments, request.revision_date)
    changed['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    _inventory(changed, request.source_sha256)
    if _canonical(root) == _canonical(ET.fromstring(parts['word/document.xml'])):
        raise GuardError('professional_no_changes')
    if request.mode == 'tracked':
        from .tracked_edits import _coalesced
        approved, _, _ = transform(parts, request.model_copy(update={'mode': 'clean'}))
        original_ids = {n.get(W + 'id') for n in ET.fromstring(parts['word/document.xml']).iter() if n.tag in {W + 'ins', W + 'del'}}
        for accept, expected in ((True, approved), (False, parts)):
            projected = deepcopy(root)
            for wrapper in list(projected.iter()):
                if wrapper.tag not in {W + 'ins', W + 'del'} or wrapper.get(W + 'id') in original_ids:
                    continue
                parent = wrapper.getparent()
                index = parent.index(wrapper)
                if (wrapper.tag == W + 'ins') == accept:
                    for run in list(wrapper):
                        for text in run.findall(W + 'delText'):
                            text.tag = W + 't'
                        parent.insert(index, run)
                        index += 1
                parent.remove(wrapper)
            if _coalesced(projected) != _coalesced(ET.fromstring(expected['word/document.xml'])):
                raise GuardError('professional_revision_projection_failed')
    after[0]['verification'] = {'requested': len(prepared), 'derived': ['run splitting and xml:space for exact ranges', 'property ordering; requested comment parts or independent numbering instance where applicable'],
                               'unexplained': [], 'preserved_parts': sum(original_parts[name] == changed[name] for name in original_parts),
                               'changed_parts': [name for name in changed if changed[name] != original_parts.get(name)],
                               'styles': 'direct formatting changed only as requested; inherited styles, themes and definitions byte-preserved',
                               'mode': request.mode, 'existing_revisions': 'non-target wrappers preserved; no global accept/reject',
                               'reference_impact': reference_impact,
                               'fields': {'preserved': len(fields), 'cross_references': sum(row['target'] is not None for row in fields),
                                          'results_refreshed': False, 'pagination_may_change': True,
                                          'warning': 'Existing field results are preserved, not verified or refreshed; PAGEREF/PAGE/TOC may be stale after reflow. Review in Word before final delivery.'},
                               'coverage': 'document.xml exact planned transform; all other package bytes unchanged; not full OOXML semantic or Word visual proof'}
    if request.editing_profile != 'general':
        from .technical_profile import report
        texts = [entry['text'] for entry in {entry['location']: entry for entry, _, _ in prepared}.values()]
        applied = [literal for literal in request.protected_literals if any(literal in text for text in texts)]
        after[0]['verification']['technical_profile'] = {**report(texts, applied), 'all_modified_paragraphs_checked': True}
    if sum(len(json.dumps(value, ensure_ascii=False)) for value in before + after) > 60000:
        raise GuardError('professional_capacity')
    return changed, before, after


def verify(before, after, request):
    expected, _, _ = transform(before, request)
    if set(expected) != set(after) or any(expected[name] != after[name] for name in expected if name != 'word/document.xml'):
        raise GuardError('professional_non_target_part_changed')
    if _canonical(ET.fromstring(expected['word/document.xml'])) != _canonical(ET.fromstring(after['word/document.xml'])):
        raise GuardError('professional_non_target_structure_changed')


def now():
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
