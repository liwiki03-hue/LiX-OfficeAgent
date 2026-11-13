from copy import deepcopy
from difflib import SequenceMatcher
import json
from typing import Literal

from lxml import etree as ET
from pydantic import Field

from .professional_edits import W, _inventory, _canonical, _text, _protect, _hash_text
from .range_edits import RangeDocx, RangeOperation, _ordinary, _islands, _inline, _boundary
from .range_fields import analyze, impact
from .writes import Strict, GuardError
from .rag import SECRET


class BlockOperation(Strict):
    action: Literal['paragraphs_replace', 'paragraphs_merge', 'chapter_move', 'bookmark_text']
    locator_id: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    locator_ids: list[str] = Field(min_length=1, max_length=12)
    destination_locator: str | None = Field(default=None, pattern=r'^loc:[a-f0-9]{64}$')
    placement: Literal['before', 'after'] = 'before'
    new_texts: list[str] = Field(default_factory=list, max_length=6)
    start: int = Field(default=0, ge=0, le=8000)
    end: int | None = Field(default=None, ge=0, le=8000)
    quote: str = Field(default='', max_length=8000)
    text: str = Field(default='', max_length=2000)
    reason: str = Field(min_length=1, max_length=400)
    inheritance: Literal['reject_mixed', 'start', 'distributed'] = 'reject_mixed'
    bookmark_name: str | None = Field(default=None, pattern=r'^[A-Za-z_][A-Za-z0-9_]{0,39}$')
    acknowledge_stale_fields: bool = False


class BlockIntent(Strict):
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    operations: list[BlockOperation] = Field(min_length=1, max_length=4)
    mode: Literal['clean'] = 'clean'
    preserve_numbers: bool = True
    preserve_citations: bool = True
    preserve_formulas: bool = True
    protected_literals: list[str] = Field(default_factory=list, max_length=30)


class BlockDocx(RangeDocx):
    kind: Literal['block_docx'] = 'block_docx'
    operations: list[BlockOperation] = Field(min_length=1, max_length=4)
    mode: Literal['clean'] = 'clean'


def _heading(entry):
    import re
    match = re.fullmatch(r'Heading([1-9])', entry['style'] or '', re.I)
    return int(match[1]) if match else None


def _chapter(entry, entries, body):
    if not _heading(entry) or entry['node'].getparent() is not body:
        raise GuardError('professional_block_heading')
    begin = body.index(entry['node'])
    stop = len(body) - 1 if len(body) and body[-1].tag == W + 'sectPr' else len(body)
    for other in entries:
        if other['node'].getparent() is body and body.index(other['node']) > begin and _heading(other) and _heading(other) <= _heading(entry):
            stop = body.index(other['node'])
            break
    return list(body)[begin:stop]


def _edit(entry, start, end, replacement, op, request):
    old = _text(entry['node'])
    if not 0 <= start <= end <= len(old) or not _boundary(old, start) or not _boundary(old, end):
        raise GuardError('professional_invalid_range')
    if len(replacement) > 2000:
        raise GuardError('professional_capacity')
    _protect(old[start:end], replacement, request)
    for kind, a, b, c, d in reversed(SequenceMatcher(None, old[start:end], replacement, autojunk=False).get_opcodes()):
        if kind == 'equal':
            continue
        left, right = start + a, start + b
        islands = [island for island in _islands(entry) if island['start'] <= left <= right <= island['end']]
        if len(islands) != 1:
            raise GuardError('professional_protected_boundary')
        change = RangeOperation(locator_id=entry['locator_id'], expected_text_sha256=entry['text_sha256'],
            action='insert' if left == right else 'delete' if c == d else 'replace', island_id=islands[0]['id'], start=left, end=right,
            quote=old[left:right], text=replacement[c:d], reason=op.reason, inheritance=op.inheritance)
        _inline(entry, change, islands[0], request, [0])


def transform(parts, request):
    if request.editing_profile != 'general':
        raise GuardError('professional_profile_operation')
    root, entries = _inventory(parts, request.source_sha256)
    _, fields = analyze(parts)
    body = root.find(W + 'body')
    targets = {entry['locator_id']: entry for entry in entries}
    occupied, before, after = set(), [], []
    original_locations = {entry['node']: entry['location'] for entry in entries}
    for op in request.operations:
        if SECRET.search(op.model_dump_json()) or len(set(op.locator_ids)) != len(op.locator_ids) or op.locator_id != op.locator_ids[0] or any(key not in targets for key in op.locator_ids):
            raise GuardError('professional_stale_locator')
        selected = [targets[key] for key in op.locator_ids]
        nodes = [entry['node'] for entry in selected]
        if occupied.intersection(nodes):
            raise GuardError('professional_overlapping_changes')
        occupied.update(nodes)
        if len(occupied) > 12:
            raise GuardError('professional_capacity')
        before.append({'location': selected[0]['location'], 'locations': [entry['location'] for entry in selected],
                       'texts': [_text(node) for node in nodes], 'action': op.action})
        if op.action == 'bookmark_text':
            if len(selected) != 1 or not op.acknowledge_stale_fields or len(op.new_texts) != 1:
                raise GuardError('professional_block_reference_acknowledgment')
            node = nodes[0]
            matching = [row for row in fields if row['target'] == op.bookmark_name]
            starts = [n for n in node.findall(W + 'bookmarkStart') if n.get(W + 'name') == op.bookmark_name]
            if not matching or len(starts) != 1 or any(n.tag not in {W + 'pPr', W + 'r', W + 'bookmarkStart', W + 'bookmarkEnd'} for n in node):
                raise GuardError('professional_block_bookmark_scope')
            ends = [n for n in node.findall(W + 'bookmarkEnd') if n.get(W + 'id') == starts[0].get(W + 'id')]
            if len(ends) != 1 or len(node.findall(W + 'bookmarkStart')) != 1 or node.find('./' + W + 'pPr/' + W + 'sectPr') is not None:
                raise GuardError('professional_block_bookmark_scope')
            if any(child.tag not in {W + 'rPr', W + 't'} for run in node.findall(W + 'r') for child in run):
                raise GuardError('professional_block_bookmark_scope')
            left = sum(len(_text(n)) for n in list(node)[:node.index(starts[0])])
            right = sum(len(_text(n)) for n in list(node)[:node.index(ends[0])])
            if left == right:
                raise GuardError('professional_block_bookmark_scope')
            _edit(selected[0], left, right, op.new_texts[0], op, request)
        else:
            if any(node.getparent() is not body for node in nodes) or any(nodes[i].getnext() is not nodes[i + 1] for i in range(len(nodes) - 1)):
                raise GuardError('professional_block_contiguous')
            for entry in selected:
                _ordinary(entry, parts)
            if op.action == 'chapter_move':
                if any(node.find('./' + W + 'pPr/' + W + 'sectPr') is not None for node in body.findall(W + 'p')):
                    raise GuardError('professional_block_section_boundary')
                if _chapter(selected[0], entries, body) != nodes:
                    raise GuardError('professional_block_incomplete_chapter')
                destination = targets.get(op.destination_locator)
                if not destination or destination['node'] in occupied or _heading(destination) != _heading(selected[0]):
                    raise GuardError('professional_block_destination')
                _ordinary(destination, parts)
                anchor = destination['node']
                if op.placement == 'after':
                    destination_nodes = _chapter(destination, entries, body)
                    if any(node in nodes for node in destination_nodes):
                        raise GuardError('professional_block_destination')
                    anchor = destination_nodes[-1]
                for node in nodes:
                    body.remove(node)
                position = body.index(anchor) + (op.placement == 'after')
                for node in nodes:
                    body.insert(position, node)
                    position += 1
            elif op.action == 'paragraphs_merge':
                if len(nodes) < 2 or len(nodes) > 6 or len(op.new_texts) != 1 or op.start or op.end is not None:
                    raise GuardError('professional_block_boundary_policy')
                properties = [_canonical(n.find(W + 'pPr')) if n.find(W + 'pPr') is not None else b'' for n in nodes]
                if len(set(properties)) != 1:
                    raise GuardError('professional_block_style_conflict')
                combined = ' '.join(_text(node) for node in nodes)
                _protect(combined, op.new_texts[0], request)
                for node in nodes[1:]:
                    run = ET.SubElement(nodes[0], W + 'r')
                    ET.SubElement(run, W + 't').text = ' '
                    for run in list(node.findall(W + 'r')):
                        nodes[0].append(run)
                    body.remove(node)
                _edit(selected[0], 0, len(combined), op.new_texts[0], op, request)
            else:
                if not 2 <= len(selected) <= 6 or len(op.new_texts) != len(selected):
                    raise GuardError('professional_block_boundary_policy')
                for index, (entry, text) in enumerate(zip(selected, op.new_texts)):
                    start = op.start if index == 0 else 0
                    end = op.end if index == len(selected) - 1 and op.end is not None else len(entry['text'])
                    _edit(entry, start, end, text, op, request)
        after.append({'location': selected[0]['location'], 'action': op.action, 'texts': op.new_texts,
                      'destination_locator': op.destination_locator, 'placement': op.placement, 'reason': op.reason,
                      'destination_location': targets[op.destination_locator]['location'] if op.destination_locator else None,
                      'mode': 'clean', 'native_structure_revision': False})
    updated = dict(parts)
    updated['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    analyze(updated)
    if _canonical(root) == _canonical(ET.fromstring(parts['word/document.xml'])):
        raise GuardError('professional_no_changes')
    current_root, current_entries = _inventory(updated, request.source_sha256)
    by_path = {current_root.getroottree().getpath(entry['node']): entry['location'] for entry in current_entries}
    mapping = {location: by_path.get(root.getroottree().getpath(node)) if node.getparent() is not None else None for node, location in original_locations.items()}
    original_root, original_entries = _inventory(parts, request.source_sha256)
    after[0]['verification'] = {'non_target_parts_preserved': True, 'position_map': mapping,
                              'reference_impact': impact(original_root, original_entries, fields, request.operations, request.source_sha256),
                              'stored_fields_refreshed': False, 'native_structure_revision': False,
                              'coverage': 'exact deterministic operation replay and package preservation; no dynamic Word field or semantic proof'}
    if len(json.dumps([before, after], ensure_ascii=False)) > 60000:
        raise GuardError('professional_capacity')
    return updated, before, after


def render_anchors(parts, request):
    _, entries = _inventory(parts, request.source_sha256)
    updated, before, after = transform(parts, request)
    _, current = _inventory(updated, request.source_sha256)
    mapping = after[0]['verification']['position_map']
    by_id = {entry['locator_id']: entry for entry in entries}
    by_location = {entry['location']: entry for entry in current}
    anchors = []
    for op in request.operations:
        entry = by_id[op.locator_id]
        target = by_location.get(mapping.get(entry['location']))
        anchors.append({'location': op.locator_id, 'reason': op.reason, 'action': op.action, 'before': entry['text'][:200],
                        'after': target['text'][:200] if target else '', 'contexts': {
                            'before': {'text': entry['text'][:200]}, 'after': {'text': target['text'][:200] if target else ''}}})
    return anchors


def verify(before, after, request):
    expected, _, _ = transform(before, request)
    if set(expected) != set(after) or any(expected[name] != after[name] for name in expected if name != 'word/document.xml'):
        raise GuardError('professional_non_target_part_changed')
    if _canonical(ET.fromstring(expected['word/document.xml'])) != _canonical(ET.fromstring(after['word/document.xml'])):
        raise GuardError('professional_non_target_structure_changed')
