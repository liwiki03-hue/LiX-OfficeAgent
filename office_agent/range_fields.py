import re

from lxml import etree as ET

from .professional_edits import W, _text, _canonical
from .provenance import identifier
from .rag import SECRET
from .writes import GuardError


def _instruction(value):
    value = value.strip()
    if value in {'PAGE', 'NUMPAGES', 'TOC \\o "1-3" \\h \\z \\u'}:
        return value.split()[0], None
    match = re.fullmatch(r'(REF|PAGEREF)\s+([A-Za-z_][A-Za-z0-9_]{0,39})((?:\s+\\[hp])*)', value)
    if match:
        kind, target, switches = match.groups()
        flags = switches.split()
        if len(flags) == len(set(flags)) and (kind == 'PAGEREF' or '\\p' not in flags) and not SECRET.search(target):
            return kind, target
    raise GuardError('professional_field_scope')


def analyze(parts):
    rows = []
    main = ET.fromstring(parts['word/document.xml'])
    for name, data in parts.items():
        if not name.startswith('word/') or not name.endswith('.xml'):
            continue
        root = main if name == 'word/document.xml' else ET.fromstring(data)
        paragraphs = list(root.iter(W + 'p'))
        positions = {node: index for index, node in enumerate(paragraphs)}
        stack = []

        def paragraph(node):
            return next((parent for parent in node.iterancestors(W + 'p')), None)

        def record(instruction, node, end, nested=False, result=True):
            kind, target = _instruction(instruction)
            if target and (name != 'word/document.xml' or paragraph(node) is not paragraph(end) or nested or not result):
                raise GuardError('professional_field_scope')
            rows.append({'part': name, 'type': kind, 'target': target,
                         'paragraph': positions.get(paragraph(node)), 'node': node})
            if len(rows) > 200:
                raise GuardError('professional_capacity')

        for node in root.iter():
            if node.tag == W + 'fldSimple':
                if stack or any(n.tag in {W + 'fldSimple', W + 'fldChar', W + 'instrText'} for n in node.iterdescendants()):
                    raise GuardError('professional_field_scope')
                record(node.get(W + 'instr', ''), node, node)
            elif node.tag == W + 'fldChar':
                kind = node.get(W + 'fldCharType')
                if kind == 'begin':
                    if len(stack) >= 16:
                        raise GuardError('professional_capacity')
                    if stack:
                        stack[-1]['nested'] = True
                    stack.append({'node': node, 'code': [], 'result': False, 'nested': False})
                elif kind == 'separate' and stack and not stack[-1]['result']:
                    stack[-1]['result'] = True
                elif kind == 'end' and stack:
                    field = stack.pop()
                    record(''.join(field['code']), field['node'], node, field['nested'], field['result'])
                else:
                    raise GuardError('professional_field_scope')
            elif node.tag == W + 'instrText':
                if not stack or stack[-1]['result']:
                    raise GuardError('professional_field_scope')
                stack[-1]['code'].append(node.text or '')
                if sum(map(len, stack[-1]['code'])) > 2000:
                    raise GuardError('professional_capacity')
        if stack:
            raise GuardError('professional_field_scope')
    paragraphs = list(main.iter(W + 'p'))
    ordered = {node: index for index, node in enumerate(main.iter())}
    for row in rows:
        if not row['target']:
            continue
        starts = [n for n in main.iter(W + 'bookmarkStart') if n.get(W + 'name') == row['target']]
        if len(starts) != 1 or not (starts[0].get(W + 'id') or '').isdigit():
            raise GuardError('professional_reference_target')
        start = starts[0]
        identity = start.get(W + 'id')
        if sum(n.get(W + 'id') == identity for n in main.iter(W + 'bookmarkStart')) != 1:
            raise GuardError('professional_reference_target')
        ends = [n for n in main.iter(W + 'bookmarkEnd') if n.get(W + 'id') == identity]
        parent = next(start.iterancestors(W + 'p'), None)
        if len(ends) != 1 or parent is None or next(ends[0].iterancestors(W + 'p'), None) is not parent or ordered[start] >= ordered[ends[0]]:
            raise GuardError('professional_reference_target')
        if SECRET.search(_text(parent)):
            raise GuardError('professional_sensitive_or_control_text')
        row['target_paragraph'] = paragraphs.index(parent)
    return main, rows


def public_fields(rows, paragraph_index):
    return [{'type': row['type'], 'target': row['target'], 'editable': False,
             'stored_result_only': True, 'target_locked': row['target'] is not None}
            for row in rows if row['part'] == 'word/document.xml' and row['paragraph'] == paragraph_index]


def guard_targets(root, rows, prepared):
    paragraphs = list(root.iter(W + 'p'))
    locked = {paragraphs[row['target_paragraph']] for row in rows if row.get('target_paragraph') is not None}
    for entry, operation, _ in prepared:
        node = entry['node']
        if operation.action == 'merge_next' and node.getnext() in locked:
            raise GuardError('professional_reference_target_locked')
        if operation.action == 'row_delete' and any(n in locked for n in node.getparent().getparent().iter(W + 'p')):
            raise GuardError('professional_reference_target_locked')
        if node in locked and operation.action not in {'comment', 'format', 'paragraph_format', 'insert_before', 'insert_after'}:
            raise GuardError('professional_reference_target_locked')


def impact(root, entries, rows, operations, version):
    paragraphs = list(root.iter(W + 'p'))
    registered = {paragraphs.index(entry['node']): entry for entry in entries}
    changes = {op.locator_id: op.action for op in operations}
    paginating = any(op.action != 'comment' for op in operations)
    records = []
    for index, row in enumerate(rows):
        field = registered.get(row['paragraph']) if row['part'] == 'word/document.xml' else None
        target = registered.get(row.get('target_paragraph'))
        nearby = field is not None and field['locator_id'] in changes
        target_changed = target is not None and target['locator_id'] in changes
        reasons = []
        if nearby:
            reasons.append('reference_surroundings')
        if target_changed:
            reasons.append('target_presentation')
        if paginating and row['type'] in {'PAGEREF', 'PAGE', 'NUMPAGES', 'TOC'}:
            reasons.append('possible_pagination')
        records.append({'field_id': 'field:' + identifier(version, row['part'], str(index), _canonical(row['node']).hex()),
            'type': row['type'], 'target_name': row['target'], 'part': row['part'],
            'field_locator': field['locator_id'] if field else None, 'field_location': field['location'] if field else None,
            'target_locator': target['locator_id'] if target else None, 'target_location': target['location'] if target else None,
            'reasons': reasons or ['no_direct_intersection'], 'certainty': 'structural_dependency_only',
            'result_refreshed': False, 'needs_review': bool(reasons),
            'scope_complete': field is not None and (row['target'] is None or target is not None)})
    return {'source_sha256': version, 'fields_total': len(rows), 'items': records[:40],
            'truncated': len(records) > 40, 'needs_review_total': sum(r['needs_review'] for r in records),
            'complete_dependency_graph': False, 'coverage': 'admitted package fields; registered document paragraphs only; no external resources',
            'dynamic_word_pagination_verified': False}
