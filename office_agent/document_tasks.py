from collections import Counter
from copy import deepcopy
import json
import re
import time
import uuid

from lxml import etree as ET
from pydantic import Field
from typing import Literal

from .advanced_writes import _parts
from .professional_edits import W, _inventory, _canonical, _hash_text
from .range_edits import RangeDocx, RangeOperation, ParagraphFormat, _islands, transform, now
from .range_fields import analyze
from .rag import SECRET
from .writes import Strict, GuardError, digest
from .provenance import identifier
from .block_edits import BlockOperation, BlockDocx

MAX_ITEMS = 96
BLOCK_ACTIONS = {'paragraphs_replace', 'paragraphs_merge', 'chapter_move', 'bookmark_text'}


def extra_text_regions(xml):
    for namespace, tag in (('http://schemas.openxmlformats.org/drawingml/2006/main', 'p'),
                           ('http://schemas.openxmlformats.org/officeDocument/2006/math', 'oMath')):
        for index, node in enumerate(xml.iter('{' + namespace + '}' + tag)):
            yield tag + ':' + str(index), ''.join(n.text or '' for n in node.iter('{' + namespace + '}t'))


class TaskRequirement(Strict):
    key: str = Field(pattern=r'^[a-z][a-z0-9_-]{0,30}$')
    goal: str = Field(min_length=1, max_length=400)
    kind: Literal['terminology', 'format', 'ranges', 'blocks']
    scope: Literal['document', 'chapter', 'selected'] = 'document'
    section_locator: str | None = Field(default=None, pattern=r'^loc:[a-f0-9]{64}$')
    locators: list[str] = Field(default_factory=list, max_length=96)
    find: str = Field(default='', max_length=120)
    replacement: str = Field(default='', max_length=200)
    paragraph_format: ParagraphFormat | None = None
    operations: list[RangeOperation] = Field(default_factory=list, max_length=96)
    block_operations: list[BlockOperation] = Field(default_factory=list, max_length=4)
    acceptance: str = Field(min_length=1, max_length=400)


class TaskSpec(Strict):
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    request: str = Field(min_length=1, max_length=2000)
    requirements: list[TaskRequirement] = Field(min_length=1, max_length=8)
    ambiguities: list[str] = Field(default_factory=list, max_length=8)
    protected_literals: list[str] = Field(default_factory=list, max_length=30)
    mode: Literal['clean', 'tracked'] = 'clean'
    preserve_numbers: bool = True
    preserve_citations: bool = True
    preserve_formulas: bool = True
    editing_profile: Literal['general', 'technical_report_v1'] = 'general'
    profile_rule_sha256: str | None = Field(default=None, pattern=r'^[a-f0-9]{64}$')


class NavigationArgs(Strict):
    query: str = Field(default='', max_length=120)
    offset: int = Field(default=0, ge=0, le=4000)
    limit: int = Field(default=20, ge=1, le=20)


class RewriteItem(Strict):
    locator_id: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    new_text: str = Field(min_length=1, max_length=2000)
    reason: str = Field(min_length=1, max_length=400)


class RewriteBatch(Strict):
    items: list[RewriteItem] = Field(default_factory=list, max_length=6)


def load(editor, path):
    source = editor.space.file(str(path))
    version = digest(source)
    parts = _parts(source, range_edit=True)
    root, entries = _inventory(parts, version)
    if digest(source) != version:
        raise GuardError('professional_source_changed')
    return version, parts, root, entries


def level(entry):
    match = re.fullmatch(r'Heading([1-9])', entry['style'] or '', re.I)
    return int(match[1]) if match else None


def navigation(editor, path, query='', offset=0, limit=20):
    args = NavigationArgs(query=query, offset=offset, limit=limit)
    if SECRET.search(query):
        raise GuardError('professional_sensitive_or_control_text')
    version, parts, root, entries = load(editor, path)
    matches = [entry for entry in entries if not args.query or args.query in entry['text']]
    rows = []
    for entry in matches[offset:offset + limit]:
        sensitive = SECRET.search(entry['text']) or len(entry['text']) > 8000
        position = entry['text'].find(query) if query else 0
        rows.append({'locator_id': entry['locator_id'], 'location': entry['location'], 'text_sha256': entry['text_sha256'],
                     'excerpt': '' if sensitive else entry['text'][max(0, position - 50):position + 200],
                     'style': entry['style'], 'read_only_reason': 'sensitive_or_oversize' if sensitive else entry['refusal'],
                     'occurrences': entry['text'].count(query) if query and not sensitive else None})
    outline = [{'locator_id': entry['locator_id'], 'location': entry['location'], 'level': level(entry), 'title': entry['text'][:120]}
               for entry in entries if level(entry) and not SECRET.search(entry['text'])]
    _, fields = analyze(parts)
    paragraphs = list(root.iter(W + 'p'))
    bookmarks = []
    for field in fields:
        target = next((entry for entry in entries if field.get('target_paragraph') is not None and entry['node'] is paragraphs[field['target_paragraph']]), None)
        if target and not any(row['name'] == field['target'] for row in bookmarks):
            bookmarks.append({'name': field['target'], 'locator_id': target['locator_id'], 'location': target['location']})
    return {'source_sha256': version, 'outline': outline[:40], 'outline_total': len(outline), 'outline_truncated': len(outline) > 40, 'reference_targets': bookmarks[:40],
            'registered_targets': len(entries), 'matching_targets': len(matches), 'items': rows,
            'next_offset': offset + len(rows) if offset + len(rows) < len(matches) else None,
            'statistics': {'tables': len(list(root.iter(W + 'tbl'))), 'revisions': sum(1 for n in root.iter() if n.tag in {W + 'ins', W + 'del'}),
                           'unsupported_registered': sum(bool(e['refusal']) for e in entries)},
            'coverage': 'bounded main-document paragraphs and physical cells; header/footer/note text separately checked by task compilation',
            'scan_complete': True, 'read_only': True, 'rendered_page_numbers': False}


def natural_spec(editor, path, request):
    if type(request) is not str or not 1 <= len(request) <= 2000 or SECRET.search(request):
        raise GuardError('professional_task_request')
    version = digest(editor.space.file(str(path)))
    match = re.fullmatch(r'\s*(?:将|把)(?:全文|全篇)(?:所有)?[“「"]([^”」"]+)[”」"](?:统一)?(?:替换|改)(?:成|为)[“「"]([^”」"]+)[”」"](?:[，,].*)?\s*', request)
    if not match:
        raise GuardError('professional_task_needs_specification')
    tail = request[match.end(2) + 1:].strip()
    if tail and tail not in {'。', '.'} and not re.fullmatch(r'[，,](?:保留|保持)(?:数字|单位|公式|引文|引用|和|、|及|\s)+[。.]?', tail):
        raise GuardError('professional_task_needs_specification')
    return TaskSpec(source_sha256=version, request=request, requirements=[TaskRequirement(key='terminology', kind='terminology',
        goal='全篇精确术语一致性', find=match[1], replacement=match[2], acceptance='授权文档已扫描文本中旧术语无残留；受保护或未支持位置单独列出')])


def selected(requirement, entries):
    if requirement.scope == 'document':
        return entries
    by_id = {entry['locator_id']: entry for entry in entries}
    if requirement.scope == 'selected':
        if not requirement.locators or len(set(requirement.locators)) != len(requirement.locators) or any(key not in by_id for key in requirement.locators):
            raise GuardError('professional_task_scope')
        return [e for e in entries if e['locator_id'] in requirement.locators]
    heading = by_id.get(requirement.section_locator)
    if not heading or not level(heading) or '/table/' in heading['location']:
        raise GuardError('professional_task_scope')
    start = entries.index(heading)
    end = next((i for i in range(start + 1, len(entries)) if level(entries[i]) and level(entries[i]) <= level(heading)), len(entries))
    return entries[start:end]


def compile_task(editor, path, raw):
    spec = TaskSpec.model_validate(raw)
    if SECRET.search(spec.model_dump_json()) or len({r.key for r in spec.requirements}) != len(spec.requirements):
        raise GuardError('professional_task_request')
    version, parts, root, entries = load(editor, path)
    if version != spec.source_sha256:
        raise GuardError('professional_source_changed')
    from .technical_profile import RULE_SHA256, identity
    if spec.editing_profile != 'general' and spec.profile_rule_sha256 is None:
        spec = spec.model_copy(update={'profile_rule_sha256': RULE_SHA256})
    identity(spec.editing_profile, spec.profile_rule_sha256)
    items, coverage = [], []
    order = {entry['location']: i for i, entry in enumerate(entries)}
    for requirement in spec.requirements:
        scope = selected(requirement, entries)
        coverage.append({'requirement': requirement.key, 'scope': requirement.scope, 'scanned': len(scope), 'complete': True,
                         'acceptance': requirement.acceptance, 'locations': [entry['location'] for entry in scope] if requirement.scope != 'document' else None})
        candidates = []
        if requirement.kind == 'terminology':
            if not requirement.find or not requirement.replacement or requirement.find == requirement.replacement or requirement.find in requirement.replacement:
                raise GuardError('professional_task_terminology')
            for entry in scope:
                position = 0
                while True:
                    start = entry['text'].find(requirement.find, position)
                    if start < 0:
                        break
                    end = start + len(requirement.find)
                    islands = [island for island in _islands(entry) if island['start'] <= start < end <= island['end']]
                    candidates.append((entry, {'locator_id': entry['locator_id'], 'expected_text_sha256': entry['text_sha256'],
                        'action': 'replace', 'island_id': islands[0]['id'] if len(islands) == 1 else None, 'start': start, 'end': end,
                        'quote': requirement.find, 'text': requirement.replacement, 'reason': requirement.goal, 'inheritance': 'distributed' if len(requirement.find) == len(requirement.replacement) else 'reject_mixed'}))
                    position = end
                    if len(candidates) > MAX_ITEMS:
                        raise GuardError('professional_task_capacity')
        elif requirement.kind == 'format':
            if requirement.scope == 'document' or requirement.paragraph_format is None:
                raise GuardError('professional_task_explicit_format_scope')
            candidates = [(entry, {'locator_id': entry['locator_id'], 'expected_text_sha256': entry['text_sha256'], 'action': 'paragraph_format',
                         'paragraph_format': requirement.paragraph_format.model_dump(exclude_none=True), 'reason': requirement.goal}) for entry in scope]
        elif requirement.kind == 'blocks':
            by_id = {entry['locator_id']: entry for entry in scope}
            for operation in requirement.block_operations:
                if any(key not in by_id for key in operation.locator_ids):
                    raise GuardError('professional_task_scope')
                candidates.append((by_id[operation.locator_id], operation.model_dump()))
        else:
            by_id = {entry['locator_id']: entry for entry in scope}
            for operation in requirement.operations:
                if operation.locator_id not in by_id or operation.action not in {'replace', 'insert', 'delete', 'format', 'paragraph_format', 'comment', 'list_continue', 'list_restart'}:
                    raise GuardError('professional_task_operation_scope')
                candidates.append((by_id[operation.locator_id], operation.model_dump()))
            represented = {entry['locator_id'] for entry, _ in candidates}
            for entry in scope:
                if entry['locator_id'] not in represented:
                    items.append({'id': uuid.uuid4().hex, 'requirement': requirement.key, 'location': entry['location'],
                                  'operation': None, 'state': 'no_change' if level(entry) else 'restricted',
                                  'reason': 'heading_preserved' if level(entry) else 'professional_task_no_proposal',
                                  'receipt': None, 'position': order[entry['location']]})
        for entry, operation in candidates:
            item = {'id': uuid.uuid4().hex, 'requirement': requirement.key, 'location': entry['location'], 'operation': operation,
                    'state': 'pending', 'reason': '', 'receipt': None, 'position': order[entry['location']]}
            try:
                if operation['action'] == 'replace' and operation.get('quote') == operation.get('text') and entry['text'][operation.get('start', 0):operation.get('end', 0)] == operation.get('quote'):
                    item.update(state='no_change', reason='exact_text_already_present')
                    items.append(item)
                    continue
                if operation['action'] in BLOCK_ACTIONS:
                    operation['quote'] = ' ¶ '.join(row['text'] for row in entries if row['locator_id'] in operation['locator_ids'])[:200]
                    operation['text'] = ' ¶ '.join(operation['new_texts'])[:200] if operation['action'] != 'chapter_move' else '移动至宿主目标章节' + operation['placement']
                    if spec.mode != 'clean':
                        raise GuardError('professional_revision_scope')
                    from .block_edits import transform as block_transform
                    request = BlockDocx(path=str(path), source_sha256=version, operations=[operation], revision_date=now(), **spec.model_dump(include={'mode', 'preserve_numbers', 'preserve_citations', 'preserve_formulas', 'protected_literals', 'editing_profile', 'profile_rule_sha256'}))
                    block_transform(parts, request)
                else:
                    request = RangeDocx(path=str(path), source_sha256=version, operations=[operation], revision_date=now(), **spec.model_dump(include={'mode', 'preserve_numbers', 'preserve_citations', 'preserve_formulas', 'protected_literals', 'editing_profile', 'profile_rule_sha256'}))
                    transform(parts, request)
            except GuardError as error:
                item['state'] = 'no_change' if str(error) == 'professional_no_changes' else 'restricted'
                item['reason'] = str(error).split(':')[0]
            items.append(item)
        if requirement.kind == 'terminology' and requirement.scope == 'document':
            registered = {entry['node'] for entry in entries}
            for name, data in parts.items():
                if not name.startswith('word/') or not name.endswith('.xml'):
                    continue
                xml = root if name == 'word/document.xml' else ET.fromstring(data)
                for index, paragraph in enumerate(xml.iter(W + 'p')):
                    if name == 'word/document.xml' and paragraph in registered:
                        continue
                    count = ''.join(n.text or '' for n in paragraph.iter(W + 't')).count(requirement.find)
                    if count:
                        items.append({'id': uuid.uuid4().hex, 'requirement': requirement.key, 'location': name + '/unregistered-paragraph:' + str(index),
                                      'operation': None, 'state': 'restricted', 'reason': 'professional_task_unregistered_text',
                                      'receipt': None, 'position': len(entries), 'occurrences': count})
                for region, text in extra_text_regions(xml):
                    count = text.count(requirement.find)
                    if count:
                        items.append({'id': uuid.uuid4().hex, 'requirement': requirement.key, 'location': name + '/unsupported-' + region,
                                      'operation': None, 'state': 'restricted', 'reason': 'professional_task_unregistered_text',
                                      'receipt': None, 'position': len(entries), 'occurrences': count})
        if not candidates and not any(item['requirement'] == requirement.key for item in items):
            items.append({'id': uuid.uuid4().hex, 'requirement': requirement.key, 'location': None, 'operation': None,
                          'state': 'no_change', 'reason': 'no_matching_target', 'receipt': None, 'position': -1})
    if len(items) > MAX_ITEMS:
        raise GuardError('professional_task_capacity')
    for item in items:
        item['source_location'] = item['location']
    executable = [item for item in items if item['state'] == 'pending']
    for index, item in enumerate(executable):
        first = item['operation']
        if first['action'] not in {'replace', 'insert', 'delete'}:
            continue
        for other in executable[index + 1:]:
            second = other['operation']
            if item['location'] != other['location'] or second['action'] not in {'replace', 'insert', 'delete'}:
                continue
            if max(first['start'], second['start']) < min(first['end'], second['end']) or first['start'] == second['start']:
                raise GuardError('professional_task_conflicting_requirements')
    executable.sort(key=lambda item: (2 if item['operation']['action'] in BLOCK_ACTIONS else int(item['operation']['action'] in {'paragraph_format', 'comment', 'list_continue', 'list_restart'}), -item['position'], -item['operation'].get('start', 0)))
    groups, group, cost = [], [], 0
    for item in executable:
        entry_size = len(next(entry['text'] for entry in entries if entry['location'] == item['location']))
        item_cost = 2 * (entry_size + len(item['operation'].get('text', '')) + len(item['operation'].get('quote', '')) + 800)
        conflict = any(other['location'] == item['location'] and (other['operation']['action'] not in {'replace', 'insert', 'delete', 'format'} or item['operation']['action'] not in {'replace', 'insert', 'delete', 'format'}) for other in group)
        if len(group) == 12 or conflict or (group and (cost + item_cost > 24000 or item['operation']['action'] in BLOCK_ACTIONS or group[0]['operation']['action'] in BLOCK_ACTIONS)):
            groups.append([other['id'] for other in group])
            group, cost = [], 0
        group.append(item)
        cost += item_cost
    if group:
        groups.append([item['id'] for item in group])
    if len(groups) > 8:
        raise GuardError('professional_task_capacity')
    if len(path.name) + 15 * len(groups) > 220:
        raise GuardError('professional_task_filename_budget')
    profile_report = None
    if spec.editing_profile != 'general':
        from .technical_profile import report
        locations = {item['location'] for item in items if item['operation'] is not None}
        profile_report = report([entry['text'] for entry in entries if entry['location'] in locations], spec.protected_literals)
    task = {'id': uuid.uuid4().hex, 'spec': spec.model_dump(), 'source_sha256': version, 'current_sha256': version,
            'source_name': path.name, 'current_name': path.name, 'items': items, 'groups': groups, 'group_index': 0,
            'coverage': coverage, 'receipts': [], 'created': time.monotonic(), 'generation': 1,
            'anchor_map': {entry['locator_id']: entry['location'] for entry in entries},
            'expected': {entry['location']: {'text': entry['text_sha256'], 'node': _hash_text(_canonical(entry['node']).hex())} for entry in entries},
            'metrics': {'model_calls': 0, 'tool_calls': 0, 'input_tokens': 0, 'output_tokens': 0},
            'approval': None, 'uncertain': False, 'checks': [], 'profile_report': profile_report}
    check_postconditions(task, parts)
    return task


def check_postconditions(task, parts):
    _, entries = _inventory(parts, task['current_sha256'])
    checks = [check for check in task.get('checks', []) if check['kind'] != 'exact_old_term_absence']
    for requirement in task['spec']['requirements']:
        if requirement['kind'] != 'terminology':
            continue
        coverage = next(row for row in task['coverage'] if row['requirement'] == requirement['key'])
        if requirement['scope'] == 'document':
            paragraphs = [paragraph for name, data in parts.items() if name.startswith('word/') and name.endswith('.xml')
                          for paragraph in ET.fromstring(data).iter(W + 'p')]
            remaining = sum(''.join(n.text or '' for n in paragraph.iter(W + 't')).count(requirement['find']) for paragraph in paragraphs)
            remaining += sum(text.count(requirement['find']) for name, data in parts.items() if name.startswith('word/') and name.endswith('.xml')
                             for region, text in extra_text_regions(ET.fromstring(data)))
        else:
            remaining = sum(entry['text'].count(requirement['find']) for entry in entries if entry['location'] in coverage['locations'])
        checks.append({'requirement': requirement['key'], 'kind': 'exact_old_term_absence', 'remaining': remaining,
                       'passed': remaining == 0, 'coverage': 'OOXML paragraph w:t and registered DrawingML/math text regions; opaque/unknown parts and dynamic fields not semantically evaluated'})
    task['checks'] = checks


def group_request(editor, path, task):
    if task['uncertain'] or task['spec']['ambiguities']:
        raise GuardError('professional_task_unresolved')
    if len(path.name) + 15 > 220:
        raise GuardError('professional_task_filename_budget')
    version, parts, root, entries = load(editor, path)
    if version != task['current_sha256']:
        raise GuardError('professional_source_changed')
    current = {entry['location']: entry for entry in entries}
    if set(current) != set(task['expected']) or any(task['expected'][key]['node'] != _hash_text(_canonical(entry['node']).hex()) for key, entry in current.items()):
        raise GuardError('professional_task_mapping_changed')
    if task['group_index'] >= len(task['groups']):
        raise GuardError('professional_task_no_pending_group')
    ids = task['groups'][task['group_index']]
    rows = [next(item for item in task['items'] if item['id'] == key) for key in ids]
    operations = []
    for item in rows:
        if item['state'] != 'pending':
            continue
        entry = current[item['location']]
        op = deepcopy(item['operation'])
        op.update(locator_id=entry['locator_id'], expected_text_sha256=entry['text_sha256'])
        if op['action'] in BLOCK_ACTIONS:
            op.pop('expected_text_sha256')
            def remap(key):
                location = task['anchor_map'].get(key)
                if location not in current:
                    raise GuardError('professional_task_dependency')
                return current[location]['locator_id']
            op['locator_ids'] = [remap(key) for key in op['locator_ids']]
            if op.get('destination_locator'):
                op['destination_locator'] = remap(op['destination_locator'])
        if op['action'] in {'replace', 'insert', 'delete', 'format', 'comment'}:
            islands = [island for island in _islands(entry) if island['start'] <= op['start'] <= op['end'] <= island['end']]
            if len(islands) != 1:
                raise GuardError('professional_protected_boundary')
            op['island_id'] = islands[0]['id']
        if op.get('numbering_locator_id'):
            original = next((i for i in task['items'] if i['operation'] and i['operation']['locator_id'] == op['numbering_locator_id']), None)
            if not original or original['location'] not in current:
                raise GuardError('professional_task_dependency')
            op['numbering_locator_id'] = current[original['location']]['locator_id']
        operations.append(op)
    if not operations:
        raise GuardError('professional_task_no_pending_group')
    from .technical_profile import identity
    identity(task['spec'].get('editing_profile', 'general'), task['spec'].get('profile_rule_sha256'))
    fields = TaskSpec.model_validate(task['spec']).model_dump(include={'mode', 'preserve_numbers', 'preserve_citations', 'preserve_formulas', 'protected_literals', 'editing_profile', 'profile_rule_sha256'})
    if operations[0]['action'] in BLOCK_ACTIONS:
        from .block_edits import transform as block_transform
        request = BlockDocx(path=str(path), source_sha256=version, operations=operations, revision_date=now(), **fields)
        block_transform(parts, request)
    else:
        request = RangeDocx(path=str(path), source_sha256=version, operations=operations, revision_date=now(), **fields)
        transform(parts, request)
    return request


def public_task(task, offset=0):
    if type(offset) is not int or not 0 <= offset <= MAX_ITEMS:
        raise GuardError('professional_invalid_page')
    states = Counter(item['state'] for item in task['items'])
    requirements = [{key: value for key, value in row.items() if key not in {'operations', 'locators', 'block_operations'}} |
                    {'operation_count': len(row['operations']), 'block_operation_count': len(row['block_operations']), 'selected_count': len(row['locators'])} for row in task['spec']['requirements']]
    return {'task_id': task['id'], 'request': task['spec']['request'], 'requirements': requirements, 'mode': task['spec']['mode'],
            'protection_profile': task.get('profile_report') or {'profile': task['spec'].get('editing_profile', 'general'), 'rule_sha256': task['spec'].get('profile_rule_sha256'), 'items': [], 'manual_literals': task['spec']['protected_literals'], 'scope': 'stored protection policy; restored tasks require fresh plan'},
            'ambiguities': task['spec']['ambiguities'], 'source_sha256': task['source_sha256'], 'current_sha256': task['current_sha256'],
            'source_name': task['source_name'], 'current_name': task['current_name'], 'coverage': [{k: v for k, v in row.items() if k != 'locations'} for row in task['coverage']], 'states': dict(states),
            'return_context': {'baseline_name': task['source_name'], 'delivered_name': task['current_name'], 'receipts': list(task['receipts'])} if task['receipts'] else None,
            'items': [{key: value for key, value in item.items() if key not in {'position', 'operation'}} |
                      {'action': item['operation']['action'] if item['operation'] else None,
                       'before': item['operation'].get('quote', '')[:200] if item['operation'] else '',
                       'after': item['operation'].get('text', '')[:200] if item['operation'] else ''} for item in task['items'][offset:offset + 20]],
            'next_offset': offset + 20 if offset + 20 < len(task['items']) else None,
            'total_items': len(task['items']), 'groups_total': len(task['groups']), 'group_index': task['group_index'],
            'historical_receipts_only': bool(task['receipts']) and task['receipts'][-1] in task.get('historical_receipts', []),
            'group_sizes': [len(group) for group in task['groups']], 'receipts': task['receipts'], 'generation': task['generation'],
            'deterministic_complete': not task['uncertain'] and not task['spec']['ambiguities'] and all(item['state'] in {'completed', 'no_change'} for item in task['items']) and all(check['passed'] for check in task['checks']),
            'verification_checks': task['checks'],
            'semantic_review': 'not_proven; user acceptance still required', 'uncertain': task['uncertain'],
            'multi_transaction_atomic': False, 'metrics': task['metrics'],
            'metrics_scope': 'tool_calls counts group-plan calls only; token usage is provider-reported when available, not billing; full workflow telemetry is separate',
            'restores_approval': False}
