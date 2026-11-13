import hashlib
from io import BytesIO
import json
import zipfile

from docx import Document
from lxml import etree as ET

from .document_tasks import TaskSpec, load, natural_spec
from .professional_edits import W, _canonical, _hash_text, _protect
from .range_edits import RangeDocx, RangeOperation, _inline, _islands, now
from .tracked_edits import _coalesced
from .writes import GuardError, WordBatchReplace, WordTableCellReplace, _batch_replace, _replace_runs, _table_paragraph, digest


def package_projection(parts):
    return {name: hashlib.sha256(_coalesced(ET.fromstring(data)) if name == 'word/document.xml' else data).hexdigest()
            for name, data in parts.items()}


def verify_package(path, expected):
    from .advanced_writes import _parts
    if package_projection(_parts(path, range_edit=True)) != expected:
        raise GuardError('professional_route_projection_changed')


def candidate(editor, path, task):
    fallback = {'selected': 'bounded_ranges', 'reason': '保留精确选区、格式/编号、跨段或多需求的原有有界计划；未证明旧原语语义等价。',
                'candidates': [{'primitive': 'bounded_ranges', 'eligible': True}], 'covered_items': 0}
    spec = TaskSpec.model_validate(task['spec'])
    if spec.editing_profile != 'general':
        fallback['reason'] = '已选择技术报告字面保护；使用携带完整保护规则的选区事务。'
        return fallback, None
    if task['uncertain'] or spec.ambiguities:
        raise GuardError('professional_task_unresolved')
    if spec.mode != 'clean' or len(spec.requirements) != 1:
        return fallback, None
    requirement = spec.requirements[0]
    if requirement.kind != 'terminology' or requirement.scope != 'document':
        return fallback, None
    try:
        canonical = natural_spec(editor, path, spec.request)
    except GuardError:
        canonical = None
    if canonical is None or requirement != canonical.requirements[0] or not all((spec.preserve_numbers, spec.preserve_citations, spec.preserve_formulas)):
        fallback['reason'] = '请求或验收条件超出已证明等价的标准全文字面任务；保留原有完整需求和选区路径。'
        return fallback, None
    version, parts, root, entries = load(editor, path)
    if version != task['current_sha256']:
        raise GuardError('professional_source_changed')
    current = {entry['location']: entry for entry in entries}
    if set(current) != set(task['expected']) or any(task['expected'][key]['node'] != _hash_text(_canonical(entry['node']).hex()) for key, entry in current.items()):
        raise GuardError('professional_task_mapping_changed')
    pending = [item for item in task['items'] if item['state'] == 'pending']
    body = [item for item in pending if current[item['location']]['node'].getparent().tag == W + 'body']
    if not body and len(pending) != 1:
        return fallback, None
    rows = body or pending[:1]
    if not rows:
        return fallback, None
    pending_ids = {item['id'] for item in pending}
    selected_ids = {item['id'] for item in rows}
    fallback_transactions = sum(bool(set(group) & pending_ids) for group in task['groups'])
    routed_transactions = 1 + sum(bool(set(group) & (pending_ids - selected_ids)) for group in task['groups'])
    if routed_transactions > fallback_transactions:
        fallback['reason'] = '旧字面原语会增加独立事务及渲染次数，保留原有组合选区计划。'
        return fallback, None
    try:
        doc = Document(path)
        if body:
            actual = sum(paragraph.text.count(requirement.find) for paragraph in doc.paragraphs)
            if actual != len(body):
                raise GuardError('professional_route_scope_not_equivalent')
            request = WordBatchReplace(path=str(path), old=requirement.find, new=requirement.replacement, max_occurrences=actual)
            _batch_replace(doc, request)
        else:
            node = current[rows[0]['location']]['node']
            coordinates = [(ti, ri, ci, pi) for ti, table in enumerate(doc.tables)
                           for ri, row in enumerate(table.rows) for ci, cell in enumerate(row.cells)
                           for pi, paragraph in enumerate(cell.paragraphs) if paragraph._p.getroottree().getpath(paragraph._p) == node.getroottree().getpath(node)]
            if len(coordinates) != 1:
                raise GuardError('professional_route_scope_not_equivalent')
            ti, ri, ci, pi = coordinates[0]
            request = WordTableCellReplace(path=str(path), old=requirement.find, new=requirement.replacement,
                table_index=ti, row_index=ri, cell_index=ci, paragraph_index=pi)
            paragraph = _table_paragraph(doc, request)
            if paragraph.text.count(requirement.find) != 1:
                raise GuardError('professional_route_scope_not_equivalent')
            _replace_runs(paragraph, request.old, request.new)
        render_operations = []
        for item in sorted(rows, key=lambda item: (item['position'], item['operation']['start']), reverse=True):
            entry = current[item['location']]
            operation = RangeOperation.model_validate({**item['operation'], 'locator_id': entry['locator_id'], 'expected_text_sha256': entry['text_sha256']})
            islands = [island for island in _islands(entry) if island['start'] <= operation.start < operation.end <= island['end']]
            if len(islands) != 1:
                raise GuardError('professional_protected_boundary')
            if len(render_operations) < 12:
                render_operations.append(operation.model_copy(update={'island_id': islands[0]['id']}))
            _inline(entry, operation, islands[0], spec, [0])
        for location in {item['location'] for item in rows}:
            entry = current[location]
            _protect(entry['text'], ''.join(n.text or '' for n in entry['node'].iter(W + 't')), spec)
        expected_parts = {**parts, 'word/document.xml': ET.tostring(root)}
        stream = BytesIO()
        doc.save(stream)
        with zipfile.ZipFile(stream) as archive:
            actual_parts = {name: archive.read(name) for name in archive.namelist()}
        expected = package_projection(expected_parts)
        if package_projection(actual_parts) != expected:
            raise GuardError('professional_route_format_not_equivalent')
        if digest(path) != version:
            raise GuardError('professional_source_changed')
        render_request = RangeDocx(path=str(path), source_sha256=version, operations=render_operations, revision_date=now(),
            **spec.model_dump(include={'mode', 'preserve_numbers', 'preserve_citations', 'preserve_formulas', 'protected_literals'}))
        from .range_alignment import anchors_for
        approved_starts = {(op.locator_id, op.start, op.end): op.start + sum(len(item['operation']['text']) - (item['operation']['end'] - item['operation']['start'])
                           for item in rows if current[item['location']]['locator_id'] == op.locator_id and item['operation']['start'] < op.start)
                           for op in render_operations}
        render_anchors = anchors_for(parts, render_request, approved_parts=expected_parts, approved_starts=approved_starts)
        details = {'selected': request.kind, 'reason': '现有字面原语与宿主选区计划的全文文字/直接格式投影及非目标部件一致；保留逐项账本和最终残留检查。',
                   'covered_items': len(rows), 'primitive_operations': 1, 'range_transaction_limit': 12,
                   'estimated_remaining_transactions': {'native_then_fallback': routed_transactions, 'bounded_ranges': fallback_transactions},
                   'candidates': [{'primitive': request.kind, 'eligible': True}, {'primitive': 'bounded_ranges', 'eligible': True}],
                   'page_alignment_items': len(render_operations), 'page_alignment_total': len(rows),
                   'page_alignment_complete': len(render_operations) == len(rows),
                   'comparison_scope': '只读内存投影，不在真实文件试写；不证明最终继承样式或Word视觉等价'}
        return details, {'request': request, 'ids': [item['id'] for item in rows], 'projection': expected, 'render_request': render_request, 'render_anchors': render_anchors}
    except GuardError as error:
        if str(error) == 'professional_source_changed':
            raise
        fallback['candidates'].append({'primitive': 'native_literal', 'eligible': False, 'reason': '范围、格式、保护项或结构不能证明等价'})
        return fallback, None


def preview_rows(task, ids):
    rows = [item for item in task['items'] if item['id'] in ids]
    return (json.dumps({'items': [{'location': item['location'], 'text': item['operation']['quote']} for item in rows]}, ensure_ascii=False),
            json.dumps({'items': [{'action': item['operation']['action'], 'text': item['operation']['text'], 'reason': item['operation']['reason']} for item in rows]}, ensure_ascii=False))
