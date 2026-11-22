import io
import uuid

from docx import Document

from ..paged_documents import GuardError, W, sha


def reference(editor, target):
    if set(target) != {'nodes', 'whole_table'} or type(target['whole_table']) is not bool:
        raise GuardError('表格引用字段无效')
    nodes = target['nodes']
    if not isinstance(nodes, list) or not 1 <= len(nodes) <= 8 or any(not isinstance(row, dict) or set(row) != {'node_id', 'start', 'end'} for row in nodes):
        raise GuardError('表格引用限1—8个段落，请缩小范围')
    if any(not isinstance(row['node_id'], str) for row in nodes):
        raise GuardError('表格节点身份无效')
    if len({row['node_id'] for row in nodes}) != len(nodes):
        raise GuardError('表格引用节点重复')
    document = Document(io.BytesIO(editor.data))
    paragraphs = list(document.element.body.iter(W + 'p'))
    by_id = {row['id']: (row, paragraph) for row, paragraph in zip(editor.manifest, paragraphs)}
    tables = list(document.element.body.iter(W + 'tbl'))
    selected = []
    table = None
    from .paged_editor import utf16_slice
    for intent in nodes:
        row, paragraph = by_id.get(intent['node_id'], (None, None))
        if row is None:
            raise GuardError('表格节点已失效')
        ancestors = list(paragraph.iterancestors())
        current = next((node for node in ancestors if node.tag == W + 'tbl'), None)
        cell = next((node for node in ancestors if node.tag == W + 'tc'), None)
        if current is None or cell is None or table is not None and current is not table:
            raise GuardError('请引用同一张表格内的文字')
        table = current
        old = utf16_slice(row['text'], intent['start'], intent['end'])
        formats = {key for start, end, key in editor.selection_formats.get(row['id'], []) if start < intent['end'] and end > intent['start']}
        complex_table = table.getparent() is not document.element.body or any(node.tag == W + 'tbl' for node in cell.iterdescendants())
        merged = any(node.tag in {W + 'gridSpan', W + 'vMerge', W + 'hMerge'} for node in cell.iter())
        complex_text = any(node.tag in {W + 'drawing', W + 'fldChar', W + 'instrText', W + 'hyperlink', W + 'sdt', W + 'ins', W + 'del'} for node in paragraph.iter())
        locator = editor.locator_map.get(row['id'])
        protected = bool(editor.floating.protected and (locator not in editor.floating.locators or editor.floating.locators[locator] in editor.floating.protected))
        selected.append({**intent, 'cell_id': 'table-' + str(tables.index(table)) + '!cell-' + str(list(table.iter(W + 'tc')).index(cell)), 'old': old, 'text_sha256': sha(old.encode('utf8')), 'draft_locator_id': editor.draft_locators[row['id']]['id'],
                         'source_host_locator_id': locator, 'editable': not (complex_table or merged or complex_text or len(formats) > 1 or protected),
                         'readonly_reason': '复杂、合并、混合格式或受保护内容仅供分析' if complex_table or merged or complex_text or len(formats) > 1 or protected else ''})
    if target['whole_table']:
        expected = {row['id'] for row, paragraph in by_id.values() if table in paragraph.iterancestors() and row['text']}
        if expected != {row['node_id'] for row in selected} or any(row['start'] != 0 or row['end'] != len(by_id[row['node_id']][0]['text'].encode('utf-16-le')) // 2 for row in selected):
            raise GuardError('整表范围与当前草稿不符，请重新引用')
    old = '\n'.join(row['old'] for row in selected)
    if not old.strip() or len(old) > 4000:
        raise GuardError('表格引用限1—4000字，请缩小范围')
    return {'id': uuid.uuid4().hex, 'kind': 'table', 'document_id': editor.document_id, 'source_version': editor.source_version,
            'session_epoch': editor.floating.epoch, 'draft_revision': editor.revision, 'table_id': 'table-' + str(tables.index(table)),
            'cells': selected, 'old': old, 'context': old, 'text_sha256': sha(old.encode('utf8')), 'coordinate': 'UTF-16 code units',
            'node_id': selected[0]['node_id'], 'start': selected[0]['start'], 'end': selected[0]['end']}
