import base64
from types import SimpleNamespace

import pytest
from docx import Document
from docx.text.paragraph import Paragraph

from office_agent.paged_documents import W
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from tests.test_document_conversation import FixedAnswer, attach, finish, send


def table_document(tmp_path, mode):
    source = tmp_path / '表格定位.docx'
    book = Document()
    book.add_paragraph('非目标正文起点。')
    for i in range(2 if mode == 'multiple' else 1):
        table = book.add_table(rows=2, cols=2)
        for row, values in zip(table.rows, [('分组', '平均延迟'), ('测量', '287 ms' if i == 0 else '399 ms')]):
            for cell, value in zip(row.cells, values):
                cell.text = value
        if mode == 'empty':
            for cell in table.rows[-1].cells:
                cell.text = ''
        elif mode == 'merged':
            table.cell(0, 0).merge(table.cell(0, 1))
    book.add_paragraph('正文末尾独立标记。')
    book.save(source)
    api = SimpleNamespace(_app=WebviewApp(), _auth=SimpleNamespace(status=lambda: {'authenticated': True}),
                          _memory_key=lambda: 'table-test', _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating = FloatingSession(api)
    assert floating.admit(str(source))['ok'] and floating.read(floating.read_token, True)['ok']
    return book, floating, source


def select_second_table_header(book, floating, original):
    service = floating.paged
    assert service.open(floating.document['id'], floating.document['version'])['ok']
    manifest = [{'id': f'table-node-{i}', 'text': Paragraph(node, book).text}
                for i, node in enumerate(book.element.body.iter(W + 'p'))]
    service.sync(service.document_id, 0, base64.b64encode(original).decode(), manifest, True)
    selected_id = [item['id'] for item in manifest if item['text'] == '分组'][-1]
    return service.reference(service.document_id, 0, {'node_id': selected_id, 'start': 0, 'end': 2})['reference']['id']


def test_word_table_tail_and_body_tail_use_distinct_material(tmp_path):
    _, floating, source = table_document(tmp_path, 'single')
    original = source.read_bytes()
    for identifier, question, answer, expected, excluded, location in [
        ('table-tail', '表格最后一行的延迟是多少？', '平均延迟为287 ms。', ('平均延迟', '287 ms'), '正文末尾独立标记', '表格1第2行'),
        ('body-tail', '整篇文档最后一段是什么？', '正文末尾独立标记。', ('正文末尾独立标记',), '非目标正文起点', '正文第2段'),
    ]:
        model = FixedAnswer({'content': answer})
        attach(floating, model)
        assert send(floating, identifier, question)['handled']
        finish(floating)
        turn = floating.conversation.turns[identifier]
        assert len(model.calls) == 1 and all(value in model.calls[0] for value in expected)
        assert excluded not in model.calls[0]
        assert turn['status'] == 'complete' and turn['answer'] == answer + '\n（已读范围：' + location + '）'
        assert location in floating.conversation.answer_provenance[identifier]['labels']
        if identifier == 'table-tail':
            assert [item['location'] for item in turn['references']] == ['表格1第1行', '表格1第2行']
        else:
            assert [item['location'] for item in turn['references']] == ['正文段落2']
    assert source.read_bytes() == original and not floating.outputs and not floating.access.state()['id']


@pytest.mark.parametrize('question,select', [('表格最后一行的延迟是多少？', True), ('第二张表最后一行的延迟是多少？', False)])
def test_multiple_tables_resolve_explicit_number_or_bound_table_cell(tmp_path, question, select):
    book, floating, source = table_document(tmp_path, 'multiple')
    original = source.read_bytes()
    reference = select_second_table_header(book, floating, original) if select else None
    model = FixedAnswer({'content': '平均延迟为399 ms。'})
    attach(floating, model)
    assert send(floating, 'multiple-table', question, reference=reference)['handled']
    finish(floating)
    row = floating.conversation.turns['multiple-table']
    assert row['answer'].endswith('（已读范围：表格2第2行）') and row['status'] == 'complete'
    assert len(model.calls) == 1 and '399 ms' in model.calls[0] and '287 ms' not in model.calls[0]
    assert source.read_bytes() == original and not floating.outputs and not floating.access.state()['id']


@pytest.mark.parametrize('missing_old_scope', [False, True])
def test_explicit_selection_cannot_receive_table_tail_from_chat_history(tmp_path, missing_old_scope):
    book, floating, source = table_document(tmp_path, 'multiple')
    original = source.read_bytes()
    reference = select_second_table_header(book, floating, original)
    model = FixedAnswer({'content': '平均延迟为399 ms。'})
    attach(floating, model)
    send(floating, 'table-data', '表格最后一行的延迟是多少？', reference=reference)
    finish(floating)
    earlier = floating.conversation.turns['table-data']
    assert len(model.calls) == 1 and '399 ms' in model.calls[0] and '399 ms' in earlier['answer']
    if missing_old_scope:
        earlier.pop('references')
    model = FixedAnswer({'content': '分组。'})
    attach(floating, model)
    send(floating, 'only-header', '只看我选中的文字写了什么？', reference=reference)
    finish(floating)
    row = floating.conversation.turns['only-header']
    assert len(model.calls) == 1 and '分组' in model.calls[0] and '399 ms' not in model.calls[0]
    assert row['status'] == 'complete' and row['answer'].endswith('表格2第1行第1格第1段）')
    assert source.read_bytes() == original and not floating.outputs and not floating.access.state()['id']


@pytest.mark.parametrize('mode,question,explanation', [
    ('multiple', '表格最后一行的延迟是多少？', '表号'),
    ('multiple', '表格3最后一行的延迟是多少？', '找不到表格3'),
    ('empty', '表格最后一行的延迟是多少？', '没有可读取的文字'),
    ('merged', '表格最后一行的延迟是多少？', '合并'),
])
def test_missing_table_evidence_does_not_send_document(tmp_path, mode, question, explanation):
    _, floating, source = table_document(tmp_path, mode)
    original = source.read_bytes()
    model = FixedAnswer({'content': '不应调用'})
    attach(floating, model)
    assert send(floating, 'missing-table', question)['handled']
    finish(floating)
    row = floating.conversation.turns['missing-table']
    assert explanation in row['text'] and not row['references']
    assert not model.calls and source.read_bytes() == original and not floating.outputs and not floating.access.state()['id']
