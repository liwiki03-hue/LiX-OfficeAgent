from docx import Document
import pytest

from office_agent.advanced_writes import _parts
from office_agent.range_alignment import anchors_for, locate
from office_agent.range_edits import view, transform
from office_agent.writes import Editor
from tests.test_range_edits import operation, request


def test_short_cell_uses_approved_row_not_repeated_body_or_old_neighbor(tmp_path):
    path = tmp_path / 'source.docx'
    doc = Document()
    doc.add_paragraph('正文多次规范术语；规范术语并非该表行。')
    table = doc.add_table(rows=1, cols=3)
    for cell, text in zip(table.rows[0].cells, ['干燥样本', '旧术语', '20 ms']):
        cell.text = text
    doc.save(path)
    editor = Editor(tmp_path)
    data = view(editor, path)
    selected = [item for item in data['items'] if item['text'] in {'干燥样本', '旧术语'}]
    req = request((path, editor, data), [operation(item, text='潮湿样本' if item['text'] == '干燥样本' else '规范术语') for item in selected])
    before = _parts(path, range_edit=True)
    after = transform(before, req)[0]
    anchors = anchors_for(before, req, after)
    target = next(item for item in anchors if item['before'] == '旧术语')
    assert locate([{'text': '正文旧术语。干燥样本 旧术语 20 ms'}], target['contexts']['before'])['pages'] == [1]
    assert '潮湿样本' in target['contexts']['after']['prefix']
    assert locate([{'text': '规范术语 规范术语。'}, {'text': '潮湿样本 规范术语 20 ms'}], target['contexts']['after'])['pages'] == [2]
    assert anchors_for(before, req)[1]['contexts']['after'] == target['contexts']['after']
    assert locate([{'text': '潮湿样本 规范术语 20 ms 潮湿样本 规范术语 20 ms'}], target['contexts']['after'])['status'] == 'ambiguous'


@pytest.mark.parametrize('complexity', ['multiple_paragraphs', 'merged'])
def test_merged_and_multiline_cells_do_not_guess_row_order(tmp_path, complexity):
    path = tmp_path / 'source.docx'
    doc = Document()
    table = doc.add_table(rows=1, cols=3)
    table.cell(0, 0).text = '行标题'
    table.cell(0, 2).text = '旧术语'
    if complexity == 'multiple_paragraphs':
        table.cell(0, 0).add_paragraph('多段')
    else:
        table.cell(0, 0).merge(table.cell(0, 1))
    doc.save(path)
    editor = Editor(tmp_path)
    data = view(editor, path)
    row = next(item for item in data['items'] if item['text'] == '旧术语')
    req = request((path, editor, data), [operation(row, text='规范术语')])
    target = anchors_for(_parts(path, range_edit=True), req)[0]
    assert locate([{'text': '行标题 多段 旧术语'}], target['contexts']['before'])['status'] == 'not_located'
