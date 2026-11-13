from copy import deepcopy

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.professional_edits import W
from office_agent.writes import digest
from tests.test_return_continuation import prepared, parts


def compared(service, payload):
    reply = service.compare_return({**payload, 'include_plain_tables': True})
    assert reply['ok'], reply
    return reply['result']


def table_request(result, **extra):
    row = next(row for row in result['items'] if '/table/' in (row['current_location'] or ''))
    return {'comparison_id': result['comparison_id'], 'row_id': row['continuation_id'], 'find': '表格候选',
            'replacement': '表格续编', 'reason': '只调整普通单元格措辞', 'acknowledge_current': True, **extra}


@pytest.mark.parametrize('mode', ['clean', 'tracked'])
def test_table_opt_in_new_task_cancel_copy_geometry_and_undo(tmp_path, monkeypatch, mode):
    baseline, delivered, current, editor, service, _, _, payload, old = prepared(tmp_path, monkeypatch)
    assert next(row for row in old['items'] if '/table/' in (row['current_location'] or ''))['continuation_id'] is None
    result = compared(service, payload)
    row = next(row for row in result['items'] if row['continuation_family'] == 'ordinary_table_cell')
    assert row['continuation_id'] and row['continuation_blocked'] is None
    versions = [digest(path) for path in (baseline, delivered, current)]
    reply = service.continue_return(table_request(result, mode=mode, editing_profile='technical_report_v1'))
    assert reply['ok'], reply
    task = reply['result']; assert task['source_sha256'] == versions[-1] and not task['receipts']
    plan = service.plan_group(task['task_id'])['result']
    assert service.confirm_group(task['task_id'], plan['id'], False)['result']['status'] == 'cancelled'
    plan = service.plan_group(task['task_id'])['result']
    saved = service.confirm_group(task['task_id'], plan['id'], True)
    assert saved['ok'], saved
    output = tmp_path / saved['result']['task']['current_name']
    before, after = parts(current), parts(output)
    assert all(after[name] == value for name, value in before.items() if name != 'word/document.xml')
    def geometry(data):
        root = ET.fromstring(data['word/document.xml'])
        table = root.find('.//' + W + 'tbl')
        for paragraph in table.iter(W + 'p'):
            paragraph.clear()
        return ET.tostring(table, method='c14n')
    assert geometry(before) == geometry(after)
    assert versions == [digest(path) for path in (baseline, delivered, current)]
    assert service.undo_last(task['task_id'], True)['result']['status'] == 'rolled_back'
    assert not output.exists()


@pytest.mark.parametrize('change', ['width', 'merge', 'nested', 'paragraph', 'marker'])
def test_changed_or_unsupported_table_never_gets_continuation_ticket(tmp_path, monkeypatch, change):
    _, _, current, _, service, _, _, payload, _ = prepared(tmp_path, monkeypatch)
    contents = parts(current); root = ET.fromstring(contents['word/document.xml']); table = root.find('.//' + W + 'tbl')
    cell = table.find('.//' + W + 'tc')
    if change == 'width':
        table.find(W + 'tblGrid/' + W + 'gridCol').set(W + 'w', '123')
    elif change == 'merge':
        ET.SubElement(cell.find(W + 'tcPr'), W + 'gridSpan', {W + 'val': '2'})
    elif change == 'nested':
        ET.SubElement(cell, W + 'tbl')
    elif change == 'paragraph':
        ET.SubElement(cell, W + 'p')
    else:
        ET.SubElement(cell.find(W + 'p'), W + 'bookmarkStart', {W + 'id': '99', W + 'name': 'test'})
    contents['word/document.xml'] = ET.tostring(root)
    import zipfile
    with zipfile.ZipFile(current, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in contents.items():archive.writestr(name, data)
    reply = service.compare_return({**payload, 'include_plain_tables': True})
    if reply['ok']:
        rows = [row for row in reply['result']['items'] if '/table/' in row['delivered_location']]
        assert rows and all(row['continuation_id'] is None for row in rows)
    else:
        assert reply['reason'].startswith('professional_')


@pytest.mark.parametrize('rows,columns,allowed', [(5, 20, True), (1, 21, False), (101, 1, False)])
def test_table_admission_budget_is_not_silently_expanded(tmp_path, rows, columns, allowed):
    from office_agent.document_tasks import load
    from office_agent.return_tables import plain_current
    from office_agent.writes import Editor
    path = tmp_path / 'budget.docx'
    document = Document()
    document.add_table(rows=rows, cols=columns).cell(0, 0).text = '明确目标'
    document.save(path)
    entry = next(entry for entry in load(Editor(tmp_path), path)[3] if entry['text'] == '明确目标')
    assert plain_current(entry) is allowed
