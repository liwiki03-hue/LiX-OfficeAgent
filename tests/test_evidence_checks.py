import pytest
from docx import Document
from openpyxl import Workbook

from office_agent.evidence_checks import table_observations, check_numeric
from office_agent.writes import Workspace, GuardError


def fixture(tmp_path, suffix):
    path = tmp_path / ('observations' + suffix)
    rows = [['Sample', 'Baseline (%)', 'Current (%)', 'Latency (ms)', 'Missing'], ['dry', '40', '50', '20', None]]
    if suffix == '.docx':
        doc = Document()
        table = doc.add_table(rows=2, cols=5)
        for ri, row in enumerate(rows):
            for ci, value in enumerate(row):
                table.cell(ri, ci).text = '' if value is None else value
        doc.save(path)
    else:
        book = Workbook()
        for row in rows:
            book.active.append(row)
        book.save(path)
        book.close()
    records = table_observations(Workspace(tmp_path), path.name)
    return records, {r['header']: r for r in records if '/row:1/' in r['location']}


@pytest.mark.parametrize('suffix', ['.docx', '.xlsx'])
def test_units_relative_percent_points_and_missing(tmp_path, suffix):
    records, rows = fixture(tmp_path, suffix)
    latency = rows['Latency (ms)']
    def check(value, unit, **extra):
        return check_numeric(records, dict(observation_id=latency['id'], value=value, unit=unit, **extra))
    assert check('.02', 's')['numeric_status'] == 'match'
    assert check('20', 's')['numeric_status'] == 'mismatch'
    assert check('20', 'kg')['numeric_status'] == 'unit_mismatch'
    assert rows['Missing']['missing'] and rows['Missing']['value'] is None
    args = dict(observation_id=rows['Current (%)']['id'], baseline_id=rows['Baseline (%)']['id'])
    assert check_numeric(records, dict(**args, value='25', unit='%', operation='relative_increase'))['numeric_status'] == 'match'
    assert check_numeric(records, dict(**args, value='10', unit='百分点', operation='percentage_points'))['numeric_status'] == 'match'
    assert check_numeric(records, dict(**args, value='10', unit='%', operation='percentage_points'))['numeric_status'] == 'unit_mismatch'
    assert check_numeric(records, dict(**args, value='25', unit='%', operation='relative_decrease'))['numeric_status'] == 'mismatch'
    assert check('.02', 's')['status'] == 'needs_manual_judgment'
    assert check('20', '°C')['status'] == 'insufficient'


def test_different_object_and_missing_baseline_are_not_contradiction(tmp_path):
    records, rows = fixture(tmp_path, '.docx')
    baseline = dict(rows['Baseline (%)'], id='obs:' + 'a' * 64, row_label='wet')
    records.append(baseline)
    result = check_numeric(records, dict(observation_id=rows['Current (%)']['id'], baseline_id=baseline['id'], value='25', unit='%', operation='relative_increase'))
    assert result['status'] == 'condition_mismatch'
    assert check_numeric(records, dict(observation_id=rows['Current (%)']['id'], value='25', unit='%', operation='relative_increase'))['status'] == 'insufficient'


def test_merged_table_uncertain_and_formula_not_a_value(tmp_path):
    path = tmp_path / 'merged.xlsx'
    book = Workbook()
    sheet = book.active
    sheet.append(['Merged', None])
    sheet.merge_cells('A1:B1')
    sheet.append(['20 ms', '=1+2'])
    book.save(path)
    book.close()
    rows = table_observations(Workspace(tmp_path), path.name)
    assert all(r['uncertain_structure'] for r in rows)
    row = next(r for r in rows if r['raw'] == '20 ms')
    assert check_numeric(rows, dict(observation_id=row['id'], value='20', unit='ms'))['status'] == 'insufficient'
    assert next(r for r in rows if r['raw'] == '=1+2')['value'] is None
