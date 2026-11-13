import pytest
from docx import Document
from openpyxl import Workbook

from office_agent.evidence_checks import table_observations, check_numeric
from office_agent.writes import Workspace


def material(tmp_path, suffix, note='Note [a]: conditions=dry; time=2024', duplicate=False):
    rows = [['Sample', 'Mean (ms)', 'Baseline (ms)', 'Conditions', 'Time'],
            ['A', '20 ± 2[a]', '10[a]', 'dry', '2024'],
            ['B', '0', '10', 'wet', '2023']]
    path = tmp_path / ('table' + suffix)
    if suffix == '.docx':
        doc = Document()
        table = doc.add_table(rows=len(rows), cols=len(rows[0]))
        for ri, row in enumerate(rows):
            for ci, value in enumerate(row):
                table.cell(ri, ci).text = value
        if note:
            doc.add_paragraph(note)
            if duplicate:
                doc.add_paragraph(note)
        doc.save(path)
    else:
        book = Workbook()
        sheet = book.active
        for row in rows:
            sheet.append(row)
        if note:
            sheet.append([note])
            if duplicate:
                sheet.append([note])
        book.save(path)
        book.close()
    records = table_observations(Workspace(tmp_path), path.name)
    selected = next(row for row in records if row['raw'] == '20 ± 2[a]')
    return records, selected


@pytest.mark.parametrize('suffix', ['.docx', '.xlsx'])
def test_explicit_notes_context_uncertainty_and_zero_preserve_source(tmp_path, suffix):
    records, row = material(tmp_path, suffix)
    assert row['raw'] == '20 ± 2[a]' and row['value'] == '20' and row['uncertainty'] == '2'
    assert row['unit'] == 'ms' and row['sample'] == 'A' and row['conditions'] == 'dry' and row['time'] == '2024'
    assert row['footnote_records'][0]['marker'] == 'a' and row['context_sources'] and not row['context_flags']
    assert any(r['value'] == '0' for r in records)
    result = check_numeric(records, dict(observation_id=row['id'], value='20', unit='ms', statistic='mean', sample='A', conditions='dry', time='2024'))
    assert result['numeric_status'] == 'match' and result['status'] == 'needs_manual_judgment'
    assert not result['user_reviewed'] and result['evidence']['uncertainty'] == '2'


@pytest.mark.parametrize('key,value', [('sample', 'B'), ('conditions', 'wet'), ('time', '2023')])
def test_claim_context_mismatch_is_not_arithmetic_agreement(tmp_path, key, value):
    records, row = material(tmp_path, '.xlsx')
    result = check_numeric(records, dict(observation_id=row['id'], value='20', unit='ms', statistic='mean', **{key: value}))
    assert result['status'] == 'condition_mismatch' and result['numeric_status'] == 'insufficient'
    assert {'field': key, 'status': 'condition_mismatch'} in result['context_issues']


@pytest.mark.parametrize('note,duplicate,flag', [('', False, 'missing_footnote'),
    ('Note [a]: conditions=dry', True, 'ambiguous_footnote'),
    ('Note [a]: conditions=wet', False, 'conflicting_conditions'),
    ('Note [a]: conditions=dry; not comparable across runs', False, 'unstructured_footnote_requires_review')])
def test_unresolved_or_restrictive_footnotes_block_comparison(tmp_path, note, duplicate, flag):
    records, row = material(tmp_path, '.docx', note, duplicate)
    assert flag in row['context_flags']
    assert check_numeric(records, dict(observation_id=row['id'], value='20', unit='ms', statistic='mean'))['status'] == 'insufficient'


def test_statistics_and_baseline_context_are_checked(tmp_path):
    records, row = material(tmp_path, '.xlsx')
    assert check_numeric(records, dict(observation_id=row['id'], value='20', unit='ms', statistic='uncertainty'))['numeric_status'] == 'statistic_mismatch'
    baseline = next(r for r in records if r['raw'] == '0')
    result = check_numeric(records, dict(observation_id=row['id'], baseline_id=baseline['id'], value='20', unit='%', statistic='mean', operation='relative_increase'))
    assert result['status'] == 'condition_mismatch'


def test_notes_from_next_table_do_not_authorize_previous_table(tmp_path):
    path = tmp_path / 'scope.docx'
    doc = Document()
    first = doc.add_table(rows=2, cols=1)
    first.cell(0, 0).text = 'Value (ms)'
    first.cell(1, 0).text = '20[a]'
    doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Other table'
    doc.add_paragraph('Note [a]: conditions=dry')
    doc.save(path)
    row = next(r for r in table_observations(Workspace(tmp_path), path.name) if r['raw'] == '20[a]')
    assert row['footnote_records'] == [] and 'missing_footnote' in row['context_flags']


def test_interval_is_not_a_point_value(tmp_path):
    path = tmp_path / 'range.xlsx'
    book = Workbook()
    book.active.append(['Value (ms)'])
    book.active.append(['18–22 ms'])
    book.save(path)
    book.close()
    rows = table_observations(Workspace(tmp_path), path.name)
    row = rows[1]
    assert row['range'] == ['18', '22'] and row['value'] is None and not row['missing']
    assert check_numeric(rows, dict(observation_id=row['id'], value='20', unit='ms'))['status'] == 'insufficient'
