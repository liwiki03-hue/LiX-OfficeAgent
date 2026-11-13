import pytest

from office_agent.evidence_checks import check_numeric, observations_from_tables
from office_agent.pdf_table import payload, validate_grid
from office_agent.rag_sources import SourceError
from test_pdf_table import host

GRID = {'page': 1, 'rows': [.16, .29, .41, .62], 'columns': [.02, .42, .98], 'units': [], 'header_rows': 2}
ROWS = [['Sample', 'Latency'], ['Identifier', '(ms)'], ['A', '20']]


def header_pdf(rows=None):
    rows = rows or ROWS
    parts = []
    for ri, row in enumerate(rows):
        for ci, raw in enumerate(row):
            text = raw.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')
            parts.append(f'BT /F1 12 Tf {20 + ci * 130} {150 - ri * 25} Td ({text}) Tj ET')
    content = ' '.join(parts).encode()
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>', b'<< /Type /Pages /Kids [5 0 R] /Count 1 >>',
               b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
               f'<< /Length {len(content)} >>\nstream\n'.encode() + content + b'\nendstream',
               b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 3 0 R >> >> /Contents 4 0 R >>']
    data, offsets = b'%PDF-1.4\n', []
    for index, obj in enumerate(objects, 1):
        offsets.append(len(data))
        data += f'{index} 0 obj\n'.encode() + obj + b'\nendobj\n'
    start = len(data)
    data += b'xref\n0 6\n0000000000 65535 f \n' + b''.join(f'{offset:010} 00000 n \n'.encode() for offset in offsets)
    return data + f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{start}\n%%EOF\n'.encode()


def records(rows=ROWS, header_rows=2):
    return observations_from_tables('table.pdf', 'owned-source-hash', [('pdf-page:1/grid:test', rows, False, [], False)], header_rows=header_rows)


def test_two_headers_keep_source_coordinates_and_context(tmp_path):
    api, authenticated = host(tmp_path)
    (tmp_path / 'table.pdf').write_bytes(header_pdf())
    anchor = api.read_structure_anchors()['result']['anchors'][0]['locator_id']
    preview = api.preview_pdf_numeric('table.pdf', 1)['result']
    parsed = api.read_pdf_numeric(preview['pdf_table_id'], GRID)
    assert parsed['ok'], parsed
    parsed = parsed['result']
    assert parsed['rows'] == ROWS and parsed['header_rows'] == 2
    selected = next(row for row in parsed['observations'] if row['raw'] == '20')
    assert selected['header'] == 'Latency / (ms)' and selected['unit'] == 'ms'
    assert selected['sample'] == 'A' and '/row:2/column:1' in selected['location']
    assert [row['raw'] for row in selected['header_sources']] == ['Latency', '(ms)']
    assert [row['location'].split('/row:')[-1] for row in selected['header_sources']] == ['0/column:1', '1/column:1']
    assert selected['char_indices'] and not selected['layout_confirmed']
    ready = api.acknowledge_pdf_numeric(parsed['pdf_table_id'], True)['result']
    assert not ready['user_reviewed']
    result = api.verify_paragraph_numeric(anchor, {'observation_id': selected['id'], 'value': '20', 'unit': 'ms', 'sample': 'A'})
    assert result['result']['numeric_status'] == 'match', result
    assert result['result']['user_reviewed'] is False
    proof = api._professional._numeric_reviews[result['result']['numeric_review_id']]['proof']
    assert proof['grid']['header_rows'] == 2 and proof['observation_id'] == selected['id']
    mismatch = api.verify_paragraph_numeric(anchor, {'observation_id': selected['id'], 'value': '20', 'unit': 'ms', 'sample': 'B'})
    assert mismatch['result']['status'] == 'condition_mismatch'


@pytest.mark.parametrize('count', [True, 0, 3, '2'])
def test_invalid_header_selection_rejected(count):
    with pytest.raises(SourceError):
        validate_grid({**GRID, 'header_rows': count})


def test_two_headers_require_data_and_no_empty_header_cell(tmp_path):
    with pytest.raises(SourceError):
        validate_grid({**GRID, 'rows': [.16, .29, .41]})
    path = tmp_path / 'table.pdf'
    path.write_bytes(header_pdf([['Sample', 'Latency'], ['', '(ms)'], ['A', '20']]))
    with pytest.raises(SourceError):
        payload(path, 1, GRID)


@pytest.mark.parametrize('headers,flag', [(['Latency (s)', '(ms)'], 'ambiguous_header_unit'),
                                       (['Mean', 'Uncertainty (ms)'], 'ambiguous_header_statistic')])
def test_ambiguous_header_units_or_statistic_never_match(headers, flag):
    rows = [['Sample', headers[0]], ['Identifier', headers[1]], ['A', '20 ms']]
    values = records(rows)
    selected = values[-1]
    assert flag in selected['context_flags']
    assert check_numeric(values, {'observation_id': selected['id'], 'value': '20', 'unit': 'ms'})['numeric_status'] == 'insufficient'


def test_header_numbers_not_data_and_default_stays_single_row():
    rows = [['Sample', '2024'], ['Identifier', '(ms)'], ['A', '20']]
    values = records(rows)
    assert values[1]['is_header'] and values[1]['value'] is None
    old = records(ROWS, header_rows=1)
    assert old[-1]['unit'] is None and 'header_sources' not in old[-1]


def test_repeated_unit_and_conflicting_context_headers():
    values = records([['Sample', 'Latency (ms)'], ['Conditions', '(ms)'], ['A', '20']])
    assert values[-1]['unit'] == 'ms'
    assert 'ambiguous_context_header' in values[-1]['context_flags']
    assert check_numeric(values, {'observation_id': values[-1]['id'], 'value': '20', 'unit': 'ms'})['numeric_status'] == 'insufficient'
