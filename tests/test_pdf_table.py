from copy import deepcopy
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.pdf_table import payload, read_table, validate_grid
from office_agent.rag_sources import SourceError
from office_agent.webui.bridge import Api
from office_agent.writes import Editor, Workspace

GRID = {'page': 1, 'rows': [.16, .35, .62], 'columns': [.02, .42, .98], 'units': []}


def grid_pdf(note=None, crop=False):
    text = 'BT /F1 12 Tf 20 150 Td (Sample) Tj ET BT /F1 12 Tf 150 150 Td (Value \\(ms\\)) Tj ET BT /F1 12 Tf 20 100 Td (A) Tj ET BT /F1 12 Tf 150 100 Td (20) Tj ET'
    if note:
        text = text.replace('(20)', '(20[a])') + f' BT /F1 12 Tf 20 50 Td ({note}) Tj ET'
    content = text.encode()
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>', b'<< /Type /Pages /Kids [5 0 R] /Count 1 >>',
               b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
               f'<< /Length {len(content)} >>\nstream\n'.encode() + content + b'\nendstream',
               b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 3 0 R >> >> /Contents 4 0 R >>']
    if crop:
        objects[-1] = objects[-1].replace(b'/Resources', b'/CropBox [10 0 300 200] /Resources')
    data, offsets = b'%PDF-1.4\n', []
    for index, obj in enumerate(objects, 1):
        offsets.append(len(data))
        data += f'{index} 0 obj\n'.encode() + obj + b'\nendobj\n'
    start = len(data)
    data += b'xref\n0 6\n0000000000 65535 f \n' + b''.join(f'{offset:010} 00000 n \n'.encode() for offset in offsets)
    return data + f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{start}\n%%EOF\n'.encode()


def host(tmp_path):
    path = tmp_path / 'report.docx'
    document = Document()
    document.add_paragraph('Measured latency is 20 ms.')
    document.save(path)
    (tmp_path / 'table.pdf').write_bytes(grid_pdf())
    authenticated = [True]
    app = SimpleNamespace(path=path, editor=Editor(tmp_path, require_confirmation=True), busy=False)
    auth = SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': authenticated[0]})
    return Api(app, auth=auth, settings_root=tmp_path), authenticated


def test_native_grid_exact_cells_and_owned_worker_cleanup(tmp_path):
    path = tmp_path / 'table.pdf'
    path.write_bytes(grid_pdf())
    result = read_table(Workspace(tmp_path), path.name, 1, GRID)
    assert result['rows'] == [['Sample', 'Value (ms)'], ['A', '20']]
    assert all(job['empty_confirmed'] for job in result['processes'])
    assert not result['ocr_performed'] and not result['automatic_table_detection']
    assert result['cells'][-1]['char_indices'] and result['cells'][-1]['box'] == [.42, .35, .98, .62]


@pytest.mark.parametrize('change', ['boolean', 'unsorted', 'nan', 'capacity', 'unit_map', 'extra'])
def test_invalid_grids_rejected(change):
    grid = deepcopy(GRID)
    if change == 'boolean':
        grid['rows'][0] = True
    elif change == 'unsorted':
        grid['columns'].reverse()
    elif change == 'nan':
        grid['rows'][0] = float('nan')
    elif change == 'capacity':
        grid['rows'] = [i / 21 for i in range(22)]
        grid['columns'] = [i / 12 for i in range(13)]
    elif change == 'unit_map':
        grid['units'] = [[1, 9]]
    else:
        grid['path'] = 'unowned.pdf'
    with pytest.raises(SourceError):
        validate_grid(grid)


def test_clipped_text_and_path_refused(tmp_path):
    path = tmp_path / 'table.pdf'
    path.write_bytes(grid_pdf())
    clipped = {**GRID, 'columns': [.02, .52, .98]}
    with pytest.raises(SourceError):
        payload(path, 1, clipped)
    with pytest.raises(SourceError):
        read_table(Workspace(tmp_path), '../table.pdf', 1, GRID)


def test_cropped_page_origin_refused_instead_of_misaligned_grid(tmp_path):
    path = tmp_path / 'cropped.pdf'
    path.write_bytes(grid_pdf(crop=True))
    with pytest.raises(SourceError):
        payload(path, 1, GRID)


def test_overlapping_note_region_refused():
    with pytest.raises(SourceError):
        validate_grid({**GRID, 'note_box': [.1, .2, .9, .7]})


def test_layout_ack_required_and_source_change_retires_numeric(tmp_path):
    api, authenticated = host(tmp_path)
    anchor = api.read_structure_anchors()['result']['anchors'][0]['locator_id']
    preview = api.preview_pdf_numeric('table.pdf', 1)['result']
    parsed = api.read_pdf_numeric(preview['pdf_table_id'], GRID)
    assert parsed['ok'], parsed
    parsed = parsed['result']
    row = next(row for row in parsed['observations'] if row['raw'] == '20')
    args = {'observation_id': row['id'], 'value': '20', 'unit': 'ms'}
    assert not api.verify_paragraph_numeric(anchor, args)['ok']
    assert not row['layout_confirmed'] and row['uncertain_structure']
    ready = api.acknowledge_pdf_numeric(parsed['pdf_table_id'], True)['result']
    assert ready['layout_confirmed'] and not ready['user_reviewed'] and not ready['write_approval_granted']
    assert api.verify_paragraph_numeric(anchor, args)['result']['numeric_status'] == 'match'
    assert not api.acknowledge_pdf_numeric(parsed['pdf_table_id'], True)['ok']
    (tmp_path / 'table.pdf').write_bytes(grid_pdf() + b'\n')
    assert api.verify_paragraph_numeric(anchor, args)['reason'] == 'professional_pdf_source_changed'


def test_expiry_cancel_and_logout_no_restoration(tmp_path):
    api, authenticated = host(tmp_path)
    first = api.preview_pdf_numeric('table.pdf', 1)['result']
    api._professional._pdf_candidate['created'] -= 301
    assert api.read_pdf_numeric(first['pdf_table_id'], GRID)['reason'] == 'professional_pdf_expired'
    first = api.preview_pdf_numeric('table.pdf', 1)['result']
    parsed = api.read_pdf_numeric(first['pdf_table_id'], GRID)['result']
    assert api.acknowledge_pdf_numeric(parsed['pdf_table_id'], False)['result']['status'] == 'cancelled'
    assert not api.acknowledge_pdf_numeric(parsed['pdf_table_id'], True)['ok']
    first = api.preview_pdf_numeric('table.pdf', 1)['result']
    authenticated[0] = False
    assert api.read_pdf_numeric(first['pdf_table_id'], GRID)['reason'] == 'not_authenticated'
    authenticated[0] = True
    assert api.read_pdf_numeric(first['pdf_table_id'], GRID)['reason'] == 'professional_pdf_stale'


def test_pdf_change_after_comment_plan_rejected_before_write(tmp_path):
    api, authenticated = host(tmp_path)
    structure = api.read_structure_anchors()['result']
    anchor = structure['anchors'][0]
    preview = api.preview_pdf_numeric('table.pdf', 1)['result']
    parsed = api.read_pdf_numeric(preview['pdf_table_id'], GRID)['result']
    ready = api.acknowledge_pdf_numeric(parsed['pdf_table_id'], True)['result']
    row = next(row for row in ready['observations'] if row['raw'] == '20')
    checked = api.verify_paragraph_numeric(anchor['locator_id'], {'observation_id': row['id'], 'value': '20', 'unit': 'ms'})['result']
    start = anchor['text'].index('20')
    plan = api.plan_structure_edit({'source_sha256': structure['document']['source_sha256'],
        'comments': [{'anchor': anchor['locator_id'], 'start': start, 'end': start + 2, 'quote': '20', 'text': 'Check the PDF source.'}],
        'numeric_reviews': [checked['numeric_review_id']]})
    assert plan['ok'], plan
    assert 'pdf_numeric_review' in plan['result']['after']
    (tmp_path / 'table.pdf').write_bytes(grid_pdf() + b'\n')
    rejected = api.confirm_professional_edit(plan['result']['id'], True)
    assert rejected['reason'] == 'professional_pdf_source_changed', rejected
    assert not list(tmp_path.glob('*.agent-*.docx')) and not api._professional._disabled


@pytest.mark.parametrize('note,expected', [('Note [a]: conditions=dry', 'match'), ('Note [a]: At 95 percent confidence', 'insufficient')])
def test_explicit_note_region_and_conservative_unstructured_context(tmp_path, note, expected):
    api, authenticated = host(tmp_path)
    (tmp_path / 'table.pdf').write_bytes(grid_pdf(note))
    preview = api.preview_pdf_numeric('table.pdf', 1)['result']
    parsed = api.read_pdf_numeric(preview['pdf_table_id'], {**GRID, 'note_box': [.02, .66, .98, .82]})
    assert parsed['ok'], parsed
    adopted = api.acknowledge_pdf_numeric(parsed['result']['pdf_table_id'], True)['result']
    row = next(row for row in adopted['observations'] if row['raw'] == '20[a]')
    assert row['pdf_note_text'] == note and row['pdf_note_char_indices']
    assert not row['context_outside_selection_included']
    anchor = api.read_structure_anchors()['result']['anchors'][0]['locator_id']
    checked = api.verify_paragraph_numeric(anchor, {'observation_id': row['id'], 'value': '20', 'unit': 'ms'})
    assert checked['result']['numeric_status'] == expected, checked
