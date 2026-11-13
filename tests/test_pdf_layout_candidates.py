from copy import deepcopy

import pytest

from office_agent.pdf_layout import propose, validate_box
from office_agent.pdf_table import read_table
from office_agent.rag_sources import SourceError
from office_agent.writes import Workspace
from test_pdf_header_rows import header_pdf, ROWS
from test_pdf_table import host

BOX = [.02, .16, .98, .62]


def test_alignment_candidate_preserves_cells_but_requires_new_adoption(tmp_path):
    api, auth = host(tmp_path)
    (tmp_path / 'table.pdf').write_bytes(header_pdf())
    response = api.propose_pdf_layout('table.pdf', 1, BOX)
    assert response['ok'], response
    candidate = response['result']
    assert candidate['candidate_rows'] == ROWS and candidate['candidate_only']
    assert candidate['header_candidates'] == [1, 2]
    assert not candidate['layout_confirmed'] and not candidate['support_assessed']
    assert not api._professional._observations
    parsed = api.read_pdf_numeric(candidate['pdf_table_id'], {**candidate['proposal'], 'header_rows': 2})
    assert parsed['ok'], parsed
    assert parsed['result']['rows'] == ROWS
    assert not parsed['result']['observations'][-1]['layout_confirmed']
    assert parsed['result']['observations'][-1]['unit'] == 'ms'
    assert all(job['empty_confirmed'] for job in candidate['processes'])


@pytest.mark.parametrize('box', [[True, .1, .9, .9], [0, 0, float('nan'), 1], [.9, .1, .1, .9], [0, 0, 2, 1], [0, 1, 1]])
def test_invalid_explicit_rectangles_refused(box):
    with pytest.raises(SourceError):
        validate_box(box)


def test_clipped_and_unaligned_candidates_do_not_fill_unknown_cells(tmp_path):
    path = tmp_path / 'table.pdf'
    path.write_bytes(header_pdf())
    result = propose(path, 1, [.1, .16, .98, .62])
    assert result['proposal'] is None and 'rectangle_clips_glyphs' in result['uncertainties']
    rows = deepcopy(ROWS)
    rows[1][0] = ''
    path.write_bytes(header_pdf(rows))
    result = propose(path, 1, BOX)
    assert result['proposal'] is None
    assert 'inconsistent_columns_or_possible_merged_cells' in result['uncertainties']


def test_proposal_cannot_supply_grid_or_escape_selected_file(tmp_path):
    (tmp_path / 'table.pdf').write_bytes(header_pdf())
    with pytest.raises(SourceError):
        read_table(Workspace(tmp_path), '../table.pdf', 1, proposal_box=BOX)
    with pytest.raises(SourceError):
        read_table(Workspace(tmp_path), 'table.pdf', 1, grid={'page': 1, 'rows': [.1,.3,.8], 'columns': [0,.5,1], 'units': []}, proposal_box=BOX)
