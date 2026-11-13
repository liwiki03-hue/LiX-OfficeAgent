import pytest

from office_agent.rag_sources import SourceError
from office_agent.rag_workspace import WorkspaceRag
from office_agent.research_pdf import PdfResearchSource
from office_agent.writes import Workspace
from test_fidelity import pdf_bytes
from test_rag_workspace import SyntheticEmbedder, SESSION


def test_actual_pdf_text_layer_worker_index_and_coordinate_context(tmp_path):
    source = tmp_path / 'fixture.pdf'
    source.write_bytes(pdf_bytes(text='EXPT-02 sample 30 ms'))
    original = source.read_bytes()
    extractor = PdfResearchSource(Workspace(tmp_path))
    result = extractor.extract(source.name)
    assert result['blocks'] and 'pdf-page:1' in result['blocks'][0]['location']
    assert '/box:' in result['blocks'][0]['location']
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    plan = service.plan_research_index(source.name, SESSION)
    assert service.confirm_index(plan['plan_id'], SESSION, True)['status'] == 'saved'
    query = service.plan_research_query(source.name, 'EXPT-02', SESSION)
    saved = service.confirm_query(query['plan_id'], SESSION, True)
    assert saved['status'] == 'saved'
    context = service.read_research_context(saved['save_id'], saved['hits'][0]['id'], SESSION)
    assert 'pdf-page:1' in context['sections'][0]['location'] and not context['writable']
    assert source.read_bytes() == original


@pytest.mark.parametrize('data', [pdf_bytes(text=''), pdf_bytes(pages=6)])
def test_pdf_without_text_or_over_page_budget_refuses_entire_index(tmp_path, data):
    source = tmp_path / 'fixture.pdf'
    source.write_bytes(data)
    with pytest.raises(SourceError):
        PdfResearchSource(Workspace(tmp_path)).extract(source.name)
    assert not list(tmp_path.glob('.office-*'))
