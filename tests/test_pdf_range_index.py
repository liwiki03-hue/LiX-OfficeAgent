import json

import pytest

from office_agent.rag_sources import SourceError
from office_agent.rag_workspace import WorkspaceRag
from office_agent.research_pdf import PdfResearchSource
from office_agent.vectorstore import VectorStoreError
from office_agent.writes import Workspace
from test_fidelity import pdf_bytes
from test_rag_workspace import SyntheticEmbedder, SESSION
from test_research_reading import long_pdf
from test_rag_api import make


def test_late_range_transaction_reopen_query_context_replace_and_undo(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    original = path.read_bytes()
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    plan = service.plan_research_index(path.name, SESSION, start_page=6, end_page=7)
    assert not embedder.calls and not service.snapshot()['exists']
    assert plan['pdf_coverage'] == {'start_page': 6, 'end_page': 7, 'total_pages': 7, 'complete': False, 'text_order': 'content_order_unreviewed'}
    assert service.confirm_index(plan['plan_id'], SESSION, False)['status'] == 'rejected'
    assert not service.snapshot()['exists']
    plan = service.plan_research_index(path.name, SESSION, start_page=6, end_page=7)
    service.confirm_index(plan['plan_id'], SESSION, True)
    reopened = WorkspaceRag(Workspace(tmp_path), embedder)
    assert reopened.research_status(path.name, SESSION, start_page=6, end_page=7)['status'] == 'current'
    query = reopened.plan_research_query(path.name, 'measured latency', SESSION, top_k=10, start_page=6, end_page=7)
    assert not query['pdf_coverage']['complete']
    saved = reopened.confirm_query(query['plan_id'], SESSION, True)
    assert any('pdf-page:7' in hit['location'] for hit in saved['hits'])
    hit = next(hit for hit in saved['hits'] if 'pdf-page:7' in hit['location'])
    assert '20 ms' in json.dumps(reopened.read_research_context(saved['save_id'], hit['id'], SESSION))
    assert reopened.research_status(path.name, SESSION, start_page=1, end_page=7)['status'] == 'profile_mismatch'
    with pytest.raises(VectorStoreError, match='index_profile_mismatch'):
        reopened.plan_research_query(path.name, 'latency', SESSION, start_page=1, end_page=7)
    plan = reopened.plan_research_index(path.name, SESSION, start_page=1, end_page=7)
    assert plan['pdf_coverage']['complete']
    indexed = reopened.confirm_index(plan['plan_id'], SESSION, True)
    assert reopened.research_status(path.name, SESSION, start_page=1, end_page=7)['status'] == 'current'
    undo = reopened.plan_index_rollback(indexed['save_id'], SESSION)
    assert reopened.confirm_index_rollback(undo['rollback_id'], SESSION, True)['uncertain'] is False
    assert reopened.research_status(path.name, SESSION, start_page=6, end_page=7)['status'] == 'current'
    assert path.read_bytes() == original
    assert '20 ms' in json.dumps(reopened.read_research_context(saved['save_id'], hit['id'], SESSION))
    path.write_bytes(pdf_bytes(text='replacement'))
    with pytest.raises(ValueError):
        reopened.read_research_context(saved['save_id'], hit['id'], SESSION)


@pytest.mark.parametrize('start,end', [(True, 2), (0, 2), (7, 6), (1, 21), (1, None), (1.0, 2), (2000, 2001)])
def test_invalid_ranges_rejected_before_parsing_or_writing(tmp_path, monkeypatch, start, end):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    monkeypatch.setattr('office_agent.research_pdf.read_pdf', lambda *a, **k: pytest.fail('must validate before parse'))
    with pytest.raises(SourceError):
        PdfResearchSource(Workspace(tmp_path)).extract(path.name, start_page=start, end_page=end)
    assert set(tmp_path.iterdir()) == {path}


def test_failed_page_truncation_and_parser_change_keep_published_index(tmp_path, monkeypatch):
    import office_agent.research_pdf as module
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    plan = service.plan_research_index(path.name, SESSION, start_page=6, end_page=7)
    service.confirm_index(plan['plan_id'], SESSION, True)
    before = service.snapshot()
    real = module.read_pdf
    def broken(*args, **kwargs):
        report = real(*args, **kwargs)
        report['pages'][0]['text_truncated'] = True
        return report
    monkeypatch.setattr(module, 'read_pdf', broken)
    with pytest.raises(SourceError, match='pdf_text_truncated'):
        service.plan_research_index(path.name, SESSION, start_page=1, end_page=7)
    assert service.snapshot() == before
    monkeypatch.setattr(module, 'read_pdf', real)
    monkeypatch.setattr(module, '_ADAPTER_VERSION', 'fixture-adapter-next')
    assert service.research_status(path.name, SESSION, start_page=6, end_page=7)['status'] == 'profile_mismatch'
    monkeypatch.setattr(module, 'PDF_PARSER_VERSION', 'fixture:new-parser')
    assert service.research_status(path.name, SESSION, start_page=6, end_page=7)['status'] == 'profile_mismatch'
    assert service.snapshot() == before


def test_empty_page_and_source_change_stop_before_embedding(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(pdf_bytes(pages=7, text=''))
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    with pytest.raises(SourceError, match='pdf_text_missing'):
        service.plan_research_index(path.name, SESSION, start_page=7, end_page=7)
    path.write_bytes(long_pdf())
    plan = service.plan_research_index(path.name, SESSION, start_page=6, end_page=7)
    path.write_bytes(pdf_bytes(text='changed'))
    result = service.confirm_index(plan['plan_id'], SESSION, True)
    assert result['status'] == 'rejected' and result['wrote_nothing'] and not result['applied']
    assert not embedder.calls and not service.snapshot()['exists']


def test_authenticated_api_range_schema_and_cancel(tmp_path):
    api, app, auth, root, source, embedder, factories = make(tmp_path)
    app.path = root / 'paper.pdf'
    app.path.write_bytes(long_pdf())
    assert api.get_research_status(6, 7)['research']['pdf_coverage']['end_page'] == 7
    assert not embedder.calls
    for args in [(0, 1), (1, 21), (True, 3)]:
        assert api.plan_research_index(*args)['reason'] == 'rag_invalid'
    plan = api.plan_research_index(6, 7)['plan']
    assert api.confirm_rag_index(plan['plan_id'], False)['result']['wrote_nothing']
    auth.authenticated = False
    assert not api.get_research_status(6, 7)['ok']


def test_actual_character_budget_and_end_outside_document_do_not_publish(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(pdf_bytes(text='X' * 20100, width=600))
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    with pytest.raises(SourceError, match='pdf_text_truncated'):
        service.plan_research_index(path.name, SESSION, start_page=1, end_page=1)
    path.write_bytes(long_pdf())
    with pytest.raises(SourceError, match='invalid'):
        service.plan_research_index(path.name, SESSION, start_page=6, end_page=8)
    assert not service.snapshot()['exists']


def test_twenty_page_boundary_is_explicit_partial_coverage(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(pdf_bytes(pages=25, text='bounded development fixture'))
    result = PdfResearchSource(Workspace(tmp_path)).extract(path.name, start_page=1, end_page=20)
    assert result['pdf_coverage']['total_pages'] == 25 and not result['pdf_coverage']['complete']
    assert {block['group'] for block in result['blocks']} == {f'pdf-page:{page}' for page in range(1, 21)}
    assert len(result['parser_version']) <= 64
