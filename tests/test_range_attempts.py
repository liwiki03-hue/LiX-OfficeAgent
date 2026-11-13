import pytest

from office_agent.rag import RagError
from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import Workspace
from test_rag_workspace import SyntheticEmbedder, SESSION
from test_research_reading import long_pdf


def test_failed_preview_retires_plan_and_reports_unprocessed_pages(tmp_path, monkeypatch):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    plan = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    extract = service._pdf_source.extract
    def parse(name, *, start_page, end_page):
        if start_page == 3:
            raise RagError('pdf_text_missing')
        return extract(name, start_page=start_page, end_page=end_page)
    monkeypatch.setattr(service._pdf_source, 'extract', parse)
    with pytest.raises(RagError, match='pdf_text_missing'):
        service.plan_pdf_ranges(path.name, [[1, 4]], SESSION)
    result = service.pdf_ranges_status(path.name, SESSION)
    attempt = result['attempt']
    assert attempt['parsed_pages'] == [1, 2] and attempt['failed_pages'] == [3]
    assert attempt['unprocessed_pages'] == [4] and attempt['state'] == 'failed'
    assert attempt['reason'] == 'pdf_text_missing' and attempt['source_is_current']
    assert result['indexed_pages'] == [] and result['attempt_storage'] == 'session_memory_only'
    assert service.confirm_index(plan['plan_id'], SESSION, True)['wrote_nothing']
    assert not service._rag.embedder.calls
    (tmp_path / 'other.pdf').write_bytes(long_pdf())
    assert service.pdf_ranges_status('other.pdf', SESSION)['attempt'] is None
    path.write_bytes(path.read_bytes() + b'\n')
    assert not service.pdf_ranges_status(path.name, SESSION)['attempt']['source_is_current']
    restarted = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    assert restarted.pdf_ranges_status(path.name, SESSION)['attempt'] is None


def test_attempt_reuse_cancel_expiry_and_status_isolation(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    plan = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    assert service.confirm_index(plan['plan_id'], SESSION, True)['status'] == 'saved'
    plan = service.plan_pdf_ranges(path.name, [[1, 2]], SESSION)
    attempt = service.pdf_ranges_status(path.name, SESSION)['attempt']
    assert attempt['reused_pages'] == [1] and attempt['parsed_pages'] == [1, 2]
    attempt['parsed_pages'].clear()
    assert service.pdf_ranges_status(path.name, SESSION)['attempt']['parsed_pages'] == [1, 2]
    assert service.confirm_index(plan['plan_id'], SESSION, False)['wrote_nothing']
    assert service.pdf_ranges_status(path.name, SESSION)['attempt']['state'] == 'cancelled'
    plan = service.plan_pdf_ranges(path.name, [[1, 2]], SESSION)
    service._ranges.pending[plan['plan_id']]['created'] -= 301
    assert service.pdf_ranges_status(path.name, SESSION)['attempt']['state'] == 'expired'
    assert service.pdf_ranges_status(path.name, SESSION)['indexed_pages'] == [1]
