import json

import pytest

from office_agent.rag_workspace import WorkspaceRag
from office_agent.rag import RagError
from office_agent.writes import Workspace
from test_research_reading import long_pdf
from test_rag_workspace import SyntheticEmbedder, SESSION


def test_ranges_increment_overlap_restart_remove_query_and_undo(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    original = path.read_bytes()
    embedding = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedding)
    plan = service.plan_pdf_ranges(path.name, [[1, 2]], SESSION)
    assert not embedding.calls
    assert service.confirm_index(plan['plan_id'], SESSION, False)['wrote_nothing']
    plan = service.plan_pdf_ranges(path.name, [[1, 2]], SESSION)
    first = service.confirm_index(plan['plan_id'], SESSION, True)
    assert first['status'] == 'saved'
    service = WorkspaceRag(Workspace(tmp_path), embedding)
    plan = service.plan_pdf_ranges(path.name, [[1, 2], [2, 2], [6, 7]], SESSION)
    assert plan['reused_pages'] == 2 and plan['new_pages'] == 2
    added = service.confirm_index(plan['plan_id'], SESSION, True)
    assert added['status'] == 'saved'
    status = service.pdf_ranges_status(path.name, SESSION)
    assert status['pages'] == [1, 2, 6, 7] and status['gaps'] == [3, 4, 5]
    assert status['native_pages'] == [1, 2, 6, 7] and not status['ocr_pages'] and not status['complete']
    calls = len(embedding.calls)
    with pytest.raises(RagError, match='rag_index_current'):
        service.plan_pdf_ranges(path.name, [[1, 2], [6, 7]], SESSION)
    assert len(embedding.calls) == calls
    plan = service.query_pdf_ranges(path.name, 'measured latency', SESSION, 10)
    saved = service.confirm_query(plan['plan_id'], SESSION, True)
    assert any('pdf-page:1' in h['location'] for h in saved['hits'])
    assert any('pdf-page:7' in h['location'] for h in saved['hits'])
    count = service.snapshot()['chunk_count']
    remove = service.plan_pdf_ranges(path.name, [[6, 7]], SESSION)
    assert remove['removed_pages'] == 2 and remove['new_pages'] == 0
    removed = service.confirm_index(remove['plan_id'], SESSION, True)
    assert removed['status'] == 'saved' and len(embedding.calls) == calls + 1
    assert service.snapshot()['chunk_count'] < count
    undo = service.plan_index_rollback(removed['save_id'], SESSION)
    service.confirm_index_rollback(undo['rollback_id'], SESSION, True)
    assert service.pdf_ranges_status(path.name, SESSION)['pages'] == [1, 2, 6, 7]
    assert '20 ms' in service.read_evidence(saved['save_id'], SESSION)['text']
    assert path.read_bytes() == original


def test_stale_and_expired_collection_never_embed(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    embedding = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedding)
    plan = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    service._ranges.pending[plan['plan_id']]['created'] -= 301
    assert service.confirm_index(plan['plan_id'], SESSION, True)['wrote_nothing']
    plan = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    path.write_bytes(path.read_bytes() + b'\n')
    assert service.confirm_index(plan['plan_id'], SESSION, True)['wrote_nothing']
    assert not embedding.calls and not service.snapshot()['exists']


@pytest.mark.parametrize('ranges', [[[1, 21]], [[True, 2]], [[0, 1]], [[2, 1]]])
def test_ranges_budget_and_types(tmp_path, ranges):
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    with pytest.raises(RagError):
        service.plan_pdf_ranges('paper.pdf', ranges, SESSION)


def test_persistent_receipt_assistant_does_not_reembed_or_replan(tmp_path):
    from types import SimpleNamespace
    from docx import Document
    from office_agent.webui.bridge import Api
    from office_agent.writes import Editor
    from test_research_assist import Model
    root = tmp_path
    path = root / 'draft.docx'
    doc = Document()
    doc.add_paragraph('The experiment measured latency.')
    doc.save(path)
    (root / 'paper.pdf').write_bytes(long_pdf())
    app = SimpleNamespace(path=path, editor=Editor(root, require_confirmation=True), busy=False)
    auth = SimpleNamespace(root=root / 'auth', status=lambda: {'authenticated': True})
    api = Api(app, auth=auth, settings_root=root)
    embedding = SyntheticEmbedder()
    api._rag._factory = lambda check: WorkspaceRag(app.editor.space, embedding, scope_check=check)
    plan = api.plan_pdf_ranges([[1, 2], [6, 7]], 'paper.pdf')['plan']
    assert api.confirm_rag_index(plan['plan_id'], True)['result']['status'] == 'saved'
    query = api.query_pdf_ranges('latency', 10, 'paper.pdf')['plan']
    receipt = api.confirm_rag_query(query['plan_id'], True)['result']['save_id']
    source = api.read_research_assist()['result']
    class SavedModel(Model):
        def plan(self, *args):
            pytest.fail('persistent receipt must not re-enter temporary retrieval')
    model = SavedModel()
    api._research_assist.model_factory = lambda: model
    api._research_assist.retriever_factory = lambda: pytest.fail('no temporary embedder')
    before = len(embedding.calls)
    payload = {'question': 'Review the measured latency.', 'mode': 'critique', 'source_sha256': source['document']['source_sha256'],
               'locator_ids': [source['items'][0]['locator_id']], 'capture_ids': [], 'evidence_save_id': receipt,
               'protected_terms': [], 'outbound_approved': True}
    started = api.start_research_assist(payload)
    assert started['ok'], started
    job = api._research_assist.jobs[started['result']['id']]
    job['thread'].join(5)
    assert not job['thread'].is_alive() and job['status'] == 'completed'
    assert len(model.inputs) == 1 and len(embedding.calls) == before


def test_aggregate_parse_budget_does_not_reset_per_page(tmp_path, monkeypatch):
    from copy import deepcopy
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    embedding = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedding)
    parsed = service._pdf_source.extract(path.name, start_page=1, end_page=1)
    clock = [1.0]
    def extract(name, *, start_page, end_page):
        clock[0] += 31
        result = deepcopy(parsed)
        for block in result['blocks']:
            block['group'] = 'pdf-page:' + str(start_page)
            block['location'] = block['location'].replace('pdf-page:1', 'pdf-page:' + str(start_page))
        return result
    monkeypatch.setattr(service._pdf_source, 'extract', extract)
    monkeypatch.setattr('office_agent.research_ranges.time.monotonic', lambda: clock[0])
    with pytest.raises(RagError, match='rag_capacity'):
        service.plan_pdf_ranges(path.name, [[1, 2]], SESSION)
    assert not embedding.calls and not service.snapshot()['exists']


def test_new_preview_retires_same_session_old_range_plan(tmp_path):
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    first = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    second = service.plan_pdf_ranges(path.name, [[1, 2]], SESSION)
    assert len(service._ranges.pending) == 1
    assert service.confirm_index(first['plan_id'], SESSION, True)['wrote_nothing']
    assert service.confirm_index(second['plan_id'], SESSION, True)['status'] == 'saved'


def test_api_model_name_does_not_reuse_different_endpoint_or_credential_scope(tmp_path):
    from types import SimpleNamespace
    path = tmp_path / 'paper.pdf'
    path.write_bytes(long_pdf())
    embedding = SyntheticEmbedder()
    embedding.identity = {'channel': 'api', 'model': 'same-model-alias', 'dimension': 2}
    embedding.config = SimpleNamespace(base_url='https://first.example.test/v1', api_key='fixture-only-credential-A')
    service = WorkspaceRag(Workspace(tmp_path), embedding)
    plan = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    assert service.confirm_index(plan['plan_id'], SESSION, True)['status'] == 'saved'
    previous = service.snapshot()
    embedding.config = SimpleNamespace(base_url='https://second.example.test/v1', api_key='fixture-only-credential-A')
    assert service.pdf_ranges_status(path.name, SESSION)['status'] == 'stale'
    plan = service.plan_pdf_ranges(path.name, [[1, 1]], SESSION)
    assert plan['reused_pages'] == 0 and plan['new_pages'] == 1
    assert 'fixture-only-credential' not in json.dumps(plan)
    calls = len(embedding.calls)
    embedding.config = SimpleNamespace(base_url='https://second.example.test/v1', api_key='fixture-only-credential-B')
    assert service.confirm_index(plan['plan_id'], SESSION, True)['wrote_nothing']
    assert len(embedding.calls) == calls and service.snapshot() == previous
    manifest = service._store.corpus_manifest(service._ranges.bind(path.name)[0])
    assert 'fixture-only-credential' not in json.dumps(manifest)
