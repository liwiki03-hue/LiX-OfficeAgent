from docx import Document

from office_agent.writes import Editor
from test_rag_api import make, index
from test_professional_edits import fake_render
from test_research_assist import Model, request, complete


def test_old_evidence_rollback_grant_and_rag_only_output_dependency(tmp_path, monkeypatch):
    api, app, auth, root, source, embedder, factories = make(tmp_path)
    document = Document()
    document.add_paragraph('The experiment measured latency under controlled conditions.')
    path = root / 'draft.docx'
    document.save(path)
    app.path = path
    app.editor = Editor(root, require_confirmation=True, render_validation=True)
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    index(api)
    query = api.plan_rag_query('latency', 1)['plan']
    saved_evidence = api.confirm_rag_query(query['plan_id'], True)['result']
    existing_grant = api.plan_rag_rollback('evidence', saved_evidence['save_id'])['plan']
    service = api._research_assist
    service.model_factory = lambda: Model()
    task = complete(service, request(service, evidence_save_id=saved_evidence['save_id']))
    assert task['status'] == 'completed', task
    assert api.confirm_rag_rollback('evidence', existing_grant['rollback_id'], True)['reason'] == 'rag_evidence_in_use'
    assert existing_grant['rollback_id'] not in api._rag._store._evidence._plans
    proposal = task['result']['suggestions'][0]
    plan = service.plan({'task_id': task['id'], 'edits': [{'locator_id': proposal['locator_id'], 'new_text': proposal['new_text'], 'reason': proposal['reason'], 'reviewed': True}]})
    assert plan['ok'], plan
    saved = service.confirm(plan['result']['id'], True)
    assert saved['result']['status'] == 'saved', saved
    lease = api._rag._research_dependencies[saved['result']['save_id']]
    assert lease['expires_at'] is None
    assert service.cancel(task['id'])['ok']
    assert api.plan_rag_rollback('evidence', saved_evidence['save_id'])['reason'] == 'rag_evidence_in_use'
    assert service.rollback(saved['result']['save_id'], True)['result']['status'] == 'rolled_back'
    rollback = api.plan_rag_rollback('evidence', saved_evidence['save_id'])
    assert rollback['ok'], rollback
    assert api.confirm_rag_rollback('evidence', rollback['plan']['rollback_id'], True)['result']['restored']
