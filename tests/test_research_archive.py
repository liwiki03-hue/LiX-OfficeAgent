from pathlib import Path

from test_research_assist import setup, request, complete


def test_task_save_cancel_restore_without_authority_and_owned_undo(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    task = complete(service, request(service))
    proposal = task['result']['suggestions'][0]
    drafts = [{'locator_id': proposal['locator_id'], 'new_text': proposal['new_text'] + ' Human draft.', 'accepted': True, 'reviewed': True}]
    plan = service.save_task({'task_id': task['id'], 'drafts': drafts})
    assert plan['ok'], plan
    assert service.confirm_task(plan['result']['plan_id'], False)['result']['status'] == 'rejected'
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    plan = service.save_task({'task_id': task['id'], 'drafts': drafts})['result']
    saved = service.confirm_task(plan['plan_id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    before = len(model.inputs)
    restored = service.restore_task(saved['result']['save_id'])
    assert restored['ok'], restored
    data = restored['result']
    assert data['drafts'][0]['reviewed'] is True and data['task']['request']['outbound_approved'] is False
    assert data['task']['status'] == 'restored' and data['resumable_task']
    assert not data['restores_approval'] and not data['restores_rollback_authority']
    assert len(model.inputs) == before
    current = data['resumable_task']
    assert current['result']['suggestions'][0]['assessment'] == 'unverified'
    assert service.plan({'task_id': current['id'], 'edits': [{'locator_id': proposal['locator_id'], 'new_text': proposal['new_text'], 'reason': proposal['reason'], 'reviewed': False}]})['reason'] == 'professional_assist_review_required'
    assert service.rollback_task(saved['result']['save_id'], True)['result']['status'] == 'rolled_back'


def test_restart_restores_comparison_and_source_changes_stale_without_delete_grant(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    task = complete(service, request(service))
    plan = service.save_task({'task_id': task['id'], 'drafts': []})['result']
    saved = service.confirm_task(plan['plan_id'], True)['result']
    service.invalidate()
    restored = service.restore_task(saved['save_id'])
    assert restored['ok'] and restored['result']['resumable_task'], restored
    assert not service.rollback_task(saved['save_id'], True)['ok']
    from docx import Document
    document = Document(path)
    document.paragraphs[0].text += ' External change.'
    document.save(path)
    restored = service.restore_task(saved['save_id'])
    assert restored['ok'] and 'source_changed' in restored['result']['stale_reasons']
    assert restored['result']['resumable_task'] is None
    directory = tmp_path / ('.office-rag-evidence-' + saved['save_id'])
    (directory / 'research-task.json').write_text('{}')
    assert not service.restore_task(saved['save_id'])['ok']
