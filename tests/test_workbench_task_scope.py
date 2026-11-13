import json

import pytest

from office_agent.writes import digest
from test_pdf_header_rows import header_pdf
from test_project_history import save
from test_project_workbench import CSL
from test_research_assist import Model
from test_workbench_research import prepared


def completed(api):
    model = Model()
    api._research_assist.model_factory = lambda: model
    reply = api.start_workbench_assist({'question': 'Polish this selected observation.', 'mode': 'polish', 'outbound_approved': True})
    assert reply['ok'], reply
    job = api._research_assist.jobs[reply['result']['id']]
    job['thread'].join(5)
    assert not job['thread'].is_alive() and job['status'] == 'completed'
    return job, model


@pytest.mark.parametrize('change', ['new', 'restore', 'target', 'document'])
def test_old_task_and_save_plan_cannot_cross_workbench_context(tmp_path, change):
    api, auth, key, evidence = prepared(tmp_path)
    if change == 'target':
        from docx import Document
        document = Document(tmp_path / 'report.docx')
        document.add_paragraph('A different target in the same document.')
        document.save(tmp_path / 'report.docx')
        anchor = api.read_structure_anchors()['result']['anchors'][0]
        assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 0})['ok']
        assert api.workbench_action('capture', {'id': key, 'page': 0, 'limit': 1})['ok']
    receipt = save(api, 'Scoped task')
    job, model = completed(api)
    plan = api.plan_workbench_task(job['id'])
    assert plan['ok'], plan
    identifier = plan['result']['plan_id']
    if change == 'new':
        assert api.workbench_action('new_project', {'name': 'Another project'})['ok']
    elif change == 'restore':
        assert api.restore_project(receipt)['ok']
    elif change == 'target':
        anchors = api.read_structure_anchors()['result']['anchors']
        other = next(row for row in anchors if row['locator_id'] != api._workbench.target['locator_id'])
        assert api.workbench_action('target', {'locator_id': other['locator_id'], 'offset': 0})['ok']
    else:
        from docx import Document
        document = Document(tmp_path / 'report.docx')
        document.add_paragraph('Externally changed source.')
        document.save(tmp_path / 'report.docx')
        assert api.workbench_snapshot()['ok']
    assert not api.confirm_workbench_task(identifier, True)['ok']
    assert not api.plan_workbench_task(job['id'])['ok']
    assert not api.workbench_action('adopt_suggestion', {'task_id': job['id'], 'operation': 'replace'})['ok']
    assert identifier not in api._research_assist._task_plans
    assert identifier not in api._research_assist._task_archive._plans
    assert not api._workbench.project['tasks'] and not api._workbench.live_drafts
    assert len(model.inputs) == 1


def test_task_registration_uses_original_materials_not_later_capture(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    original = digest(tmp_path / 'report.docx')
    job, model = completed(api)
    response = api.workbench_action('import_csl', {'content': json.dumps([{**CSL[0], 'title': 'A later material', 'DOI': '10.1000/later'}])})
    assert response['ok'], response
    later = next(row for row in api._workbench.project['entries'] if row != key)
    assert api.workbench_action('review', {'id': later, 'decision': 'include', 'reason': 'Owned later fixture.', 'tags': []})['ok']
    (tmp_path / 'later.pdf').write_bytes(header_pdf([['Sample', 'Latency'], ['Identifier', '(ms)'], ['B', '40 ms']]))
    assert api.workbench_action('associate', {'id': later, 'filename': 'later.pdf'})['ok']
    assert api.workbench_action('capture', {'id': later, 'page': 0, 'limit': 1})['ok']
    assert {row['record_id'] for row in api._workbench.evidence.values()} == {key, later}
    plan = api.plan_workbench_task(job['id'])
    assert plan['ok'], plan
    result = api.confirm_workbench_task(plan['result']['plan_id'], True)
    assert result['ok'] and result['result']['status'] == 'saved', result
    task = api._workbench.project['tasks'][-1]
    assert task['record_ids'] == [key]
    assert task['source_stamps'] == {key: api._workbench.project['entries'][key]['link']['sha256']}
    assert task['target_version'] == original and digest(tmp_path / 'report.docx') == original
    assert len(model.inputs) == 1


def test_project_change_retires_save_preview_and_cancel_does_not_register(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    job, model = completed(api)
    plan = api.plan_workbench_task(job['id'])['result']['plan_id']
    assert api.workbench_action('rename', {'name': 'Changed metadata'})['ok']
    assert not api.confirm_workbench_task(plan, True)['ok']
    assert plan not in api._research_assist._task_plans
    assert plan not in api._research_assist._task_archive._plans
    fresh = api.plan_workbench_task(job['id'])
    assert fresh['ok'], fresh
    cancelled = api.confirm_workbench_task(fresh['result']['plan_id'], False)
    assert cancelled['ok'] and cancelled['result']['status'] == 'rejected', cancelled
    assert not api._workbench.project['tasks'] and len(model.inputs) == 1


def test_running_task_context_retirement_does_not_leave_permanent_busy(tmp_path):
    import threading
    api, auth, key, evidence = prepared(tmp_path)
    entered, release = threading.Event(), threading.Event()
    class Blocking(Model):
        def suggest(self, payload, cancelled):
            entered.set()
            assert release.wait(3)
            return super().suggest(payload, cancelled)
    api._research_assist.model_factory = Blocking
    started = api.start_workbench_assist({'question': 'Polish the selected observation.', 'mode': 'polish', 'outbound_approved': True})
    assert started['ok'], started
    job = api._research_assist.jobs[started['result']['id']]
    try:
        assert entered.wait(2)
        assert api.workbench_action('new_project', {'name': 'Next context'})['ok']
        assert job['cancel'].is_set() and job['status'] == 'cancelled'
        assert job['id'] not in api._rag._research_dependencies
    finally:
        release.set()
        job['thread'].join(5)
        assert not job['thread'].is_alive()
    anchor = api.read_structure_anchors()['result']['anchors'][0]
    assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 0})['ok']
    fresh, model = completed(api)
    assert fresh['id'] != job['id'] and len(model.inputs) == 1
