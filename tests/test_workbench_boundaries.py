from collections import Counter

from office_agent.research_capture_retrieval import retrieve_captures
from test_professional_browser import _browser, STATIC
from test_workbench_research import prepared
from test_research_assist import Model


def test_known_paper_versions_do_not_fill_context_and_explicit_context_is_retained():
    evidence = {}
    for paper in range(3):
        for fragment in range(8):
            key = f'{paper}:{fragment}'
            evidence[key] = {'id': key, 'original_text': 'latency experiment 20 ms', 'location': key,
                'source_sha256': str(paper), 'title': str(paper), 'page': 1,
                'capture_id': key, 'diversity_key': 'doi:' + str(paper)}
    selected, report = retrieve_captures(evidence, 'latency', ['latency'], [])
    assert Counter(row['diversity_key'] for row in selected.values()) == {'doi:0': 3, 'doi:1': 3, 'doi:2': 3}
    assert report['diversity_policy'] == 'known_paper_max_3_except_explicit_context'
    selected, _ = retrieve_captures(evidence, 'latency', ['latency'], ['0:7'])
    assert '0:7' in selected


def test_document_change_retires_body_and_numeric_proof_but_preserves_library(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    workbench = api._workbench
    workbench.last_numeric = {'private': True}
    workbench.numeric_record_id = key
    workbench.document_token = None
    result = api.workbench_snapshot()
    assert result['ok'], result
    assert key in workbench.project['entries']
    assert not workbench.evidence and not workbench.captures and not workbench.live_drafts
    assert workbench.last_numeric is None and workbench.numeric_record_id is None


def test_cached_image_retry_cannot_resurrect_after_source_clear_or_logout(tmp_path, monkeypatch):
    browser = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, replies = next(browser)
        page.add_script_tag(path=str(STATIC / 'pdf_preview.js'))
        result = page.evaluate('''async () => {
          const image=document.createElement('img');document.body.append(image);
          const canvas=document.createElement('canvas');canvas.width=2;canvas.height=2;
          const data=canvas.toDataURL();let current=true;
          const first=await OfficePdfPreview.show(image,data,()=>current);
          const ignored=!(await OfficePdfPreview.show(image,data,()=>false))&&image.dataset.previewState==='ready'&&!image.hidden;
          const pending=OfficePdfPreview.retry(image);OfficePdfPreview.clear(image);
          const cleared=!(await pending)&&!image.hasAttribute('src')&&image.hidden;
          await OfficePdfPreview.show(image,data,()=>current);
          const late=OfficePdfPreview.retry(image);current=false;
          window.dispatchEvent(new Event('officeauth:unauthenticated'));
          return {first,ignored,cleared,logout:!(await late)&&!image.hasAttribute('src')&&image.hidden};
        }''')
        assert result == {'first': True, 'ignored':True, 'cleared': True, 'logout': True}
        assert calls == []
    finally:
        browser.close()


def test_workbench_model_planning_keeps_original_conditions_and_never_approves(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    class Planner(Model):
        def plan(self, payload, cancelled):
            return {'queries': ['latency'], 'read_ids': []}
    model = Planner()
    api._research_assist.model_factory = lambda: model
    api._research_assist.retriever_factory = lambda: None
    question = 'Do not generalize 20 ms in 2024 beyond sample A'
    started = api.start_workbench_assist({'question': question, 'mode': 'polish', 'outbound_approved': True})
    assert started['ok'], started
    job = api._research_assist.jobs[started['result']['id']]
    job['thread'].join(5)
    assert not job['thread'].is_alive() and job['status'] == 'completed', job.get('error')
    assert job['retrieval']['queries'] == [question]
    assert job['calls'] <= 3
    adopted = api.workbench_action('adopt_suggestion', {'task_id': job['id'], 'operation': 'replace'})
    assert adopted['ok'], adopted
    assert adopted['result']['drafts'][-1]['status'] == 'pending'


def test_relation_history_and_retraction_cannot_become_body_permission(tmp_path):
    import json
    from test_project_workbench import CSL
    from office_agent.project_library import validate
    api, auth, key, evidence = prepared(tmp_path)
    response = api.workbench_action('import_csl', {'content': json.dumps([{**CSL[0], 'title': 'Correction notice', 'DOI': '10.1000/correction'}])})
    assert response['ok'], response
    other = next(value for value in api._workbench.project['entries'] if value != key)
    related = api.workbench_action('relate', {'id':other, 'target_id':key, 'relation':'retraction_of', 'reason':'Explicit developer fixture relation.'})
    assert related['ok'], related
    assert validate(api._workbench.project)['entries'][key]['item']['retraction_status'] == 'reported_retracted'
    denied = api.workbench_action('capture', {'id':key, 'page':0, 'limit':1})
    assert denied['reason'] == 'professional_project_not_included'


def test_saved_material_project_guard_rechecks_decision_without_capture_lease(tmp_path, monkeypatch):
    import pytest
    from office_agent.writes import GuardError
    api, auth, key, evidence = prepared(tmp_path)
    api._workbench.captures.clear()
    api._workbench.capture_windows.clear()
    api._workbench.evidence[evidence['id']]['capture_id'] = None
    callbacks = {}
    def started(payload, **kwargs):
        callbacks.update(kwargs)
        return {'ok': True, 'result': {'status':'started'}}
    monkeypatch.setattr(api._research_assist, 'start', started)
    reply = api.start_workbench_assist({'question':'Check selected evidence','mode':'critique','outbound_approved':True})
    assert reply['ok'], reply
    callbacks['project_guard']()
    api._workbench.project['entries'][key]['decision'] = 'exclude'
    with pytest.raises(GuardError, match='professional_project_not_included'):
        callbacks['project_guard']()
