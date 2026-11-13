import json

from docx import Document

from test_pdf_header_rows import header_pdf
from test_pdf_table import host
from test_project_workbench import imported
from test_professional_edits import fake_render


def prepared(tmp_path):
    api, auth = host(tmp_path)
    key = imported(api)
    assert api.workbench_action('review', {'id': key, 'decision': 'include', 'reason': 'Owned source.', 'tags': ['baseline']})['ok']
    (tmp_path / 'table.pdf').write_bytes(header_pdf([['Sample','Latency'],['Identifier','(ms)'],['A','20 ms']]))
    assert api.workbench_action('associate', {'id': key, 'filename': 'table.pdf'})['ok']
    anchor = api.read_structure_anchors()['result']['anchors'][0]
    assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 0})['ok']
    captured = api.workbench_action('capture', {'id': key, 'page': 0, 'limit': 1})
    assert captured['ok'], captured
    return api, auth, key, captured['result']['evidence'][0]


def review_comment(api, evidence):
    response = api.workbench_action('queue', {'operation':'comment','text':'Check the source condition.','evidence_ids':[evidence['id']]})
    assert response['ok'], response
    row = response['result']['drafts'][-1]
    review = api.workbench_action('queue_review', {'id':row['id'],'accepted':True,'reviewed':True,'relation':'supports',
        'applicability':'Sample A','limitation':'Developer fixture only','text':row['text']})
    assert review['ok'], review
    return row


def test_capture_is_not_human_reading_and_search_has_two_round_budget(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    entry = api._workbench.project['entries'][key]
    assert entry['parsed_pages'] == [1] and entry['read_pages'] == []
    assert api.workbench_action('mark_read', {'id':key,'reviewed':True})['ok']
    assert api._workbench.project['entries'][key]['read_pages'] == [1]
    question = 'Do not infer latency below 20 ms in 2024 under wet conditions'
    for index in (1,2):
        reply = api.workbench_action('search', {'question':question,'focus':'A','dense':False,'read_ids':[]})
        assert reply['ok'], reply
        report = reply['result']
        assert report['effective_query'].startswith(question)
        assert report['rounds_used'] == index and report['model_calls'] == 0
        assert not report['queries_are_evidence'] and not report['absence_of_contradictions_proven']
    denied = api.workbench_action('search', {'question':question,'focus':'other','dense':False,'read_ids':[]})
    assert denied['reason'] == 'professional_project_search_budget'


def test_source_change_compares_registered_quote_without_rewriting_history(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    row = review_comment(api, evidence)
    original = json.dumps(api._workbench.project['claims'])
    (tmp_path / 'table.pdf').write_bytes(header_pdf([['Sample','Latency'],['Identifier','(ms)'],['A','25 ms']]))
    response = api.workbench_action('impact', {'record_ids':[key]})
    assert response['ok'], response
    claims = [row for row in response['result']['impacts'] if row['kind']=='claim']
    assert claims and claims[0]['content_comparison']['status']=='quote_changed_or_moved_requires_review'
    assert response['result']['content_windows_read']==1
    assert not response['result']['documents_rewritten'] and not response['result']['history_deleted']
    assert json.dumps(api._workbench.project['claims'])==original
    assert api._workbench.project['drafts'][0]['status']=='recheck'
    assert not api.plan_workbench_queue()['ok']


def test_restored_queue_requires_fresh_material_and_explicit_recheck(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    review_comment(api,evidence)
    plan=api.plan_project_save()['result']
    saved=api.confirm_project_save(plan['plan_id'],True)['result']['save_id']
    api._workbench.invalidate()
    restored=api.restore_project(saved)
    assert restored['ok'],restored
    assert restored['result']['drafts'][0]['status']=='recheck'
    assert not api.plan_workbench_queue()['ok']
    assert not api._workbench.live_drafts and not api._workbench.captures


def test_existing_numeric_guard_still_rejects_header_only_unit_as_direct_claim(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    (tmp_path / 'table.pdf').write_bytes(header_pdf())
    assert api.workbench_action('associate', {'id':key,'filename':'table.pdf'})['ok']
    result=api.workbench_action('capture', {'id':key,'page':0,'limit':1})
    assert result['ok'],result
    current=result['result']['evidence'][-1]
    response=api.workbench_action('queue',{'operation':'insert','text':'The reported latency is 20 ms.','evidence_ids':[current['id']]})
    assert response['ok'],response
    item=response['result']['drafts'][0]
    assert api.workbench_action('queue_review',{'id':item['id'],'accepted':True,'reviewed':True,'relation':'supports',
        'applicability':'Sample A','limitation':'Owned only','text':item['text']})['ok']
    target=api._workbench.target['locator_id']
    denied=api.plan_workbench_queue({'bibliography_anchor':target})
    assert denied['reason']=='professional_unsupported_fact',denied


def test_project_snapshot_removal_is_separate_and_keeps_source(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    plan=api.plan_project_save()['result']
    receipt=api.confirm_project_save(plan['plan_id'],True)['result']['save_id']
    remove=api.plan_project_removal(receipt)
    assert remove['ok'],remove
    declined=api.confirm_project_removal(remove['result']['rollback_id'],False)
    assert declined['result']['status']=='rejected'
    remove=api.plan_project_removal(receipt)['result']
    removed=api.confirm_project_removal(remove['rollback_id'],True)
    assert removed['result']['status']=='rolled_back',removed
    assert (tmp_path/'table.pdf').is_file() and (tmp_path/'report.docx').is_file()
