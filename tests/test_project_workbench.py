import json

import pytest

from office_agent.project_library import canonical_csl, parse_csl
from office_agent.writes import GuardError, digest
from test_pdf_table import host
from test_professional_edits import fake_render

CSL = [{'id': 'external-1', 'type': 'article-journal', 'title': 'Latency experiment',
        'author': [{'family': 'Chen', 'given': 'Lin'}], 'issued': {'date-parts': [[2024]]}, 'DOI': '10.1000/latency'}]


def imported(api):
    result = api.workbench_action('import_csl', {'content': json.dumps(CSL)})
    assert result['ok'], result
    return result['result']['snapshot']['entries'][0]['id']


def test_project_csl_roundtrip_duplicates_conflicts_and_no_body_read(tmp_path):
    api, auth = host(tmp_path)
    key = imported(api)
    repeated = api.workbench_action('import_csl', {'content': json.dumps(CSL)})
    assert repeated['result']['added'] == 0 and repeated['result']['identical_skipped'] == 1
    changed = [{**CSL[0], 'title': 'Latency revised experiment'}]
    response = api.workbench_action('import_csl', {'content': json.dumps(changed)})
    assert response['ok']
    group = response['result']['snapshot']['groups'][0]
    assert len(group['variants']) == 2 and 'title' in group['conflicts']
    exported = api.workbench_action('export_csl', {})['result']['csl_json']
    assert len(parse_csl(exported)) == 2
    assert canonical_csl(CSL[0])['id'] == key
    assert not api._workbench.evidence and not api._grounded._captures


@pytest.mark.parametrize('content', ['[{"type":"article","title":"A","shell":"x"}]',
                                    '[{"type":"article","title":"A","title":"B"}]',
                                    '[{"type":"article","title":"A","URL":"file:///secret"}]',
                                    '[{"type":"article","title":"A","URL":"https://example.com/?key=x"}]',
                                    '[{"type":"article","title":"A","issued":{"date-parts":[[2024,2,31]]}}]'])
def test_strict_import_rejects_unsafe_or_ambiguous_data(content):
    with pytest.raises(GuardError):
        parse_csl(content)


def test_project_save_cancel_restore_never_restores_permissions(tmp_path):
    api, auth = host(tmp_path)
    imported(api)
    plan = api.plan_project_save()
    assert plan['ok'], plan
    cancelled = api.confirm_project_save(plan['result']['plan_id'], False)
    assert cancelled['result']['status'] == 'rejected'
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    plan = api.plan_project_save()
    result = api.confirm_project_save(plan['result']['plan_id'], True)
    assert result['ok'], result
    saved = result['result']['save_id']
    api._workbench.invalidate()
    restored = api.restore_project(saved)
    assert restored['ok'], restored
    assert restored['result']['total'] == 1 and not restored['result']['permissions_restored']
    assert not api._workbench.evidence and not api._workbench.live_drafts


def test_project_change_rejects_old_save_and_old_write_plan(tmp_path, monkeypatch):
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    api, auth = host(tmp_path)
    key = imported(api)
    plan = api.plan_project_save()['result']
    api.workbench_action('rename', {'name': 'Another title'})
    assert not api.confirm_project_save(plan['plan_id'], True)['ok']
    structure = api.read_structure_anchors()['result']
    anchor = structure['anchors'][0]
    assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 0})['ok']
    queued = api.workbench_action('queue', {'operation': 'comment', 'text': 'Check the condition.', 'evidence_ids': []})
    assert queued['ok'], queued
    row = queued['result']['drafts'][0]
    review = api.workbench_action('queue_review', {'id': row['id'], 'accepted': True, 'reviewed': True,
        'relation': 'insufficient', 'applicability': '', 'limitation': '', 'text': row['text']})
    assert review['ok'], review
    result = api.plan_workbench_queue()
    assert result['ok'], result
    original = digest(tmp_path / 'report.docx')
    api.workbench_action('rename', {'name': 'Changed after plan'})
    refused = api.confirm_professional_edit(result['result']['id'], True)
    assert refused['reason'] == 'professional_project_changed', refused
    assert digest(tmp_path / 'report.docx') == original and not list(tmp_path.glob('*.agent-*.docx'))
