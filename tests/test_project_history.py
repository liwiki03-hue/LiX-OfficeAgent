from copy import deepcopy
import json

import pytest

from test_pdf_table import host
from test_project_workbench import imported


def save(api, name):
    assert api.workbench_action('rename', {'name':name})['ok']
    plan=api.plan_project_save()
    assert plan['ok'], plan
    saved=api.confirm_project_save(plan['result']['plan_id'],True)
    assert saved['ok'], saved
    return saved['result']['save_id']


def test_registered_history_is_bounded_read_only_and_does_not_scan_other_projects(tmp_path, monkeypatch):
    api, auth=host(tmp_path)
    imported(api)
    receipts=[]
    for index in range(9):
        if index==5:
            assert api.restore_project(receipts[-1])['ok']
        receipts.append(save(api,'Version '+str(index+1)))
    original=deepcopy(api._workbench.project)
    monkeypatch.setattr('office_agent.rag_sources.SourceExtractor.extract',lambda *args:pytest.fail('history must not read source bodies'))
    result=api.workbench_action('project_history',{})
    assert result['ok'], result
    data=result['result']
    assert [row['receipt'] for row in data['snapshots']]==list(reversed(receipts[1:]))
    assert data['truncated'] and len(data['snapshots'])==data['limit']==8
    assert not data['permissions_restored'] and api._workbench.project==original
    assert not api._workbench.evidence and not api._workbench.live_drafts
    assert all(set(row)=={'receipt','name','generation','materials','tasks','drafts','parent_receipt','current_saved'} for row in data['snapshots'])
    assert not api.workbench_action('project_history',{'receipt':'0'*32})['ok']


@pytest.mark.parametrize('fault',['missing','tampered','wrong_project','cycle','wrong_generation'])
def test_history_stops_on_missing_tampered_or_inconsistent_parent(tmp_path, fault):
    api, auth=host(tmp_path)
    imported(api)
    first=save(api,'First')
    latest=save(api,'Second')
    archive=api._workbench.archive
    if fault in {'missing','tampered'}:
        artifact=archive.root/('.office-rag-evidence-'+first)/'research-project.json'
        if fault=='missing':
            artifact.unlink()
        else:
            artifact.write_text('{}',encoding='utf8')
    else:
        original=archive.inspect
        def inspect(receipt):
            document,saved=original(receipt)
            if receipt==first:
                document=deepcopy(document)
                if fault=='wrong_project': document['id']='f'*32
                elif fault=='cycle': document['parent_receipt']=latest
                else: document['generation']=99
            return document,saved
        archive.inspect=inspect
    result=api.workbench_action('project_history',{})
    assert not result['ok'], result
    if fault not in {'missing','tampered'}:
        assert result['reason']=='professional_project_history_lineage'
    assert api._workbench.receipt==latest and not api._workbench.evidence


def test_history_requires_auth_and_restore_rechecks_selected_archive(tmp_path):
    api, auth=host(tmp_path)
    imported(api)
    receipt=save(api,'Saved')
    assert api.workbench_action('project_history',{})['ok']
    artifact=api._workbench.archive.root/('.office-rag-evidence-'+receipt)/'research-project.json'
    artifact.write_text('{}',encoding='utf8')
    assert not api.restore_project(receipt)['ok']
    auth[0]=False
    result=api.workbench_action('project_history',{})
    assert result['reason']=='not_authenticated'


def test_listed_snapshot_rewritten_with_matching_audit_still_requires_new_view(tmp_path):
    import hashlib
    api,auth=host(tmp_path)
    imported(api)
    receipt=save(api,'Original')
    assert api.workbench_action('project_history',{})['ok']
    archive=api._workbench.archive
    directory=archive.root/('.office-rag-evidence-'+receipt)
    artifact=directory/'research-project.json'
    document=json.loads(artifact.read_text(encoding='utf8'))
    document['name']='External replacement'
    raw=json.dumps(document,ensure_ascii=False).encode('utf8')
    artifact.write_bytes(raw)
    audit=json.loads((directory/'audit-verified.json').read_text(encoding='utf8'))
    audit['payload_sha256']=hashlib.sha256(raw).hexdigest()
    for name in ('audit-intent.json','audit-verified.json'):
        (directory/name).write_text(json.dumps(audit),encoding='utf8')
    assert archive.inspect(receipt)[0]['name']=='External replacement'
    assert not api.restore_project(receipt)['ok']
    assert api._workbench.project['name']=='Original'
    assert api.workbench_action('project_history',{})['ok']
    assert api.restore_project(receipt)['ok']
    assert not api._workbench.evidence and not api._workbench.live_drafts
