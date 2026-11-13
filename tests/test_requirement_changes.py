from copy import deepcopy
from types import SimpleNamespace

import pytest

from office_agent.instruction_runtime import InstructionRuntime
from office_agent.requirement_history import active_requirements, change_requirement, event_digest, runtime_requirements
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from tests.test_paged_editor import draft, sync
from tests.test_paged_inbox import receive, reference


def change(service, row=None, action='replace', text='精简实验描述，保留全部数字和单位。', client='edit', expected=None):
    state=service.state();row=row or state['requirements'][0]
    return service.change_requirement(service.document_id,service.source_version,service.inbox_id,state['requirement_session_id'],
                                      client,row['id'],row['version'],state['instruction_revision'] if expected is None else expected,action,text)


def test_replace_is_immutable_versioned_and_idempotent(draft):
    service,floating,source=draft;original=source.read_bytes()
    first=receive(service,text='第一版要求。');row=service.state()['requirements'][0]
    result=change(service,row)
    duplicate=change(service,row,expected=1)
    assert duplicate['duplicate'] and duplicate['record']==result['record']
    assert floating.runtime.revision==2 and len(floating.runtime.messages)==1
    assert floating.runtime.messages[0]['text']=='第一版要求。'
    current=result['requirements'][0]
    assert current['text']=='精简实验描述，保留全部数字和单位。' and current['version']==2
    assert current['history']==[{'text':'第一版要求。','version':1,'status':'superseded'}]
    assert len(list(floating.app.editor.audit.glob('inbox/*/*/*.json')))==2
    assert source.read_bytes()==original and not floating.app.busy
    with pytest.raises(ValueError):change(service,action='cancel',text='',client='edit')


def test_all_cancelled_never_starts_or_regenerates_original_goal(draft):
    service,floating,_=draft;sync(service);target=reference(service)
    receive(service,text='不要再次执行的旧要求。')
    result=change(service,action='cancel',text='')
    assert not result['pending_requirements'] and result['requirements'][0]['status']=='cancelled'
    assert not active_requirements(floating.runtime) and not service.start()
    with pytest.raises(ValueError):service.start_again(False)
    with pytest.raises(ValueError):service.start_pending(service.document_id,1,target,2,result['requirement_session_id'])
    with pytest.raises(ValueError):change(service,client='twice',action='cancel',text='')
    receive(service,'fresh','一条新的有效要求。')
    assert [row['text'] for row in active_requirements(floating.runtime)]==['一条新的有效要求。']


def test_conflicting_versions_and_foreign_scope_do_not_change_history(draft):
    service,floating,_=draft;receive(service)
    row=service.state()['requirements'][0];inbox=service.inbox_id;session=floating.runtime.session_id
    receive(service,'second','补充要求。')
    with pytest.raises(ValueError):change(service,row,expected=1)
    for document,source,owner,actor in [('old',service.source_version,inbox,session),(service.document_id,'old',inbox,session),(service.document_id,service.source_version,'old',session),(service.document_id,service.source_version,inbox,'old')]:
        with pytest.raises(ValueError):service.change_requirement(document,source,owner,actor,'bad',row['id'],1,2,'cancel','')
    assert floating.runtime.revision==2 and not getattr(floating.runtime,'requirement_changes',[])


def test_failed_or_sensitive_change_has_no_receipt_or_revision(draft,monkeypatch):
    service,floating,_=draft;receive(service)
    original=floating.runtime.persist
    def fail(_):raise OSError('simulated disk failure')
    monkeypatch.setattr(floating.runtime,'persist',fail)
    with pytest.raises(OSError):change(service)
    assert floating.runtime.revision==1 and 'edit' not in floating.runtime.ids
    monkeypatch.setattr(floating.runtime,'persist',original)
    with pytest.raises(ValueError):change(service,text='sk-'+'a'*40)
    assert floating.runtime.revision==1 and len(runtime_requirements(floating.runtime))==1


def test_stale_plans_are_blocked_but_claimed_apply_and_publish_keep_snapshot(draft):
    service,floating,_=draft;sync(service);receive(service)
    service.proposals['old']={'id':'old','status':'pending','draft_revision':1,'instruction_revision':1}
    lease=service.prepare_apply('old',service.document_id,1)
    plan=service.plan(service.document_id,1)['plan']
    change(service)
    assert service.proposals['old']['status']=='stale' and not service.publisher.plans
    with pytest.raises(ValueError):service.publish(plan['id'],service.document_id,1,True)
    service.changed(service.document_id,2,lease['token'])
    service.applied(lease['token'],service.document_id,1,2)
    assert service.proposals['old']['status']=='adopted'
    sync(service,2);plan=service.plan(service.document_id,2)['plan']
    frozen=service.publisher.claim(plan['id'],document_id=service.document_id,draft_revision=2,confirmed=True)
    service.publishing=floating.app.busy=True
    change(service,action='cancel',text='',client='cancel')
    assert frozen['instruction_revision']==2 and frozen['draft_revision']==2
    assert service.state()['requirements'][0]['cancelled']


def test_restart_replays_latest_effective_requirements_without_permissions(draft):
    service,floating,source=draft
    receive(service,text='已被替代的要求。');change(service,text='当前有效要求。')
    receive(service,'second','稍后取消的要求。')
    row=service.state()['requirements'][1];change(service,row,action='cancel',text='',client='cancel')
    app=WebviewApp();api=SimpleNamespace(**{**vars(floating.api),'_app':app})
    restored=FloatingSession(api);restored.admit(str(source))
    tasks=restored.restore_drafts()['tasks']
    assert tasks[0]['messages']==['当前有效要求。'] and tasks[0]['cancelled_count']==1
    assert tasks[0]['changed_count']==2
    assert restored.runtime is None and not restored.document['ready']
    assert app.pending_id is None and app._read_authorized_path is None and not app.busy


def test_history_metadata_corruption_gap_and_incomplete_group_never_revive_cancelled(draft):
    service,floating,_=draft;receive(service);change(service,action='cancel',text='')
    directory=next(floating.app.editor.audit.glob('inbox/*/*'))
    target=directory/'2.json';raw=target.read_bytes()
    import json
    value=json.loads(raw);value['target_id']='altered';target.write_text(json.dumps(value),encoding='utf8')
    with pytest.raises(ValueError):floating.restore_drafts()
    target.write_bytes(b' '*20001)
    result=floating.restore_drafts();assert result['limited'] and not result['tasks']
    target.write_bytes(raw);first=directory/'1.json';first.unlink()
    with pytest.raises(ValueError):floating.restore_drafts()


def test_history_unknown_version_and_boolean_identity_are_rejected():
    runtime=InstructionRuntime('source');record=runtime.submit('one','初始要求。')
    event,_=change_requirement(runtime,'edit',record['id'],1,1,'replace','修改后的要求。')
    event=deepcopy(event);event['expected_version']=True;event['event_sha256']=event_digest(event)
    from office_agent.requirement_history import project_requirements
    with pytest.raises(ValueError):project_requirements([record,event])
    record['schema_version']=99
    with pytest.raises(ValueError):project_requirements([record])


def test_change_budget_includes_history_and_unchanged_edits_do_not_write(draft):
    service,floating,_=draft;receive(service,text='版本0')
    with pytest.raises(ValueError):change(service,text='版本0')
    for number in range(1,64):change(service,text='版本'+str(number),client='edit-'+str(number))
    assert floating.runtime.revision==64
    with pytest.raises(ValueError):change(service,text='超出预算',client='overflow')
    with pytest.raises(ValueError):receive(service,'overflow-add','新增也不得绕过预算。')


def test_cancel_before_model_worker_starts_blocks_even_model_construction(draft,monkeypatch):
    service,floating,_=draft;sync(service);target=reference(service)
    held=[];built=[]
    monkeypatch.setattr('office_agent.webui.paged_editor.threading.Thread.start',lambda thread:held.append(thread))
    floating.api._connections.capture=lambda capability:SimpleNamespace(public=lambda:{'model_id':'fixture','connection_name':'fixture'},context_window=8192)
    floating.api._connections.build_model=lambda route:built.append(True)
    receive(service,reference=target)
    assert len(held)==1 and floating.app.busy
    change(service,action='cancel',text='')
    held[0]._target(*held[0]._args,**held[0]._kwargs)
    assert not built and not floating.app.busy and floating.runtime.model_calls==0
