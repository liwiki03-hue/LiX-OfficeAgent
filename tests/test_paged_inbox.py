from types import SimpleNamespace

import pytest

from office_agent.paged_documents import GuardError
from office_agent.settings_store import SettingsError
from tests.test_paged_editor import draft, sync


def receive(service, client_id='one', text='精简文字，保留数字和单位。', reference=None, session=None):
    runtime = service.floating.runtime
    return service.submit(client_id,text,reference,service.document_id,service.source_version,
                          session if session is not None else runtime.session_id if runtime else '',service.inbox_id)


def reference(service):
    return service.reference(service.document_id,service.revision,{'node_id':'node-1','start':0,'end':9})['reference']['id']


def test_no_reference_persists_once_without_model_or_file_changes(draft):
    service,floating,source = draft
    original = source.read_bytes()
    first = receive(service)
    again = receive(service,session='')
    assert first['pending_reason']=='reference' and again['duplicate']
    assert first['record']==again['record'] and floating.runtime.revision==1
    assert len(list(floating.app.editor.audit.glob('inbox/*/*/*.json')))==1
    assert floating.runtime.applied==0 and not floating.app.busy
    assert source.read_bytes()==original
    with pytest.raises(ValueError,match='相同消息'):receive(service,text='更改内容')


def test_old_reference_waits_and_explicit_processing_binds_requirement_version(draft,monkeypatch):
    service,floating,_ = draft
    sync(service);old = reference(service)
    service.changed(service.document_id,2)
    result = receive(service,reference=old)
    assert result['pending_reason']=='reference'
    sync(service,2);current = reference(service)
    starts=[]
    monkeypatch.setattr(service,'start',lambda:starts.append(True) or True)
    receive(service,'two','保留原有年份。')
    with pytest.raises(GuardError):service.start_pending(service.document_id,2,current,1,result['session_id'])
    assert not starts
    assert service.start_pending(service.document_id,2,current,2,result['session_id'])['started']
    assert starts==[True]


def test_missing_model_retains_receipt_and_never_retries_duplicate(draft):
    service,floating,_ = draft
    sync(service);target=reference(service);calls=[]
    def capture(_):
        calls.append(True)
        raise SettingsError('model_selection_invalid')
    floating.api._connections.capture=capture
    result=receive(service,reference=target)
    assert result['pending_reason']=='model' and not floating.app.busy
    assert receive(service,reference=target)['duplicate'] and calls==[True]
    result=service.start_pending(service.document_id,1,target,1,result['session_id'])
    assert not result['started'] and result['pending_reason']=='model' and calls==[True,True]


def test_failed_persistence_does_not_acknowledge_or_advance_revision(draft,monkeypatch):
    service,floating,_=draft
    def fail(_):raise OSError('simulated disk failure')
    monkeypatch.setattr(floating,'persist',fail)
    with pytest.raises(OSError):receive(service)
    assert floating.runtime.messages==[] and floating.runtime.revision==0
    assert 'one' not in floating.runtime.ids


def test_old_document_round_source_and_external_change_cannot_receive(draft):
    service,floating,source=draft
    result=receive(service)
    with pytest.raises(GuardError):service.submit('two','新要求',None,'old-doc',service.source_version,result['session_id'],service.inbox_id)
    with pytest.raises(GuardError):service.submit('two','新要求',None,service.document_id,'old-source',result['session_id'],service.inbox_id)
    old_inbox=service.inbox_id
    service.start_again(True)
    with pytest.raises(GuardError):receive(service,'late',session=result['session_id'])
    with pytest.raises(GuardError):service.submit('late-first','旧首条要求',None,service.document_id,service.source_version,'',old_inbox)
    assert floating.runtime is None
    source.write_bytes(b'externally changed')
    with pytest.raises(ValueError):receive(service,'three')
    assert floating.runtime is None


def test_receive_invalidates_unclaimed_plan_and_pending_proposal(draft):
    service,floating,_=draft
    sync(service);receive(service)
    service.proposals['old']={'id':'old','status':'pending','draft_revision':1,'instruction_revision':1}
    plan=service.plan(service.document_id,1)['plan']
    receive(service,'two','保留温度。')
    assert service.proposals['old']['status']=='stale'
    with pytest.raises(GuardError):service.publish(plan['id'],service.document_id,1,True)
    with pytest.raises(GuardError):service.prepare_apply('old',service.document_id,1)


def test_claimed_apply_finishes_original_snapshot_then_new_requirement_waits(draft):
    service,floating,_=draft
    sync(service);receive(service)
    service.proposals['old']={'id':'old','status':'pending','draft_revision':1,'instruction_revision':1}
    grant=service.prepare_apply('old',service.document_id,1)
    result=receive(service,'two','继续精简。')
    assert result['pending_reason']=='applying'
    service.changed(service.document_id,2,grant['token'])
    result=service.applied(grant['token'],service.document_id,1,2)
    assert service.proposals['old']['status']=='adopted'
    assert result['pending_reason']=='reference' and floating.runtime.applied==0
    with pytest.raises(GuardError):service.applied(grant['token'],service.document_id,1,2)


def test_claimed_publish_is_not_replaced_and_uncertainty_blocks_model(draft):
    service,floating,_=draft
    sync(service);receive(service)
    plan=service.plan(service.document_id,1)['plan']
    frozen=service.publisher.claim(plan['id'],document_id=service.document_id,draft_revision=1,confirmed=True)
    service.publishing=floating.app.busy=True
    result=receive(service,'two','增加限定条件。')
    assert result['pending_reason']=='publishing' and frozen['instruction_revision']==1
    assert frozen['draft_revision']==1 and floating.runtime.revision==2
    service.publishing=floating.app.busy=False;service.uncertain=True
    result=receive(service,'three','暂时保留要求。')
    assert result['pending_reason']=='verify_result' and not service.start()


def test_stop_duplicate_ack_does_not_replay_and_stale_session_cannot_start(draft):
    service,floating,_=draft
    result=receive(service)
    floating.stop()
    assert receive(service,session='')['duplicate']
    assert floating.runtime.closed and not floating.app.busy
    sync(service);target=reference(service)
    with pytest.raises(GuardError):service.start_pending(service.document_id,1,target,1,result['session_id'])
    receive(service,'two')
    with pytest.raises(GuardError):service.start_pending(service.document_id,1,target,1,result['session_id'])
