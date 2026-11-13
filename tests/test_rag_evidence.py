"""Evidence publication tests: temp synthetic hits, no provider/Office/secrets."""
import hashlib
import json
import os
from pathlib import Path
import pytest

from office_agent.rag_evidence import RagEvidence, EvidenceError, MAX_PENDING, TTL
from office_agent.writes import Workspace

SESSION='a'*32
FOREIGN='b'*32
FP='c'*64


def hit(text='全文'*400, identifier=1, rank=1):
    return {'id':identifier,'doc_id':'doc-fixture','source':'fixture.docx','location':'paragraph:1',
            'text':text,'excerpt':text[:200],'model':'synthetic-vectors','source_fingerprint':FP,
            'distance':0.2,'rank':rank}


def make(tmp_path, version=None, clock=None):
    root=tmp_path/'workspace';root.mkdir()
    return RagEvidence(Workspace(root),version or (lambda doc:FP),**({'clock':clock} if clock else {})),root


def save(service, hits=None):
    plan=service.plan('格式约定',hits if hits is not None else [hit()],SESSION)
    return service.confirm(plan['plan_id'],SESSION,True)


def test_full_evidence_precedes_only_bounded_model_excerpts_and_owned_reread(tmp_path):
    service,root=make(tmp_path)
    rows=[hit(),hit('第二个完整正文'*300,2,2)]
    plan=service.plan('格式约定',rows,SESSION)
    assert plan['wrote_nothing'] and plan['confirmation_required']
    assert not list(root.iterdir()) and 'hits' not in plan and 'query' not in plan
    # Defend against callers retaining and mutating their original rows/list.
    rows[0]['text']='attacker changed original input'
    rows.clear()
    result=service.confirm(plan['plan_id'],SESSION,True)
    assert result['status']=='saved' and result['verified'] and result['applied'] is True
    path=root/result['evidence'];data=path.read_bytes();payload=json.loads(data)
    assert payload['hits'][0]['text']=='全文'*400
    assert hashlib.sha256(data).hexdigest()==result['payload_sha256']==plan['payload_sha256']
    assert len(result['hits'])==2 and all('text' not in item and len(item['excerpt'])<=200 for item in result['hits'])
    assert service.read(result['save_id'],SESSION,limit=150)['has_more'] is True
    assert len(service.read(result['save_id'],SESSION,limit=150)['text'])==150
    with pytest.raises(EvidenceError,match='session_invalid'):service.read(result['save_id'],FOREIGN)
    assert 'workspace' not in repr(service) and '全文' not in repr(service)


@pytest.mark.parametrize('approval',[False,1,'true',None,{},[]])
def test_no_io_on_nonliteral_approval_and_consumed_owner_plan(tmp_path,approval):
    service,root=make(tmp_path)
    plan=service.plan('格式约定',[hit()],SESSION)
    denied=service.confirm(plan['plan_id'],SESSION,approval)
    assert denied['status']=='rejected' and denied['applied'] is False and denied['wrote_nothing']
    assert not list(root.iterdir())
    assert service.confirm(plan['plan_id'],SESSION,True)['code']=='evidence_plan_invalid'


def test_cross_session_and_wrong_api_do_not_consume_owner_plan(tmp_path):
    service,root=make(tmp_path)
    plan=service.plan('格式约定',[hit()],SESSION)
    assert service.confirm(plan['plan_id'],FOREIGN,True)['code']=='evidence_session_invalid'
    assert service.confirm_rollback(plan['plan_id'],SESSION,True)['code']=='evidence_plan_invalid'
    result=service.confirm(plan['plan_id'],SESSION,True)
    rollback=service.plan_rollback(result['save_id'],SESSION)
    assert service.confirm(rollback['rollback_id'],SESSION,True)['code']=='evidence_plan_invalid'
    assert service.confirm_rollback(rollback['rollback_id'],SESSION,True)['restored'] is True


def test_separate_rollback_denial_replan_removes_only_owned_evidence_retains_audit(tmp_path):
    service,root=make(tmp_path)
    result=save(service);path=root/result['evidence']
    proposal=service.plan_rollback(result['save_id'],SESSION)
    assert path.exists() and proposal['wrote_nothing']
    assert service.confirm_rollback(proposal['rollback_id'],SESSION,False)['wrote_nothing']
    proposal=service.plan_rollback(result['save_id'],SESSION)
    assert service.confirm_rollback(proposal['rollback_id'],SESSION,True)['backups_audits_retained']
    assert not path.exists()
    assert (path.parent/'backup').read_bytes()==b''
    assert (path.parent/'audit-removed.json').exists()
    assert not service.confirm_rollback(proposal['rollback_id'],SESSION,True)['restored']


@pytest.mark.parametrize('field,value',[
    ('id',True),('rank',True),('rank',2),('distance',float('nan')),('distance',True),
    ('distance',3),('source_fingerprint','not-a-fingerprint'),('doc_id',[]),
    ('source','.env.production'),('text','api_key=sk-test-exposure'),('text','x'*8001),
    ('location','x'*513),('excerpt','fake excerpt'),('model','password=private-value'),
    ('source','fixture\x00.docx'),('text','\ud800'),
])
def test_strict_hit_validation_before_callbacks_or_io(tmp_path,field,value):
    calls=[]
    service,root=make(tmp_path,version=lambda doc:calls.append(doc) or FP)
    row=hit();row[field]=value
    with pytest.raises(EvidenceError):service.plan('格式约定',[row],SESSION)
    assert calls==[] and list(root.iterdir())==[]


@pytest.mark.parametrize('query',['api_key=forbiddenvalue','q'*201,'','\ud800'])
def test_invalid_query_before_callbacks(tmp_path,query):
    calls=[];service,root=make(tmp_path,version=lambda doc:calls.append(doc) or FP)
    with pytest.raises(EvidenceError):service.plan(query,[hit()],SESSION)
    assert calls==[] and list(root.iterdir())==[]


def test_duplicate_extra_field_or_too_many_hits_refused(tmp_path):
    service,root=make(tmp_path)
    for rows in ([hit(),hit(rank=2)],[{**hit(),'backend_path':'foreign'}],[hit(identifier=i,rank=i) for i in range(1,12)]):
        with pytest.raises(EvidenceError):service.plan('query',rows,SESSION)
    assert not list(root.iterdir())


def test_stale_before_approval_no_write_and_after_publication_no_excerpts(tmp_path):
    current=[FP];service,root=make(tmp_path,version=lambda doc:current[0])
    plan=service.plan('query',[hit()],SESSION)
    current[0]='d'*64
    result=service.confirm(plan['plan_id'],SESSION,True)
    assert result['code']=='evidence_source_changed' and result['wrote_nothing']
    assert not list(root.iterdir())
    current[0]=FP
    calls=[0]
    def changes_after_publication(doc):
        calls[0]+=1
        return FP if calls[0]<=3 else 'd'*64
    service._version=changes_after_publication
    result=save(service)
    assert result['status']=='uncertain' and result['applied'] is None and 'hits' not in result
    assert len(list(root.glob('*/evidence.json')))==1
    service.invalidate()
    with pytest.raises(EvidenceError,match='evidence_uncertain'):service.plan('query',[hit()],SESSION)


def test_ambiguous_replace_fuses_without_returning_excerpts(tmp_path,monkeypatch):
    service,root=make(tmp_path)
    real=os.replace
    def ambiguous(source,target):
        real(source,target)
        raise OSError('secret unsafe path from low-level exception')
    monkeypatch.setattr(os,'replace',ambiguous)
    result=save(service)
    assert result['status']=='uncertain' and result['applied'] is None
    assert 'secret unsafe' not in str(result) and 'hits' not in result
    assert len(list(root.glob('*/evidence.json')))==1
    assert service._fused


@pytest.mark.parametrize('artifact',['evidence.json','backup','audit-intent.json','audit-verified.json'])
def test_stale_artifacts_block_read_or_rollback_without_fusing(tmp_path,artifact):
    service,root=make(tmp_path)
    result=save(service);directory=(root/result['evidence']).parent
    (directory/artifact).write_bytes(b'tampered')
    with pytest.raises(EvidenceError):service.read(result['save_id'],SESSION)
    with pytest.raises(EvidenceError):service.plan_rollback(result['save_id'],SESSION)
    assert not service._fused
    assert (directory/'evidence.json').exists()


def test_postplan_artifact_tamper_rejected_before_rollback_writes(tmp_path):
    service,root=make(tmp_path)
    result=save(service);directory=(root/result['evidence']).parent
    proposal=service.plan_rollback(result['save_id'],SESSION)
    (directory/'backup').write_bytes(b'tamper')
    denied=service.confirm_rollback(proposal['rollback_id'],SESSION,True)
    assert denied['restored'] is False and denied['wrote_nothing']
    assert not (directory/'audit-remove-intent.json').exists()


def test_lifetime_expiration_pending_and_saved_capacity_and_invalidation(tmp_path):
    clock=[0.0];service,root=make(tmp_path,clock=lambda:clock[0])
    proposal=service.plan('query',[hit()],SESSION);clock[0]=TTL
    assert service.confirm(proposal['plan_id'],SESSION,True)['code']=='evidence_expired'
    for _ in range(MAX_PENDING):service.plan('query',[hit()],SESSION)
    with pytest.raises(EvidenceError,match='capacity'):service.plan('query',[hit()],SESSION)
    service.invalidate()
    results=[save(service) for _ in range(MAX_PENDING)]
    with pytest.raises(EvidenceError,match='capacity'):service.plan('query',[hit()],SESSION)
    rollback=service.plan_rollback(results[-1]['save_id'],SESSION)
    assert service.confirm_rollback(rollback['rollback_id'],SESSION,True)['restored']
    proposal=service.plan('query',[hit()],SESSION)
    service.invalidate()
    assert service.confirm(proposal['plan_id'],SESSION,True)['wrote_nothing']
    with pytest.raises(EvidenceError):service.plan_rollback(results[0]['save_id'],SESSION)
    other=RagEvidence(Workspace(root),lambda doc:FP)
    with pytest.raises(EvidenceError):other.read(results[0]['save_id'],SESSION)


def test_root_replacement_and_hardlinked_artifact_guard(tmp_path):
    service,root=make(tmp_path)
    plan=service.plan('query',[hit()],SESSION)
    root.rename(tmp_path/'retired');root.mkdir()
    assert service.confirm(plan['plan_id'],SESSION,True)['wrote_nothing']
    service=RagEvidence(Workspace(root),lambda doc:FP)
    saved=save(service);path=root/saved['evidence']
    os.link(path,root/'linked-copy')
    with pytest.raises(EvidenceError):service.read(saved['save_id'],SESSION)
    with pytest.raises(EvidenceError):service.plan_rollback(saved['save_id'],SESSION)


def test_unknown_callback_errors_are_redacted(tmp_path):
    def broken(doc):raise OSError('api_key=private-value C:/unsafe/location')
    service,root=make(tmp_path,version=broken)
    with pytest.raises(EvidenceError) as exc:service.plan('query',[hit()],SESSION)
    assert str(exc.value)=='evidence_io_failed'
    assert not list(root.iterdir())


def test_empty_hit_evidence_is_allowed_and_offset_and_session_validated(tmp_path):
    service,root=make(tmp_path)
    result=save(service,[])
    assert result['hits']==[] and json.loads((root/result['evidence']).read_bytes())['hits']==[]
    for params in ({'offset':True},{'limit':True},{'limit':2001},{'offset':-1}):
        with pytest.raises(EvidenceError):service.read(result['save_id'],SESSION,**params)
    with pytest.raises(EvidenceError):service.plan('query',[],'forged-session')
