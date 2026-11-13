"""Real sqlite-vec/source/evidence composition, deterministic synthetic vectors only.

Synthetic topic mapping below does NOT establish real paraphrase/embedding quality.
"""
import json
import pytest

from office_agent.rag import RagError
from office_agent.rag_evidence import EvidenceError, MAX_PENDING
from office_agent.rag_sources import SourceExtractor
from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import Workspace

SESSION='a'*32
FOREIGN='b'*32


class SyntheticEmbedder:
    identity={'channel':'local','model':'test-workspace-synthetic','dimension':2}
    def __init__(self):self.calls=[]
    def embed(self,texts):
        self.calls.append(list(texts))
        return [[1.,0.] if '成本' in text or '预算' in text else [0.,1.] for text in texts]


def setup(tmp_path):
    source=tmp_path/'report.txt'
    source.write_text(('成本安排与开支说明。'*180)+'\n'+('会议时间与参加名单。'*180),encoding='utf-8')
    embedder=SyntheticEmbedder()
    service=WorkspaceRag(Workspace(tmp_path),embedder)
    return service,source,embedder


def index(service):
    plan=service.plan_index('report.txt',SESSION)
    result=service.confirm_index(plan['plan_id'],SESSION,True)
    assert result['status']=='saved' and result['applied'] is True
    return result


def test_actual_index_query_then_approved_full_evidence_excerpts_and_separate_rollbacks(tmp_path):
    service,source,embedder=setup(tmp_path)
    initial=set(tmp_path.iterdir())
    plan=service.plan_index(source.name,SESSION)
    assert plan['chunk_count']==2 and embedder.calls==[] and set(tmp_path.iterdir())==initial
    assert service.confirm_index(plan['plan_id'],FOREIGN,True)['wrote_nothing']
    result=service.confirm_index(plan['plan_id'],SESSION,True)
    assert result['applied'] is True and service.snapshot()['chunk_count']==2
    assert len(embedder.calls)==1
    # Topic vectors are fabricated, not claims about any real semantic model.
    question='预算支出'
    assert question not in source.read_text(encoding='utf-8')
    query_plan=service.plan_query(question,SESSION,top_k=1)
    assert query_plan['hit_count']==1 and 'hits' not in query_plan
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    assert service.confirm_query(query_plan['plan_id'],SESSION,False)['wrote_nothing']
    query_plan=service.plan_query(question,SESSION,top_k=1)
    query=service.confirm_query(query_plan['plan_id'],SESSION,True)
    assert query['applied'] is True and query['hits'][0]['excerpt'].startswith('成本')
    assert len(query['hits'][0]['excerpt'])==200 and 'text' not in query['hits'][0]
    evidence=tmp_path/query['evidence']
    assert len(json.loads(evidence.read_bytes())['hits'][0]['text'])>200
    assert len(service.read_evidence(query['save_id'],SESSION,limit=200)['text'])==200
    evidence_rollback=service.plan_evidence_rollback(query['save_id'],SESSION)
    assert service.confirm_evidence_rollback(evidence_rollback['rollback_id'],SESSION,True)['restored']
    assert not evidence.exists() and service.snapshot()['chunk_count']==2
    database_rollback=service.plan_index_rollback(result['save_id'],SESSION)
    restored=service.confirm_index_rollback(database_rollback['rollback_id'],SESSION,True)
    assert restored['chunk_count']==0 and restored['first_absent_restored_empty'] is True
    assert source.exists() and '成本' in source.read_text(encoding='utf-8')


@pytest.mark.parametrize('approval',[False,1,'true',None])
def test_index_denial_no_embeddings_no_writes(tmp_path,approval):
    service,source,embedder=setup(tmp_path)
    before=set(tmp_path.iterdir());plan=service.plan_index(source.name,SESSION)
    assert service.confirm_index(plan['plan_id'],SESSION,approval)['wrote_nothing']
    assert embedder.calls==[] and set(tmp_path.iterdir())==before


def test_source_changes_after_parse_before_core_plan_never_bind_old_blocks_to_new_hash(tmp_path,monkeypatch):
    service,source,embedder=setup(tmp_path)
    real=SourceExtractor.extract
    def extract_then_change(self,name):
        result=real(self,name)
        source.write_text('新版本不能用于旧正文的指纹。',encoding='utf-8')
        return result
    monkeypatch.setattr(SourceExtractor,'extract',extract_then_change)
    with pytest.raises(RagError,match='rag_source_changed'):service.plan_index(source.name,SESSION)
    assert service._rag._plans=={} and embedder.calls==[]
    assert set(p.name for p in tmp_path.iterdir())=={source.name}


def test_change_after_index_plan_rejects_before_provider_and_query_stale_refused(tmp_path):
    service,source,embedder=setup(tmp_path)
    plan=service.plan_index(source.name,SESSION)
    source.write_text('changed',encoding='utf-8')
    assert service.confirm_index(plan['plan_id'],SESSION,True)['wrote_nothing']
    assert embedder.calls==[]
    index(service)
    source.write_text('changed again',encoding='utf-8')
    with pytest.raises(RagError,match='rag_query_failed'):service.plan_query('changed',SESSION)
    assert not list(tmp_path.glob('.office-rag-evidence-*'))


def test_source_changes_before_evidence_approval_no_artifacts_or_excerpts(tmp_path):
    service,source,embedder=setup(tmp_path);index(service)
    query=service.plan_query('预算',SESSION)
    source.write_text('changed',encoding='utf-8')
    result=service.confirm_query(query['plan_id'],SESSION,True)
    assert result['applied'] is False and result['wrote_nothing'] and 'hits' not in result
    assert not list(tmp_path.glob('.office-rag-evidence-*'))


def test_restart_has_no_grants_and_requires_explicit_readonly_source_binding(tmp_path):
    service,source,embedder=setup(tmp_path);index(service)
    query=service.plan_query('预算',SESSION)
    saved=service.confirm_query(query['plan_id'],SESSION,True)
    other=WorkspaceRag(Workspace(tmp_path),embedder)
    with pytest.raises(EvidenceError):other.read_evidence(saved['save_id'],SESSION)
    with pytest.raises(RagError,match='rag_query_failed'):other.plan_query('预算',SESSION)
    before=set(tmp_path.iterdir());calls=len(embedder.calls)
    bound=other.bind_source(source.name,SESSION)
    assert bound['wrote_nothing'] and len(embedder.calls)==calls and set(tmp_path.iterdir())==before
    assert 'text' not in bound and 'blocks' not in bound
    plan=other.plan_query('预算',SESSION)
    assert other.confirm_query(plan['plan_id'],SESSION,True)['applied'] is True
    assert 'report.txt' not in repr(other)


def test_invalid_session_topk_or_sensitive_query_zero_provider_calls(tmp_path):
    service,source,embedder=setup(tmp_path)
    with pytest.raises(RagError):service.plan_index(source.name,'invalid')
    assert service._source._bindings=={}
    for question,session,top_k in [('预算','invalid',5),('api_key=blocked-value',SESSION,5),('预算',SESSION,True)]:
        with pytest.raises(RagError):service.plan_query(question,session,top_k=top_k)
    assert embedder.calls==[]


def test_evidence_fuse_or_capacity_rejected_before_new_provider_reads(tmp_path):
    service,source,embedder=setup(tmp_path);index(service)
    for _ in range(MAX_PENDING):service.plan_query('预算',SESSION)
    count=len(embedder.calls)
    with pytest.raises(EvidenceError,match='capacity'):service.plan_query('预算',SESSION)
    assert len(embedder.calls)==count
    service.invalidate()
    service._evidence._fused=True
    with pytest.raises(EvidenceError,match='uncertain'):service.plan_query('预算',SESSION)
    assert len(embedder.calls)==count


def test_invalidation_retires_index_query_and_owned_rollback_authorities(tmp_path):
    service,source,embedder=setup(tmp_path)
    proposed=service.plan_index(source.name,SESSION);service.invalidate()
    assert service.confirm_index(proposed['plan_id'],SESSION,True)['wrote_nothing']
    indexed=index(service)
    query=service.plan_query('预算',SESSION)
    rollback=service.plan_index_rollback(indexed['save_id'],SESSION)
    service.invalidate()
    assert service.confirm_query(query['plan_id'],SESSION,True)['wrote_nothing']
    from office_agent.vectorstore import VectorStoreError
    with pytest.raises(VectorStoreError):service.confirm_index_rollback(rollback['rollback_id'],SESSION,True)
    assert service.snapshot()['chunk_count']==2
