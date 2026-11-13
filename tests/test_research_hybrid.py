import json

import pytest

from office_agent.hybrid_ranking import terms, rankings, fuse
from office_agent.rag_workspace import WorkspaceRag
from office_agent.token_chunks import CHUNK_VERSION, token_count
from office_agent.vectorstore import VectorStore, VectorStoreError
from office_agent.writes import Workspace
from test_rag_workspace import SyntheticEmbedder, SESSION


def test_mixed_terms_and_exact_identifiers_are_separate_rankings():
    assert {'热', '热导', '导率', 'expt-02', '12.5'} <= set(terms('热导率 EXPT-02 12.5'))
    records = [{'id': 1, 'text': 'EXPT-02 热导率 measured 12.5 ms', 'distance': .8},
               {'id': 2, 'text': 'Unrelated semantic fixture EXPT-020', 'distance': .1}]
    lists = rankings(records, 'EXPT-02 热导率', 50)
    assert lists[0] == [1] and lists[1][0] == 1 and lists[2] == [2, 1]
    assert fuse(lists, 1) == [1]


def publish(store, doc_id, version, text, *, vector=None, kind='document', profile=CHUNK_VERSION):
    plan = store.plan_replace(doc_id, version, [{'source': doc_id + '.txt', 'location': json.dumps({'chunker': profile}),
                                               'text': text, 'embedding': vector or [1, 0], 'kind': kind}], SESSION)
    return store.confirm_replace(plan['plan_id'], SESSION, True)


def test_candidate_scope_before_ranking_and_profile_version_refusal(tmp_path):
    store = VectorStore(tmp_path, model='fixture', dimension=2)
    publish(store, 'foreign', 'foreign-version', 'EXPT-02 热导率')
    publish(store, 'memory', 'private-version', 'EXPT-02 private', kind='memory')
    publish(store, 'allowed', 'current-version', 'EXPT-02 热导率', vector=[0, 1])
    hits = store.query_scoped([1, 0], keyword='EXPT-02', sources={'allowed': 'current-version'}, top_k=1, chunk_version=CHUNK_VERSION)
    assert len(hits) == 1 and hits[0]['doc_id'] == 'allowed' and hits[0]['distance'] == 1
    with pytest.raises(VectorStoreError, match='stale_source'):
        store.query_scoped([1, 0], keyword='EXPT-02', sources={'allowed': 'old-version'}, chunk_version=CHUNK_VERSION)
    publish(store, 'allowed', 'current-version', 'EXPT-02', profile='structure-v1')
    with pytest.raises(VectorStoreError, match='index_profile_mismatch'):
        store.query_scoped([1, 0], keyword='EXPT-02', sources={'allowed': 'current-version'}, chunk_version=CHUNK_VERSION)


def test_token_index_persistence_query_evidence_and_rollback(tmp_path):
    source = tmp_path / 'report.txt'
    source.write_text(('成本安排与开支说明 EXPT-02。' * 200), encoding='utf-8')
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    plan = service.plan_research_index(source.name, SESSION)
    assert plan['chunk_count'] > 1 and embedder.calls == []
    assert set(tmp_path.iterdir()) == {source}
    saved = service.confirm_index(plan['plan_id'], SESSION, True)
    assert saved['status'] == 'saved' and len(embedder.calls) == 1
    assert all(token_count(text) <= 384 for text in embedder.calls[0])
    reopened = WorkspaceRag(Workspace(tmp_path), embedder)
    query = reopened.plan_research_query(source.name, 'EXPT-02 成本', SESSION, top_k=2)
    assert len(embedder.calls) == 2 and len(embedder.calls[1]) == 1
    result = reopened.confirm_query(query['plan_id'], SESSION, True)
    assert result['status'] == 'saved' and len(result['hits']) == 2
    assert all('chunker' in hit['location'] for hit in result['hits'])
    evidence = json.loads((tmp_path / result['evidence']).read_bytes())
    assert all(token_count(hit['text']) <= 384 for hit in evidence['hits'])
    source.write_text('changed', encoding='utf-8')
    with pytest.raises(VectorStoreError, match='stale_source'):
        reopened.plan_research_query(source.name, 'EXPT-02', SESSION)
    rollback = service.plan_index_rollback(saved['save_id'], SESSION)
    assert service.confirm_index_rollback(rollback['rollback_id'], SESSION, True)['chunk_count'] == 0


def test_missing_index_and_uncertainty_refuse_before_embedding(tmp_path):
    from office_agent.rag import RagError

    source = tmp_path / 'report.txt'
    source.write_text('An authorized synthetic experiment.', encoding='utf-8')
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    with pytest.raises(VectorStoreError, match='document_missing'):
        service.plan_research_query(source.name, 'experiment', SESSION)
    assert embedder.calls == [] and set(tmp_path.iterdir()) == {source}
    service._rag._disabled = True
    with pytest.raises(RagError, match='rag_uncertain'):
        service.plan_research_query(source.name, 'experiment', SESSION)
    assert embedder.calls == []
