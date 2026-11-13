import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from office_agent.diagram_process import OwnedRun
from office_agent.rag import RagError
from office_agent.research_reranker import OfflineReranker
from office_agent.retrieval_pipeline import subqueries
from office_agent.retrieval_pipeline import retrieve
from test_research_corpus import setup, sync
from test_rag_workspace import SESSION


def test_literal_query_decomposition_is_bounded_and_keeps_numeric_constraints():
    question = '比较 thermal conductivity 20 ms 以及 battery cycle 30 ms'
    values = subqueries(question, True)
    assert values[0] == question and 1 < len(values) <= 3
    assert all('20' in value and '30' in value for value in values)
    assert subqueries('比较热导率以及电池循环，排除30 ms结果', True) == ['比较热导率以及电池循环，排除30 ms结果']


@pytest.mark.parametrize('outcome', ['completed', 'timeout', 'invalid_response', 'cleanup_failed'])
def test_offline_reranker_bound_process_result_and_temp_cleanup(tmp_path, monkeypatch, outcome):
    monkeypatch.setattr('office_agent.research_reranker.Embedder._manifest', lambda *args: 'fixture-manifest')
    observed = []

    def runner(command, **kwargs):
        request, target = Path(command[-2]), Path(command[-1])
        observed.append(request.parent)
        assert json.loads(request.read_bytes())['texts'] == ['first', 'second']
        assert kwargs['timeout_ms'] == 1000
        if outcome == 'timeout':
            return OwnedRun('render_timeout', None, 0, 1, True)
        if outcome == 'cleanup_failed':
            return OwnedRun('process_completed', 0, 0, 0, False)
        target.write_text(json.dumps({'order': [1, 0] if outcome == 'completed' else [0, 0]}), encoding='utf-8')
        return OwnedRun('process_completed', 0, 0, 0, True)

    ranker = OfflineReranker(str(tmp_path), 1, runner=runner)
    hits = [{'id': 1, 'text': 'first'}, {'id': 2, 'text': 'second'}]
    ranked, status = ranker.rank('question', hits)
    assert status == outcome
    assert [hit['id'] for hit in ranked] == ([2, 1] if outcome == 'completed' else [1, 2])
    assert observed and all(not path.exists() for path in observed)


def test_corpus_rerank_reorders_existing_records_and_reports_degraded_state(tmp_path):
    service, embedder = setup(tmp_path)
    sync(service)
    seen = []

    def rank(question, hits):
        seen.append(hits)
        return [dict(hit, text='This fabricated text must never become evidence') for hit in reversed(hits)], 'completed'

    service._reranker = SimpleNamespace(rank=rank)
    query = service.plan_corpus_query('a.docx', '比较 EXPT-01 以及 EXPT-02', SESSION, rerank=True, expand=True)
    assert 1 < query['retrieval']['queries'] <= 3
    assert query['retrieval']['rerank'] == 'completed' and len(seen[0]) <= 20
    result = service.confirm_query(query['plan_id'], SESSION, True)
    assert result['status'] == 'saved'
    payload = json.loads((tmp_path / result['evidence']).read_bytes())
    assert payload['retrieval']['rerank'] == 'completed'
    assert all('fabricated' not in hit['text'] for hit in payload['hits'])
    service._reranker = SimpleNamespace(rank=lambda query, hits: (hits, 'timeout'))
    query = service.plan_corpus_query('a.docx', 'EXPT-01', SESSION, rerank=True)
    assert query['retrieval']['rerank'] == 'timeout'
    service._reranker = SimpleNamespace(rank=lambda query, hits: (hits, 'cleanup_failed'))
    with pytest.raises(RagError, match='rag_uncertain'):
        service.plan_corpus_query('a.docx', 'EXPT-01', SESSION, rerank=True)
    assert service._rag._disabled


def test_search_normalization_preserves_original_evidence():
    from office_agent.hybrid_ranking import terms, rankings

    original = 'Thermal con-\nductivity 热导\n率 ＥＸＰＴ－０２'
    assert {'conductivity', '导率', 'expt-02'} <= set(terms(original))
    rows = [{'id': 1, 'text': original, 'distance': 1}, {'id': 2, 'text': 'unrelated', 'distance': 0}]
    assert rankings(rows, 'conductivity 热导率', 50)[1] == [1]
    assert rows[0]['text'] == original


def test_duplicate_removal_keeps_independent_sources_and_original_text():
    rows = [dict(id=index, doc_id=doc, source_fingerprint='version', text='exact original')
            for index, doc in [(1, 'a'), (2, 'a'), (3, 'b')]]
    store = SimpleNamespace(query_scoped=lambda *args, **kwargs: rows)
    rag = SimpleNamespace(embedder=SimpleNamespace(embed=lambda values: [[1, 0] for value in values]),
                          _vectors=lambda values, count: values)
    hits, report = retrieve(store, rag, 'original', {'a': 'version', 'b': 'version'})
    assert [hit['id'] for hit in hits] == [1, 3]
    assert report['duplicates_removed'] == 1 and report['candidates'] == 2
    assert all(hit['text'] == 'exact original' for hit in hits)


def test_lexical_cache_reuses_only_same_published_authorized_snapshot(tmp_path, monkeypatch):
    import office_agent.hybrid_ranking as module
    from office_agent.token_chunks import CHUNK_VERSION
    from office_agent.vectorstore import VectorStore
    from test_research_hybrid import publish

    store = VectorStore(tmp_path, model='fixture', dimension=2)
    publish(store, 'a', 'v1', 'thermal conductivity')
    publish(store, 'b', 'v1', 'private unrelated')
    calls, original = [], module.terms
    monkeypatch.setattr(module, 'terms', lambda text: (calls.append(text) if text != 'thermal' else None, original(text))[1])

    def query(sources):
        return store.query_scoped([1, 0], keyword='thermal', sources=sources, chunk_version=CHUNK_VERSION)

    assert query({'a': 'v1'})[0]['doc_id'] == 'a'
    assert calls == ['thermal conductivity']
    query({'a': 'v1'})
    assert len(calls) == 1
    assert query({'b': 'v1'})[0]['doc_id'] == 'b'
    assert calls[-1] == 'private unrelated'
    publish(store, 'a', 'v2', 'new conductivity')
    assert query({'a': 'v2'})[0]['text'] == 'new conductivity'
    assert calls[-1] == 'new conductivity'
    store.invalidate()
    assert store._lexical_cache is None
