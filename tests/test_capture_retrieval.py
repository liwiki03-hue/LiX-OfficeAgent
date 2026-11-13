from copy import deepcopy

from office_agent.research_capture_retrieval import material_pool, retrieve_captures


def evidence(count=80):
    return {f'ev:{index}': {'id': f'ev:{index}', 'original_text': f'Original paragraph {index}.', 'capture_id': 'a' if index < count - 1 else 'b',
                           'page': 1, 'location': f'pdf-page:1/char:{index * 80}', 'source_sha256': 'a' if index < count - 1 else 'b', 'title': 'Synthetic source'} for index in range(count)}


def test_balanced_materials_keep_independent_small_source():
    rows = evidence(400)
    pooled = material_pool(rows, ['跨语言问题'], 16)
    assert len(pooled) == 16 and any(row['capture_id'] == 'b' for row in pooled.values())


def test_dense_materials_include_beyond_lexical_candidate_fifty_and_real_neighbors():
    rows = evidence()
    rows['ev:70']['original_text'] = 'Thermal conductivity observation.'

    class Embedder:
        identity = {'channel': 'local', 'model': 'synthetic-unit-fixture', 'dimension': 2}

        def embed(self, texts):
            assert len(texts) == 80
            return [[1., 0.] if 'Thermal' in text else [0., 1.] for text in texts]

        def embed_queries(self, texts):
            return [[1., 0.] for text in texts]

    selected, report = retrieve_captures(rows, '热导率', ['热导率'], ['ev:70'], (Embedder(), None))
    assert len(report['vectorized_evidence_ids']) == 80 and not report['candidate_pool_truncated']
    assert 'ev:70' in selected and {'ev:69', 'ev:71'} <= set(selected)
    assert {value['evidence_id'] for value in selected['ev:70']['context']} == {'ev:69', 'ev:71'}
    assert rows['ev:70'].get('context') is None


def test_saved_retrieval_does_not_mix_neighboring_different_sources():
    rows = evidence(2)
    for index, row in enumerate(rows.values()):
        row.update(capture_id=None, page=None, context=[{'role': 'parent', 'text': f'Approved context {index}'}])
    original = deepcopy(rows)
    selected, report = retrieve_captures(rows, 'Original', ['Original'], list(rows))
    assert all(row['context'] == original[key]['context'] for key, row in selected.items())
    assert report['embedding_identity'] is None and not report['vectorized_evidence_ids']


def test_cancel_after_document_embedding_prevents_query_embedding():
    calls = []
    cancelled = [False]

    class Embedder:
        identity = {'channel': 'local', 'model': 'synthetic-unit-fixture', 'dimension': 2}

        def embed(self, texts):
            calls.append('documents')
            cancelled[0] = True
            return [[1., 0.] for text in texts]

        def embed_queries(self, texts):
            calls.append('query')
            raise AssertionError('late query must not execute')

    def check():
        if cancelled[0]:
            raise RuntimeError('cancelled')

    import pytest
    with pytest.raises(RuntimeError, match='cancelled'):
        retrieve_captures(evidence(2), 'fixture', ['fixture'], [], (Embedder(), None), check)
    assert calls == ['documents']
