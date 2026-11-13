from pathlib import Path

from docx import Document
import pytest

from office_agent.rag import RagError
from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import Workspace
from test_rag_workspace import SyntheticEmbedder, SESSION


def document(path, text):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(path)


def setup(tmp_path):
    document(tmp_path / 'a.docx', 'EXPT-01 成本 30 ms；常温条件。')
    document(tmp_path / 'b.docx', 'EXPT-02 测量 20 ms；高温条件不适用。')
    embedder = SyntheticEmbedder()
    return WorkspaceRag(Workspace(tmp_path), embedder), embedder


def sync(service):
    plan = service.plan_corpus('a.docx', SESSION)
    result = service.confirm_corpus(plan['plan_id'], SESSION, True)
    assert result['status'] == 'saved', result
    return result


def test_incremental_restart_delete_and_atomic_rollback(tmp_path):
    service, embedder = setup(tmp_path)
    originals = {path.name: path.read_bytes() for path in tmp_path.glob('*.docx')}
    status = service.corpus_status('a.docx', SESSION)
    assert status['changed'] == 2 and not embedder.calls
    plan = service.plan_corpus('a.docx', SESSION)
    assert service.confirm_corpus(plan['plan_id'], SESSION, False)['wrote_nothing']
    assert not list(tmp_path.glob('.office-*'))
    sync(service)
    assert len(embedder.calls) == 1 and len(embedder.calls[0]) == 2
    with pytest.raises(RagError, match='rag_index_current'):
        service.plan_corpus('a.docx', SESSION)
    restarted = WorkspaceRag(Workspace(tmp_path), embedder)
    assert restarted.corpus_status('a.docx', SESSION)['status'] == 'current'
    query = restarted.plan_corpus_query('a.docx', 'EXPT-02', SESSION, top_k=2, rerank=True)
    assert query['retrieval']['rerank'] == 'not_configured'
    result = restarted.confirm_query(query['plan_id'], SESSION, True)
    assert result['status'] == 'saved' and {hit['source'] for hit in result['hits']} == {'a.docx', 'b.docx'}
    assert len(embedder.calls) == 2 and len(embedder.calls[-1]) == 1
    for hit in result['hits']:
        assert restarted.read_research_context(result['save_id'], hit['id'], SESSION)['source'] == hit['source']
    (tmp_path / 'b.docx').unlink()
    document(tmp_path / 'a.docx', 'EXPT-01 新版本 成本 31 ms。')
    plan = restarted.plan_corpus('a.docx', SESSION)
    assert plan['changed'] == ['a.docx'] and plan['removed'] == ['b.docx'] and plan['chunk_count'] == 1
    saved = restarted.confirm_corpus(plan['plan_id'], SESSION, True)
    assert saved['status'] == 'saved' and len(embedder.calls[-1]) == 1
    query = restarted.plan_corpus_query('a.docx', 'EXPT-02', SESSION)
    result = restarted.confirm_query(query['plan_id'], SESSION, True)
    assert all(hit['source'] != 'b.docx' for hit in result['hits'])
    rollback = restarted.plan_index_rollback(saved['save_id'], SESSION)
    assert restarted.confirm_index_rollback(rollback['rollback_id'], SESSION, True)['chunk_count'] == 2
    with pytest.raises(RagError, match='rag_corpus_stale'):
        restarted.plan_corpus_query('a.docx', 'EXPT-02', SESSION)
    for name, data in originals.items():
        (tmp_path / name).write_bytes(data)
    assert restarted.corpus_status('a.docx', SESSION)['status'] == 'current'


def test_inventory_change_before_embedding_and_before_evidence_confirmation(tmp_path):
    service, embedder = setup(tmp_path)
    plan = service.plan_corpus('a.docx', SESSION)
    document(tmp_path / 'new.docx', 'A newly admitted experiment.')
    assert service.confirm_corpus(plan['plan_id'], SESSION, True)['wrote_nothing']
    assert embedder.calls == []
    sync(service)
    query = service.plan_corpus_query('a.docx', 'EXPT-01', SESSION)
    document(tmp_path / 'another.docx', 'Another new experiment.')
    assert service.confirm_query(query['plan_id'], SESSION, True)['wrote_nothing']
    assert not list(tmp_path.glob('.office-rag-evidence-*'))


def test_directory_limit_nonrecursive_and_selected_file_index_separation(tmp_path):
    service, embedder = setup(tmp_path)
    nested = tmp_path / 'nested'
    nested.mkdir()
    document(nested / 'hidden.docx', 'Never admitted.')
    for index in range(21):
        document(tmp_path / f'fixture-{index:02}.docx', f'Fixture {index}')
    status = service.corpus_status('a.docx', SESSION)
    assert status['documents'] == 20
    assert any(item.get('reason') == 'document_limit' for item in status['inventory'])
    assert all(item['source'] != 'hidden.docx' for item in status['inventory'])
    plan = service.plan_research_index('a.docx', SESSION)
    service.confirm_index(plan['plan_id'], SESSION, True)
    before = service.snapshot()['chunk_count']
    sync(service)
    assert service.snapshot()['chunk_count'] == before + 20
    assert service.research_status('a.docx', SESSION)['status'] == 'current'


def test_post_publication_failure_fuses_corpus_and_survives_restart(tmp_path, monkeypatch):
    service, embedder = setup(tmp_path)
    plan = service.plan_corpus('a.docx', SESSION)
    monkeypatch.setattr(service._store, '_finish_audit', lambda *args: (_ for _ in ()).throw(RuntimeError('private-data')))
    result = service.confirm_corpus(plan['plan_id'], SESSION, True)
    assert result['status'] == 'uncertain' and result['applied'] is None
    with pytest.raises(RagError, match='rag_uncertain'):
        service.corpus_status('a.docx', SESSION)
    reopened = WorkspaceRag(Workspace(tmp_path), embedder)
    assert reopened.snapshot()['uncertain']


def test_embedding_profile_switch_isolated_and_returning_reuses_published_corpus(tmp_path):
    service, original = setup(tmp_path)
    sync(service)
    first = service._store._db

    class OtherEmbedder:
        identity = {'channel': 'local', 'model': 'different-controlled-fixture', 'dimension': 3}

        def embed(self, texts):
            return [[1, 0, 0] for _ in texts]

    other = WorkspaceRag(Workspace(tmp_path), OtherEmbedder())
    assert other.snapshot()['chunk_count'] == 0 and other._store._db != first
    sync(other)
    assert other.snapshot()['dimension'] == 3
    returned = WorkspaceRag(Workspace(tmp_path), original)
    assert returned._store._db == first
    assert returned.corpus_status('a.docx', SESSION)['status'] == 'current'
    assert len(original.calls) == 1


@pytest.mark.parametrize('damage', ['uncertain', 'duplicate_metadata', 'missing_vector'])
def test_model_switch_cannot_bypass_damaged_or_uncertain_original_database(tmp_path, damage):
    import sqlite3
    from office_agent.vectorstore import VectorStoreError

    service, original = setup(tmp_path)
    sync(service)
    with sqlite3.connect(service._store._db) as connection:
        if damage == 'uncertain':
            connection.execute('UPDATE metadata SET uncertain=1')
        elif damage == 'duplicate_metadata':
            connection.execute('INSERT INTO metadata SELECT * FROM metadata')
        else:
            connection.execute('DELETE FROM chunks WHERE id=(SELECT min(id) FROM chunks)')

    class OtherEmbedder:
        identity = {'channel': 'local', 'model': 'another-fixture', 'dimension': 3}

        def embed(self, texts):
            pytest.fail('Rejected profile must not embed')

    expected = {'uncertain': 'uncertain_state', 'duplicate_metadata': 'invalid_database', 'missing_vector': 'verification_failed'}
    with pytest.raises(VectorStoreError, match=expected[damage]):
        WorkspaceRag(Workspace(tmp_path), OtherEmbedder())
    assert list(tmp_path.glob('.office-knowledge-*.sqlite3')) == []
