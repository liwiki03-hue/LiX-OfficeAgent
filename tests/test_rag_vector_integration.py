"""Actual extension integration with synthetic vectors; no provider/Office calls."""
from contextlib import closing
import sqlite3

import pytest

from office_agent.rag import RagService, RagError
from office_agent.vectorstore import VectorStore, VectorStoreError


class Embedder:
    identity = {'channel':'local','model':'fixture-integration','dimension':2}
    def __init__(self): self.calls = []
    def embed(self, texts):
        self.calls.append(texts)
        return [[1,0] for _ in texts]


def setup(tmp_path):
    embedder = Embedder()
    store = VectorStore(tmp_path,model=embedder.identity['model'],dimension=2)
    rag = RagService(store,embedder,lambda doc_id:'a'*64,store_error_type=VectorStoreError)
    return rag,store,embedder


def plan(rag):
    return rag.plan_index('document',[{'source':'fixture.docx','location':'paragraph:1','type':'paragraph',
                                       'group':'document','text':'approved text'}],'host-session')


def test_actual_store_approved_index_receipt_query_and_owned_rollback(tmp_path):
    rag,store,embedder = setup(tmp_path)
    proposal = plan(rag)
    assert embedder.calls == [] and list(tmp_path.iterdir()) == []
    result = rag.confirm_index(proposal['plan_id'],'host-session',True)
    assert result['status'] == 'saved' and result['applied'] is True
    assert rag._disabled is False and store.snapshot()['chunk_count'] == 1
    hits = rag.query('approved text')
    assert hits[0]['text'] == 'approved text' and len(hits[0]['excerpt']) <= 200
    rollback = store.plan_rollback(result['save_id'],'host-session')
    restored = store.confirm_rollback(rollback['rollback_id'],'host-session',True)
    assert restored['chunk_count'] == 0 and restored['first_absent_restored_empty'] is True


@pytest.mark.parametrize('sql',[
    "CREATE TRIGGER injected AFTER INSERT ON chunks BEGIN UPDATE chunks SET text='injected sentinel' WHERE id=NEW.id; END",
    'CREATE VIEW injected AS SELECT text FROM chunks',
    'CREATE INDEX injected ON chunks(text)',
    'CREATE TABLE injected(secret TEXT)',
    'ALTER TABLE chunks ADD COLUMN unapproved TEXT',
],ids=['trigger','view','index','table','column'])
def test_unexpected_schema_rejected_on_reopen_before_approved_payload_can_change(tmp_path,sql):
    rag,store,embedder = setup(tmp_path)
    proposal = plan(rag)
    assert rag.confirm_index(proposal['plan_id'],'host-session',True)['applied'] is True
    with closing(sqlite3.connect(store._db)) as connection:
        connection.execute(sql)
        connection.commit()
    with pytest.raises(VectorStoreError,match='invalid_database'):
        VectorStore(tmp_path,model=store.model,dimension=2)


def test_model_mismatch_rejects_before_any_embedding_or_database_creation(tmp_path):
    embedder = Embedder()
    store = VectorStore(tmp_path,model='wrong-fixture-model',dimension=2)
    rag = RagService(store,embedder,lambda doc_id:'a'*64)
    with pytest.raises(RagError,match='rag_model_mismatch'): plan(rag)
    with pytest.raises(RagError,match='rag_model_mismatch'): rag.query('query')
    assert embedder.calls == [] and list(tmp_path.iterdir()) == []


def test_actual_ambiguous_publication_preserves_safe_receipt_for_independent_rollback(tmp_path,monkeypatch):
    import office_agent.vectorstore as module
    rag,store,embedder = setup(tmp_path)
    proposal = plan(rag)
    actual = module.os.replace
    def replace_then_raise(source,target):
        actual(source,target)
        raise OSError('fixture-private-publication-error')
    with monkeypatch.context() as patch:
        patch.setattr(module.os,'replace',replace_then_raise)
        result = rag.confirm_index(proposal['plan_id'],'host-session',True)
    assert result['status'] == 'uncertain' and result['applied'] is None
    assert 'fixture-private-publication-error' not in repr(result)
    assert len(result['save_id']) == 32 and rag._disabled is True
    rollback = store.plan_rollback(result['save_id'],'host-session')
    restored = store.confirm_rollback(rollback['rollback_id'],'host-session',True)
    assert restored['chunk_count'] == 0
    with pytest.raises(RagError,match='rag_uncertain'): rag.query('query')
