import json
import os

from docx import Document
import pytest

from office_agent.rag import RagError
from office_agent.rag_evidence import EvidenceError
from office_agent.rag_sources import SourceExtractor
from office_agent.rag_workspace import WorkspaceRag
from office_agent.token_chunks import token_count, token_structure_chunks
from office_agent.research_context import expand_contexts
from office_agent.vectorstore import VectorStoreError
from office_agent.writes import Workspace
from test_rag_workspace import SyntheticEmbedder, SESSION


def test_parse_cache_reuses_only_verified_bytes_and_never_returns_mutable_storage(tmp_path, monkeypatch):
    import office_agent.rag_sources as module

    path = tmp_path / 'report.txt'
    path.write_text('Original data 30 ms.', encoding='utf-8')
    extractor = SourceExtractor(Workspace(tmp_path))
    calls = []
    original = extractor._parse
    monkeypatch.setattr(extractor, '_parse', lambda *args: (calls.append(True), original(*args))[1])
    first = extractor.extract(path.name)
    first['blocks'][0]['text'] = 'Caller corruption'
    assert extractor.extract(path.name)['blocks'][0]['text'] == 'Original data 30 ms.'
    assert len(calls) == 1
    stamp = path.stat()
    path.write_text('Modified data 31 ms.', encoding='utf-8')
    os.utime(path, ns=(stamp.st_atime_ns, stamp.st_mtime_ns))
    changed = extractor.extract(path.name)
    assert changed['fingerprint'] != first['fingerprint'] and len(calls) == 2
    monkeypatch.setattr(module, 'PARSER_VERSION', 'fixture-parser-next')
    extractor.extract(path.name)
    assert len(calls) == 3
    extractor.clear_cache()
    assert not extractor._cache and extractor._cache_characters == 0


def test_parse_cache_eviction_and_path_guard_before_reuse(tmp_path, monkeypatch):
    import office_agent.rag_sources as module
    monkeypatch.setattr(module, 'CACHE_CHARACTERS', 100)
    extractor = SourceExtractor(Workspace(tmp_path))
    for index in range(10):
        name = f'file-{index}.txt'
        (tmp_path / name).write_text('x' * 80, encoding='utf-8')
        extractor.extract(name)
    assert len(extractor._cache) == 1 and extractor._cache_characters == 80
    (tmp_path / name).unlink()
    with pytest.raises(ValueError):
        extractor.extract(name)


def fixture(tmp_path):
    source = tmp_path / 'report.docx'
    doc = Document()
    doc.add_heading('实验结果与适用条件', 1)
    doc.add_paragraph('测量在 20 °C 条件下进行。')
    doc.add_paragraph('EXPT-02 结果 12.5 ms；样本30个。' + '附带原文解释。' * 180)
    doc.add_paragraph('限制：此实验不适用于高温条件。')
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = '条件 (°C)'
    table.cell(0, 1).text = '耗时 (ms)'
    table.cell(1, 0).text = '20'
    table.cell(1, 1).text = '12.5'
    doc.save(source)
    return source


def test_context_preserves_exact_hit_and_adds_bound_parent_heading_and_table_candidate(tmp_path):
    path = fixture(tmp_path)
    source = SourceExtractor(Workspace(tmp_path)).extract(path.name)
    chunks = token_structure_chunks(source['blocks'])
    long = next(chunk for chunk in chunks if chunk['text'].startswith('EXPT-02'))
    row = next(chunk for chunk in chunks if chunk['text'] == '20 | 12.5')
    hits = [{**chunk, 'id': index, 'doc_id': source['doc_id'], 'source_fingerprint': source['fingerprint']}
            for index, chunk in enumerate([long, row], 1)]
    contexts = expand_contexts(source, hits)
    assert [item['text'] for item in hits] == [long['text'], row['text']]
    assert contexts[0]['sections'][0]['text'].startswith(long['text'])
    assert contexts[0]['sections'][0]['truncated']
    assert any(section['role'] == 'heading' for section in contexts[0]['sections'])
    assert any(section['role'] == 'table_first_row' and 'ms' in section['text'] for section in contexts[1]['sections'])
    for context in contexts:
        assert not context['writable'] and context['tokens'] <= 768
        assert context['tokens'] == sum(token_count(section['text']) for section in context['sections'])
    corrupt = [dict(hits[0], text='Unrelated replacement')]
    with pytest.raises(RagError, match='source_changed'):
        expand_contexts(source, corrupt)


def test_profile_preflight_reuse_and_approved_context_lifecycle(tmp_path):
    source = fixture(tmp_path)
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    assert service.research_status(source.name, SESSION)['status'] == 'missing'
    assert not embedder.calls and not list(tmp_path.glob('.office-*'))
    plan = service.plan_research_index(source.name, SESSION)
    service.confirm_index(plan['plan_id'], SESSION, True)
    before = service.snapshot()
    assert service.research_status(source.name, SESSION)['status'] == 'current'
    with pytest.raises(RagError, match='rag_index_current'):
        service.plan_research_index(source.name, SESSION)
    assert service.snapshot() == before and len(embedder.calls) == 1
    query = service.plan_research_query(source.name, 'EXPT-02', SESSION, top_k=1)
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    saved = service.confirm_query(query['plan_id'], SESSION, True)
    hit = saved['hits'][0]
    context = service.read_research_context(saved['save_id'], hit['id'], SESSION)
    assert context['source_fingerprint'] == hit['source_fingerprint'] and context['tokens'] <= 768
    assert len(embedder.calls) == 2
    payload = json.loads((tmp_path / saved['evidence']).read_bytes())
    assert payload['version'] == 2 and payload['hits'][0]['text'] not in {'', None}
    assert payload['contexts'][0] == context
    source.write_bytes(source.read_bytes() + b'changed')
    with pytest.raises(EvidenceError, match='source_changed'):
        service.read_research_context(saved['save_id'], hit['id'], SESSION)


def test_stale_and_wrong_profile_are_rejected_before_query_embedding(tmp_path):
    source = tmp_path / 'report.txt'
    source.write_text('Controlled EXPT-02 data.', encoding='utf-8')
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    plan = service.plan_index(source.name, SESSION)
    service.confirm_index(plan['plan_id'], SESSION, True)
    assert service.research_status(source.name, SESSION)['status'] == 'profile_mismatch'
    with pytest.raises(VectorStoreError, match='index_profile_mismatch'):
        service.plan_research_query(source.name, 'EXPT-02', SESSION)
    assert len(embedder.calls) == 1
    source.write_text('Updated EXPT-02 data.', encoding='utf-8')
    assert service.research_status(source.name, SESSION)['status'] == 'stale'
    with pytest.raises(VectorStoreError, match='stale_source'):
        service.plan_research_query(source.name, 'EXPT-02', SESSION)
    assert len(embedder.calls) == 1


def test_context_is_not_granted_on_cancel_and_is_session_bound_after_save(tmp_path):
    source = fixture(tmp_path)
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    index = service.plan_research_index(source.name, SESSION)
    service.confirm_index(index['plan_id'], SESSION, True)
    query = service.plan_research_query(source.name, 'EXPT-02', SESSION, top_k=1)
    assert service.confirm_query(query['plan_id'], SESSION, False)['wrote_nothing']
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    with pytest.raises(EvidenceError, match='plan_invalid'):
        service.read_research_context(query['plan_id'], 1, SESSION)
    query = service.plan_research_query(source.name, 'EXPT-02', SESSION, top_k=1)
    saved = service.confirm_query(query['plan_id'], SESSION, True)
    hit_id = saved['hits'][0]['id']
    with pytest.raises(EvidenceError, match='session_invalid'):
        service.read_research_context(saved['save_id'], hit_id, 'b' * 32)
    service.invalidate()
    assert not service._source._cache
    with pytest.raises(EvidenceError, match='plan_invalid'):
        service.read_research_context(saved['save_id'], hit_id, SESSION)


def test_context_payload_tampering_never_returns_modified_text(tmp_path):
    source = fixture(tmp_path)
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    index = service.plan_research_index(source.name, SESSION)
    service.confirm_index(index['plan_id'], SESSION, True)
    query = service.plan_research_query(source.name, 'EXPT-02', SESSION, top_k=1)
    saved = service.confirm_query(query['plan_id'], SESSION, True)
    path = tmp_path / saved['evidence']
    document = json.loads(path.read_bytes())
    document['contexts'][0]['sections'][0]['text'] = 'Unapproved forged evidence'
    path.write_text(json.dumps(document), encoding='utf-8')
    with pytest.raises(EvidenceError, match='path_changed'):
        service.read_research_context(saved['save_id'], saved['hits'][0]['id'], SESSION)


def test_context_bridge_rejects_unknown_fields_and_authentication_before_factory(tmp_path):
    from types import SimpleNamespace
    from test_rag_api import make

    api, app, auth, root, source, embedder, factories = make(tmp_path)
    auth.authenticated = False
    assert api.get_research_context('a' * 32, 1)['reason'] == 'not_authenticated'
    assert factories == []
    auth.authenticated = True
    backend = SimpleNamespace(invalidate=lambda: None, read_research_context=lambda *args: {
        'writable': False, 'sections': [{'text': 'Unapproved raw content'}], 'tokens': 1, 'api_key': 'synthetic-private-value'})
    api._rag._factory = lambda check: backend
    result = api.get_research_context('a' * 32, 1)
    assert result['reason'] == 'rag_failed' and 'synthetic-private-value' not in json.dumps(result)
