import json
from pathlib import Path
import time

import pytest

from office_agent.embeddings import Embedder, EmbeddingConfig, EmbeddingError
from office_agent.research_onnx import bundle_identity


def bundle(tmp_path, monkeypatch):
    monkeypatch.setattr('office_agent.research_onnx.version', lambda dependency: 'test-version')
    for name in ('model.onnx', 'tokenizer.json', 'config.json', 'MODEL_CARD.md'):
        (tmp_path / name).write_bytes(b'fixture-not-real-weights')
    manifest = {'format': 'office-onnx-v1', 'task': 'embedding', 'repo': 'test', 'revision': 'fixed', 'license': 'mit',
                'dimension': 384, 'max_tokens': 512, 'normalization': 'l2', 'pooling': 'attention_mean',
                'query_prefix': 'query: ', 'document_prefix': 'passage: ', 'artifacts': []}
    (tmp_path / 'research_model.json').write_text(json.dumps(manifest))
    return manifest


def test_identity_pins_weights_tokenizer_runtime_and_conventions(tmp_path, monkeypatch):
    manifest = bundle(tmp_path, monkeypatch)
    identity = lambda: bundle_identity(str(tmp_path), 'embedding', time.monotonic() + 5)[0]
    first = identity()
    (tmp_path / 'model.onnx').write_bytes(b'changed-weights-same-directory')
    second = identity()
    assert first != second
    (tmp_path / 'tokenizer.json').write_bytes(b'changed-tokenizer')
    third = identity()
    assert second != third
    monkeypatch.setattr('office_agent.research_onnx.version', lambda dependency: 'new-version')
    assert third != identity()
    manifest['query_prefix'] = 'unexpected prefix'
    (tmp_path / 'research_model.json').write_text(json.dumps(manifest))
    with pytest.raises(ValueError, match='onnx_metadata'):
        identity()


def test_queries_and_documents_use_distinct_host_modes_and_change_refuses(tmp_path, monkeypatch):
    bundle(tmp_path, monkeypatch)
    seen = []

    def run(model, task, texts, timeout, **kwargs):
        seen.append(kwargs['query_mode'])
        return [[1.] * 384 for _ in texts]

    monkeypatch.setattr('office_agent.research_onnx.run_owned', run)
    embedder = Embedder(EmbeddingConfig('local', str(tmp_path), 384))
    embedder.embed(['original'])
    embedder.embed_queries(['question'])
    assert seen == [False, True]
    (tmp_path / 'model.onnx').write_bytes(b'changed')
    with pytest.raises(EmbeddingError, match='embedding_local_changed'):
        embedder.embed_queries(['question'])


def test_executable_or_linked_bundle_rejected_and_cleanup_failure_sticks(tmp_path, monkeypatch):
    bundle(tmp_path, monkeypatch)
    forbidden = tmp_path / 'custom.py'
    forbidden.write_text('raise RuntimeError()')
    with pytest.raises(ValueError, match='onnx_bundle_files'):
        bundle_identity(str(tmp_path), 'embedding', time.monotonic() + 5)
    forbidden.unlink()
    calls = []

    def failed(*args, **kwargs):
        calls.append(True)
        raise ValueError('onnx_cleanup_failed')

    monkeypatch.setattr('office_agent.research_onnx.run_owned', failed)
    embedder = Embedder(EmbeddingConfig('local', str(tmp_path), 384))
    for _ in range(2):
        with pytest.raises(EmbeddingError):
            embedder.embed(['safe fixture'])
    assert len(calls) == 1 and embedder._disabled


def test_temp_cleanup_exception_does_not_hide_owned_process_uncertainty(tmp_path, monkeypatch):
    from office_agent.diagram_process import OwnedRun
    from office_agent.research_onnx import run_owned

    bundle(tmp_path, monkeypatch)
    scratch = tmp_path / 'scratch'
    scratch.mkdir()

    class Scratch:
        def __init__(self, **kwargs):
            pass

        def __enter__(self):
            return str(scratch)

        def __exit__(self, *args):
            raise OSError('temporary-file-locked')

    monkeypatch.setattr('office_agent.research_onnx.bundle_identity', lambda *args, **kwargs: ('fixture', {}))
    monkeypatch.setattr('office_agent.research_onnx.tempfile.TemporaryDirectory', Scratch)
    monkeypatch.setattr('office_agent.diagram_process._run_owned', lambda *args, **kwargs: OwnedRun('process_failed', 1, 0, 0, False))
    with pytest.raises(ValueError, match='onnx_cleanup_failed'):
        run_owned(str(tmp_path), 'embedding', ['fixture'], 1)
