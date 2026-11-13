"""Host adapter fixtures: synthetic temporary .env only, no providers or Office."""
import os
import traceback
from unittest.mock import Mock

import pytest

from office_agent import rag_backend as m
from office_agent.settings_store import MAX_ENV_BYTES


CONFIG = (
    'OFFICE_EMBEDDING_CHANNEL=api\n'
    'OFFICE_EMBEDDING_MODEL=text-embedding-v3\n'
    'OFFICE_EMBEDDING_DIMENSION=64\n'
    'DASHSCOPE_API_KEY=fixture-private-key\n'
    'DASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1\n'
)


@pytest.fixture
def constructors(monkeypatch):
    embedder = Mock(name='host_embedder')
    backend = Mock(name='workspace_backend')
    monkeypatch.setattr(m, 'Embedder', embedder)
    monkeypatch.setattr(m, 'WorkspaceRag', backend)
    return embedder, backend


def test_host_composition_preserves_scope_and_canonical_values(tmp_path, constructors, capsys):
    path = tmp_path / '.env'
    path.write_text(CONFIG, encoding='utf-8')
    original = path.read_bytes()
    environment = dict(os.environ)
    workspace = object()
    scope_check = lambda: False
    embedder, backend = constructors
    result = m.make_rag_backend(workspace, tmp_path, scope_check=scope_check)
    assert result is backend.return_value
    config = embedder.call_args.args[0]
    assert config.channel == 'api' and config.model == 'text-embedding-v3'
    assert config.dimension == 64 and config.api_key == 'fixture-private-key'
    backend.assert_called_once_with(workspace, embedder.return_value, scope_check=scope_check)
    assert path.read_bytes() == original and list(tmp_path.iterdir()) == [path]
    assert dict(os.environ) == environment
    assert capsys.readouterr() == ('', '')


def test_default_scope_callback(tmp_path, constructors):
    (tmp_path / '.env').write_text(CONFIG, encoding='utf-8')
    m.make_rag_backend(object(), tmp_path)
    assert constructors[1].call_args.kwargs['scope_check']() is True


@pytest.mark.parametrize('data,code', [
    (None, 'settings_io_failed'),
    (b'\xff', 'settings_file_invalid'),
    (b'\x00fixture-private-key', 'settings_file_invalid'),
    (b'x' * (MAX_ENV_BYTES + 1), 'settings_file_invalid'),
    ((CONFIG + 'OFFICE_EMBEDDING_MODEL=duplicate\n').encode(), 'settings_file_invalid'),
    (CONFIG.replace('text-embedding-v3', '"fixture-private-key').encode(), 'settings_file_invalid'),
], ids=['missing', 'invalid-utf8', 'nul', 'oversized', 'duplicate', 'malformed-quote'])
def test_missing_or_invalid_canonical_file_is_guarded(tmp_path, constructors, data, code):
    if data is not None:
        (tmp_path / '.env').write_bytes(data)
    with pytest.raises(m.SettingsError) as failure:
        m.make_rag_backend(object(), tmp_path)
    assert str(failure.value) == code and failure.value.code == code
    assert 'fixture-private-key' not in ''.join(traceback.format_exception(failure.value))
    for constructor in constructors:
        constructor.assert_not_called()


@pytest.mark.parametrize('data', [
    '',
    CONFIG.replace('OFFICE_EMBEDDING_DIMENSION=64\n', ''),
    CONFIG.replace('64', 'fixture-private-key'),
    CONFIG.replace('api\n', 'unknown\n'),
    CONFIG.replace('DASHSCOPE_API_KEY=fixture-private-key\n', ''),
    CONFIG.replace('DASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1\n', ''),
    CONFIG.replace('text-embedding-v3', 'fixture-private-key'),
    CONFIG + 'OFFICE_EMBEDDING_TIMEOUT=fixture-private-key\n',
], ids=['empty', 'missing-dimension', 'invalid-dimension', 'invalid-channel',
        'missing-key', 'missing-url', 'invalid-model', 'invalid-timeout'])
def test_invalid_embedding_values_fail_redacted_before_construction(tmp_path, constructors, data):
    (tmp_path / '.env').write_text(data, encoding='utf-8')
    with pytest.raises(m.EmbeddingError) as failure:
        m.make_rag_backend(object(), tmp_path)
    assert str(failure.value) == 'embedding_not_configured'
    assert 'fixture-private-key' not in ''.join(traceback.format_exception(failure.value))
    for constructor in constructors:
        constructor.assert_not_called()


def test_no_environment_fallback_or_interpolation(tmp_path, constructors, monkeypatch):
    monkeypatch.setenv('DASHSCOPE_API_KEY', 'fixture-ambient-key')
    monkeypatch.setenv('OFFICE_EMBEDDING_DIMENSION', '64')
    path = tmp_path / '.env'
    path.write_text(CONFIG.replace('OFFICE_EMBEDDING_DIMENSION=64',
                                   'OFFICE_EMBEDDING_DIMENSION=${OFFICE_EMBEDDING_DIMENSION}'), encoding='utf-8')
    environment = dict(os.environ)
    with pytest.raises(m.EmbeddingError, match='^embedding_not_configured$'):
        m.make_rag_backend(object(), tmp_path)
    path.write_text(CONFIG.replace('DASHSCOPE_API_KEY=fixture-private-key\n', ''), encoding='utf-8')
    with pytest.raises(m.EmbeddingError, match='^embedding_not_configured$'):
        m.make_rag_backend(object(), tmp_path)
    assert dict(os.environ) == environment
    for constructor in constructors:
        constructor.assert_not_called()


def test_directory_in_place_of_file_is_rejected(tmp_path, constructors):
    (tmp_path / '.env').mkdir()
    with pytest.raises(m.SettingsError, match='^settings_file_invalid$'):
        m.make_rag_backend(object(), tmp_path)
    for constructor in constructors:
        constructor.assert_not_called()


def test_missing_root_is_redacted(tmp_path, constructors):
    with pytest.raises(m.SettingsError, match='^settings_path_changed$'):
        m.make_rag_backend(object(), tmp_path / 'fixture-private-key')
    for constructor in constructors:
        constructor.assert_not_called()


def test_unexpected_embedder_failure_is_redacted(tmp_path, constructors):
    (tmp_path / '.env').write_text(CONFIG, encoding='utf-8')
    constructors[0].side_effect = RuntimeError('fixture-private-key')
    with pytest.raises(m.EmbeddingError) as failure:
        m.make_rag_backend(object(), tmp_path)
    assert str(failure.value) == 'embedding_failed'
    assert 'fixture-private-key' not in ''.join(traceback.format_exception(failure.value))
    constructors[1].assert_not_called()
