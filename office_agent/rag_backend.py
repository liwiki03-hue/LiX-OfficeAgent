"""Host-only RAG construction from the host-bound canonical settings directory.

Never register this factory as a model tool or pass browser-selected roots to it.
SettingsStore._read and _values are intentional private dependencies: they retain
its bounded, identity-checked read and non-interpolating canonical .env parser.
"""
from .embeddings import EmbeddingConfig, EmbeddingError, Embedder
from .rag_workspace import WorkspaceRag
from .settings_store import SettingsError, SettingsStore


def make_rag_backend(workspace, settings_root, *, scope_check=lambda: True):
    """Compose a backend using trusted host bindings, without ambient env fallback."""
    try:
        settings = SettingsStore(settings_root)
        source = settings._read(settings.root / '.env')
        values = settings._values(source)
    except SettingsError as error:
        # Hide any underlying I/O/parser exception text, including config values.
        raise SettingsError(error.code) from None
    except Exception:
        raise SettingsError('settings_file_invalid') from None

    try:
        embedder = Embedder(EmbeddingConfig.from_values(values))
    except EmbeddingError as error:
        # from_values may retain a conversion exception containing input text.
        raise EmbeddingError(error.code) from None
    except Exception:
        raise EmbeddingError('embedding_failed') from None
    backend = WorkspaceRag(workspace, embedder, scope_check=scope_check)
    if values.get('OFFICE_RERANKER_MODEL'):
        from .research_reranker import OfflineReranker
        try:
            backend._reranker = OfflineReranker(values['OFFICE_RERANKER_MODEL'], float(values.get('OFFICE_RERANKER_TIMEOUT') or 8))
        except (ValueError, TypeError):
            backend._reranker = OfflineReranker('', 8)
    return backend
