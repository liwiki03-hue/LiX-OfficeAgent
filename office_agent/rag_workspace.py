"""Host workspace composition: parsed snapshot -> approved index -> saved evidence.

No client-selected root, authentication, UI, provider configuration, auto-indexing,
historical sessions or model tool registration here. The authenticated host
issues sessions and asks independently for each write; text grants no authority.
"""
import re

from .rag import RagService, RagError, _text
from .rag_sources import SourceExtractor
from .rag_evidence import RagEvidence
from .vectorstore import VectorStore, VectorStoreError


def _session(session):
    if type(session) is not str or not re.fullmatch('[a-f0-9]{32}', session):
        raise RagError('rag_invalid')


class WorkspaceRag:
    def __init__(self, workspace, embedder, *, scope_check=lambda: True):
        self._workspace = workspace
        self._scope_check = scope_check
        self._source = SourceExtractor(workspace)
        from .research_pdf import PdfResearchSource
        self._pdf_source = PdfResearchSource(workspace)
        # Validate identity before it is used to construct any database adapter.
        validator = RagService(None, embedder, self._version)
        identity = validator._identity()
        try:
            self._store = VectorStore(workspace.root, model=identity['model'], dimension=identity['dimension'])
        except VectorStoreError as error:
            if error.code != 'model_mismatch':
                raise
            from .research_corpus import checksum
            self._store = VectorStore(workspace.root, model=identity['model'], dimension=identity['dimension'], profile=checksum(identity))
        self._rag = RagService(self._store, embedder, self._version, store_error_type=VectorStoreError)
        self._evidence = RagEvidence(workspace, self._version)
        self._aliases = {}
        from .research_corpus import ResearchCorpus
        self._corpus = ResearchCorpus(workspace, self._source, self._store, self._rag, self._evidence, self._check, self._register_alias)
        self._reranker = None
        from .research_ranges import PdfRanges
        self._ranges = PdfRanges(self)

    def _register_alias(self, alias, original):
        self._aliases[alias] = original

    def _check(self):
        if self._scope_check() is not True:
            raise RagError('rag_source_changed')

    def _version(self, doc_id):
        self._check()
        original = self._aliases.get(doc_id, doc_id)
        source = self._pdf_source if original in self._pdf_source._bindings else self._source
        return source.fingerprint(original)

    def corpus_status(self, name, session):
        _session(session)
        return self._corpus.status(name)

    def plan_corpus(self, name, session):
        _session(session)
        return self._corpus.plan(name, session)

    def confirm_corpus(self, plan_id, session, confirmed):
        _session(session)
        self._check()
        return self._corpus.confirm(plan_id, session, confirmed)

    def plan_corpus_query(self, name, question, session, *, top_k=5, rerank=False, expand=False):
        _session(session)
        self._corpus.reranker = self._reranker
        return self._corpus.query(name, question, session, top_k=top_k, rerank=rerank, expand=expand)

    def __repr__(self):
        return '<WorkspaceRag redacted>'

    def bind_source(self, name, session):
        """Explicit authorized read for reopening an index; never auto-scans sources."""
        _session(session)
        self._check()
        source = self._source.extract(name)
        return {'doc_id': source['doc_id'], 'source_fingerprint': source['fingerprint'],
                'block_count': len(source['blocks']), 'wrote_nothing': True}

    def plan_index(self, name, session):
        _session(session)
        self._check()
        source = self._source.extract(name)
        plan = self._rag.plan_index(source['doc_id'], source['blocks'], session)
        try:
            # The core accepts trusted pre-parsed blocks; bind them to the exact
            # extracted bytes, not a newer version observed during plan creation.
            if (plan['source_fingerprint'] != source['fingerprint'] or
                    self._version(source['doc_id']) != source['fingerprint']):
                raise RagError('rag_source_changed')
            return plan
        except Exception:
            self._rag.confirm_index(plan['plan_id'], session, False)
            raise RagError('rag_source_changed') from None

    def confirm_index(self, plan_id, session, confirmed):
        _session(session)
        self._check()
        if plan_id in self._ranges.pending:
            return self._ranges.confirm(plan_id, session, confirmed)
        return self._rag.confirm_index(plan_id, session, confirmed)

    def pdf_ranges_status(self, name, session):
        _session(session)
        return self._ranges.status(name)

    def plan_pdf_ranges(self, name, ranges, session):
        _session(session)
        return self._ranges.plan(name, ranges, session)

    def query_pdf_ranges(self, name, question, session, top_k=5):
        _session(session)
        return self._ranges.query(name, question, session, top_k)

    def plan_research_index(self, name, session, *, start_page=None, end_page=None):
        from .token_chunks import token_structure_chunks

        _session(session)
        self._check()
        if self._rag._disabled:
            raise RagError('rag_uncertain')
        source, profile = self._research_source(name, start_page=start_page, end_page=end_page)
        if profile['status'] == 'current':
            raise RagError('rag_index_current')
        from .rag_sources import PARSER_VERSION
        plan = self._rag.plan_index(source['doc_id'], source['blocks'], session,
                                   chunker=lambda blocks: token_structure_chunks(blocks, parser_version=source.get('parser_version', PARSER_VERSION)))
        if plan['source_fingerprint'] != source['fingerprint'] or self._version(source['doc_id']) != source['fingerprint']:
            self._rag.confirm_index(plan['plan_id'], session, False)
            raise RagError('rag_source_changed')
        return {**plan, **({'pdf_coverage': source['pdf_coverage']} if 'pdf_coverage' in source else {})}

    def _research_source(self, name, *, start_page=None, end_page=None):
        from .rag_sources import PARSER_VERSION
        from .token_chunks import CHUNK_VERSION, TOKENIZER_VERSION

        self._check()
        if self._rag._disabled:
            raise RagError('rag_uncertain')
        extractor = self._pdf_source if type(name) is str and name.lower().endswith('.pdf') else self._source
        if start_page is not None or end_page is not None:
            if extractor is not self._pdf_source:
                raise RagError('rag_invalid')
            source = extractor.extract(name, start_page=start_page, end_page=end_page)
        else:
            source = extractor.extract(name)
        profile = self._store.research_profile(source['doc_id'], source['fingerprint'], chunk_version=CHUNK_VERSION,
                                               tokenizer_version=TOKENIZER_VERSION, parser_version=source.get('parser_version', PARSER_VERSION))
        if self._version(source['doc_id']) != source['fingerprint']:
            raise RagError('rag_source_changed')
        return source, profile

    def research_status(self, name, session, *, start_page=None, end_page=None):
        _session(session)
        source, profile = self._research_source(name, start_page=start_page, end_page=end_page)
        return {**profile, 'source': name, 'source_fingerprint': source['fingerprint'], 'wrote_nothing': True,
                **({'pdf_coverage': source['pdf_coverage']} if 'pdf_coverage' in source else {})}

    def plan_research_query(self, name, question, session, *, top_k=5, start_page=None, end_page=None):
        from .token_chunks import CHUNK_VERSION

        _session(session)
        self._check()
        if self._rag._disabled:
            raise RagError('rag_uncertain')
        _text(question, 200)
        if type(top_k) is not int or not 1 <= top_k <= 10:
            raise RagError()
        self._evidence.ensure_available(session)
        source, profile = self._research_source(name, start_page=start_page, end_page=end_page)
        if profile['status'] != 'current':
            raise VectorStoreError({'missing': 'document_missing', 'stale': 'stale_source',
                                    'profile_mismatch': 'index_profile_mismatch'}[profile['status']])
        version = source['fingerprint']
        identity = self._rag._identity()
        snapshot = self._rag._snapshot()
        self._rag._compatible(snapshot, identity)
        if snapshot.get('exists') is not True or snapshot['chunk_count'] == 0:
            raise VectorStoreError('document_missing')
        vectors = self._rag._vectors(getattr(self._rag.embedder, 'embed_queries', self._rag.embedder.embed)([question]), 1)
        if self._version(source['doc_id']) != version or self._rag._snapshot() != snapshot or self._rag._identity() != identity:
            raise RagError('rag_source_changed')
        hits = self._store.query_scoped(vectors[0], keyword=question, sources={source['doc_id']: version},
                                        candidates=50, top_k=top_k, chunk_version=CHUNK_VERSION)
        if self._version(source['doc_id']) != version or self._rag._snapshot() != snapshot or self._rag._identity() != identity:
            raise RagError('rag_source_changed')
        from .research_context import expand_contexts
        contexts = expand_contexts(source, hits)
        plan = self._evidence.plan(question, hits, session, contexts=contexts)
        return {**plan, **({'pdf_coverage': source['pdf_coverage']} if 'pdf_coverage' in source else {})}

    def read_research_context(self, save_id, hit_id, session):
        _session(session)
        self._check()
        return self._evidence.read_context(save_id, hit_id, session)

    def plan_query(self, question, session, *, top_k=5):
        _session(session)
        self._check()
        # Internal full hits are never returned. Provider work is a retrieval
        # read, but full evidence persistence still needs a separate human True.
        self._evidence.ensure_available(session)
        hits = self._rag.query(question, top_k=top_k)
        return self._evidence.plan(question, hits, session)

    def plan_directory_query(self, name, question, session, *, top_k=5):
        from .directory_search import DirectoryIndex

        _session(session)
        self._check()
        self._evidence.ensure_available(session)
        _text(question, 200)
        if type(top_k) is not int or not 1 <= top_k <= 10:
            raise RagError()
        if type(name) is not str or '/' in name or '\\' in name or not name:
            raise RagError()
        try:
            index = DirectoryIndex(self._workspace, self._workspace.root / name,
                                   embedder=self._rag.embedder).build()
            report = index.query(question, top_k=top_k, _full=True)
            hits = report['hits']
            for source in index.documents:
                binding = self._source.extract(source['name'])
                if binding['fingerprint'] != source['fingerprint']:
                    raise RagError('rag_source_changed')
            self._check()
            plan = self._evidence.plan(question, hits, session)
            directory = {key: report[key] for key in ('document_count', 'skipped_document_count',
                                                       'corpus_truncated', 'coverage_complete')}
            directory['inventory'] = [{key: item.get(key) for key in ('source', 'status', 'reason')}
                                      for item in report['inventory']]
            return {**plan, 'directory': directory}
        except RagError:
            raise
        except Exception:
            raise RagError('rag_query_failed') from None

    def confirm_query(self, plan_id, session, confirmed):
        _session(session)
        self._check()
        if type(plan_id) is str and plan_id in self._ranges.queries:
            return self._ranges.confirm_query(plan_id, session, confirmed)
        if type(plan_id) is str and plan_id in self._corpus.queries:
            return self._corpus.confirm_query(plan_id, session, confirmed)
        return self._evidence.confirm(plan_id, session, confirmed)

    def read_evidence(self, save_id, session, *, offset=0, limit=2000):
        _session(session)
        self._check()
        return self._evidence.read(save_id, session, offset=offset, limit=limit)

    def plan_evidence_rollback(self, save_id, session):
        _session(session)
        self._check()
        return self._evidence.plan_rollback(save_id, session)

    def confirm_evidence_rollback(self, rollback_id, session, confirmed):
        _session(session)
        self._check()
        return self._evidence.confirm_rollback(rollback_id, session, confirmed)

    def plan_index_rollback(self, save_id, session):
        _session(session)
        self._check()
        return self._store.plan_rollback(save_id, session)

    def confirm_index_rollback(self, rollback_id, session, confirmed):
        _session(session)
        self._check()
        return self._store.confirm_rollback(rollback_id, session, confirmed)

    def snapshot(self):
        return self._store.snapshot()

    def invalidate(self):
        self._rag.invalidate()
        self._evidence.invalidate()
        self._source.clear_cache()
        self._pdf_source.clear_cache()
        self._corpus.invalidate()
        self._ranges.invalidate()
        # Retire grants, preserve artifacts, the parsed identity bindings, and fuses.
