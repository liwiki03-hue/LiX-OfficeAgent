import hashlib
import json
import time
import uuid

from .directory_search import DirectoryIndex
from .rag import RagError, _text
from .rag_sources import PARSER_VERSION
from .token_chunks import CHUNK_VERSION, TOKENIZER_VERSION, token_structure_chunks
from .vectorstore import VectorStoreError
from .research_context import expand_contexts


def checksum(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode('utf-8')).hexdigest()


class ResearchCorpus:
    def __init__(self, workspace, source, store, rag, evidence, check, register):
        self.space, self.source, self.store, self.rag, self.evidence = workspace, source, store, rag, evidence
        self.check, self.register = check, register
        self.pending = {}
        self.queries = {}
        self.inspections = {}

    def invalidate(self):
        self.pending.clear()
        self.queries.clear()
        self.inspections.clear()

    def inventory(self, name):
        self.check()
        if self.rag._disabled:
            raise RagError('rag_uncertain')
        if type(name) is not str or '/' in name or '\\' in name:
            raise RagError('rag_invalid')
        selected = self.space.file(str(self.space.root / name))
        index = DirectoryIndex(self.space, selected, extractor=self.source, inspection_cache=self.inspections).build()
        self.inspections = {item['name']: self.inspections[item['name']] for item in index.documents if item['name'] in self.inspections}
        root = self.space.root.stat()
        scope = checksum([str(self.space.root), root.st_dev, root.st_ino])
        sources, chunks, total = {}, {}, 0
        for document in index.documents:
            source = self.source.extract(document['name'])
            if source['fingerprint'] != document['fingerprint']:
                raise RagError('rag_source_changed')
            original = source['doc_id']
            doc_id = checksum(['directory', scope, document['name'].casefold()])
            self.register(doc_id, original)
            source = {**source, 'doc_id': doc_id}
            records = token_structure_chunks(source['blocks'])
            for record in records:
                where = json.loads(record['location'])
                where['corpus_scope'] = scope
                record['location'] = json.dumps(where, ensure_ascii=False, separators=(',', ':'))
                if len(record['location']) > 512:
                    raise RagError('rag_capacity')
            total += sum(len(row['text']) for row in records)
            if sum(map(len, chunks.values())) + len(records) > 500 or total > 1_000_000:
                raise RagError('rag_capacity')
            sources[doc_id], chunks[doc_id] = source, records
        names = [[item['source'], item['status'], item.get('reason')] for item in index.inventory
                 if not item['source'].startswith('.office-')]
        inventory_hash = checksum([index.anchor_sha256, names, {key: value['fingerprint'] for key, value in sources.items()}])
        self.check()
        return {'scope': scope, 'sources': sources, 'chunks': chunks, 'hash': inventory_hash,
                'inventory': index.inventory, 'anchor': index.anchor_sha256}

    def inspect(self, name):
        current = self.inventory(name)
        old = self.store.corpus_manifest(current['scope'])
        changed = []
        for key, source in current['sources'].items():
            previous = old.get(key)
            valid = previous is not None and previous['fingerprint'] == source['fingerprint'] and all(
                where.get('chunker') == CHUNK_VERSION and where.get('tokenizer') == TOKENIZER_VERSION and
                where.get('parser') == PARSER_VERSION for where in previous['locations'])
            if not valid:
                changed.append(key)
        removed = sorted(set(old) - set(current['sources']))
        return current, changed, removed

    def status(self, name):
        current, changed, removed = self.inspect(name)
        return {'status': 'current' if not changed and not removed else 'needs_sync',
                'documents': len(current['sources']), 'changed': len(changed), 'removed': len(removed),
                'unchanged': len(current['sources']) - len(changed), 'inventory': current['inventory'],
                'corpus_sha256': current['hash'], 'wrote_nothing': True}

    def plan(self, name, session):
        current, changed, removed = self.inspect(name)
        if not changed and not removed:
            raise RagError('rag_index_current')
        self.pending = {key: value for key, value in self.pending.items() if time.monotonic() - value['created'] < 300}
        if len(self.pending) >= 4:
            raise RagError('rag_capacity')
        token = uuid.uuid4().hex
        self.pending[token] = {'created': time.monotonic(), 'session': session, 'name': name, 'current': current,
                               'changed': changed, 'removed': removed, 'snapshot': self.store.snapshot(),
                               'identity': self.rag._identity()}
        return {'plan_id': token, 'operation': 'sync_directory_index', 'confirmation_required': True,
                'changed': [current['sources'][key]['blocks'][0]['source'] for key in changed],
                'removed': [self.store.corpus_manifest(current['scope'])[key]['source'] for key in removed],
                'unchanged': len(current['sources']) - len(changed), 'chunk_count': sum(len(current['chunks'][key]) for key in changed),
                'inventory': current['inventory'], 'embedding': self.rag._identity(), 'expires_in_seconds': 300, 'wrote_nothing': True}

    def confirm(self, token, session, confirmed):
        plan = self.pending.pop(token, None) if type(token) is str else None
        denied = {'status': 'rejected', 'applied': False, 'wrote_nothing': True}
        if plan is None or plan['session'] != session or confirmed is not True or time.monotonic() - plan['created'] >= 300:
            return denied
        try:
            def fresh():
                return (self.inventory(plan['name'])['hash'] == plan['current']['hash'] and
                        self.store.snapshot() == plan['snapshot'] and self.rag._identity() == plan['identity'])
            if not fresh():
                return denied
            documents = {}
            texts = [chunk['text'] for key in plan['changed'] for chunk in plan['current']['chunks'][key]]
            vectors = self.rag._vectors(self.rag.embedder.embed(texts), len(texts)) if texts else []
            if not fresh():
                return denied
            position = 0
            for key in plan['changed']:
                records = []
                for chunk in plan['current']['chunks'][key]:
                    records.append({**chunk, 'embedding': vectors[position]})
                    position += 1
                documents[key] = {'fingerprint': plan['current']['sources'][key]['fingerprint'], 'records': records}
            inner = self.store.plan_sync(plan['current']['scope'], documents, plan['removed'], session)
        except Exception:
            return denied
        try:
            result = self.store.confirm_replace(inner['plan_id'], session, True)
            if result.get('uncertain') is not False or result.get('revision') != plan['snapshot']['revision'] + 1:
                raise ValueError()
            return {**result, 'status': 'saved', 'applied': True, 'wrote_nothing': False}
        except Exception:
            self.rag._disabled = True
            return {'status': 'uncertain', 'applied': None, 'wrote_nothing': False}

    def query(self, name, question, session, *, top_k=5, rerank=False, expand=False):
        from .retrieval_pipeline import retrieve

        _text(question, 200)
        if type(top_k) is not int or not 1 <= top_k <= 10 or type(rerank) is not bool or type(expand) is not bool:
            raise RagError('rag_invalid')
        self.evidence.ensure_available(session)
        current, changed, removed = self.inspect(name)
        if changed or removed:
            raise RagError('rag_corpus_stale')
        snapshot, identity = self.store.snapshot(), self.rag._identity()
        versions = {key: source['fingerprint'] for key, source in current['sources'].items()}
        hits, report = retrieve(self.store, self.rag, question, versions, top_k=top_k, rerank=rerank,
                                expand=expand, reranker=getattr(self, 'reranker', None))
        if self.inventory(name)['hash'] != current['hash'] or self.store.snapshot() != snapshot or self.rag._identity() != identity:
            raise RagError('rag_source_changed')
        contexts = []
        for hit in hits:
            contexts.extend(expand_contexts(current['sources'][hit['doc_id']], [hit]))
        report = {**report, 'scope': 'directory', 'documents': len(versions), 'corpus_sha256': current['hash'],
                  'coverage_complete': not any(item['status'] != 'indexed_in_memory' and item['source'].lower().endswith(('.docx', '.xlsx')) and item.get('reason') != 'generated_or_lock_file' for item in current['inventory'])}
        plan = self.evidence.plan(question, hits, session, contexts=contexts, retrieval=report)
        self.queries = {key: value for key, value in self.queries.items() if time.monotonic() - value[3] < 300}
        self.queries[plan['plan_id']] = (name, current['hash'], session, time.monotonic())
        return {**plan, 'retrieval': report}

    def confirm_query(self, token, session, confirmed):
        lease = self.queries.pop(token)
        try:
            if lease[2] != session or self.inventory(lease[0])['hash'] != lease[1]:
                confirmed = False
        except Exception:
            confirmed = False
        return self.evidence.confirm(token, session, confirmed)
