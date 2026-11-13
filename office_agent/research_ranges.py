import hashlib
import json
import time
import uuid

from .rag import RagError, _text
from .research_corpus import checksum
from .token_chunks import CHUNK_VERSION, TOKENIZER_VERSION, token_structure_chunks


def processing_stamp():
    from .research_pdf import PDF_PARSER_VERSION, _ADAPTER_VERSION
    return checksum([PDF_PARSER_VERSION, _ADAPTER_VERSION, CHUNK_VERSION, TOKENIZER_VERSION])


def range_scope(workspace, name):
    root = workspace.root.stat()
    return checksum(['pdf-ranges-v1', str(workspace.root), root.st_dev, root.st_ino, name.casefold()])[:24]


class PdfRanges:
    def __init__(self, owner):
        self.owner = owner
        self.pending = {}
        self.attempt = None
        self.cache = {}
        self.queries = {}

    def invalidate(self):
        self.pending.clear()
        self.cache.clear()
        self.queries.clear()
        self.attempt = None

    def bind(self, name):
        o = self.owner
        o._check()
        if o._rag._disabled:
            raise RagError('rag_uncertain')
        if type(name) is not str or not 1 <= len(name) <= 128 or name.lower().startswith('.env') or not name.lower().endswith('.pdf') or any(c in name for c in '/\\:'):
            raise RagError('rag_invalid')
        data, identity, version = o._pdf_source._snapshot(name)
        if not data.startswith(b'%PDF-'):
            raise RagError('rag_invalid')
        original = hashlib.sha256(name.casefold().encode()).hexdigest()
        from .rag_sources import MAX_IDENTITIES
        if original not in o._pdf_source._bindings and len(o._pdf_source._bindings) >= MAX_IDENTITIES:
            raise RagError('rag_capacity')
        o._pdf_source._bindings[original] = (name, identity)
        scope = range_scope(o._workspace, name)
        old = o._store.corpus_manifest(scope)
        for key in old:
            o._register_alias(key, original)
        return scope, original, version, old

    def parser(self, page, total):
        from .research_pdf import PDF_PARSER_VERSION, _ADAPTER_VERSION
        return f'pdf-range-v1:{page}-{page}/{total}:' + hashlib.sha256((PDF_PARSER_VERSION + _ADAPTER_VERSION).encode()).hexdigest()[:24]

    def embedding_stamp(self):
        identity = self.owner._rag._identity()
        if identity['channel'] == 'api':
            config = getattr(self.owner._rag.embedder, 'config', None)
            if config is None or type(getattr(config, 'base_url', None)) is not str or type(getattr(config, 'api_key', None)) is not str:
                raise RagError('rag_model_mismatch')
            credential_scope = hashlib.sha256(config.api_key.encode()).hexdigest()
            return checksum([identity, config.base_url, credential_scope])
        return checksum(identity)

    def valid(self, row, version):
        return row['fingerprint'] == version and all(
            w.get('chunker') == CHUNK_VERSION and w.get('tokenizer') == TOKENIZER_VERSION and
            w.get('ei') == self.embedding_stamp() and
            type(w.get('pdf_total')) is int and w.get('parser') == self.parser(int(w['group'].split(':')[1]), w['pdf_total'])
            for w in row['locations'])

    def status(self, name):
        scope, original, version, old = self.bind(name)
        pages, total, stale = [], None, False
        for row in old.values():
            stale |= not self.valid(row, version)
            for w in row['locations']:
                pages.append(int(w['group'].split(':')[1]))
                total = w['pdf_total']
        pages = sorted(set(pages))
        attempt = None
        if self.attempt and self.attempt['scope'] == scope:
            from copy import deepcopy
            attempt = deepcopy(self.attempt)
            attempt['source_is_current'] = (attempt['source_fingerprint'] == version and attempt['embedding_stamp'] == self.embedding_stamp() and attempt['processing_stamp'] == processing_stamp())
            attempt['unprocessed_pages'] = sorted(set(attempt['selected_pages']) - set(attempt['parsed_pages']) - set(attempt['failed_pages']))
            if attempt['state'] == 'awaiting_confirmation' and not any(p.get('attempt_id') == attempt['id'] and time.monotonic() - p['created'] < 300 for p in self.pending.values()):
                attempt['state'] = 'expired'
        return {'status': 'stale' if stale else 'current' if pages else 'missing', 'pages': pages,
                'selected_pages': pages, 'parsed_pages': pages if not stale else [],
                'indexed_pages': pages if not stale else [], 'failed_pages': [],
                'native_pages': pages if not stale else [], 'ocr_pages': [], 'total_pages': total,
                'gaps': [n for n in range(1, total + 1) if n not in pages] if total else [],
                'complete': bool(total and len(pages) == total and not stale),
                'generation': self.owner._store.snapshot()['revision'], 'source_fingerprint': version,
                'attempt': attempt, 'attempt_storage': 'session_memory_only', 'wrote_nothing': True}

    def plan(self, name, ranges, session):
        previous = self.attempt
        try:
            result = self._plan(name, ranges, session)
            if self.attempt is not None:
                self.attempt['state'] = 'awaiting_confirmation'
            return result
        except Exception as error:
            if self.attempt is not None and self.attempt is not previous:
                self.attempt['state'] = 'failed'
                code = getattr(error, 'code', '')
                self.attempt['reason'] = code if code in {'rag_capacity', 'rag_source_changed', 'pdf_text_missing', 'pdf_text_truncated', 'rag_invalid', 'rag_model_mismatch'} else 'processing_failed'
            raise

    def _plan(self, name, ranges, session):
        if type(ranges) is not list or not 0 <= len(ranges) <= 20:
            raise RagError('rag_invalid')
        pages = set()
        for pair in ranges:
            if type(pair) is not list or len(pair) != 2 or any(type(n) is not int for n in pair) or not 1 <= pair[0] <= pair[1] <= 2000:
                raise RagError('rag_invalid')
            pages.update(range(pair[0], pair[1] + 1))
        if len(pages) > 20:
            raise RagError('rag_capacity')
        scope, original, version, old = self.bind(name)
        stamp = self.embedding_stamp()
        self.pending = {k: p for k, p in self.pending.items() if time.monotonic() - p['created'] < 300 and p['session'] != session}
        self.attempt = {'id': uuid.uuid4().hex, 'scope': scope, 'source_fingerprint': version, 'embedding_stamp': stamp,
                        'processing_stamp': processing_stamp(),
                        'selected_pages': sorted(pages), 'parsed_pages': [], 'reused_pages': [], 'failed_pages': [],
                        'state': 'building', 'published': False}
        documents, retained, total, chars = {}, [], None, 0
        started = time.monotonic()
        for page in sorted(pages):
            if time.monotonic() - started > 60:
                raise RagError('rag_capacity')
            key = checksum([scope, page])
            row = old.get(key)
            if row and self.valid(row, version):
                retained.append(key)
                self.attempt['parsed_pages'].append(page)
                self.attempt['reused_pages'].append(page)
                chars += sum(w['char_end'] - w['char_start'] for w in row['locations'])
                total = row['locations'][0]['pdf_total']
                continue
            try:
                from .research_pdf import PDF_PARSER_VERSION, _ADAPTER_VERSION
                from copy import deepcopy
                cache_key = checksum([scope, version, page, PDF_PARSER_VERSION, _ADAPTER_VERSION, CHUNK_VERSION, TOKENIZER_VERSION, stamp])
                source = deepcopy(self.cache.get(cache_key))
                if source is None:
                    source = self.owner._pdf_source.extract(name, start_page=page, end_page=page)
                    if len(self.cache) >= 20 or sum(sum(len(b['text']) for b in v['blocks']) for v in self.cache.values()) + sum(len(b['text']) for b in source['blocks']) > 100000:
                        self.cache.clear()
                    self.cache[cache_key] = deepcopy(source)
            except Exception:
                self.attempt['failed_pages'].append(page)
                raise
            if source['fingerprint'] != version:
                self.attempt['failed_pages'].append(page)
                raise RagError('rag_source_changed')
            self.attempt['parsed_pages'].append(page)
            total = source['pdf_coverage']['total_pages']
            chunks = token_structure_chunks(source['blocks'], parser_version=source['parser_version'])
            for chunk in chunks:
                where = json.loads(chunk['location'])
                where.update(corpus_scope=scope, pdf_total=total, ei=stamp)
                chunk['location'] = json.dumps(where, ensure_ascii=False, separators=(',', ':'))
                if len(chunk['location']) > 512:
                    raise RagError('rag_capacity')
            chars += sum(len(c['text']) for c in chunks)
            documents[key] = {'fingerprint': version, 'chunks': chunks}
            if time.monotonic() - started > 60 or chars > 100000 or sum(len(d['chunks']) for d in documents.values()) + sum(len(old[k]['locations']) for k in retained) > 500:
                raise RagError('rag_capacity')
            self.owner._register_alias(key, original)
        if chars > 100000 or sum(len(d['chunks']) for d in documents.values()) + sum(len(old[k]['locations']) for k in retained) > 500:
            raise RagError('rag_capacity')
        if stamp != self.embedding_stamp():
            raise RagError('rag_source_changed')
        removals = sorted(set(old) - set(documents) - set(retained))
        if not documents and not removals:
            self.attempt = None
            raise RagError('rag_index_current')
        self.pending = {k: p for k, p in self.pending.items() if time.monotonic() - p['created'] < 300 and p['session'] != session}
        if len(self.pending) >= 4:
            raise RagError('rag_capacity')
        token = uuid.uuid4().hex
        self.pending[token] = dict(name=name, scope=scope, original=original, version=version, old=old,
                                   documents=documents, removals=removals, session=session, created=time.monotonic(),
                                   snapshot=self.owner._store.snapshot(), identity=self.owner._rag._identity(), embedding_stamp=stamp, pages=sorted(pages), attempt_id=self.attempt['id'])
        return {'plan_id': token, 'operation': 'pdf_range_collection', 'selected_pages': sorted(pages),
                'reused_pages': len(retained), 'new_pages': len(documents), 'removed_pages': len(removals),
                'removed_page_numbers': sorted({int(w['group'].split(':')[1]) for k in removals for w in old[k]['locations']}),
                'chunk_count': sum(len(d['chunks']) for d in documents.values()), 'total_pages': total,
                'generation': self.owner._store.snapshot()['revision'], 'requires_confirmation': True, 'wrote_nothing': True}

    def confirm(self, token, session, confirmed):
        p = self.pending.pop(token, None)
        denied = {'status': 'rejected', 'applied': False, 'wrote_nothing': True}
        if not p or p['session'] != session or confirmed is not True or time.monotonic() - p['created'] >= 300:
            if p and p['session'] == session and self.attempt and self.attempt['id'] == p['attempt_id']:
                self.attempt['state'] = 'expired' if time.monotonic() - p['created'] >= 300 else 'cancelled'
            return denied
        o = self.owner
        def fresh():
            return (self.bind(p['name'])[2:] == (p['version'], p['old']) and o._store.snapshot() == p['snapshot'] and o._rag._identity() == p['identity'] and self.embedding_stamp() == p['embedding_stamp'])
        try:
            if not fresh():
                if self.attempt and self.attempt['id'] == p['attempt_id']:
                    self.attempt['state'] = 'rejected'
                return denied
            chunks = [c for d in p['documents'].values() for c in d['chunks']]
            vectors = iter(o._rag._vectors(o._rag.embedder.embed([c['text'] for c in chunks]), len(chunks)) if chunks else [])
            documents = {k: {'fingerprint': d['fingerprint'], 'records': [{**c, 'embedding': next(vectors)} for c in d['chunks']]} for k, d in p['documents'].items()}
            if not fresh():
                if self.attempt and self.attempt['id'] == p['attempt_id']:
                    self.attempt['state'] = 'rejected'
                return denied
            inner = o._store.plan_sync(p['scope'], documents, p['removals'], session)
        except Exception as error:
            if self.attempt and self.attempt['id'] == p['attempt_id']:
                self.attempt['state'] = 'rejected'
            if getattr(error, 'code', None) == 'uncertain_state':
                o._rag._disabled = True
                return {'status': 'uncertain', 'applied': None, 'wrote_nothing': False}
            return denied
        try:
            result = o._store.confirm_replace(inner['plan_id'], session, True)
            if result.get('uncertain') is not False or result.get('revision') != p['snapshot']['revision'] + 1:
                raise ValueError()
            self.attempt = None
            return {**result, 'status': 'saved', 'applied': True, 'wrote_nothing': False}
        except Exception:
            o._rag._disabled = True
            return {'status': 'uncertain', 'applied': None, 'wrote_nothing': False}

    def query(self, name, question, session, top_k):
        _text(question, 200)
        if type(top_k) is not int or not 1 <= top_k <= 10:
            raise RagError('rag_invalid')
        o = self.owner
        scope, original, version, old = self.bind(name)
        stamp = self.embedding_stamp()
        if not old or any(not self.valid(row, version) for row in old.values()):
            raise RagError('rag_corpus_stale')
        o._evidence.ensure_available(session)
        snapshot, identity = o._store.snapshot(), o._rag._identity()
        vectors = o._rag._vectors(getattr(o._rag.embedder, 'embed_queries', o._rag.embedder.embed)([question]), 1)
        hits = o._store.query_scoped(vectors[0], keyword=question, sources={key: version for key in old},
                                     candidates=50, top_k=top_k, chunk_version=CHUNK_VERSION)
        if self.bind(name)[2:] != (version, old) or o._store.snapshot() != snapshot or o._rag._identity() != identity or self.embedding_stamp() != stamp:
            raise RagError('rag_source_changed')
        plan = o._evidence.plan(question, hits, session)
        self.queries = {k: v for k, v in self.queries.items() if time.monotonic() - v['created'] < 300}
        self.queries[plan['plan_id']] = {'snapshot': snapshot, 'session': session, 'created': time.monotonic()}
        return plan

    def confirm_query(self, token, session, confirmed):
        lease = self.queries.pop(token)
        if lease['session'] != session or self.owner._store.snapshot() != lease['snapshot'] or time.monotonic() - lease['created'] >= 300:
            confirmed = False
        return self.owner._evidence.confirm(token, session, confirmed)
