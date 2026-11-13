from __future__ import annotations

from collections import Counter, defaultdict
import hashlib
import itertools
import math
from pathlib import Path
import re

from .documents import inspect_file
from .rag import structure_chunks
from .rag_sources import MAX_SOURCE_BYTES, SourceError, SourceExtractor
from .writes import Workspace


def _tokens(text):
    words = re.findall(r'[a-z0-9_]+|[\u3400-\u9fff]', text.lower())
    return words + [left + right for left, right in zip(words, words[1:])
                    if len(left) == len(right) == 1 and '\u3400' <= left <= '\u9fff' and '\u3400' <= right <= '\u9fff']


class DirectoryIndex:
    def __init__(self, workspace: Workspace, selected: Path, *, embedder=None, extractor=None, inspection_cache=None):
        self.space = workspace
        self.selected = workspace.file(str(selected))
        self.anchor_sha256 = self._anchor_version()
        self.root = self.selected.parent
        self.extractor = extractor if extractor is not None else SourceExtractor(Workspace(self.root))
        self.inspection_cache = inspection_cache
        self.embedder = embedder
        self.identity = self._identity()
        self.chunks, self.documents, self.inventory, self.vectors = [], [], [], []
        self.postings = defaultdict(set)
        self.built = False

    def _identity(self):
        if self.embedder is None:
            return None
        identity = self.embedder.identity
        if (type(identity) is not dict or set(identity) != {'channel', 'model', 'dimension'}
                or identity['channel'] not in {'local', 'api'} or type(identity['model']) is not str
                or not 1 <= len(identity['model']) <= 200 or type(identity['dimension']) is not int
                or not 1 <= identity['dimension'] <= 4096):
            raise SourceError('invalid')
        return dict(identity)

    def _anchor_version(self):
        if self.selected.stat().st_size > MAX_SOURCE_BYTES:
            raise SourceError('capacity')
        with self.selected.open('rb') as stream:
            data = stream.read(MAX_SOURCE_BYTES + 1)
        if len(data) > MAX_SOURCE_BYTES:
            raise SourceError('capacity')
        return hashlib.sha256(data).hexdigest()

    def _vector(self, vector):
        if (type(vector) is not list or len(vector) != self.identity['dimension']
                or any(type(value) not in (int, float) or not math.isfinite(value) or abs(value) > 1e6 for value in vector)):
            raise SourceError('invalid')
        norm = math.sqrt(sum(value * value for value in vector))
        if not norm:
            raise SourceError('invalid')
        return [value / norm for value in vector]

    def _check(self):
        if self._identity() != self.identity:
            raise SourceError('source_changed')
        if self.space.file(str(self.selected)) != self.selected:
            raise SourceError('source_changed')
        if self._anchor_version() != self.anchor_sha256:
            raise SourceError('source_changed')
        for doc in self.documents:
            if self.space.file(str(self.root / doc['name'])) != self.root / doc['name']:
                raise SourceError('source_changed')
            if self.extractor.fingerprint(doc['doc_id']) != doc['fingerprint']:
                raise SourceError('source_changed')

    def build(self):
        self.built = False
        self.chunks, self.documents, self.inventory, self.vectors = [], [], [], []
        self.postings.clear()
        names = list(itertools.islice(self.root.iterdir(), 257))
        if len(names) > 256:
            raise SourceError('capacity')
        total = 0
        for path in sorted(names, key=lambda item: (item.name.casefold(), item.name)):
            item = {'source': path.name, 'status': 'skipped'}
            self.inventory.append(item)
            try:
                checked = self.space.file(str(path))
                if checked.parent != self.root or checked != path:
                    raise SourceError('invalid')
                if path.suffix.lower() not in {'.docx', '.xlsx'}:
                    item['reason'] = 'legacy_format' if path.suffix.lower() in {'.doc', '.xls'} else 'unsupported_format'
                    continue
                if path.name.startswith('~$') or re.search(r'\.agent-[a-f0-9]{8}\.', path.name, re.I):
                    item['reason'] = 'generated_or_lock_file'
                    continue
                if len(self.documents) >= 20:
                    item['reason'] = 'document_limit'
                    continue
                source = None
                if self.inspection_cache is not None:
                    source = self.extractor.extract(path.name)
                    cached = self.inspection_cache.get(path.name)
                    report = cached[1] if cached is not None and cached[0] == source['fingerprint'] else inspect_file(checked)
                    self.inspection_cache[path.name] = (source['fingerprint'], report)
                else:
                    report = inspect_file(checked)
                if not report['readable'] or report['risks'] or report['macros'] or report['external_links']:
                    item['reason'] = 'risky_document'
                    continue
                source = source if source is not None else self.extractor.extract(path.name)
                chunks = structure_chunks(source['blocks'])
                size = sum(len(chunk['text']) for chunk in chunks)
                if len(self.chunks) + len(chunks) > 500 or total + size > 1_000_000:
                    item['reason'] = 'index_capacity'
                    continue
                total += size
                for chunk in chunks:
                    self.chunks.append({**chunk, 'doc_id': source['doc_id'], 'fingerprint': source['fingerprint']})
                    for token in set(_tokens(chunk['text'])):
                        self.postings[token].add(len(self.chunks) - 1)
                self.documents.append({'name': path.name, 'doc_id': source['doc_id'], 'fingerprint': source['fingerprint']})
                item.update(status='indexed_in_memory', chunks=len(chunks), reason=None)
            except Exception as error:
                item['reason'] = error.code if isinstance(error, SourceError) else 'path_or_parse_rejected'
        if self.embedder is not None and self.chunks:
            raw = self.embedder.embed([chunk['text'] for chunk in self.chunks])
            if type(raw) is not list or len(raw) != len(self.chunks):
                raise SourceError('invalid')
            self.vectors = [self._vector(vector) for vector in raw]
        self._check()
        self.built = True
        return self

    def query(self, query, *, top_k=5, _full=False):
        if type(query) is not str or not query.strip() or len(query) > 200 or type(top_k) is not int or not 1 <= top_k <= 10:
            raise SourceError('invalid')
        if not self.built:
            raise SourceError('invalid')
        if type(_full) is not bool or (_full and self.identity is None):
            raise SourceError('invalid')
        self._check()
        weights = Counter(_tokens(query))
        candidates = set().union(*(self.postings.get(token, set()) for token in weights))
        lexical = {}
        for index in candidates:
            terms = Counter(_tokens(self.chunks[index]['text']))
            lexical[index] = sum(weight * math.log1p(len(self.chunks) / len(self.postings[token])) * terms[token]
                                 for token, weight in weights.items() if terms[token])
        keyword_order = sorted(lexical, key=lambda index: (-lexical[index], index))
        scores = defaultdict(float)
        for rank, index in enumerate(keyword_order, 1):
            scores[index] += 1 / (60 + rank)
        semantic = None
        if self.vectors:
            raw = getattr(self.embedder, 'embed_queries', self.embedder.embed)([query])
            if type(raw) is not list or len(raw) != 1:
                raise SourceError('invalid')
            vector = self._vector(raw[0])
            semantic = [sum(left * right for left, right in zip(vector, value)) for value in self.vectors]
            for rank, index in enumerate(sorted(range(len(semantic)), key=lambda index: (-semantic[index], index))[:50], 1):
                scores[index] += 1 / (60 + rank)
        ordered = sorted(scores, key=lambda index: (-scores[index], self.chunks[index]['source'], self.chunks[index]['location']))
        hits = [{key: self.chunks[index][key] for key in ('source', 'location')} |
                {'excerpt': self.chunks[index]['text'][:1000], 'text_truncated': len(self.chunks[index]['text']) > 1000,
                 'source_sha256': self.chunks[index]['fingerprint'], 'score': scores[index]}
                for index in ordered[:top_k]]
        if _full is True:
            hits = [{'id': index + 1, 'doc_id': self.chunks[index]['doc_id'],
                     'source': self.chunks[index]['source'], 'location': self.chunks[index]['location'],
                     'text': self.chunks[index]['text'], 'excerpt': self.chunks[index]['text'][:200],
                     'source_fingerprint': self.chunks[index]['fingerprint'], 'model': self.identity['model'],
                     'distance': max(0., min(2., 1 - semantic[index])), 'rank': rank}
                    for rank, index in enumerate(ordered[:top_k], 1)]
        self._check()
        skipped = [item for item in self.inventory if Path(item['source']).suffix.lower() in {'.docx', '.xlsx'}
                   and item['status'] != 'indexed_in_memory' and item.get('reason') != 'generated_or_lock_file']
        return {'hits': hits, 'candidate_count': len(ordered), 'truncated': len(ordered) > top_k,
                'inventory': self.inventory, 'document_count': len(self.documents), 'read_only': True,
                'corpus_truncated': any(item.get('reason') in {'document_limit', 'index_capacity', 'capacity'} for item in self.inventory),
                'coverage_complete': not skipped, 'skipped_document_count': len(skipped),
                'index_persisted': False, 'semantic_available': bool(self.vectors),
                'retrieval': 'inverted_index_and_vector_rrf' if self.vectors else 'inverted_index',
                'model_identity': self.identity, 'requires_rebuild_on_model_change': True}
