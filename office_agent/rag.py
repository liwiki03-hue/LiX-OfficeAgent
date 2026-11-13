"""Structural chunking and approval-before-embedding orchestration.

No file parsing, provider credentials or model tool registration here. Trusted
host supplies already-bounded document blocks and a live source fingerprint.
The independent VectorStore implements write backup/reopen/audit/rollback.
"""
import json
import math
import re
import time
import uuid

MAX_BLOCKS = 4000
MAX_CHUNKS = 500
MAX_TEXT = 8000
MAX_TOTAL_TEXT = 1_000_000
TARGET_TEXT = 1600
SECRET = re.compile(r'(?i)(?:api[_ -]?key|access[_ -]?token|password|secret)\s*[:=]\s*\S+|\bsk-[a-z0-9_-]{10,}|\bbearer\s+[a-z0-9._-]{10,}')
KINDS = frozenset({'paragraph', 'heading', 'table_row', 'sheet_row', 'session'})


class RagError(ValueError):
    def __init__(self, code='rag_invalid'):
        self.code = code
        super().__init__(code)


def _text(value, limit):
    if type(value) is not str or not value.strip() or len(value) > limit or '\x00' in value:
        raise RagError()
    if SECRET.search(value):
        raise RagError('rag_sensitive_content')
    return value


def structure_chunks(blocks):
    """Keep rows/paragraphs whole, pack adjacent same-group blocks; reject oversize.

    Every block must provide source, location, type, text and structural group.
    Tables/sheets use distinct groups, headings are standalone. No silent skips.
    """
    if type(blocks) is not list or not blocks or len(blocks) > MAX_BLOCKS:
        raise RagError('rag_capacity')
    chunks, texts, locations, current, total = [], [], [], None, 0
    def flush():
        nonlocal texts, locations
        if texts:
            chunks.append({'source': current[0], 'location': json.dumps({'type':current[1], 'group':current[2], 'blocks':locations}, ensure_ascii=False, separators=(',',':')),
                           'text':'\n'.join(texts), 'kind':'document'})
            texts, locations = [], []
            if len(chunks) > MAX_CHUNKS:
                raise RagError('rag_capacity')
    for block in blocks:
        if type(block) is not dict or set(block) != {'source','location','type','text','group'}:
            raise RagError()
        source = _text(block['source'], 128)
        if source.replace('\\','/').split('/')[-1].lower().startswith('.env'):
            raise RagError('rag_sensitive_source')
        kind = block['type']
        if type(kind) is not str or kind not in KINDS:
            raise RagError()
        group = _text(block['group'], 64)
        location = _text(block['location'], 128)
        text = _text(block['text'], MAX_TEXT)
        total += len(text)
        if total > MAX_TOTAL_TEXT:
            raise RagError('rag_capacity')
        key = (source, kind, group)
        proposed_location = json.dumps({'type':kind,'group':group,'blocks':[*locations,location]}, ensure_ascii=False, separators=(',',':'))
        if texts and (key != current or kind == 'heading' or
                      sum(map(len,texts))+len(texts)+len(text) > TARGET_TEXT or len(proposed_location) > 512):
            flush()
        current = key
        texts.append(text)
        locations.append(location)
        if kind == 'heading' or len(text) >= TARGET_TEXT:
            flush()
    flush()
    return chunks


class RagService:
    """In-memory bounded plans; no provider work until independent literal approval.

    store: snapshot/plan_replace/confirm_replace/query/invalidate protocol.
    embedder: identity {channel,model,dimension}, embed(list[str]) protocol.
    source_version(doc_id): fresh deterministic host fingerprint, no client path.
    """
    def __init__(self, store, embedder, source_version, *, store_error_type=None):
        if store_error_type is not None and (not isinstance(store_error_type,type) or not issubclass(store_error_type,Exception)):
            raise RagError()
        self._store_error_type = store_error_type
        self.store, self.embedder, self.source_version = store, embedder, source_version
        self._plans = {}
        self._ids = set()
        self._disabled = False

    def invalidate(self):
        self._plans.clear()
        self.store.invalidate()

    def _identity(self):
        identity = self.embedder.identity
        if (type(identity) is not dict or set(identity) != {'channel','model','dimension'} or
                identity['channel'] not in {'local','api'} or type(identity['dimension']) is not int or
                not 1 <= identity['dimension'] <= 4096):
            raise RagError()
        _text(identity['model'], 200)
        return dict(identity)

    def _snapshot(self):
        snapshot = self.store.snapshot()
        if type(snapshot) is not dict:
            raise RagError()
        encoded = json.dumps(snapshot, allow_nan=False)
        if len(encoded) > 4096:
            raise RagError('rag_capacity')
        return json.loads(encoded)

    def _compatible(self, snapshot, identity):
        if snapshot.get('model') != identity['model'] or snapshot.get('dimension') != identity['dimension']:
            raise RagError('rag_model_mismatch')
        if snapshot.get('uncertain') is True:
            self._disabled = True
            raise RagError('rag_uncertain')

    def _vectors(self, vectors, count):
        dimension = self._identity()['dimension']
        if type(vectors) is not list or len(vectors) != count:
            raise RagError('rag_embedding_failed')
        for vector in vectors:
            if (type(vector) is not list or len(vector) != dimension or
                    any(type(value) not in (int,float) or not math.isfinite(value) or abs(value) > 1e6 for value in vector) or
                    not any(value != 0 for value in vector)):
                raise RagError('rag_embedding_failed')
        return vectors

    def plan_index(self, doc_id, blocks, session_id, *, chunker=structure_chunks):
        if self._disabled:
            raise RagError('rag_uncertain')
        _text(doc_id,128); _text(session_id,128)
        now = time.monotonic()
        self._plans = {key:value for key,value in self._plans.items() if now-value['created'] <= 300}
        if len(self._plans) >= 8 or len(self._ids) >= 256:
            raise RagError('rag_capacity')
        chunks = chunker(blocks)
        fingerprint = self.source_version(doc_id)
        if type(fingerprint) is not str or not re.fullmatch('[a-f0-9]{64}', fingerprint):
            raise RagError('rag_source_changed')
        identity = self._identity()
        snapshot = self._snapshot()
        self._compatible(snapshot,identity)
        identifier = uuid.uuid4().hex
        if identifier in self._ids:
            raise RagError('rag_id_collision')
        self._ids.add(identifier)
        self._plans[identifier] = {'created':now,'session':session_id,'chunks':chunks,'doc_id':doc_id,
                                   'fingerprint':fingerprint,'identity':identity,'snapshot':snapshot}
        return {'plan_id':identifier,'doc_id':doc_id,'chunk_count':len(chunks),'source_fingerprint':fingerprint,
                'embedding':identity,'provider_call_after_confirmation':True,'requires_confirmation':True}

    def confirm_index(self, plan_id, session_id, confirmed=False):
        denied = {'status':'rejected','applied':False,'wrote_nothing':True}
        if self._disabled:
            return {'status':'uncertain','applied':None,'wrote_nothing':False}
        if type(plan_id) is not str or type(session_id) is not str:
            return denied
        plan = self._plans.get(plan_id)
        if plan is None or plan['session'] != session_id:
            return denied
        self._plans.pop(plan_id)
        if confirmed is not True or time.monotonic()-plan['created'] > 300:
            return denied
        try:
            if (self.source_version(plan['doc_id']) != plan['fingerprint'] or
                    self._identity() != plan['identity'] or self._snapshot() != plan['snapshot']):
                return denied
            # Confirmed API channel may disclose document text to configured provider.
            vectors = self._vectors(self.embedder.embed([chunk['text'] for chunk in plan['chunks']]),len(plan['chunks']))
            if (self.source_version(plan['doc_id']) != plan['fingerprint'] or
                    self._identity() != plan['identity'] or self._snapshot() != plan['snapshot']):
                return denied
            records = [{**chunk,'embedding':vector} for chunk,vector in zip(plan['chunks'],vectors)]
            inner = self.store.plan_replace(plan['doc_id'],plan['fingerprint'],records,session_id)
        except Exception:
            # No write method invoked before this point; do not echo text/errors.
            return {**denied,'reason':'rag_prewrite_failed'}
        try:
            result = self.store.confirm_replace(inner['plan_id'],session_id,True)
            if type(result) is not dict:
                raise TypeError()
            # VectorStore returns verified metadata rather than transport status.
            if 'status' not in result:
                if (result.get('model') == plan['identity']['model'] and result.get('dimension') == plan['identity']['dimension'] and
                        type(result.get('revision')) is int and result['revision'] == plan['snapshot']['revision']+1 and
                        type(result.get('chunk_count')) is int and 0 <= result['chunk_count'] <= 10000 and
                        type(result.get('audit_count')) is int and result['audit_count'] >= 1 and
                        result.get('uncertain') is False and type(result.get('save_id')) is str and
                        re.fullmatch('[a-f0-9]{32}',result['save_id'])):
                    return {**result,'status':'saved','applied':True,'wrote_nothing':False}
                raise ValueError()
            if result.get('status') == 'uncertain' or result.get('applied') is None:
                self._disabled = True
                return {**result,'status':'uncertain','applied':None,'wrote_nothing':False}
            return result
        except Exception as error:
            self._disabled = True
            result = {'status':'uncertain','applied':None,'wrote_nothing':False,'reason':'rag_uncertain'}
            if self._store_error_type is not None and type(error) is self._store_error_type and getattr(error,'code',None) == 'uncertain_state':
                save_id = getattr(error,'save_id',None)
                if type(save_id) is str and re.fullmatch('[a-f0-9]{32}',save_id):
                    result['save_id'] = save_id
            return result

    def query(self, question, *, top_k=5):
        if self._disabled:
            raise RagError('rag_uncertain')
        _text(question,200)
        if type(top_k) is not int or not 1 <= top_k <= 10:
            raise RagError()
        identity = self._identity()
        self._compatible(self._snapshot(),identity)
        try:
            vectors = self._vectors(getattr(self.embedder, 'embed_queries', self.embedder.embed)([question]),1)
            results = self.store.query(vectors[0],keyword=question,top_k=top_k,kind='document')
            if type(results) is not list or len(results) > top_k:
                raise ValueError()
            checked = []
            for item in results:
                if (type(item) is not dict or type(item.get('text')) is not str or len(item['text']) > MAX_TEXT or
                        self.source_version(item['doc_id']) != item['source_fingerprint']):
                    raise ValueError()
                checked.append({**item,'excerpt':item['text'][:200]})
            return checked  # Host evidence persistence must precede model excerpts.
        except Exception:
            raise RagError('rag_query_failed') from None
