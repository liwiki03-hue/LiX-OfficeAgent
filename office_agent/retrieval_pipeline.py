import re
import time

from .hybrid_ranking import fuse, IDENTIFIERS, VERSION
from .token_chunks import CHUNK_VERSION, token_count


def subqueries(question, expand):
    if not expand or re.search(r'排除|不得|不含|不要|\bnot\b|\bexclude\b', question, re.I):
        return [question]
    constraints = list(dict.fromkeys(IDENTIFIERS.findall(question.lower()) + re.findall(r'\d+(?:\.\d+)?', question)))
    parts = re.split(r'[；;？?]|比较|对比|以及|\s+vs\.?\s+|\s+and\s+', question, flags=re.I)
    result = [question]
    for part in parts:
        value = part.strip(' ，,。:：')
        if len(value) < 3:
            continue
        value = value + ' ' + ' '.join(term for term in constraints if term not in value.lower())
        value = value.strip()
        if len(value) <= 200 and value not in result:
            result.append(value)
        if len(result) == 3:
            break
    return result


def retrieve(store, rag, question, versions, *, top_k=5, rerank=False, expand=False, reranker=None, ranking='hybrid'):
    started = time.monotonic()
    queries = subqueries(question, expand)
    report = {'queries': len(queries), 'candidates': 0, 'rerank': 'not_requested', 'query_expansion': 'literal_clauses' if len(queries) > 1 else 'original',
              'query_text_is_evidence': False, 'ranking': ranking, 'elapsed_ms': 0, 'support_assessed': False,
              'lexical_version': VERSION, 'duplicates_removed': 0}
    if not versions:
        return [], report
    vectors = rag._vectors(getattr(rag.embedder, 'embed_queries', rag.embedder.embed)(queries), len(queries))
    rankings, records = [], {}
    for query, vector in zip(queries, vectors):
        hits = store.query_scoped(vector, keyword=query, sources=versions, candidates=50, top_k=50,
                                  chunk_version=CHUNK_VERSION, candidate_pool=True, ranking=ranking)
        rankings.append([hit['id'] for hit in hits])
        records.update({hit['id']: hit for hit in hits})
    pool = [records[key] for key in fuse(rankings, 50)]
    seen, unique = set(), []
    for hit in pool:
        key = (hit['doc_id'], hit['source_fingerprint'], hit['text'])
        if key not in seen:
            unique.append(hit)
            seen.add(key)
    report['duplicates_removed'] = len(pool) - len(unique)
    pool = unique
    report['candidates'] = len(pool)
    if rerank:
        report['rerank'] = 'not_configured'
        if reranker is not None and pool:
            chosen, remaining, tokens = [], [], 0
            for hit in pool[:20]:
                cost = token_count(hit['text']) + token_count(question)
                if tokens + cost <= 8192:
                    chosen.append(hit)
                    tokens += cost
                else:
                    remaining.append(hit)
            try:
                ordered, status = reranker.rank(question, chosen)
            except Exception:
                ordered, status = chosen, 'unavailable'
            if status == 'cleanup_failed':
                from .rag import RagError
                rag._disabled = True
                raise RagError('rag_uncertain')
            if status == 'completed':
                if type(ordered) is list and sorted(hit['id'] for hit in ordered) == sorted(hit['id'] for hit in chosen):
                    by_id = {hit['id']: hit for hit in chosen}
                    pool = [by_id[hit['id']] for hit in ordered] + remaining + pool[20:]
                else:
                    status = 'invalid_response'
            report['rerank'] = status
    report['elapsed_ms'] = round((time.monotonic() - started) * 1000)
    return [dict(hit, rank=rank) for rank, hit in enumerate(pool[:top_k], 1)], report
