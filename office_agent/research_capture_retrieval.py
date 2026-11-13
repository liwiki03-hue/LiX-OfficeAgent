from copy import deepcopy
import math

from .hybrid_ranking import rankings, fuse
from .token_chunks import token_count


def material_pool(evidence, queries, limit, forced=()):
    values = list(evidence.values())
    records = [{'id': index, 'text': row['original_text'], 'distance': 0.} for index, row in enumerate(values)]
    orders = [order for query in queries for order in rankings(records, query, len(records))[:2]]
    ordered = list(dict.fromkeys(fuse(orders, len(records)) + list(range(len(records)))))
    groups = {}
    for index in ordered:
        row = values[index]
        group = row.get('diversity_key') or row.get('capture_id') or (row.get('title'), row.get('source_sha256'))
        groups.setdefault(group, []).append(index)
    selected = [index for index, row in enumerate(values) if row['id'] in forced][:limit]
    cursor = 0
    while len(selected) < min(limit, len(values)):
        added = False
        for group in groups.values():
            if cursor < len(group) and group[cursor] not in selected:
                selected.append(group[cursor])
                added = True
                if len(selected) == limit:
                    break
        cursor += 1
        if not added and not any(cursor < len(group) for group in groups.values()):
            break
    return {values[index]['id']: values[index] for index in selected}


def configured_models(settings_root):
    from .settings_store import SettingsStore
    from .embeddings import Embedder, EmbeddingConfig
    from .research_reranker import OfflineReranker
    settings = SettingsStore(settings_root)
    values = settings._values(settings._read(settings.root / '.env'))
    config = EmbeddingConfig.from_values(values)
    if config.channel != 'local':
        return None
    embedder = Embedder(config)
    embedder.identity
    ranker = OfflineReranker(values['OFFICE_RERANKER_MODEL'], float(values.get('OFFICE_RERANKER_TIMEOUT') or 8)) if values.get('OFFICE_RERANKER_MODEL') else None
    return embedder, ranker


def retrieve_captures(evidence, question, queries, read_ids, models=None, check=lambda: None):
    check()
    original = list(evidence.values())
    original_positions = {row['id']: index for index, row in enumerate(original)}
    neighboring = {}
    for key in read_ids:
        if key not in original_positions:
            continue
        index = original_positions[key]
        row = original[index]
        neighboring[key] = [original[position] for position in (index - 1, index + 1) if row.get('capture_id') and
                            0 <= position < len(original) and original[position].get('capture_id') == row.get('capture_id') and original[position].get('page') == row.get('page')]
    forced = list(read_ids) + [row['id'] for values in neighboring.values() for row in values]
    values = list(material_pool(evidence, queries, 200, forced).values())
    records = [{'id': index, 'text': row['original_text'], 'distance': 0.} for index, row in enumerate(values)]
    orders = [order for query in queries for order in rankings(records, query, 50)[:2]]
    pool_ids = list(dict.fromkeys(fuse(orders, 50) + [row['id'] for row in records]))[:50]
    pool = records
    report = {'kind': 'ephemeral_selected_captures', 'published_index': False, 'candidate_limit': 50, 'rerank_limit': 20,
              'queries': queries, 'embedding_identity': None, 'reranker_identity': None, 'rerank': 'not_configured', 'context_read_ids': read_ids,
              'candidate_pool_truncated': len(evidence) > 200, 'available_fragments': len(evidence), 'material_limit': 200,
              'vectorized_evidence_ids': [], 'quality': 'lexical_only_no_vector_model', 'token_budget': 8192}
    if models:
        embedder, reranker = models
        try:
            identity = embedder.identity
            documents = embedder.embed([row['text'] for row in pool])
            check()
            questions = embedder.embed_queries(queries)
            check()
            if identity != embedder.identity:
                raise ValueError('changed_model')
            dense_orders = []
            for query, vector in zip(queries, questions):
                for row, document in zip(pool, documents):
                    numerator = sum(left * right for left, right in zip(vector, document))
                    denominator = math.sqrt(sum(value * value for value in vector) * sum(value * value for value in document))
                    row['distance'] = 1 - numerator / denominator
                dense_orders.extend(rankings(pool, query, 50))
            orders = dense_orders
            report.update(embedding_identity=identity, quality='real_local_dense_lexical_rrf', vectorized_evidence_ids=[row['id'] for row in values])
        except Exception:
            if getattr(embedder, '_disabled', False):
                raise ValueError('research_retrieval_cleanup_failed') from None
            report['quality'] = 'local_embedding_failed_lexical_fallback'
        selected = [records[index] for index in fuse(orders, 20)]
        if reranker and selected:
            check()
            budget, bounded = 0, []
            for row in selected:
                cost = token_count(row['text']) + token_count(question)
                if budget + cost <= 8192:
                    bounded.append(row)
                    budget += cost
            try:
                ranking_identity = reranker.identity
            except Exception:
                ranking_identity = None
            ranked, status = reranker.rank(question, bounded)
            if status == 'cleanup_failed':
                raise ValueError('research_retrieval_cleanup_failed')
            check()
            report['rerank'] = status
            if status == 'completed':
                if ranking_identity is None or reranker.identity != ranking_identity:
                    raise ValueError('reranker_changed')
                report['reranker_identity'] = ranking_identity
                selected = ranked + [row for row in selected if row['id'] not in {item['id'] for item in bounded}]
    else:
        selected = [records[index] for index in list(dict.fromkeys(fuse(orders, 16) + pool_ids))[:16]]
    requested = [index for index, row in enumerate(values) if row['id'] in read_ids]
    positions = {row['id']: index for index, row in enumerate(values)}
    neighbors = [positions[row['id']] for rows in neighboring.values() for row in rows if row['id'] in positions]
    diverse = any(row.get('diversity_key') for row in values)
    ordered = list(dict.fromkeys(requested + neighbors + [row['id'] for row in selected] + (pool_ids if diverse else [])))[:50 if diverse else 16]
    allowed = {values[index]['id'] for index in ordered}
    result, used, source_counts = {}, 0, {}
    report['diversity_policy'] = 'known_paper_max_3_except_explicit_context' if diverse else 'capture_balancing'
    explicit = set(forced)
    for index in ordered:
        if len(result) >= 16:
            break
        row = deepcopy(values[index])
        source = row.get('diversity_key')
        if source and source_counts.get(source, 0) >= 3 and row['id'] not in explicit:
            continue
        cost = len(row['original_text'])
        if used + cost > 10000:
            continue
        if row['id'] in read_ids and row.get('capture_id'):
            nearby = [value for value in neighboring.get(row['id'], []) if value['id'] in allowed]
            bounded, tokens = [], 0
            for value in nearby:
                count = token_count(value['original_text'])
                if tokens + count <= 768:
                    bounded.append({'evidence_id': value['id'], 'original_text': value['original_text'], 'location': value['location']})
                    tokens += count
            row['context'] = bounded
        result[row['id']] = row
        if source:
            source_counts[source] = source_counts.get(source, 0) + 1
        used += cost
    for row in result.values():
        if type(row.get('context')) is list:
            row['context'] = [value for value in row['context'] if 'evidence_id' not in value or value['evidence_id'] in result]
    check()
    return result, report
