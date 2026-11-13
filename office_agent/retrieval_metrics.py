import math
from statistics import mean


def ranking_metrics(ranking, relevance, *, k=10):
    if type(k) is not int or k < 1 or len(set(ranking)) != len(ranking):
        raise ValueError('invalid evaluation ranking')
    if any(type(value) is not int or not 0 <= value <= 3 for value in relevance.values()):
        raise ValueError('invalid relevance grade')
    relevant = {key for key, grade in relevance.items() if grade > 0}
    if not relevant:
        return {'recall': None, 'ndcg': None, 'mrr': None, 'unanswerable_empty_result': not ranking}
    hits = [index for index, identifier in enumerate(ranking[:k], 1) if identifier in relevant]
    dcg = sum((2 ** relevance.get(identifier, 0) - 1) / math.log2(index + 1) for index, identifier in enumerate(ranking[:k], 1))
    ideal = sum((2 ** grade - 1) / math.log2(index + 1) for index, grade in enumerate(sorted(relevance.values(), reverse=True)[:k], 1))
    return {'recall': len(hits) / len(relevant), 'ndcg': dcg / ideal if ideal else 0.,
            'mrr': 1 / hits[0] if hits else 0., 'unanswerable_empty_result': None}


def aggregate(rows):
    return {name: mean(value[name] for value in rows if value[name] is not None)
            if any(value[name] is not None for value in rows) else None
            for name in ('recall', 'ndcg', 'mrr', 'unanswerable_empty_result')}
