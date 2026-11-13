from collections import Counter
import math
import re
import unicodedata

VERSION = 'identifier-bm25-cjk-ngram-rrf-v2'
WORDS = re.compile(r'[a-z0-9]+(?:[._/+:-][a-z0-9]+)*|[\u3400-\u9fff]+')
IDENTIFIERS = re.compile(r'10\.\d{4,9}/[^\s<>]+|[a-z]+[-_]\d+(?:[-_.]\w+)*|\d+(?:\.\d+)?\s?(?:%|ms|kg|mpa|mhz)\b')


def search_text(text):
    value = unicodedata.normalize('NFKC', text)
    value = re.sub(r'(?<=[A-Za-z])-\s*\r?\n\s*(?=[A-Za-z])', '', value)
    return re.sub(r'(?<=[\u3400-\u9fff])\s*\r?\n\s*(?=[\u3400-\u9fff])', '', value)


def terms(text):
    result = []
    for word in WORDS.findall(search_text(text).casefold()):
        if '\u3400' <= word[0] <= '\u9fff':
            result.extend(word)
            result.extend(word[i:i + 2] for i in range(len(word) - 1))
        else:
            result.append(word)
    return result


def rankings(records, question, candidates, *, term_counts=None):
    query = set(terms(question))
    counts = term_counts if term_counts is not None else {row['id']: Counter(terms(row['text'])) for row in records}
    lengths = {key: sum(value.values()) for key, value in counts.items()}
    average = sum(lengths.values()) / max(1, len(lengths)) or 1
    frequencies = Counter(term for value in counts.values() for term in query if term in value)
    lexical = {}
    for key, count in counts.items():
        score = 0
        for term in query:
            frequency = count[term]
            if not frequency:
                continue
            weight = math.log(1 + (len(records) - frequencies[term] + .5) / (frequencies[term] + .5))
            score += weight * frequency * 2.2 / (frequency + 1.2 * (.25 + .75 * lengths[key] / average))
        if score > 0:
            lexical[key] = score
    identifiers = IDENTIFIERS.findall(unicodedata.normalize('NFKC', question).casefold())
    exact = {}
    for row in records:
        content = unicodedata.normalize('NFKC', row['text']).casefold()
        matches = sum(bool(re.search(r'(?<!\w)' + re.escape(value) + r'(?!\w)', content)) for value in identifiers)
        if matches:
            exact[row['id']] = matches
    order = lambda scores: sorted(scores, key=lambda key: (-scores[key], key))[:candidates]
    dense = sorted(records, key=lambda row: (row['distance'], row['id']))[:candidates]
    return [order(exact), order(lexical), [row['id'] for row in dense]]


def fuse(rankings, top_k):
    scores = {}
    for ranking in rankings:
        for position, key in enumerate(ranking, 1):
            scores[key] = scores.get(key, 0) + 1 / (60 + position)
    return sorted(scores, key=lambda key: (-scores[key], key))[:top_k]
