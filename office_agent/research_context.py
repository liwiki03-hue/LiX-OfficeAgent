import json

from .rag import RagError
from .token_chunks import CHUNK_VERSION, TOKENIZER_VERSION, token_count

CONTEXT_VERSION = 'native-parent-neighbors-v1'
CONTEXT_BUDGET = 768


def _prefix(text, budget):
    low, high, end = 0, len(text), 0
    while low <= high:
        point = (low + high) // 2
        if token_count(text[:point]) <= budget:
            end = point
            low = point + 1
        else:
            high = point - 1
    return end


def _window(text, start, end, budget):
    left, right = start, end
    while left or right < len(text):
        proposed_left, proposed_right = max(0, left - 48), min(len(text), right + 48)
        if token_count(text[proposed_left:proposed_right]) > budget:
            break
        left, right = proposed_left, proposed_right
    return left, right


def expand_contexts(source, hits):
    blocks = source['blocks']
    by_location = {block['location']: index for index, block in enumerate(blocks)}
    contexts = []
    for hit in hits:
        where = json.loads(hit['location'])
        if (hit['doc_id'] != source['doc_id'] or hit['source_fingerprint'] != source['fingerprint'] or
                where.get('chunker') != CHUNK_VERSION or where.get('tokenizer') != TOKENIZER_VERSION):
            raise RagError('rag_source_changed')
        index = by_location.get(where.get('block'))
        if index is None:
            raise RagError('rag_source_changed')
        block = blocks[index]
        start, end = where.get('char_start'), where.get('char_end')
        if (type(start) is not int or type(end) is not int or not 0 <= start < end <= len(block['text']) or
                block['text'][start:end] != hit['text'] or block['source'] != hit['source']):
            raise RagError('rag_source_changed')
        sections, used = [], 0

        def add(role, candidate, left, right):
            nonlocal used
            text = candidate['text'][left:right]
            count = token_count(text)
            if not text.strip() or used + count > CONTEXT_BUDGET:
                return
            sections.append({'role': role, 'location': candidate['location'], 'char_start': left, 'char_end': right,
                             'text': text, 'tokens': count, 'truncated': left != 0 or right != len(candidate['text'])})
            used += count

        left, right = _window(block['text'], start, end, 512)
        add('parent', block, left, right)
        heading = next((candidate for candidate in reversed(blocks[:index]) if candidate['type'] == 'heading'), None)
        if heading is not None and block['type'] != 'heading':
            add('heading', heading, 0, _prefix(heading['text'], min(64, CONTEXT_BUDGET - used)))
        if block['type'] in {'table_row', 'sheet_row'}:
            first = next(candidate for candidate in blocks if candidate['group'] == block['group'] and candidate['type'] == block['type'])
            if first is not block:
                add('table_first_row', first, 0, _prefix(first['text'], min(128, CONTEXT_BUDGET - used)))
        elif block['type'] == 'paragraph':
            for position in (index - 1, index + 1):
                if 0 <= position < len(blocks):
                    neighbor = blocks[position]
                    if neighbor['group'] == block['group'] and neighbor['type'] == 'paragraph':
                        add('adjacent', neighbor, 0, _prefix(neighbor['text'], min(96, CONTEXT_BUDGET - used)))
        contexts.append({'hit_id': hit['id'], 'source': hit['source'], 'source_fingerprint': hit['source_fingerprint'],
                         'retrieved_location': hit['location'], 'context_version': CONTEXT_VERSION,
                         'tokenizer': TOKENIZER_VERSION, 'token_budget': CONTEXT_BUDGET, 'tokens': used,
                         'sections': sections, 'writable': False})
    return contexts
