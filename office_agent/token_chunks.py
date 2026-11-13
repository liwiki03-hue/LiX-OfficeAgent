import base64
from functools import lru_cache
import hashlib
import json
from pathlib import Path

import tiktoken

from .rag import structure_chunks, RagError, MAX_CHUNKS
from .rag_sources import PARSER_VERSION

TOKENIZER_VERSION = 'cl100k_base-tiktoken-0.14.0'
CHUNK_VERSION = 'structure-token-v3'
VOCAB_SHA256 = '223921b76ee99bde995b7ff738513eef100fb51d18c93597a113bcffe865b2a7'
PATTERN = r"'(?i:[sdmt]|ll|ve|re)|[^\r\n\p{L}\p{N}]?+\p{L}++|\p{N}{1,3}+| ?[^\s\p{L}\p{N}]++[\r\n]*+|\s++$|\s*[\r\n]|\s+(?!\S)|\s"


@lru_cache(maxsize=1)
def tokenizer():
    path = Path(__file__).with_name('assets') / 'cl100k_base.tiktoken'
    raw = path.read_bytes()
    if hashlib.sha256(raw).hexdigest() != VOCAB_SHA256:
        raise RagError('rag_tokenizer_invalid')
    ranks = {base64.b64decode(word): int(rank) for word, rank in (line.split() for line in raw.splitlines())}
    return tiktoken.Encoding(name='office-cl100k-base', pat_str=PATTERN, mergeable_ranks=ranks, special_tokens={})


def token_count(text):
    return len(tokenizer().encode_ordinary(text))


def token_structure_chunks(blocks, *, max_tokens=384, parser_version=PARSER_VERSION):
    if type(max_tokens) is not int or not 32 <= max_tokens <= 1024:
        raise RagError('rag_invalid')
    structure_chunks(blocks)
    chunks = []
    for block in blocks:
        text, cursor = block['text'], 0
        while cursor < len(text):
            low, high, end = cursor + 1, len(text), cursor
            while low <= high:
                point = (low + high) // 2
                if token_count(text[cursor:point]) <= max_tokens:
                    end = point
                    low = point + 1
                else:
                    high = point - 1
            if end == cursor:
                raise RagError('rag_capacity')
            fragment = text[cursor:end]
            if token_count(fragment) > max_tokens:
                raise RagError('rag_capacity')
            location = json.dumps({'type': block['type'], 'group': block['group'], 'block': block['location'],
                                   'char_start': cursor, 'char_end': end, 'token_count': token_count(fragment),
                                   'tokenizer': TOKENIZER_VERSION, 'chunker': CHUNK_VERSION,
                                   'parser': parser_version}, ensure_ascii=False, separators=(',', ':'))
            if len(location) > 512:
                raise RagError('rag_capacity')
            chunks.append({'source': block['source'], 'location': location, 'text': fragment, 'kind': 'document'})
            cursor = end
            if len(chunks) > MAX_CHUNKS:
                raise RagError('rag_capacity')
    return chunks
