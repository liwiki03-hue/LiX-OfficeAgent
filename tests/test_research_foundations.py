import json
import math

import pytest

from office_agent.rag import RagError
from office_agent.retrieval_metrics import aggregate, ranking_metrics
from office_agent.token_chunks import tokenizer, token_count, token_structure_chunks


def block(text):
    return {'source': 'fixture.docx', 'location': 'body:2/paragraph', 'group': 'body', 'type': 'paragraph', 'text': text}


def test_real_tokenizer_offline_unicode_and_reconstruction(monkeypatch):
    import requests
    monkeypatch.setattr(requests, 'get', lambda *args, **kwargs: pytest.fail('Tokenizer must never download at runtime'))
    text = ('实验结果 β-value Δt=20ms. The measured result is repeatable. ' * 70).strip()
    chunks = token_structure_chunks([block(text)], max_tokens=64)
    assert ''.join(item['text'] for item in chunks) == text
    assert len(chunks) > 1
    assert all(token_count(item['text']) <= 64 for item in chunks)
    assert any(len(item['text']) != token_count(item['text']) for item in chunks)
    previous = 0
    for item in chunks:
        where = json.loads(item['location'])
        assert where['char_start'] == previous
        assert text[where['char_start']:where['char_end']] == item['text']
        assert where['tokenizer'] == 'cl100k_base-tiktoken-0.14.0'
        previous = where['char_end']
    assert previous == len(text)


def test_tokenizer_vocab_digest_is_enforced(monkeypatch):
    from pathlib import Path
    original = Path.read_bytes
    tokenizer.cache_clear()
    monkeypatch.setattr(Path, 'read_bytes', lambda path: b'corrupt' if path.name == 'cl100k_base.tiktoken' else original(path))
    try:
        with pytest.raises(RagError, match='tokenizer_invalid'):
            tokenizer()
    finally:
        tokenizer.cache_clear()


def test_token_chunks_explicit_bounds_and_sensitive_rejection():
    for budget in (0, 31, 1025, True):
        with pytest.raises(RagError):
            token_structure_chunks([block('safe text')], max_tokens=budget)
    with pytest.raises(RagError):
        token_structure_chunks([block('甲' * 8001)])
    with pytest.raises(RagError):
        token_structure_chunks([block('API_KEY=fixture-secret-value')])


def test_metrics_known_grades_missing_and_unanswerable():
    value = ranking_metrics(['b', 'x', 'a'], {'a': 2, 'b': 1}, k=3)
    assert value['recall'] == value['mrr'] == 1
    expected = (1 + 3 / math.log2(4)) / (3 + 1 / math.log2(3))
    assert value['ndcg'] == pytest.approx(expected)
    missing = ranking_metrics(['x'], {'a': 1})
    assert missing['recall'] == missing['mrr'] == missing['ndcg'] == 0
    empty = ranking_metrics([], {})
    assert empty['recall'] is None and empty['unanswerable_empty_result'] is True
    assert aggregate([value, empty])['recall'] == 1
    assert aggregate([value, empty])['unanswerable_empty_result'] == 1
    with pytest.raises(ValueError):
        ranking_metrics(['a', 'a'], {'a': 1})


def test_opt_in_catalogue_and_dispatch_remain_preview_only(tmp_path):
    from docx import Document
    from office_agent.app import build_tools, professional_tools, select_tools, split_read_write, build_dispatcher
    from office_agent.agent import _check_tool, _read_context
    from office_agent.writes import Editor
    path = tmp_path / 'fixture.docx'
    doc = Document()
    doc.add_paragraph('The sample has 30 observations [1].')
    doc.save(path)
    original = path.read_bytes()
    base = build_tools((path,), ocr_enabled=True)
    assert len(base) == 48
    extras = professional_tools((path,))
    for spec in extras:
        _check_tool(spec)
    read, write = split_read_write(select_tools(base + extras, 'word'))
    assert 'read_docx_evidence' in {spec.name for spec in read}
    assert 'plan_constrained_docx' in {spec.name for spec in write}
    assert 'plan_constrained_docx' not in {spec.name for spec in read}
    editor = Editor(tmp_path)
    dispatch = build_dispatcher(editor, path)
    result = dispatch('read_docx_evidence', {'path': str(path)})
    context = json.loads(_read_context('read_docx_evidence', result, {'offset': 0}))
    item = context['items'][0]
    assert item['text'] == doc.paragraphs[0].text
    dispatch('plan_constrained_docx', {'path': str(path), 'source_sha256': result['document']['source_sha256'],
        'changes': [{'locator_id': item['locator_id'], 'expected_text_sha256': item['text_sha256'],
                     'new_text': 'The sample includes 30 observations [1].', 'reason': 'Clarify wording'}]})
    assert editor.plans and not editor.audit.exists() and path.read_bytes() == original


def test_partial_model_context_does_not_expose_writable_locator(tmp_path):
    from docx import Document
    from office_agent.professional_edits import snapshot
    from office_agent.agent import _read_context
    from office_agent.writes import Editor
    path = tmp_path / 'long.docx'
    doc = Document()
    doc.add_paragraph('Long unabridged original text. ' * 200)
    doc.save(path)
    full = snapshot(Editor(tmp_path), path)
    slim = json.loads(_read_context('read_docx_evidence', full, {'offset': 0}))
    assert full['items'][0]['writable']
    assert slim['items'][0]['text_truncated']
    assert not slim['items'][0]['writable']
    assert 'locator_id' not in slim['items'][0]
