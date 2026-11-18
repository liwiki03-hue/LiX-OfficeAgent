import pytest

from office_agent.visible_answer import LeadingThoughtFilter, VisibleAnswerError, visible_answer


@pytest.mark.parametrize('text,expected', [
    ('<think>private thought</think>Answer', 'Answer'),
    (' \n<think>private thought</think>\nAnswer', '\nAnswer'),
    ('<think>first</think>  <think>second</think>Answer', 'Answer'),
    ('Answer contains <think>literal</think>.', 'Answer contains <think>literal</think>.'),
    ('```xml\n<think>literal</think>\n```', '```xml\n<think>literal</think>\n```'),
    (' \nPlain answer', ' \nPlain answer'),
    ('<thinking>literal</thinking>', '<thinking>literal</thinking>'),
    ('', ''),
])
def test_only_complete_leading_thought_blocks_are_removed(text, expected):
    assert visible_answer(text) == expected


def test_all_chunk_boundaries_hide_prefix_and_closing_tag():
    text = ' \n<think>private thought</think>\nAnswer 120 ms'
    for index in range(len(text) + 1):
        stream = LeadingThoughtFilter()
        pieces = [stream.feed(text[:index]), stream.feed(text[index:]), stream.finish()]
        assert ''.join(pieces) == '\nAnswer 120 ms'
        assert 'private' not in ''.join(pieces)
    stream = LeadingThoughtFilter()
    pieces = [stream.feed(character) for character in text]
    pieces.append(stream.finish())
    assert ''.join(pieces) == '\nAnswer 120 ms'
    assert all(character not in ''.join(pieces) for character in ('<', '>'))


@pytest.mark.parametrize('text', ['<think>private thought', '  <think>private</thin', '<thi'])
def test_unclosed_and_ambiguous_leading_blocks_fail_without_output(text):
    stream = LeadingThoughtFilter()
    assert stream.feed(text) == ''
    with pytest.raises(VisibleAnswerError) as error:
        stream.finish()
    assert 'private' not in str(error.value)
    with pytest.raises(VisibleAnswerError, match='closed'):
        stream.feed('late content')


def test_hidden_prefix_budget_does_not_limit_already_visible_content():
    stream = LeadingThoughtFilter(max_hidden_characters=20)
    assert stream.feed('Answer ' + 'x' * 100) == 'Answer ' + 'x' * 100
    assert stream.finish() == ''
    for text in (' ' * 21, '<think>' + 'private' * 10, '<think>x</think><think>y</think>'):
        stream = LeadingThoughtFilter(max_hidden_characters=20)
        with pytest.raises(VisibleAnswerError, match='limit'):
            stream.feed(text)


def test_nested_leading_thought_fails_closed():
    stream = LeadingThoughtFilter()
    with pytest.raises(VisibleAnswerError, match='nested'):
        stream.feed('<think>private<think>nested</think>still private</think>answer')


def test_literal_middle_tag_split_across_chunks_is_preserved():
    stream = LeadingThoughtFilter()
    assert stream.feed('Answer starts ') == 'Answer starts '
    assert stream.feed('<thi') == '<thi'
    assert stream.feed('nk>literal</think>') == 'nk>literal</think>'
    assert stream.finish() == ''
