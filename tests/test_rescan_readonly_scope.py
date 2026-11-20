import pytest
from docx import Document
from docx.text.paragraph import Paragraph

from office_agent.paged_documents import W
from tests.test_document_conversation import Answers, attach, finish, send
from tests.test_document_conversation_browser import configure
from tests.test_editable_copies import make_readonly, readonly
from tests.test_paged_browser import paged_browser


QUESTION = '只看我选中的那句话，温度是多少？'


def test_readonly_engine_selection_and_source_excerpt_use_checked_offsets(readonly):
    page, floating, source = readonly
    original = source.read_bytes()
    model = Answers()
    attach(floating, model)
    book = Document(source)
    texts = [Paragraph(node, book._body).text for node in book.element.body.iter(W + 'p')]
    selected = '25 ℃'
    start = texts[1].index(selected)
    target = {'source_index': 1, 'paragraphs': texts, 'start': start, 'end': start + len(selected), 'expected': selected}
    with pytest.raises(ValueError, match='与原稿不一致'):
        page.reference(page.document_id, 0, {**target, 'paragraphs': [*texts[1:], texts[0]]})
    locator = next(item['locator_id'] for item in floating.view['blocks'] if item.get('paragraph_index') == 1)
    with pytest.raises(ValueError, match='文字与原稿位置不一致'):
        page.reference(page.document_id, 0, {**target, 'expected': '其他温度'})
    for index, binding in enumerate([target, {'source_locator_id': locator, 'start': start, 'end': start + len(selected), 'expected': selected}]):
        ref = page.reference(page.document_id, 0, binding)['reference']
        assert ref['old'] == selected and ref['kind'] == 'source_text'
        turn = f'selection-{index}'
        assert send(floating, turn, QUESTION, 0, ref['id'])['handled']
        finish(floating)
        assert '正文第2段' in floating.conversation.turns[turn]['answer']
        assert selected in model.calls[-1] and '30个样本' not in model.calls[-1] and '链接显示文字' not in model.calls[-1]
    assert len(model.calls) == 2 and not page.editable and not floating.outputs
    assert not floating.access.state()['active'] and source.read_bytes() == original


def test_readonly_word_recovers_selection_from_host_excerpt(paged_browser, monkeypatch, tmp_path):
    browser, api, _, errors = paged_browser
    source = tmp_path / '只读问答.docx'
    make_readonly(source)
    original = source.read_bytes()
    floating = api._floating
    answers = configure(browser, api, monkeypatch)
    assert floating.admit(str(source))['ok']
    assert floating.read(floating.read_token, True)['ok']
    browser.evaluate('OfficeFloating.refresh()')
    browser.get_by_role('button', name='创建可编辑副本', exact=True).wait_for()
    browser.locator('#floating-input').fill(QUESTION)
    browser.locator('#floating-send').click()
    resume = browser.get_by_role('button', name='用新选区回答', exact=True)
    resume.wait_for()
    assert not answers.calls
    browser.get_by_role('button', name='引用选区', exact=True).click()
    picker = browser.get_by_role('dialog', name='从原稿摘录引用')
    picker.get_by_label('原稿段落').select_option(index=1)
    area = picker.get_by_label('原稿文字')
    area.focus()
    area.evaluate("node=>{const i=node.value.indexOf('25 ℃');node.setSelectionRange(i,i+4)}")
    picker.get_by_role('button', name='引用所选文字').click()
    picker.wait_for(state='detached')
    resume.click()
    browser.locator('.conversation-turn[data-state="complete"]').get_by_text('正文第2段').wait_for()
    assert len(answers.calls) == 1 and '25 ℃' in answers.calls[0] and '30个样本' not in answers.calls[0]
    assert not errors and not floating.paged.editable and source.read_bytes() == original
    assert not floating.access.state()['active'] and not floating.outputs
