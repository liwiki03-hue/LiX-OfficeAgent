from tests.test_document_conversation import FixedAnswer, send
from tests.test_document_conversation_browser import configure
from tests.test_paged_browser import paged_browser
from tests.test_paged_inbox_browser import select_reference


def test_clarifying_after_moving_selection_keeps_original_word_question(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    configure(page, api, monkeypatch)
    model = FixedAnswer({'content': '标题是“技术报告”。'})
    monkeypatch.setattr(api._connections, 'build_model', model.build_model)
    select_reference(page, '技术报告')
    question = '我选中的标题和整篇文档里，标题是怎么写的？'
    prompt = page.locator('#floating-input')
    submit = page.locator('#floating-send')
    prompt.fill(question)
    submit.click()
    pending = page.locator('.conversation-turn[data-state="waiting"]')
    pending.get_by_text('你想先核对当前选区，还是整篇文档？', exact=True).wait_for()
    turn_id = pending.get_attribute('data-turn-id')
    assert not model.calls
    select_reference(page)
    page.locator('#paged-reference-state').filter(has_text='正文第2段').wait_for()
    prompt.fill('只看刚才选区')
    submit.click()
    reply = page.locator(f'.conversation-turn[data-state="complete"][data-turn-id="{turn_id}"]')
    reply.get_by_text('正文第1段').wait_for()
    assert page.locator('.conversation-turn').count() == 1
    assert page.locator('.floating-message.user p').all_text_contents() == [question]
    assert list(api._floating.conversation.turns) == [turn_id] and api._floating.conversation.pending_scope is None
    assert len(model.calls) == 1 and question in model.calls[0] and '技术报告' in model.calls[0]
    assert '温度120 ℃' not in model.calls[0] and '非目标段落不变' not in model.calls[0] and '第二页正文' not in model.calls[0]
    assert '标题是“技术报告”' in reply.inner_text()
    assert source.read_bytes() == original and not api._floating.outputs and not errors


def test_eightieth_word_question_can_use_new_selection_in_original_turn(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    configure(page, api, monkeypatch)
    model = FixedAnswer({'content': '温度是 120 ℃。'})
    monkeypatch.setattr(api._connections, 'build_model', model.build_model)
    conversation = api._floating.conversation
    seed = send(api._floating, 'first-status', '进度如何？')['turn']
    with conversation.lock:
        conversation.turns.update({f'old-{i}': dict(seed, id=f'old-{i}', question=f'旧记录 {i}') for i in range(78)})
    page.wait_for_function("document.querySelectorAll('.conversation-turn').length === 79")
    question = '只看我选中的那句话，温度是多少？'
    page.locator('#floating-input').fill(question)
    page.locator('#floating-send').click()
    waiting = page.locator('.conversation-turn[data-state="waiting"]')
    reselect = waiting.get_by_role('button', name='用新选区回答', exact=True)
    reselect.wait_for()
    turn_id = waiting.get_attribute('data-turn-id')
    assert len(conversation.turns) == 80 and not model.calls
    select_reference(page)
    reselect.click()
    reply = page.locator(f'.conversation-turn[data-state="complete"][data-turn-id="{turn_id}"]')
    reply.get_by_text('正文第2段').wait_for()
    assert '温度是 120 ℃' in reply.inner_text()
    assert len(conversation.turns) == 80 and page.locator('.conversation-turn').count() == 80
    assert page.locator('.floating-message.user p').last.text_content() == question
    assert len(model.calls) == 1 and question in model.calls[0] and '温度120 ℃' in model.calls[0]
    assert '技术报告' not in model.calls[0] and '非目标段落不变' not in model.calls[0]
    assert source.read_bytes() == original and not api._floating.outputs and not errors
