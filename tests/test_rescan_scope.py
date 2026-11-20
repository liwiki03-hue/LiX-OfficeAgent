import pytest

from office_agent.webui.floating_session import FloatingApi
from tests.test_document_conversation import Answers, attach, finish, send
from tests.test_document_conversation_browser import configure
from tests.test_paged_browser import paged_browser
from tests.test_paged_editor import draft, sync
from tests.test_paged_inbox_browser import select_reference


QUESTION = '只看我选中的那句话，温度是多少？'
MIXED = '我选中的标题和整篇文档哪个提到了验收？'


def select_node(service, index):
    item = service.manifest[index]
    result = service.reference(service.document_id, service.revision, {
        'node_id': item['id'], 'start': 0, 'end': len(item['text']),
    })
    return result['reference']['id']


@pytest.mark.parametrize('reference,revision', [(None, 1), ('expired', 1), ('old-version', 0)])
def test_missing_selection_resumes_original_without_widening(draft, reference, revision):
    service, floating, source = draft
    original = source.read_bytes()
    sync(service)
    answers = Answers()
    attach(floating, answers)
    turn = send(floating, 'original', QUESTION, revision, reference)['turn']
    assert turn['status'] == 'waiting' and '重新选中' in turn['text'] and turn['question'] == QUESTION
    assert not answers.calls and not floating.conversation.readers and not turn['references']
    ref = select_node(service, 1)
    resumed = send(floating, 'followup', '用新选区回答', 1, ref)
    assert resumed['handled'] and resumed['turn']['id'] == 'original'
    finish(floating)
    assert list(floating.conversation.turns) == ['original']
    assert len(answers.calls) == 1 and QUESTION in answers.calls[0] and '温度120 ℃' in answers.calls[0]
    assert '非目标内容保持' not in answers.calls[0] and '验收报告' not in answers.calls[0]
    assert '正文第2段' in floating.conversation.turns['original']['answer']
    send(floating, 'full', '请阅读整篇文档，内容是什么？', 1, ref)
    finish(floating)
    assert len(answers.calls) == 2 and '验收报告' in answers.calls[1] and '非目标内容保持' in answers.calls[1]
    assert source.read_bytes() == original and not floating.outputs


@pytest.mark.parametrize('choice,stale', [('只看刚才选区', False), ('看整篇', False), ('只看刚才选区', True)])
def test_clarification_consumes_original_verified_scope(draft, choice, stale):
    service, floating, source = draft
    original = source.read_bytes()
    sync(service)
    answers = Answers()
    attach(floating, answers)
    ref = select_node(service, 0)
    assert send(floating, 'original', MIXED, 1, ref)['turn']['status'] == 'waiting'
    if stale:
        sync(service, 2)
    current = select_node(service, 1)
    receipt = send(floating, 'choice', choice, service.revision, current)
    assert receipt['turn']['id'] == 'original'
    if stale:
        assert '原选区已失效' in receipt['turn']['text'] and not answers.calls
        assert receipt['turn']['stage'] == 'reference'
        send(floating, 'new-ref', '用新选区回答', 2, current)
    finish(floating)
    row = floating.conversation.turns['original']
    assert row['status'] == 'complete' and floating.conversation.pending_scope is None
    assert len(answers.calls) == 1 and MIXED in answers.calls[0]
    assert ('验收报告' in answers.calls[0]) == (not stale)
    assert ('温度120 ℃' in answers.calls[0]) == (stale or choice == '看整篇')
    assert ('非目标内容保持' in answers.calls[0]) == (choice == '看整篇')
    if not stale and choice != '看整篇':
        assert '正文第1段' in row['answer']
    assert send(floating, 'choice', choice, service.revision, current)['duplicate']
    assert len(answers.calls) == 1 and list(floating.conversation.turns) == ['original']
    assert source.read_bytes() == original and not floating.outputs


def test_selection_synonyms_and_real_reference_from_older_revision(draft):
    service, floating, _ = draft
    sync(service)
    old = select_node(service, 1)
    sync(service, 2)
    model = Answers()
    attach(floating, model)
    for number, question in enumerate(['只看我选的那句话，温度是多少？', '只看我刚才选择的那句话，温度是多少？']):
        row = send(floating, 'stale-' + str(number), question, 2, old)['turn']
        assert row['status'] == 'waiting' and row['stage'] == 'reference' and '原选区已失效' in row['text']
    assert not model.calls


def test_abandoned_and_restored_clarifications_cannot_consume_a_choice(draft):
    service, floating, _ = draft
    sync(service)
    answers = Answers()
    attach(floating, answers)
    ref = select_node(service, 0)
    send(floating, 'old', MIXED, 1, ref)
    send(floating, 'independent', '整篇文档内容是什么？', 1)
    finish(floating)
    assert floating.conversation.turns['old']['status'] == 'stopped'
    with pytest.raises(ValueError, match='没有对应'):
        send(floating, 'late', '只看刚才选区', 1, ref)
    send(floating, 'new-question', MIXED, 1, ref)
    with pytest.raises(ValueError, match='没有对应'):
        FloatingApi.conversation_begin(floating.api, 'delayed', '只看刚才选区', service.document_id,
            service.source_version, 1, ref, floating.conversation.binding, False, 'old')
    floating.conversation.reset()
    floating.conversation.restore()
    assert floating.conversation.turns['new-question']['status'] == 'interrupted'
    with pytest.raises(ValueError, match='没有对应'):
        send(floating, 'reopened', '只看刚才选区', 1, ref)
    assert len(answers.calls) == 1


def test_scope_reselection_button_updates_the_original_answer(paged_browser, monkeypatch):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    answers = configure(page, api, monkeypatch)
    page.locator('#floating-input').fill(QUESTION)
    page.locator('#floating-send').click()
    button = page.get_by_role('button', name='用新选区回答', exact=True)
    button.wait_for()
    assert not answers.calls
    select_reference(page)
    button.click()
    page.locator('.conversation-turn[data-state="complete"]').get_by_text('正文第2段').wait_for()
    assert page.locator('.conversation-turn').count() == 1
    assert page.locator('.floating-message.user p').all_text_contents() == [QUESTION]
    assert len(answers.calls) == 1 and '温度120 ℃' in answers.calls[0] and '非目标段落不变' not in answers.calls[0]
    assert not errors and not api._floating.outputs and source.read_bytes() == original
