import pytest
from docx import Document

from office_agent.instruction_runtime import InstructionRuntime
from office_agent.webui.document_conversation import intent
from tests.test_document_conversation import Answers, attach, finish, send
from tests.test_l3_l4_editing import sync_book
from tests.test_paged_editor import draft


@pytest.mark.parametrize('in_progress', [False, True])
def test_advice_and_edit_history_have_no_write_handoff(draft, in_progress):
    page, floating, source = draft
    model = Answers()
    attach(floating, model)
    original = source.read_bytes()
    if in_progress:
        question = '把第二段改成两句'
        assert not send(floating, 'editing', question)['handled']
        floating.runtime = InstructionRuntime(floating.document['version'])
        floating.runtime.submit('editing', question)
    before_access = floating.access.state()
    for identifier, question in [('advice', '先别改写这段，给点意见。'), ('history', '这段修改过吗？')]:
        receipt = send(floating, identifier, question)
        assert receipt['handled'] and receipt['turn']['intent'] == 'read'
        finish(floating)
        assert floating.conversation.turns[identifier]['status'] == 'complete'
    assert len(model.calls) == 1 and '给点意见' in model.calls[0]
    assert '无法确认' in floating.conversation.turns['history']['answer']
    assert floating.access.state() == before_access and not floating.access.state()['active']
    assert page.revision == 0 and not page.proposals and not floating.outputs
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))
    if in_progress:
        assert floating.conversation.active_task == 'editing' and floating.runtime.revision == 1
    else:
        assert floating.conversation.active_task is None and floating.runtime is None
    for text in ['能帮我把这段改成两句吗？', '不要改标题，把第二段压成两句']:
        assert intent(text, editing=in_progress) == 'edit'
    assert not send(floating, 'explicit-rewrite', '能帮我把这段改成两句吗？')['handled']


def test_history_compares_only_the_selected_paragraph_with_its_original_version(draft):
    page, floating, source = draft
    model = Answers()
    attach(floating, model)
    original = source.read_bytes()
    book = Document(source)
    ids = ['title', 'body', 'tail']
    sync_book(page, book, ids, 0, True)
    book.paragraphs[1].runs[0].text = '这一段现在有新的温度记录。'
    page.changed(page.document_id, 1)
    sync_book(page, book, ids, 1)
    for node, number, expected in [('body', 2, '可以确认这一段改过'), ('title', 1, '无法确认中途是否改过')]:
        ref = page.reference(page.document_id, 1, {'node_id': node, 'start': 0, 'end': 2})['reference']
        assert send(floating, node, '这段修改过吗？', 1, ref['id'])['handled']
        finish(floating)
        turn = floating.conversation.turns[node]
        assert turn['status'] == 'complete' and expected in turn['answer']
        assert turn['references'][0]['location'] == f'正文第{number}段'
    assert not model.calls and not floating.access.state()['active'] and not floating.outputs
    assert page.revision == 1 and source.read_bytes() == original
