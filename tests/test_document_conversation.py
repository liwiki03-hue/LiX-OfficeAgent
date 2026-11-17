import threading
import time
from types import SimpleNamespace

import pytest

from office_agent.model_connections import ModelRoute
from office_agent.webui.document_conversation import intent
from office_agent.webui.floating_session import FloatingApi, FloatingSession
from tests.test_editable_copies import readonly
from tests.test_office_canvas import canvas


class Answers:
    def __init__(self, entered=None, release=None):
        self.calls = []
        self.entered, self.release = entered, release
        self.route = ModelRoute('read-test', '隔离回答', 'simulated', 'agent', 'v1', 'openai', 'https://example.invalid/v1', 32768, 'test-only', 0)

    def capture(self, capability):
        assert capability == 'agent'
        return self.route

    def build_model(self, route):
        owner = self
        class Model:
            def bind_tools(self, tools):
                assert not any('propose' in item['function']['name'] or 'write' in item['function']['name'] for item in tools)
                return self

            def invoke(self, messages):
                owner.calls.append(str(messages))
                if owner.entered:
                    owner.entered.set()
                    assert owner.release.wait(5)
                return {'content': '第一点说明实验条件。第二点依据是120 ms的延迟记录。'}
        return Model()


def attach(floating, connections):
    floating.api._floating = floating
    floating.api._floating_call = lambda action, *args: action(*args)
    floating.api._connections = connections


def send(floating, identifier, text, revision=0, reference=None):
    return FloatingApi.conversation_begin(floating.api, identifier, text, floating.document['id'], floating.document['version'], revision, reference)


def finish(floating):
    deadline = time.monotonic() + 8
    while floating.conversation.readers:
        assert time.monotonic() < deadline
        time.sleep(.01)


@pytest.mark.parametrize('text,expected', [
    ('最后一段总结了什么？', 'read'), ('这张表的120 ms表示什么？', 'read'),
    ('给我两种更简洁的写法', 'read'), ('把第二种换进去', 'edit'),
    ('把最后一段改成两句话，保留全部数值', 'edit'), ('帮我处理一下这个结尾', 'clarify'),
    ('进度如何？', 'status'), ('到哪了？', 'status'), ('停一下', 'stop'),
    ('在文章末尾添加一段总结', 'edit'), ('直接在这里回答，不用修改', 'read')])
def test_intent_keeps_questions_out_of_writes(text, expected):
    assert intent(text) == expected


def test_style_instruction_and_numeric_question_during_edit_have_distinct_routes():
    assert intent('第一段语气自然，保留20 ms与30组。') == 'edit'
    assert intent('这张表的数值是什么？', editing=True) == 'read'
    assert intent('补充表格文字说明，保留120 ms与表格结构。') == 'edit'
    assert intent('为什么要修改最后一段？') == 'read'
    assert intent('如何把最后一段改成两句话？') == 'read'


def test_readonly_answer_no_copy_authority_render_or_requirement(readonly, monkeypatch):
    service, floating, source = readonly
    connections = Answers()
    attach(floating, connections)
    before = source.read_bytes()
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **k: pytest.fail('Question must not render a copy'))
    reply = send(floating, 'ask-last', '最后一段的总结是什么意思？')
    assert reply['handled'] and reply['turn']['intent'] == 'read'
    assert send(floating, 'ask-last', '最后一段的总结是什么意思？')['duplicate']
    finish(floating)
    row = floating.conversation.state()['turns'][0]
    assert row['status'] == 'complete' and '120 ms' in row['answer']
    assert row['references'][0]['scope'] == '文末正文'
    assert not floating.access.state()['id'] and floating.runtime is None
    assert not service.publishing and not service.preparing and not floating.outputs
    assert len(connections.calls) == 1 and source.read_bytes() == before
    assert not list(source.parent.glob('*.agent-*.docx'))


def test_followup_and_public_snapshot_restore_without_execution(readonly):
    _, floating, source = readonly
    connections = Answers()
    attach(floating, connections)
    send(floating, 'two-options', '给我两种更简洁的写法')
    finish(floating)
    send(floating, 'followup', '第二点依据在哪？')
    finish(floating)
    assert '第二点依据是120 ms' in connections.calls[-1]
    assert len(floating.conversation.state()['turns']) == 2
    new_api = SimpleNamespace(_app=type(floating.app)(), _auth=floating.api._auth, _memory_key=lambda:'self-created-test', _connections=connections)
    restored = FloatingSession(new_api)
    restored.admit(str(source));restored.read(restored.read_token, True)
    rows = restored.conversation.state()['turns']
    assert len(rows) == 2 and all(row['status'] == 'complete' and row['answer'] for row in rows)
    assert not restored.access.state()['id'] and restored.runtime is None and len(connections.calls) == 2


def test_progress_stop_and_other_question_are_independent(readonly):
    _, floating, _ = readonly
    entered, release = threading.Event(), threading.Event()
    connections = Answers(entered, release)
    attach(floating, connections)
    try:
        send(floating, 'editing', '把最后一段改成两句')
        floating.notice('publication_phase', {'phase':'rendering', 'text':'正在校验工作文件。'})
        result = send(floating, 'progress', '到哪了？')
        assert result['turn']['status'] == 'complete' and '校验工作文件' in result['turn']['answer']
        assert '校验工作文件' not in floating.conversation.history()
        send(floating, 'question', '最后一段是什么意思？')
        assert entered.wait(3)
        result = send(floating, 'stop', '停一下')
        assert result['turn']['status'] == 'complete'
        assert floating.conversation.turns['editing']['status'] == 'stopped'
    finally:
        release.set();finish(floating)
    assert floating.conversation.turns['question']['status'] == 'stopped'
    assert len(connections.calls) == 1


@pytest.mark.parametrize('canvas', ['.xlsx', '.pptx'], indirect=True)
def test_canvas_plain_answer_and_read_reference_do_not_change_draft(canvas):
    service, floating, source, result, _ = canvas
    connections = Answers()
    attach(floating, connections)
    original = source.read_bytes()
    target = 'B2' if source.suffix == '.xlsx' else 'run-0'
    item = 'sheet-0' if source.suffix == '.xlsx' else 'slide-0'
    reference = service.ai.select(service.document_id, 0, item, target, False, True)['reference']
    assert reference['read_only'] and not reference['allow_numeric_change']
    assert service.ai.reference is None
    send(floating, 'explain', '这段内容表示什么？', reference=reference['id'])
    finish(floating)
    row = floating.conversation.turns['explain']
    assert row['status'] == 'complete' and row['answer']
    assert not service.ai.proposals and service.revision == 0 and not floating.outputs
    assert not floating.access.state()['id'] and source.read_bytes() == original


def test_addition_after_published_snapshot_is_not_reported_as_applied(readonly):
    _, floating, _ = readonly
    attach(floating, Answers())
    send(floating, 'first', '修改最后一段')
    send(floating, 'second', '改成两句')
    floating.conversation.turns['second']['instruction_revision'] = 2
    floating.notice('draft_published', {'receipt': {'id':'copy-one', 'output':'first.docx', 'kind':'paged_candidate', 'output_sha256':'a'*64},
                                     'draft_revision':1, 'instruction_revision':1, 'pending_updates':True})
    assert floating.conversation.turns['second']['status'] == 'waiting'
    assert floating.conversation.turns['first']['stage'] == 'version_complete'
    floating.notice('draft_published', {'receipt': {'id':'copy-two', 'output':'second.docx', 'kind':'paged_candidate', 'output_sha256':'b'*64},
                                     'draft_revision':2, 'instruction_revision':2})
    assert floating.conversation.turns['second']['status'] == 'complete'
    assert len(floating.conversation.turns['first']['results']) == 2


def test_copy_preparation_accepts_extra_and_stop_without_blocking_snapshot(readonly, monkeypatch):
    from office_agent.editable_copies import EditableCopyPublisher
    service, floating, source = readonly
    attach(floating, Answers())
    entered, release = threading.Event(), threading.Event()
    prepare = EditableCopyPublisher.plan_copy
    def held(publisher, **values):
        entered.set()
        assert release.wait(5)
        return prepare(publisher, **values)
    monkeypatch.setattr(EditableCopyPublisher, 'plan_copy', held)
    monkeypatch.setattr(EditableCopyPublisher, 'execute', lambda *a, **k: pytest.fail('Stopped preparation must not publish'))
    text = '把最后一段改成两句，保留全部数值'
    send(floating, 'edit', text)
    service.submit('edit', text, None, service.document_id, service.source_version, '', service.inbox_id)
    runtime = floating.runtime
    owned_before = set(threading.enumerate())
    try:
        result = service.automatic_editable_copy(service.document_id, 0, service.inbox_id, runtime.session_id, runtime.revision, floating.access.state()['id'])
        assert result['preparing'] and entered.wait(3)
        assert floating.collect()['paged']['preparing']
        send(floating, 'extra', '改成两句')
        service.submit('extra', '改成两句', None, service.document_id, service.source_version, runtime.session_id, service.inbox_id)
        assert runtime.revision == 2
        send(floating, 'stop', '停一下')
        assert runtime.closed and not service.preparing
    finally:
        release.set()
        for thread in set(threading.enumerate()) - owned_before:
            if '_prepare_copy' in thread.name:
                thread.join(6)
                assert not thread.is_alive()
    assert not floating.outputs and not list(source.parent.glob('*.agent-*.docx'))
    assert not service.copy_publisher.plans


def test_failed_answer_resumes_same_turn_without_copy(readonly):
    _, floating, source = readonly
    class Recovering(Answers):
        def __init__(self):
            super().__init__()
            self.attempts = 0
        def build_model(self, route):
            self.attempts += 1
            if self.attempts == 1:
                raise TimeoutError('simulated provider timeout')
            return super().build_model(route)
    connections = Recovering()
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'retry-me', '最后一段是什么意思？')
    finish(floating)
    assert floating.conversation.turns['retry-me']['status'] == 'failed'
    assert FloatingApi.conversation_retry(floating.api, 'retry-me', floating.document['id'], floating.document['version'])['ok']
    finish(floating)
    assert floating.conversation.turns['retry-me']['status'] == 'complete'
    assert len(floating.conversation.turns) == 1 and connections.attempts == 2
    assert not floating.outputs and source.read_bytes() == original


def test_copy_switch_interrupts_old_read_without_duplicate_request(readonly):
    _, floating, _ = readonly
    entered, release = threading.Event(), threading.Event()
    connections = Answers(entered, release)
    attach(floating, connections)
    try:
        send(floating, 'old-read', '最后一段是什么意思？')
        assert entered.wait(3)
        saved = floating.conversation.state()
        origin = floating.document['version']
        floating.conversation.reset()
        floating.document = {**floating.document, 'id': 'verified-working-copy', 'version': 'b'*64}
        floating.conversation.rebind(saved)
        row = floating.conversation.turns['old-read']
        assert row['status'] == 'interrupted' and not row['answer']
        assert row['document_id'] == 'verified-working-copy'
        assert row['working_version'] == origin
    finally:
        release.set()
        finish(floating)
    assert floating.conversation.turns['old-read']['status'] == 'interrupted'
    assert len(connections.calls) == 1 and not floating.outputs


def test_unresolved_ordinal_does_not_start_copy_or_guess_first_paragraph(readonly, monkeypatch):
    service, floating, source = readonly
    attach(floating, Answers())
    original = source.read_bytes()
    text = '修改第2段'
    send(floating, 'ordinal', text)
    service.submit('ordinal', text, None, service.document_id, service.source_version, '', service.inbox_id)
    runtime = floating.runtime
    monkeypatch.setattr('office_agent.editable_copies.EditableCopyPublisher.plan_copy', lambda *a, **k: pytest.fail('Unresolved scope must not prepare a copy'))
    with pytest.raises(ValueError, match='明确修改范围'):
        service.automatic_editable_copy(service.document_id, 0, service.inbox_id, runtime.session_id, runtime.revision, floating.access.state()['id'])
    assert not service.preparing and not floating.outputs and source.read_bytes() == original
    assert floating.conversation.turns['ordinal']['stage'] == 'scope'


def test_final_reply_uses_verified_changes_instead_of_unapplied_proposal(readonly):
    _, floating, _ = readonly
    attach(floating, Answers())
    send(floating, 'edit', '在文末添加一段总结')
    floating.notice('draft_proposal', {'old':'提案旧文', 'new':'未采用的提案'})
    floating.notice('draft_published', {'receipt': {'id':'verified-output', 'output':'copy.docx', 'kind':'paged_candidate',
        'output_sha256':'a'*64, 'changes':[{'kind':'paragraph','operation':'insert','before':'','after':'实际已写入正文'}]},
        'draft_revision':1, 'instruction_revision':1})
    row = floating.conversation.turns['edit']
    assert row['status'] == 'complete' and '已新增1段正文' in row['text']
    assert row['changes'] == [{'before':'', 'after':'实际已写入正文'}]


def test_skill_reference_history_persists_without_changing_reply_or_replaying_execution(readonly):
    _,floating,source=readonly
    connections=Answers();attach(floating,connections)
    original=source.read_bytes()
    send(floating,'skill-history','稍后润色最后一段')
    floating.conversation.active_task='skill-history'
    before=floating.conversation.state()['turns'][0]
    receipt={'id':'user:reference','display_name':'本机参考','body_sha256':'a'*64,'characters':32,
        'instruction_revision':0,'draft_revision':0}
    floating.notice('skill_loaded',{**receipt,'raw_body':'unretained material'})
    floating.notice('skill_loaded',receipt)
    floating.notice('skill_loaded',{**receipt,'body_sha256':'invalid'})
    row=floating.conversation.state()['turns'][0]
    assert row['skills']==[receipt]
    assert (row['stage'],row['text'],row['steps'])==(before['stage'],before['text'],before['steps'])
    new_api=SimpleNamespace(_app=type(floating.app)(),_auth=floating.api._auth,_memory_key=lambda:'self-created-test',_connections=connections)
    restored=FloatingSession(new_api)
    restored.admit(str(source));restored.read(restored.read_token,True)
    assert restored.conversation.state()['turns'][0]['skills']==[receipt]
    assert not restored.access.state()['id'] and restored.runtime is None and not connections.calls
    assert source.read_bytes()==original and not list(source.parent.glob('*.agent-*.docx'))
