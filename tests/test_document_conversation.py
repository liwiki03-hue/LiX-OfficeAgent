import threading
import time
import json
from types import SimpleNamespace

import pytest

from office_agent.model_connections import ModelRoute
from office_agent.webui.document_conversation import intent
from office_agent.webui.floating_session import FloatingApi, FloatingSession
from tests.test_editable_copies import readonly
from tests.test_office_canvas import canvas
from tests.test_paged_editor import draft, snapshot


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


def retry(floating, identifier):
    return FloatingApi.conversation_retry(floating.api, identifier, floating.document['id'], floating.document['version'])


class RetryAnswers(Answers):
    def __init__(self, answer='这是所选标记。', *, entered=None, release=None):
        super().__init__(entered, release)
        self.answer = answer
        self.builds = 0

    def build_model(self, route):
        self.builds += 1
        if self.builds == 1:
            raise TimeoutError('fixture: model construction')
        return self._model()

    def _model(self):
        def invoke(messages):
            self.calls.append(str(messages))
            if self.entered:
                self.entered.set()
                assert self.release.wait(5)
            return {'content': self.answer}
        model = SimpleNamespace(bind_tools=lambda _: model, invoke=invoke)
        return model


def send(floating, identifier, text, revision=0, reference=None):
    return FloatingApi.conversation_begin(floating.api, identifier, text, floating.document['id'], floating.document['version'], revision, reference)


def finish(floating):
    deadline = time.monotonic() + 8
    while floating.conversation.readers:
        assert time.monotonic() < deadline
        time.sleep(.01)


def test_word_retry_keeps_original_range_after_selection_changes_and_rejects_new_draft(draft):
    service, floating, source = draft
    selected = '原选区标记甲。'
    service.changed(service.document_id, 1)
    service.sync(service.document_id, 1, *snapshot(service, selected))
    ref = service.reference(service.document_id, 1, {'node_id': 'node-1', 'start': 0, 'end': len(selected)})['reference']
    entered, release = threading.Event(), threading.Event()
    connections = RetryAnswers(entered=entered, release=release)
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'read-selected', '所选内容是什么？', 1, ref['id'])
    finish(floating)
    assert floating.conversation.turns['read-selected']['status'] == 'failed'
    assert connections.builds == 1 and not connections.calls
    service.reference(service.document_id, 1, {'node_id': 'node-2', 'start': 0, 'end': 6})
    try:
        retry(floating, 'read-selected')
        assert entered.wait(3)
        with pytest.raises(ValueError, match='不能重试|未释放'):
            retry(floating, 'read-selected')
    finally:
        release.set()
        finish(floating)
    assert len(connections.calls) == 1 and selected in connections.calls[0]
    assert '非目标内容保持' not in connections.calls[0] and '验收报告' not in connections.calls[0]
    assert '正文第2段' in floating.conversation.turns['read-selected']['answer']
    assert floating.conversation.metrics_state('read-selected')['attempt_count'] == 2
    connections = RetryAnswers()
    attach(floating, connections)
    ref = service.reference(service.document_id, 1, {'node_id': 'node-1', 'start': 0, 'end': len(selected)})['reference']
    send(floating, 'stale-selected', '所选内容是什么？', 1, ref['id'])
    finish(floating)
    service.changed(service.document_id, 2)
    service.sync(service.document_id, 2, *snapshot(service, '草稿的新文字。'))
    retry(floating, 'stale-selected')
    assert '原选区已失效' in floating.conversation.turns['stale-selected']['text']
    assert connections.builds == 1 and not connections.calls
    assert not service.proposals and not floating.outputs and source.read_bytes() == original


@pytest.mark.parametrize('canvas', ['.xlsx', '.pptx'], indirect=True)
def test_canvas_retry_retains_original_read_only_reference(canvas):
    service, floating, source, _, _ = canvas
    item, target, other, selected, outside = (('sheet-0', 'B2', 'A4', '120', '合并说明') if source.suffix == '.xlsx'
                                             else ('slide-0', 'run-0', 'run-1', '实验汇报', '温度120'))
    connections = RetryAnswers()
    attach(floating, connections)
    original = source.read_bytes()
    ref = service.ai.select(service.document_id, 0, item, target, False, True)['reference']
    send(floating, 'canvas-selected', '所选内容是什么？', 0, ref['id'])
    finish(floating)
    assert connections.builds == 1 and not connections.calls
    replacement = service.ai.select(service.document_id, 0, item, other, False, True)['reference']
    assert replacement['id'] != ref['id']
    retry(floating, 'canvas-selected')
    finish(floating)
    row = floating.conversation.turns['canvas-selected']
    assert row['status'] == 'complete' and row['references'][0]['location'] == ref['item_name'] + ' · ' + ref['location']
    assert len(connections.calls) == 1 and selected in connections.calls[0] and outside not in connections.calls[0]
    assert not floating.access.state()['id'] and not service.ai.proposals and not floating.outputs and source.read_bytes() == original


def test_word_full_scope_survives_reopen_but_selected_scope_requires_fresh_selection(draft):
    service, floating, source = draft
    service.changed(service.document_id, 1)
    service.sync(service.document_id, 1, *snapshot(service))
    ref = service.reference(service.document_id, 1, {'node_id':'node-1', 'start':0, 'end':4})['reference']
    first = RetryAnswers()
    attach(floating, first)
    send(floating, 'saved-selection', '所选内容是什么？', 1, ref['id'])
    finish(floating)
    attach(floating, RetryAnswers())
    send(floating, 'saved-whole', '整篇文档内容是什么？', 1)
    finish(floating)
    current = RetryAnswers()
    current.builds = 1
    api = SimpleNamespace(_app=type(floating.app)(), _auth=floating.api._auth, _memory_key=lambda:'self-created-test')
    reopened = FloatingSession(api)
    assert reopened.admit(str(source))['ok'] and reopened.read(reopened.read_token, True)['ok']
    attach(reopened, current)
    retry(reopened, 'saved-selection')
    assert '原选区已失效' in reopened.conversation.turns['saved-selection']['text']
    assert current.builds == 1 and not current.calls
    retry(reopened, 'saved-whole')
    finish(reopened)
    assert reopened.conversation.turns['saved-whole']['status'] == 'complete'
    assert len(current.calls) == 1 and '非目标内容保持' in current.calls[0]
    assert not reopened.outputs and not reopened.access.state()['id']


def test_selected_title_and_document_end_have_distinct_model_material(draft):
    from tests.test_round_undo import export
    from docx import Document
    service, floating, source = draft
    book = Document(source)
    book.paragraphs[-1].text = '最终结论'
    service.changed(service.document_id, 1)
    service.sync(service.document_id, 1, *export(book))
    reference = service.reference(service.document_id, 1, {'node_id': 'round-node-0', 'start': 0, 'end': 4})['reference']
    for index, (question, answer, include, exclude, where) in enumerate([
        ('所选标题是什么？', '验收报告', '验收报告', '最终结论', '正文第1段'),
        ('我选中的标题最后两个字是什么？', '报告', '验收报告', '最终结论', '正文第1段'),
        ('整篇文档最后一段写了什么？', '最终结论', '最终结论', '验收报告', '正文第3段'),
    ]):
        connections = RetryAnswers(answer)
        connections.builds = 1
        attach(floating, connections)
        identifier = 'scope-' + str(index)
        send(floating, identifier, question, 1, reference['id'])
        finish(floating)
        turn = floating.conversation.turns[identifier]
        assert turn['status'] == 'complete' and turn['answer'].startswith(answer + '\n') and where in turn['answer']
        assert len(connections.calls) == 1 and include in connections.calls[0] and exclude not in connections.calls[0]
        assert not floating.outputs and not floating.access.state()['id']


def test_turn_limit_controls_rollover_migration_and_archived_output(draft, monkeypatch):
    from copy import deepcopy
    from pathlib import Path
    from tests.test_round_undo import make_round
    service, floating, source, receipt, _ = make_round(draft, monkeypatch, real_render=True)
    conversation = floating.conversation
    output = Path(receipt['output'])
    assert receipt['render_diff']['inputs_unchanged'] and output.is_file()
    conversation.turns['round-a']['results'] = [{'id': receipt['id'], 'name': output.name, 'verified': True,
                                                'sha256': receipt['output_sha256']}]
    attach(floating, RetryAnswers())
    seed = send(floating, 'first-status', '进度如何？')['turn']
    conversation.turns.update({f'migration-{index}': dict(deepcopy(seed), id=f'migration-{index}') for index in range(77)})
    assert len(conversation.turns) == 79
    entered, release = threading.Event(), threading.Event()
    connections = RetryAnswers(entered=entered, release=release)
    connections.builds = 1
    attach(floating, connections)
    held = service.data
    try:
        send(floating, 'eightieth', '请解释这份报告。', revision=2)
        assert entered.wait(3)
        document = floating.document
        context = conversation.binding
        with pytest.raises(ValueError, match='开始新对话'):
            send(floating, 'ordinary-extra', '另一条问题？', revision=2)
        with pytest.raises(ValueError, match='上一对话仍在'):
            FloatingApi.conversation_new(floating.api, context, document['id'], document['version'])
        progress = send(floating, 'overflow-progress', '进度如何？')['turn']
        assert progress['transient'] and '正在' in progress['answer']
        stop = send(floating, 'overflow-stop', '停止')['turn']
        assert stop['transient'] and '已停止' in stop['answer'] and len(conversation.turns) == 80
        assert conversation.turns['eightieth']['status'] == 'stopped'
    finally:
        release.set()
        finish(floating)
    assert len(connections.calls) == 1
    legacy = json.loads(conversation._target().read_bytes())
    assert legacy['schema'] == 1 and len(legacy['turns']) == 80 and 'control_turn' not in legacy
    state = FloatingApi.conversation_new(floating.api, context, document['id'], document['version'])['conversation']
    assert not state['turns'] and state['context_id'] != context
    assert service.data == held and output.is_file() and len(floating.outputs) == 1
    with pytest.raises(ValueError, match='对话已切换'):
        FloatingApi.conversation_new(floating.api, context, document['id'], document['version'])
    send(floating, 'new-question', '整篇文档内容是什么？', revision=2)
    finish(floating)
    assert conversation.turns['new-question']['status'] == 'complete' and len(connections.calls) == 2
    api = SimpleNamespace(_app=type(floating.app)(), _auth=floating.api._auth, _memory_key=lambda:'self-created-test')
    reopened = FloatingSession(api)
    assert reopened.admit(str(source))['ok'] and reopened.read(reopened.read_token, True)['ok']
    attach(reopened, connections)
    assert list(reopened.conversation.turns) == ['new-question']
    assert reopened.conversation.binding == state['context_id']
    active = reopened.document
    archive, = FloatingApi.conversation_archive(api, active['id'], active['version'])['archives']
    archived = FloatingApi.conversation_archive(api, active['id'], active['version'], archive['id'])['turns']
    assert len(archived) == 80 and archived[-1]['status'] == 'stopped'
    opened = FloatingApi.conversation_output(api, receipt['id'], active['id'], active['version'], archive['id'])['output']
    assert Path(opened['output']) == output and 'authorization' not in opened
    assert not reopened.outputs
    assert reopened.select_history_output(receipt['id'])['ok']
    assert reopened.app.path == output and not reopened.document['ready']
    assert not reopened.access.state()['id'] and reopened.runtime is None
    assert len(connections.calls) == 2 and len(list(source.parent.glob('*.agent-*.docx'))) == 1


def test_full_legacy_file_recovers_and_archives_without_loading_old_approvals(draft):
    from copy import deepcopy
    service, first, source = draft
    attach(first, RetryAnswers())
    row = send(first, 'legacy-seed', '进度如何？')['turn']
    first.conversation.turns.update({f'legacy-{index}': dict(deepcopy(row), id=f'legacy-{index}') for index in range(79)})
    first.conversation.persist()
    version = first.document['version']
    api = SimpleNamespace(_app=type(first.app)(), _auth=first.api._auth, _memory_key=lambda:'self-created-test')
    restored = FloatingSession(api)
    assert restored.admit(str(source))['ok'] and restored.read(restored.read_token, True)['ok']
    attach(restored, RetryAnswers())
    conv = restored.conversation
    assert len(conv.turns) == 80 and not restored.access.state()['id']
    assert send(restored, 'legacy-progress', '进度如何？')['turn']['transient']
    assert FloatingApi.conversation_new(api, conv.binding, restored.document['id'], version)['ok']
    assert len(conv.archived(conv.archives()[0]['id'])) == 80 and not conv.turns


def test_consultation_during_edit_does_not_join_older_requirements(draft):
    from office_agent.instruction_runtime import InstructionRuntime
    service, floating, source = draft
    connections = RetryAnswers()
    connections.builds = 1
    attach(floating, connections)
    edit = '能帮我把这段改成两句吗？'
    send(floating, 'edit-in-progress', edit)
    runtime = InstructionRuntime(floating.document['version'])
    floating.runtime = runtime
    runtime.submit('edit-in-progress', edit)
    revision = runtime.revision
    reply = send(floating, 'consult-in-progress', '这段需要修改吗？')
    finish(floating)
    assert reply['handled'] and floating.conversation.turns['consult-in-progress']['status'] == 'complete'
    assert floating.conversation.active_task == 'edit-in-progress' and runtime.revision == revision
    assert len(connections.calls) == 1 and not floating.outputs and not service.proposals


@pytest.mark.parametrize('text', ['这段需要修改吗？', '这段需要修改吗?', '这段有没有必要改', '给我修改建议，先别动原文'])
def test_consultation_only_does_not_become_an_active_edit(text):
    assert intent(text) == intent(text, editing=True) == 'read'
    assert intent('能帮我把这段改成两句吗？') == 'edit'


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


class FixedAnswer(Answers):
    def __init__(self, response):
        super().__init__()
        self.response = response

    def build_model(self, route):
        owner = self
        class Model:
            def bind_tools(self, tools):
                assert tools == []
                return self

            def invoke(self, messages):
                owner.calls.append(str(messages))
                return owner.response
        return Model()


def test_authorized_selection_answers_once_with_units_location_and_no_copy(draft):
    service, floating, source = draft
    text = '实验温度25 ℃，30个样本的平均延迟120 ms[1]。'
    service.changed(service.document_id, 1)
    encoded, manifest = snapshot(service, text)
    service.sync(service.document_id, 1, encoded, manifest)
    reference = service.reference(service.document_id, 1, {'node_id':'node-1', 'start':0, 'end':len(text)})['reference']
    answer = '平均延迟为120 ms（毫秒），依据当前引用[1]。'
    connections = FixedAnswer({'content':answer, 'usage_metadata':{'input_tokens':84, 'output_tokens':18, 'total_tokens':102}})
    attach(floating, connections)
    original = source.read_bytes()
    before = service.data
    send(floating, 'selection-once', '平均延迟是多少，单位是什么？', 1, reference['id'])
    finish(floating)
    row = floating.conversation.turns['selection-once']
    metrics = floating.conversation.turn_metrics['selection-once'].snapshot()
    assert row['answer'] == answer and row['status'] == 'complete'
    assert row['references'] == [{'scope':'仅当前引用', 'draft_revision':1}]
    assert '非目标内容保持' not in connections.calls[0]
    assert text in connections.calls[0] and len(connections.calls) == 1
    assert metrics['counts']['decisions'] == metrics['counts']['model_attempts'] == 1
    assert metrics['counts']['tool_requests'] == metrics['counts']['tool_executions'] == 0
    assert metrics['usage']['server']['input_tokens'] == 84
    assert metrics['marks_seconds']['first_effective_answer'] <= metrics['marks_seconds']['completed']
    assert {'context', 'model', 'validation'} <= {item['stage'] for item in metrics['spans']}
    assert not any(item['active'] for item in metrics['spans'])
    public = json.dumps(metrics, ensure_ascii=False)
    assert text not in public and '120 ms' not in public and source.name not in public
    assert 'turn_metrics' not in floating.conversation.state()
    assert not service.proposals and not floating.outputs and service.data == before
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))


@pytest.mark.parametrize('name,args', [
    ('read_excel_region', {'sheet':'不相关工作表', 'start':'A1', 'end':'B2'}),
    ('write_word', {'text':'不应写入的内容'}),
])
def test_readonly_answer_rejects_unoffered_tools_without_second_model_request(readonly, name, args):
    service, floating, source = readonly
    connections = FixedAnswer({'content':'尚未执行', 'tool_calls':[{'name':name, 'args':args, 'id':'unoffered-call'}]})
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'unoffered-tool', '平均延迟是多少，单位是什么？')
    finish(floating)
    row = floating.conversation.turns['unoffered-tool']
    metrics = floating.conversation.turn_metrics['unoffered-tool'].snapshot()
    assert row['status'] == 'failed' and not row['answer']
    assert len(connections.calls) == 1 and metrics['counts']['model_attempts'] == 1
    assert metrics['counts']['tool_executions'] == 0
    assert metrics['marks_seconds']['first_effective_answer'] is None
    assert not service.proposals and not floating.outputs and floating.runtime is None
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))


def test_host_status_and_stop_need_no_model_or_document_tool(readonly, monkeypatch):
    _, floating, source = readonly
    connections = Answers()
    attach(floating, connections)
    monkeypatch.setattr(connections, 'capture', lambda *_: pytest.fail('Host control must not capture a model'))
    original = source.read_bytes()
    status = send(floating, 'host-status', '到哪了？')
    stopped = send(floating, 'host-stop', '停止')
    assert status['turn']['status'] == stopped['turn']['status'] == 'complete'
    for identifier in ('host-status', 'host-stop'):
        metrics = floating.conversation.turn_metrics[identifier].snapshot()
        assert metrics['counts']['model_attempts'] == metrics['counts']['tool_executions'] == 0
        assert metrics['marks_seconds']['completed'] is not None
    assert not connections.calls and not floating.outputs and source.read_bytes() == original


def test_explicit_answer_retry_keeps_separate_complete_anonymous_attempts(readonly, monkeypatch):
    from office_agent.turn_metrics import TurnMetrics
    _, floating, source = readonly
    ticks = iter(range(100000))
    monkeypatch.setattr('office_agent.webui.document_conversation.TurnMetrics',
                        lambda: TurnMetrics(clock=lambda: next(ticks) / 100))
    class Recovering(Answers):
        def build_model(self, route):
            owner = self
            class Model:
                def bind_tools(self, tools):
                    assert tools == []
                    return self
                def invoke(self, messages):
                    owner.calls.append(str(messages))
                    if len(owner.calls) == 1:
                        raise TimeoutError('simulated first request failure')
                    return {'content':'平均延迟为120 ms，依据当前材料[1]。'}
            return Model()
    connections = Recovering()
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'retry-metrics', '平均延迟是多少，单位是什么？')
    finish(floating)
    first = floating.conversation.metrics_state('retry-metrics')
    assert floating.conversation.turns['retry-metrics']['status'] == 'failed'
    assert first['attempt_count'] == first['counts']['model_attempts'] == 1
    assert first['current']['marks_seconds']['first_effective_answer'] is None
    assert FloatingApi.conversation_retry(floating.api, 'retry-metrics', floating.document['id'], floating.document['version'])['ok']
    finish(floating)
    result = floating.conversation.metrics_state('retry-metrics')
    current, previous = result['current'], result['previous'][0]
    assert floating.conversation.turns['retry-metrics']['status'] == 'complete'
    assert result['attempt_count'] == result['counts']['model_attempts'] == len(connections.calls) == 2
    assert current['counts']['model_attempts'] == previous['counts']['model_attempts'] == 1
    assert previous['turn_id'] == first['current']['turn_id'] != current['turn_id']
    assert previous['marks_seconds']['completed'] is not None
    assert current['marks_seconds']['first_effective_answer'] <= current['marks_seconds']['completed']
    assert result['omitted_previous_attempts'] == 0
    assert all(not item['active'] and item['start_seconds'] + item['duration_seconds'] <= current['elapsed_seconds'] + 1e-9
               for item in current['spans'])
    assert all(item['duration_seconds'] > 0 for item in current['spans'] if item['stage'] in {'model', 'validation'})
    assert result['attempt_elapsed_seconds_sum'] == previous['elapsed_seconds'] + current['elapsed_seconds']
    serialized = json.dumps(result, ensure_ascii=False)
    assert source.name not in serialized and '120 ms' not in serialized and '平均延迟' not in serialized
    assert not floating.conversation.answer_queues and not floating.outputs and source.read_bytes() == original
    floating.conversation.reset()
    assert not floating.conversation.turn_metrics and not floating.conversation.metric_history and not floating.conversation.metric_totals
