from copy import deepcopy
from pathlib import Path
import threading
from types import SimpleNamespace

from docx import Document
import pytest
from pydantic import BaseModel, ConfigDict

from office_agent.agent import OfficeAgent, ToolSpec
from office_agent.instruction_runtime import InstructionRuntime
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from office_agent.writes import WordReplace, digest


class Query(BaseModel):
    model_config = ConfigDict(extra='forbid', strict=True)
    term: str


def session(tmp_path):
    path = tmp_path / '报告.docx'
    doc = Document()
    doc.add_paragraph('实验延迟为20 ms，样本为30组。')
    doc.add_paragraph('第二段应保持原样。')
    doc.save(path)
    app = WebviewApp()
    api = SimpleNamespace(_app=app, _auth=SimpleNamespace(status=lambda: {'authenticated': True}),
                          _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    service = FloatingSession(api)
    assert service.admit(str(path))['ok']
    assert service.read(service.read_token, True)['ok']
    return service, app, path


def test_inflight_old_model_action_is_not_dispatched_and_latest_survives_compression(tmp_path):
    entered, release = threading.Event(), threading.Event()
    calls, inputs, events = [], [], []
    runtime = InstructionRuntime('v1', notify=lambda kind, data: events.append((kind, data)))
    runtime.submit('first', '精简第一段，保留数字')
    class Model:
        def bind_tools(self, tools):
            return self
        def invoke(self, messages):
            inputs.append(deepcopy(messages))
            if len(inputs) == 1:
                entered.set()
                assert release.wait(5)
                return {'tool_calls': [{'name': 'lookup', 'args': {'term': 'old'}, 'id': 'old'}]}
            return {'content': '已按最新要求形成建议'}
    agent = OfficeAgent(Model(), lambda name, args: calls.append(args),
                        [ToolSpec(name='lookup', description='lookup', schema_model=Query)], tmp_path,
                        runtime_boundary=runtime.boundary, persist=False, decision_limit=4,
                        token_estimator=lambda text: 700 if 'obsolete' in text else 100,
                        context_tokens=1000, reply_tokens=100)
    state = agent.new_state('精简第一段，保留数字', constraints=['源文件不变'])
    state.history.append({'role': 'assistant', 'content': 'obsolete'})
    worker = threading.Thread(target=lambda: agent.run(state))
    worker.start()
    assert entered.wait(5)
    runtime.submit('second', '第二段不动')
    runtime.submit('third', '第一段更口语，保留20 ms和30组')
    release.set()
    worker.join(5)
    assert not worker.is_alive()
    assert calls == [] and state.status == 'done' and state.compressions == 1
    assert runtime.applied == runtime.revision == 3
    assert '第二段不动' in str(inputs[-1]) and '第一段更口语' in str(inputs[-1])
    assert state.required.constraints == ['源文件不变']
    assert any(kind == 'stale_action_blocked' for kind, _ in events)


def test_update_during_read_confirmation_blocks_tool(tmp_path):
    runtime = InstructionRuntime('v1')
    runtime.submit('1', '读取')
    dispatched = []
    class FileQuery(BaseModel):
        model_config = ConfigDict(extra='forbid', strict=True)
        path: str
    source = tmp_path / 'file.txt'
    source.write_text('self-created')
    replies = iter([{'tool_calls': [{'name': 'lookup', 'args': {'path': str(source)}, 'id': '1'}]}, {'content': '停止旧读取'}])
    model = SimpleNamespace(bind_tools=lambda tools: SimpleNamespace(invoke=lambda messages: next(replies)))
    def confirm(plan):
        runtime.submit('2', '先不要读')
        return True
    agent = OfficeAgent(model, lambda *args: dispatched.append(args),
        [ToolSpec(name='lookup', description='read', schema_model=FileQuery, access='read', allowed_paths=(source,))],
        tmp_path, confirm=confirm, runtime_boundary=runtime.boundary, persist=False, decision_limit=3)
    result = agent.run(agent.new_state('读取', file_versions={str(source): digest(source)}))
    assert result.status == 'done' and dispatched == []
    assert not any(row.get('role') == 'tool' for row in result.history)


def test_receipt_dedup_failure_and_confirmation_order(tmp_path):
    service, app, path = session(tmp_path)
    started = []
    app._start_task = lambda *args, **kwargs: started.append(args)
    first = service.submit('', 'client1', '精简第一段')
    repeated = service.submit(first['session_id'], 'client1', '精简第一段')
    assert repeated['duplicate'] and len(started) == 1
    runtime = service.runtime
    runtime.request_revision = runtime.applied = runtime.revision
    plan = app.editor.plan(WordReplace(path=str(path), paragraph_index=0, old='实验', new='试验'))
    assert runtime.offer(plan.public())
    service.submit(runtime.session_id, 'client2', '第二段不动')
    executed = []
    app.execute = lambda: executed.append(True)
    with pytest.raises(ValueError, match='失效'):
        service.confirm(runtime.session_id, plan.id, 1, runtime.document_version, True)
    assert executed == [] and digest(path) == service.document['version']
    runtime.persist = lambda record: (_ for _ in ()).throw(OSError('journal unavailable'))
    with pytest.raises(OSError):
        runtime.submit('client3', '不能虚报收到')
    assert runtime.revision == 2 and 'client3' not in runtime.ids


def test_executing_transaction_receives_update_without_mutating_frozen_request(tmp_path):
    service, app, path = session(tmp_path)
    app._start_task = lambda *args, **kwargs: None
    service.submit('', '1', '实验改为试验')
    runtime = service.runtime
    runtime.applied = runtime.request_revision = 1
    plan = app.editor.plan(WordReplace(path=str(path), paragraph_index=0, old='实验', new='试验'))
    runtime.offer(plan.public())
    app.editor.render_validation = False
    entered, release, done = threading.Event(), threading.Event(), threading.Event()
    emit = app.emit
    def observed_emit(kind, payload):
        emit(kind, payload)
        if kind == 'idle':
            done.set()
    app.emit = observed_emit
    execute = app.editor.execute
    def gated(*args, **kwargs):
        entered.set()
        assert release.wait(5)
        return execute(*args, **kwargs)
    app.editor.execute = gated
    assert service.confirm(runtime.session_id, plan.id, 1, runtime.document_version, True)['ok']
    assert entered.wait(5)
    ack = service.submit(runtime.session_id, '2', '下一轮再调整最后一句')
    assert ack['record']['revision'] == 2 and runtime.applied == 1
    assert plan.request.new == '试验'
    release.set()
    assert done.wait(5)
    snapshot = service.collect()
    assert len(snapshot['outputs']) == 1
    output = snapshot['outputs'][0]
    assert output['status'] == 'verified'
    assert Document(output['output']).paragraphs[0].text.startswith('试验')
    assert digest(path) == service.document['version']
    assert service.select_output(output['id'])['ok']
    assert service.queued_for_next_version == ['下一轮再调整最后一句']
    assert not service.document['ready'] and service.runtime is None


def test_protected_paragraph_plan_source_change_and_read_denial(tmp_path):
    service, app, path = session(tmp_path)
    view = service.read(service.read_token, True)['view']
    locator = view['blocks'][1]['locator_id']
    service.protect(locator, True)
    safe = app.editor.plan(WordReplace(path=str(path), paragraph_index=0, old='实验', new='试验'))
    service.check_plan(safe.public())
    bad = app.editor.plan(WordReplace(path=str(path), paragraph_index=1, old='第二段', new='错误'))
    with pytest.raises(ValueError, match='保护'):
        service.check_plan(bad.public())
    with pytest.raises(ValueError):
        service.read('forged', True)
    assert service.read(service.read_token, False)['denied']
    doc = Document(path)
    doc.add_paragraph('external')
    doc.save(path)
    with pytest.raises(ValueError, match='改变'):
        service.check_source()


@pytest.mark.parametrize('name,content', [('fake.docx', b'executable'), ('shortcut.lnk', b'x'), ('file.exe', b'x')])
def test_drop_type_spoofing_never_admits(tmp_path, name, content):
    service, app, path = session(tmp_path)
    candidate = tmp_path / name
    candidate.write_bytes(content)
    before = app.path
    with pytest.raises(ValueError):
        service.admit(str(candidate))
    assert app.path == before and app._read_authorized_path == path


def test_ambiguous_protection_pauses_before_model_and_explicit_target_resumes(tmp_path):
    service, app, path = session(tmp_path)
    started = []
    app._start_task = lambda *args, **kwargs: started.append(args)
    response = service.submit('', '1', '第二段完全别动，第一段自然一点')
    assert started == [] and service.runtime.needs_clarification
    assert service.runtime.applied == 0
    service.protect(service.view['blocks'][1]['locator_id'], True)
    assert len(started) == 1 and not service.runtime.needs_clarification
    assert service.runtime.protected_targets == [{'document_version': digest(path), 'paragraph_index': 1}]
    assert service.runtime.revision == 2


def test_restart_restores_only_explicit_draft_and_no_permissions(tmp_path):
    service, app, path = session(tmp_path)
    app._start_task = lambda *args, **kwargs: None
    service.submit('', '1', '把实验改为试验')
    new_app = WebviewApp()
    restored = FloatingSession(SimpleNamespace(_app=new_app, _auth=service.api._auth))
    restored.admit(str(path))
    records = restored.restore_drafts()
    assert records['tasks'][0]['messages'] == ['把实验改为试验']
    assert restored.runtime is None and not restored.document['ready']
    assert new_app.pending_id is None and new_app._read_authorized_path is None and not new_app.busy
    service.api._auth.account_key = lambda: 'different-account'
    assert restored.restore_drafts()['tasks'] == []


def test_late_read_after_file_switch_cannot_authorize_or_publish_old_text(tmp_path, monkeypatch):
    service, app, path = session(tmp_path)
    token = service.read_token
    second = tmp_path / '第二份.docx'
    doc = Document()
    doc.add_paragraph('新的文件内容')
    doc.save(second)
    entered, release, done = threading.Event(), threading.Event(), threading.Event()
    from office_agent.webui import floating_session as module
    original = module.read_word_text
    def gated(*args, **kwargs):
        result = original(*args, **kwargs)
        entered.set()
        assert release.wait(5)
        return result
    monkeypatch.setattr(module, 'read_word_text', gated)
    errors = []
    def read():
        try:
            service.read(token, True)
        except ValueError as exc:
            errors.append(str(exc))
        finally:
            done.set()
    worker = threading.Thread(target=read)
    worker.start()
    assert entered.wait(5)
    service.admit(str(second))
    release.set()
    assert done.wait(5)
    worker.join(5)
    assert errors and '归属' in errors[0]
    assert not service.document['ready'] and service.view is None and service.locators == {}
    assert app._read_authorized_path is None


def test_budget_failure_does_not_mark_unsent_requirements_applied(tmp_path):
    runtime = InstructionRuntime('v1')
    runtime.submit('1', '完整保留要求')
    calls = []
    model = SimpleNamespace(bind_tools=lambda tools: SimpleNamespace(invoke=lambda messages: calls.append(messages)))
    agent = OfficeAgent(model, lambda *args: None, [], tmp_path, persist=False,
                        runtime_boundary=runtime.boundary, context_tokens=1000, reply_tokens=100,
                        token_estimator=lambda text: 1000)
    state = agent.run(agent.new_state('完整保留要求'))
    assert state.status == 'budget' and calls == []
    assert runtime.revision == 1 and runtime.applied == runtime.model_calls == 0


def test_expired_plan_and_replayed_confirmation_cannot_start_write(tmp_path, monkeypatch):
    service, app, path = session(tmp_path)
    app._start_task = lambda *args, **kwargs: None
    service.submit('', '1', '实验改为试验')
    runtime = service.runtime
    runtime.applied = runtime.request_revision = 1
    plan = app.editor.plan(WordReplace(path=str(path), paragraph_index=0, old='实验', new='试验'))
    runtime.offer(plan.public())
    runtime.plans[plan.id]['created'] -= 301
    calls = []
    app.execute = lambda: calls.append(True) or {'ok': True}
    with pytest.raises(ValueError, match='失效'):
        service.confirm(runtime.session_id, plan.id, 1, runtime.document_version, True)
    assert calls == [] and digest(path) == runtime.document_version
    runtime.offer(plan.public())
    assert service.confirm(runtime.session_id, plan.id, 1, runtime.document_version, True)['ok']
    with pytest.raises(ValueError, match='失效'):
        service.confirm(runtime.session_id, plan.id, 1, runtime.document_version, True)
    assert calls == [True]


def test_file_admission_cannot_interleave_output_selection(tmp_path):
    service, app, path = session(tmp_path)
    assert service.admission_lock.acquire(blocking=False)
    try:
        with pytest.raises(ValueError, match='正在选择或准入'):
            service.select_output('any')
        with pytest.raises(ValueError, match='正在选择或准入'):
            service.choose()
    finally:
        service.admission_lock.release()
    assert app.path == path and service.document['ready']


def test_retired_worker_notice_cannot_restart_new_task_and_initial_protection_is_bound(tmp_path):
    service, app, path = session(tmp_path)
    started = []
    app._start_task = lambda *args, **kwargs: started.append(args)
    service.protect(service.view['blocks'][1]['locator_id'], True)
    service.submit('', '1', '实验改为试验')
    previous = service.runtime
    assert previous.protected_targets == [{'document_version': digest(path), 'paragraph_index': 1}]
    service.collect()
    service.start_again(True)
    service.submit('', '2', '先核对现在的要求')
    service.collect()
    count = len(started)
    previous.notify('worker_finished', {})
    previous.notify('instruction_applied', {'revision': 99})
    result = service.collect(service.seq)
    assert len(started) == count and result['applied_revision'] == 0
    assert not any(row['event']['type'] in {'worker_finished', 'instruction_applied'} for row in result['events'])


@pytest.mark.parametrize('text', ['修改措辞' * 1000, '📄' * 4000], ids=['chinese', 'emoji'])
def test_maximum_unicode_requirement_fits_durable_restore_limit(tmp_path, text):
    service, app, path = session(tmp_path)
    app._start_task = lambda *args, **kwargs: None
    service.submit('', 'unicode', text)
    rows = service.restore_drafts()['tasks']
    assert rows[0]['messages'] == [text]
    receipts = list((app.editor.audit / 'inbox').glob('*/*/*.json'))
    assert len(receipts) == 1 and receipts[0].stat().st_size <= 20000
    assert service.runtime.applied == service.runtime.model_calls == 0
