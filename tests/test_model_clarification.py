"""Synthetic model/control integration; no provider, Office or real account."""
from copy import deepcopy
import json
from types import SimpleNamespace
import pytest

from office_agent.agent import OfficeAgent
from office_agent.clarification import (
    MODEL_CLARIFICATION_SCHEMA, ModelClarification, model_clarification_card,
)
from office_agent.webui import bridge
from office_agent.writes import Workspace


def payload():
    return {"preamble": "生成计划前先确定意图", "questions": [
        {"id": "scope", "question": "B2 的替换规则应如何理解？"}]}


def response():
    return {"tool_calls": [{"name": "clarify", "args": payload(), "id": "model-1"},
                           {"name": "unsafe_write", "args": {}}], "content": ""}


class Model:
    def __init__(self, replies, before=None):
        self.replies = iter(replies)
        self.schemas = []
        self.messages = []
        self.before = before
    def bind_tools(self, schemas):
        self.schemas.append(deepcopy(schemas))
        return self
    def invoke(self, messages):
        self.messages.append(deepcopy(messages))
        if self.before:
            callback, self.before = self.before, None
            callback()
        return next(self.replies)


def names(schemas):
    return {s['function']['name'] for s in schemas}


def test_model_card_adds_visible_canonical_questions_and_defaults():
    raw = payload()
    raw['preamble'] = '<script>literal</script>'
    result = model_clarification_card(raw, 'diagram')
    assert len(result['questions']) == 3 and result['source'] == 'model'
    assert result['defaults']['goal'] == 'preview'
    assert result['defaults']['diagram_kind'] == 'flowchart'
    assert result['questions'][1]['question'] == raw['questions'][0]['question']
    assert '回答不是写入、生成或导出批准' in result['preamble']
    raw['questions'][0]['question'] = 'mutated after validation'
    assert result['questions'][1]['question'] != raw['questions'][0]['question']


@pytest.mark.parametrize('raw', [
    {}, {'preamble': '', 'questions': []},
    {'preamble': 'x' * 201, 'questions': payload()['questions']},
    {'preamble': 'nul\x00', 'questions': payload()['questions']},
    {**payload(), 'approved': True}, {**payload(), 'skip_label': '跳过并批准'},
    {**payload(), 'questions': []}, {**payload(), 'questions': payload()['questions'] * 4},
    {**payload(), 'questions': payload()['questions'] * 2},
    {**payload(), 'questions': [{'id': 'unknown', 'question': 'x'}]},
    {**payload(), 'questions': [{'id': [], 'question': 'x'}]},
    {**payload(), 'questions': [{'id': 'diagram_kind', 'question': 'x'}]},
    {**payload(), 'questions': [{'id': 'scope', 'question': 'x', 'approved': True}]},
    {**payload(), 'questions': [{'id': 'goal', 'question': 'x', 'default': 'plan'}]},
    {**payload(), 'questions': [{'id': 'scope', 'question': 'x', 'default': 'B2'}]},
    {**payload(), 'questions': [{'id': 'scope', 'question': 'x', 'options': [{'value': 'write', 'label': 'approved'}]}]},
    {**payload(), 'questions': [{'id': 'scope', 'question': 'x', 'allow_free_text': False}]},
])
def test_untrusted_model_card_fails_closed(raw):
    with pytest.raises(ValueError):
        model_clarification_card(raw, 'write')


def test_core_stops_batch_before_dispatch_and_preserves_status_contract(tmp_path):
    model = Model([SimpleNamespace(**response())])
    calls = []
    agent = OfficeAgent(model, lambda *args: calls.append(args), [], tmp_path,
                        persist=False, clarification_family='write')
    state = agent.new_state('把 B2 的原值 A 改为 B')
    overhead = agent.fixed_overhead()
    with pytest.raises(ModelClarification) as pause:
        agent.run(state)
    assert pause.value.card['source'] == 'model'
    assert names(model.schemas[0]) == {'clarify'}
    assert calls == [] and state.decisions == 1 and state.status == 'running'
    assert state.pending_plan is None and not state.write_tools_requested
    assert state.required.evidence == [] and state.required.errors == []
    assert agent.fixed_overhead() < overhead  # control schema was budgeted, now removed
    assert 'clarify' not in names(agent._tool_schemas())


def test_control_schema_copies_and_resume_cannot_reoffer(tmp_path):
    model = Model([{'content': 'safe completion', 'tool_calls': []}])
    agent = OfficeAgent(model, lambda *args: pytest.fail('dispatch'), [], tmp_path,
                        persist=False, clarification_family='image')
    visible = agent._tool_schemas()
    visible[0]['function']['name'] = 'mutated'
    assert MODEL_CLARIFICATION_SCHEMA['function']['name'] == 'clarify'
    state = agent.new_state('synthetic checkpoint')
    state.decisions = 1
    agent.run(state)
    assert 'clarify' not in names(model.schemas[0])


@pytest.fixture
def controller(tmp_path, monkeypatch):
    path = tmp_path / 'synthetic.txt'
    path.write_text('synthetic input', encoding='utf-8')
    app = bridge.WebviewApp()
    app.path = path
    app.editor = SimpleNamespace(space=Workspace(tmp_path), audit=tmp_path / 'audit')
    starts, dispatches = [], []
    class Thread:
        def __init__(self, *, target, args, daemon): self.args = args
        def start(self): starts.append(self.args)
    monkeypatch.setattr(bridge.threading, 'Thread', Thread)
    monkeypatch.setattr(bridge._app, 'build_tools', lambda approved, *, ocr_enabled=False: [])
    monkeypatch.setattr(bridge._app, 'select_tools', lambda tools, kind: tools)
    monkeypatch.setattr(bridge._app, 'classify_document', lambda path: SimpleNamespace(kind='word'))
    monkeypatch.setattr(bridge._app, 'build_dispatcher', lambda *args, **kw: lambda *args: dispatches.append(args))
    return app, starts, dispatches


def run_initial(controller, monkeypatch, *, before=None, question='把 B2 的原值 A 改为 B'):
    app, starts, dispatches = controller
    model = Model([response(), {'content': '合成建议', 'tool_calls': []}], before=before)
    monkeypatch.setattr(app, '_build_model', lambda: model)
    assert app.ask(question) == {'ok': True}
    assert len(starts) == 1
    app._run_task(*starts[0])
    return model, app.poll()


def test_actual_model_response_enters_once_only_host_card(controller, monkeypatch):
    app, starts, dispatches = controller
    model, events = run_initial(controller, monkeypatch)
    event = next(e for e in events if e['type'] == 'clarify')
    assert event['source'] == 'model' and len(event['questions']) == 2
    assert names(model.schemas[0]) == {'clarify', 'request_read_tools'}
    assert not app.busy and app.model_agent is None and app.pending_id is None
    assert app._read_authorized_path is None and dispatches == []
    assert not app.execute()['ok']
    event['defaults']['goal'] = 'approved'
    assert app.answer_clarification(event['request_id'], {}, True)['ok']
    assert starts[-1][3:5] == (True, True)
    assert not app.answer_clarification(event['request_id'], {}, True)['ok']
    followup = Model([{'content': '合成只读建议', 'tool_calls': []}])
    monkeypatch.setattr(app, '_build_model', lambda: followup)
    app._run_task(*starts[-1])
    assert 'clarify' not in names(followup.schemas[0])
    assert dispatches == [] and app.pending_id is None and app._read_authorized_path is None
    metadata = [json.loads(p.read_text(encoding='utf-8')) for p in app.editor.audit.glob('sessions/*/session.json')]
    # History shows the user's wording; the clarification payload lives in `goal` only.
    resolved = next(m for m in metadata if '规划前澄清' in m.get('goal', ''))
    assert '规划前澄清' not in resolved['question']
    assert 'defaults_used' in resolved['goal']
    assert '"goal": "preview"' in resolved['goal'] and '"planning_mode": "read_only"' in resolved['goal']
    assert not ({'approved', 'confirmed', 'plan_id'} & set(resolved))


@pytest.mark.parametrize('change', ['logout', 'file_changed', 'file_deleted'])
def test_late_model_response_cannot_bind_new_card(controller, monkeypatch, change):
    app, starts, dispatches = controller
    path = app.path
    callback = app._clear_private_state if change == 'logout' else (lambda: path.write_text('changed', encoding='utf-8')) if change == 'file_changed' else path.unlink
    _, events = run_initial(controller, monkeypatch, before=callback)
    assert not any(e['type'] == 'clarify' for e in events)
    assert app._pending_clarification is None and app.pending_id is None
    assert dispatches == [] and app._read_authorized_path is None


@pytest.mark.parametrize('change', ['file_changed', 'file_deleted'])
def test_changed_file_permanently_consumes_pending_model_card(controller, monkeypatch, change):
    app, starts, dispatches = controller
    _, events = run_initial(controller, monkeypatch)
    request = next(e for e in events if e['type'] == 'clarify')['request_id']
    original = app.path.read_bytes()
    if change == 'file_deleted':
        app.path.unlink()
    else:
        app.path.write_bytes(b'changed')
    assert app.answer_clarification(request, {}, True)['reason'] == 'stale_clarification'
    assert app._pending_clarification is None
    app.path.write_bytes(original)
    assert app.answer_clarification(request, {}, True)['reason'] == 'stale_clarification'
    assert len(starts) == 1 and dispatches == []


def test_read_task_does_not_offer_or_accept_model_control(controller, monkeypatch):
    app, starts, dispatches = controller
    model, events = run_initial(controller, monkeypatch, question='总结文档')
    assert all('clarify' not in names(s) for s in model.schemas)
    assert not any(e['type'] == 'clarify' for e in events)
    assert app._pending_clarification is None and dispatches == []


def test_answered_plan_never_enters_second_clarification(controller, monkeypatch):
    app, starts, dispatches = controller
    _, events = run_initial(controller, monkeypatch)
    request = next(e for e in events if e['type'] == 'clarify')['request_id']
    assert app.answer_clarification(request, {'goal': 'plan', 'scope': 'B2 原值 A 换 B'})['ok']
    assert starts[-1][3:5] == (False, True)
    # The stored/displayed wording stays the user's own; the merged payload is model-only.
    assert starts[-1][5] == '把 B2 的原值 A 改为 B'
    assert '规划前澄清' not in starts[-1][5]
    followup = Model([response(), {'content': '待确认建议', 'tool_calls': []}])
    monkeypatch.setattr(app, '_build_model', lambda: followup)
    app._run_task(*starts[-1])
    assert all('clarify' not in names(s) for s in followup.schemas)
    assert not any(e['type'] == 'clarify' for e in app.poll())
    assert dispatches == [] and app.pending_id is None
