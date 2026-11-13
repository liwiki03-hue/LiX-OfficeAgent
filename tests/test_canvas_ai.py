from copy import deepcopy
from dataclasses import replace
import json
from pathlib import Path
import threading
import time

import pytest

from office_agent.model_connections import ModelRoute
from office_agent.office_canvas import view
from office_agent.webui.canvas_ai import CanvasSuggestion
from office_agent.webui.floating_session import FloatingApi
from office_agent.writes import GuardError
from tests.test_office_canvas import canvas, operation, wait_finished


class Connections:
    def __init__(self, service, new, entered=None, release=None):
        self.service, self.new = service, new
        self.entered, self.release = entered, release
        self.calls, self.contexts = 0, []
        self.route = ModelRoute('local-test', '隔离模拟接口', 'simulated-model', 'agent', 'v1', 'openai', 'https://example.invalid/v1', 8192, 'test-only', 0)

    def capture(self, capability):
        assert capability == 'agent'
        return self.route

    def build_model(self, route):
        owner = self
        reference = deepcopy(self.service.ai.reference)
        class Model:
            def __init__(self): self.calls = 0
            def bind_tools(self, tools):
                assert len(tools) == 1 and tools[0]['function']['name'] == 'propose_canvas_edit'
                return self
            def invoke(self, messages):
                owner.calls += 1
                self.calls += 1
                owner.contexts.append(str(messages))
                if owner.entered and owner.calls == 1:
                    owner.entered.set()
                    assert owner.release.wait(10)
                if not any(isinstance(row, dict) and row.get('role') == 'tool' for row in messages):
                    return {'tool_calls': [{'name': 'propose_canvas_edit', 'id': 'bounded-call', 'args':
                            {'reference_id': reference['id'], 'old': reference['old'], 'new': owner.new, 'reason': '按所选目标明确修改，保留其他内容'}}]}
                return {'content': '建议已形成，尚未采用或保存。'}
        return Model()


def finish(ai):
    end = time.monotonic() + 15
    while True:
        with ai.canvas.lock:
            if ai.running is None: return
        assert time.monotonic() < end
        time.sleep(.01)


def prepare(canvas, monkeypatch, entered=None, release=None):
    service, floating, source, result, op = canvas
    service.apply(service.document_id, 0, op)
    numeric = source.suffix == '.xlsx'
    target = 'B2' if numeric else 'run-0'
    item = 'sheet-0' if numeric else 'slide-0'
    reference = service.ai.select(service.document_id, 1, item, target, numeric)['reference']
    connections = Connections(service, '130' if numeric else reference['old'] + '修订', entered, release)
    floating.api._connections = connections
    return service, floating, source, reference, connections


def send(service, reference, client='first'):
    return service.ai.submit(service.document_id, service.revision, reference['id'],
                             service.ai.runtime.session_id if service.ai.runtime else '', client, '按当前选定目标修改，保留其他内容')


def test_current_draft_proposal_adoption_and_existing_copy_transaction(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    original = source.read_bytes()
    reply = send(service, reference)
    assert reply['started'] and not reply['duplicate']
    finish(service.ai)
    assert connections.calls == 2 and reference['old'] in connections.contexts[0]
    proposal = next(iter(service.ai.proposals.values()))
    assert proposal['status'] == 'pending' and source.read_bytes() == original and service.revision == 1
    service.ai.adopt(service.document_id, 1, proposal['id'])
    assert service.revision == 2 and proposal['status'] == 'adopted'
    with pytest.raises(GuardError): service.ai.adopt(service.document_id, 2, proposal['id'])
    plan = service.plan(service.document_id, 2)['plan']
    assert plan['instruction_binding']['revision'] == 1
    service.publish(service.document_id, 2, plan['id'], False)
    assert not list(source.parent.glob('*.agent-*'))
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', lambda *a, **k: {'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = service.plan(service.document_id, 2)['plan']
    service.publish(service.document_id, 2, plan['id'], True)
    wait_finished(service)
    output = Path(plan['output'])
    assert output.is_file() and output.read_bytes() == service.data and source.read_bytes() == original
    assert floating.outputs[plan['id']]['instruction_revision'] == 1
    audit = json.loads((floating.app.editor.audit / plan['id'] / 'audit.json').read_text('utf8'))
    assert audit['status'] == 'verified' and audit['instruction_binding'] == plan['instruction_binding']
    floating.app.editor.rollback(plan['id'], confirmed=True)
    assert not output.exists() and source.read_bytes() == original


def test_duplicate_and_late_model_cannot_overwrite_manual_draft(canvas, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    service, floating, source, reference, connections = prepare(canvas, monkeypatch, entered, release)
    try:
        send(service, reference)
        assert entered.wait(5)
        assert send(service, reference)['duplicate'] and connections.calls == 1
        numeric = source.suffix == '.xlsx'
        op = operation(reference['target'], '127' if numeric else '用户手改', reference['old'], reference['item_id'], 'number' if numeric else 'text')
        service.apply(service.document_id, 1, op)
        current = service.data
    finally:
        release.set()
        finish(service.ai)
    assert service.data == current and service.revision == 2 and not service.ai.proposals
    assert connections.calls == 1 and not floating.app.busy


def test_config_source_stop_and_protection_invalidate_proposals(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    send(service, reference); finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    original_route = connections.route
    connections.route = replace(connections.route, model_id='another-model')
    with pytest.raises(GuardError, match='配置'): service.ai.adopt(service.document_id, 1, proposal['id'])
    connections.route = original_route
    if source.suffix == '.xlsx':
        floating.protected_cells.add(('实验', 'B2'))
        with pytest.raises(GuardError, match='保护'): service.ai.adopt(service.document_id, 1, proposal['id'])
        floating.protected_cells.clear()
    plan = service.plan(service.document_id, 1)['plan']
    service.ai.stop(service.document_id, 1)
    assert not service.plans and proposal['status'] == 'stale'
    with pytest.raises(GuardError): service.publish(service.document_id, 1, plan['id'], True)
    with pytest.raises(GuardError): service.ai.adopt(service.document_id, 1, proposal['id'])
    source.write_bytes(source.read_bytes() + b'changed')
    with pytest.raises(ValueError): service.ai.select(service.document_id, 1, reference['item_id'], reference['target'], source.suffix == '.xlsx')


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_formula_and_implicit_numeric_change_are_refused(canvas):
    service, floating, source, result, op = canvas
    with pytest.raises(GuardError, match='数值'): service.ai.select(service.document_id, 0, 'sheet-0', 'B2')
    with pytest.raises(GuardError, match='公式'): service.ai.select(service.document_id, 0, 'sheet-0', 'C2', True)
    with pytest.raises(GuardError, match='仅适用'): service.ai.select(service.document_id, 0, 'sheet-0', 'A2', True)
    with pytest.raises(GuardError): service.ai.select(service.document_id, 0, 'sheet-0', 'B4')
    with pytest.raises(Exception): CanvasSuggestion.model_validate({'reference_id': 'x', 'old': 'x', 'new': 'y', 'reason': 'r', 'path': 'x'})


def test_new_requirement_invalidates_old_plan_and_proposal(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    send(service, reference); finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    plan = service.plan(service.document_id, 1)['plan']
    row = service.ai.runtime.messages[0]
    service.ai.change(service.document_id, 1, service.ai.runtime.session_id, 'cancel', row['id'], 1, 1, 'cancel', '')
    assert proposal['status'] == 'stale' and not service.plans and connections.calls == 2
    with pytest.raises(GuardError): service.ai.adopt(service.document_id, 1, proposal['id'])
    with pytest.raises(GuardError): service.publish(service.document_id, 1, plan['id'], True)


def test_privacy_reset_does_not_restore_old_session(canvas, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    service, floating, source, reference, connections = prepare(canvas, monkeypatch, entered, release)
    old_ai = service.ai
    try:
        send(service, reference); assert entered.wait(5)
        floating.reset()
    finally:
        release.set(); finish(old_ai)
    assert not service.active() and not service.ai.proposals and service.ai.runtime is None
    assert not floating.app.busy and floating.runtime is None and connections.calls == 1


def test_appended_requirement_starts_once_at_next_safe_boundary(canvas, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    service, floating, source, reference, connections = prepare(canvas, monkeypatch, entered, release)
    try:
        send(service, reference); assert entered.wait(5)
        update = service.ai.submit(service.document_id, 1, reference['id'], service.ai.runtime.session_id,
                                   'follow-up', '另加要求：保留未选位置和所有单位')
        assert not update['started'] and connections.calls == 1
    finally:
        release.set(); finish(service.ai)
    pending = [row for row in service.ai.proposals.values() if row['status'] == 'pending']
    assert connections.calls == 3 and len(pending) == 1
    assert pending[0]['instruction_revision'] == 2 and '保留未选位置' in connections.contexts[-1]
    assert service.revision == 1 and not floating.app.busy


def test_expired_proposal_and_changed_model_plan_do_not_write(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    send(service, reference); finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    proposal['expires'] = time.monotonic() - 1
    with pytest.raises(GuardError, match='过期'): service.ai.adopt(service.document_id, 1, proposal['id'])
    proposal['expires'] = time.monotonic() + 300
    service.ai.adopt(service.document_id, 1, proposal['id'])
    plan = service.plan(service.document_id, 2)['plan']
    connections.route = replace(connections.route, revision='v2')
    with pytest.raises(GuardError, match='配置'): service.publish(service.document_id, 2, plan['id'], True)
    assert not list(source.parent.glob('*.agent-*'))


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_text_numeric_protection_is_checked_before_proposal(canvas, monkeypatch):
    service, floating, source, result, op = canvas
    reference = service.ai.select(service.document_id, 0, 'sheet-0', 'A2')['reference']
    connections = Connections(service, '121 ℃')
    floating.api._connections = connections
    send(service, reference); finish(service.ai)
    assert not service.ai.proposals and service.revision == 0
    assert service.ai.error and connections.calls <= 8


def test_stop_while_model_waits_does_not_retry(canvas, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    service, floating, source, reference, connections = prepare(canvas, monkeypatch, entered, release)
    try:
        send(service, reference); assert entered.wait(5)
        result = service.ai.stop(service.document_id, 1)
        assert result['remote_cancelled'] is False
    finally:
        release.set(); finish(service.ai)
    assert connections.calls == 1 and not service.ai.proposals and service.revision == 1


def test_registered_host_api_and_global_stop_invalidate_plan(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    api = FloatingApi.__new__(FloatingApi)
    api._floating = floating
    reply = api.canvas_ai_submit(service.document_id, 1, reference['id'], '', 'host-call', '修改当前目标')
    assert reply['ok']
    finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    assert api.canvas_ai_adopt(service.document_id, 1, proposal['id'])['ok']
    plan = api.canvas_plan(service.document_id, 2)['plan']
    floating.stop()
    assert not service.plans
    assert not api.canvas_publish(service.document_id, 2, plan['id'], True)['ok']


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_same_number_different_unit_is_not_a_safe_wording_edit(canvas):
    service, floating, source, result, op = canvas
    reference = service.ai.select(service.document_id, 0, 'sheet-0', 'A2')['reference']
    with pytest.raises(GuardError, match='单位'): service.ai._operation(reference, '120 ℉')
    assert service.ai._operation(reference, '温度120 ℃')['value'] == '温度120 ℃'


def test_disabled_model_invalidates_old_confirmation_but_preserves_manual_save(canvas, monkeypatch):
    from office_agent.settings_store import SettingsError
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    send(service, reference); finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    service.ai.adopt(service.document_id, 1, proposal['id'])
    old = service.plan(service.document_id, 2)['plan']
    def unavailable(capability): raise SettingsError('model_selection_invalid')
    connections.capture = unavailable
    with pytest.raises(GuardError, match='配置'): service.publish(service.document_id, 2, old['id'], True)
    current = service.plan(service.document_id, 2)['plan']
    assert current['instruction_binding']['model']['unavailable']
    assert not list(source.parent.glob('*.agent-*')) and service.state()['dirty']
