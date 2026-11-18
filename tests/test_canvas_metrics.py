import json
from pathlib import Path
import threading

import pytest

from tests.test_canvas_ai import finish, prepare, send
from tests.test_office_canvas import canvas, operation, wait_finished


def begin(service, floating, reference, client='canvas-metrics'):
    floating.conversation.receive(client, '修改当前选定目标，保留其他内容', service.document_id, service.revision, reference)
    reply = send(service, reference, client)
    return client, reply


def snapshot(floating, client):
    return floating.conversation.metrics_state(client)['current']


def adopted(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    client, reply = begin(service, floating, reference)
    finish(service.ai)
    proposal = next(iter(service.ai.proposals.values()))
    service.ai.adopt(service.document_id, service.revision, proposal['id'])
    return service, floating, source, reference, connections, client, reply


def successful_render(*args, **kwargs):
    return {'inputs_unchanged': True, 'render': {'status': 'different'}}


def test_canvas_turn_spans_continue_until_verified_publication(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    original = source.read_bytes()
    client, reply = begin(service, floating, reference)
    assert reply['started']
    finish(service.ai)
    before = snapshot(floating, client)
    assert connections.calls == 2 and before['counts']['model_attempts'] == 2
    assert before['marks_seconds']['completed'] is None
    assert before['marks_seconds']['first_effective_answer'] is None
    assert floating.conversation.turns[client]['status'] == 'running'
    proposal = next(iter(service.ai.proposals.values()))
    service.ai.adopt(service.document_id, service.revision, proposal['id'])
    assert snapshot(floating, client)['marks_seconds']['completed'] is None
    plan = service.plan(service.document_id, service.revision)['plan']
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', successful_render)
    service.publish(service.document_id, service.revision, plan['id'], True)
    wait_finished(service)
    after = snapshot(floating, client)
    assert floating.conversation.turns[client]['status'] == 'complete'
    assert after['marks_seconds']['completed'] is not None
    assert after['marks_seconds']['first_effective_answer'] is not None
    assert after['counts']['model_attempts'] == 2 and connections.calls == 2
    assert all(not row['active'] for row in after['spans'])
    stages = {row['stage'] for row in after['spans']}
    assert {'queue', 'context', 'model', 'validation', 'apply', 'authorization', 'reopen', 'render', 'publication'} <= stages
    assert sum(row['stage'] == 'reopen' for row in after['spans']) == 2
    assert Path(plan['output']).read_bytes() == service.data and source.read_bytes() == original
    serialized = json.dumps(after, ensure_ascii=False)
    assert str(source) not in serialized and source.name not in serialized
    if not reference['old'].isdecimal():
        assert reference['old'] not in serialized
    assert connections.route.model_id not in serialized and 'old' not in after and 'reason' not in after


def test_manual_canvas_publish_does_not_finish_active_conversation(canvas, monkeypatch):
    service, floating, source, result, change = canvas
    service.apply(service.document_id, 0, change)
    client = 'manual-must-not-bind'
    floating.conversation.receive(client, '修改当前内容', service.document_id, service.revision, None)
    plan = service.plan(service.document_id, service.revision)['plan']
    events = []
    original_notice = floating.notice

    def notice(kind, payload):
        events.append((kind, payload))
        original_notice(kind, payload)

    monkeypatch.setattr(floating, 'notice', notice)
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', successful_render)
    service.publish(service.document_id, service.revision, plan['id'], True)
    wait_finished(service)
    assert floating.conversation.turns[client]['status'] != 'complete'
    published = next(payload for kind, payload in events if kind == 'draft_published')
    assert published['conversation_id'] is None and published['metric_attempt_id'] is None
    assert Path(plan['output']).is_file()


def test_manual_change_discards_late_model_metrics_binding(canvas, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    service, floating, source, reference, connections = prepare(canvas, monkeypatch, entered, release)
    original = source.read_bytes()
    try:
        client, _ = begin(service, floating, reference)
        assert entered.wait(5)
        numeric = source.suffix == '.xlsx'
        change = operation(reference['target'], '127' if numeric else '手工更新', reference['old'], reference['item_id'], 'number' if numeric else 'text')
        service.apply(service.document_id, service.revision, change)
        data = service.data
    finally:
        release.set()
        finish(service.ai)
    measured = snapshot(floating, client)
    assert connections.calls == 1 and not service.ai.proposals
    assert service.data == data and source.read_bytes() == original
    assert measured['marks_seconds']['first_effective_answer'] is None
    assert measured['marks_seconds']['completed'] is None
    assert all(not row['active'] for row in measured['spans'])
    assert service._metric_draft_binding is None


def test_failed_canvas_render_ends_as_failure_without_success_answer(canvas, monkeypatch):
    service, floating, source, reference, connections, client, _ = adopted(canvas, monkeypatch)
    plan = service.plan(service.document_id, service.revision)['plan']
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', lambda *args, **kwargs: {'inputs_unchanged': True, 'render': {'status': 'error'}})
    service.publish(service.document_id, service.revision, plan['id'], True)
    wait_finished(service)
    measured = snapshot(floating, client)
    assert floating.conversation.turns[client]['status'] == 'failed'
    assert measured['marks_seconds']['first_effective_answer'] is None
    assert measured['marks_seconds']['completed'] is not None
    assert all(not row['active'] for row in measured['spans'])
    assert any(row['stage'] == 'render' and row['outcome'] == 'error' for row in measured['spans'])
    assert not Path(plan['output']).exists() and connections.calls == 2


def test_new_requirement_during_publish_cannot_complete_old_round(canvas, monkeypatch):
    service, floating, source, reference, connections, client, _ = adopted(canvas, monkeypatch)
    plan = service.plan(service.document_id, service.revision)['plan']
    entered, release = threading.Event(), threading.Event()

    def render(*args, **kwargs):
        entered.set()
        assert release.wait(10)
        return successful_render()

    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', render)
    try:
        service.publish(service.document_id, service.revision, plan['id'], True)
        assert entered.wait(5)
        extra, duplicate = floating.conversation.receive('late-requirement', '追加修改要求，保留最新内容', service.document_id, service.revision, None)
        assert not duplicate and extra['parent_id'] == client
        floating.conversation.turns['late-requirement']['instruction_revision'] = None
    finally:
        release.set()
        wait_finished(service)
    measured = snapshot(floating, client)
    assert floating.conversation.turns[client]['status'] != 'complete'
    assert measured['marks_seconds']['completed'] is None
    assert measured['marks_seconds']['first_effective_answer'] is None
    assert all(not row['active'] for row in measured['spans'])
    assert Path(plan['output']).is_file() and connections.calls == 2
    assert floating.conversation.turns['late-requirement']['status'] == 'waiting'
    assert service.ai.runtime.closed


def test_stop_during_canvas_publication_does_not_mark_success(canvas, monkeypatch):
    service, floating, source, reference, connections, client, _ = adopted(canvas, monkeypatch)
    plan = service.plan(service.document_id, service.revision)['plan']
    entered, release = threading.Event(), threading.Event()

    def render(*args, **kwargs):
        entered.set()
        assert release.wait(10)
        return successful_render()

    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', render)
    try:
        service.publish(service.document_id, service.revision, plan['id'], True)
        assert entered.wait(5)
        floating.stop()
    finally:
        release.set()
        wait_finished(service)
    measured = snapshot(floating, client)
    assert floating.conversation.turns[client]['status'] == 'stopped'
    assert measured['marks_seconds']['first_effective_answer'] is None
    assert all(not row['active'] for row in measured['spans'])
    assert service._metric_draft_binding is None
    assert Path(plan['output']).is_file() and connections.calls == 2


def test_undo_redo_does_not_rebind_old_canvas_turn(canvas, monkeypatch):
    service, floating, source, reference, connections, client, _ = adopted(canvas, monkeypatch)
    service.undo(service.document_id, service.revision)
    service.redo(service.document_id, service.revision)
    assert service._metric_draft_binding is None
    plan = service.plan(service.document_id, service.revision)['plan']
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', successful_render)
    service.publish(service.document_id, service.revision, plan['id'], True)
    wait_finished(service)
    measured = snapshot(floating, client)
    assert floating.conversation.turns[client]['status'] != 'complete'
    assert measured['marks_seconds']['first_effective_answer'] is None
    assert measured['marks_seconds']['completed'] is None and connections.calls == 2
    assert Path(plan['output']).is_file()
