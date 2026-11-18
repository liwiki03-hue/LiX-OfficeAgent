import json
from types import SimpleNamespace

import pytest

from office_agent.edit_metrics import EditMetrics, for_edit, measure
from office_agent.editable_copies import EditableCopyPublisher
from office_agent.paged_documents import CandidatePublisher, GuardError
from office_agent.turn_metrics import TurnMetrics
from tests.test_editable_copies import readonly
from tests.test_paged_editor import draft, snapshot


def begin(service, floating, identifier='edit-metrics'):
    row, duplicate = floating.conversation.receive(identifier, '把最后一段改成两句话，保留数值。', floating.document['id'], service.revision)
    assert not duplicate and row['intent'] == 'edit'
    trace = for_edit(floating, identifier)
    assert trace is not None
    return trace


def test_edit_lifecycle_does_not_complete_at_proposal_and_closes_on_publication(draft):
    service, floating, source = draft
    trace = begin(service, floating)
    trace.collector.count('model_attempts', 2)
    token = trace.start('apply')
    floating.conversation.event('draft_proposal', {'conversation_id': trace.conversation_id, 'metric_attempt_id': trace.id,
                                                 'old': 'private-source-text', 'new': 'private-result-text'})
    report = trace.collector.snapshot()
    assert report['marks_seconds']['completed'] is None and trace.current()
    floating.conversation.event('draft_published', {'conversation_id': trace.conversation_id, 'metric_attempt_id': trace.id,
        'receipt': {'id': 'verified-output', 'output': str(source), 'kind': 'paged_candidate', 'changes': []}})
    report = trace.collector.snapshot()
    assert report['counts']['model_attempts'] == 2
    assert report['marks_seconds']['first_effective_answer'] <= report['marks_seconds']['completed']
    assert not trace.current() and not any(span['active'] for span in report['spans'])
    assert 'private-' not in json.dumps(report) and str(source) not in json.dumps(report)
    trace.stop(token)


def test_explicit_unbound_or_wrong_attempt_cannot_complete_an_active_edit(draft):
    service, floating, source = draft
    trace = begin(service, floating)
    receipt = {'id': 'manual-output', 'kind': 'paged_candidate', 'output': str(source), 'changes': []}
    floating.conversation.event('draft_published', {'conversation_id': None, 'receipt': receipt})
    floating.conversation.event('draft_published', {'conversation_id': trace.conversation_id,
        'metric_attempt_id': 'old-attempt', 'receipt': receipt})
    row = floating.conversation.turns[trace.conversation_id]
    assert row['status'] == 'running' and not row['results']
    assert trace.collector.snapshot()['marks_seconds']['completed'] is None


def test_child_requirement_shares_the_parent_collector_without_double_counting(draft):
    service, floating, _ = draft
    trace = begin(service, floating)
    child, _ = floating.conversation.receive('extra-metrics', '保留全部数字和单位', floating.document['id'], service.revision)
    assert child['parent_id'] == trace.conversation_id
    selected = for_edit(floating, child['id'])
    assert selected is trace and selected.conversation_id == 'edit-metrics'
    selected.collector.count('model_attempts')
    assert floating.conversation.metrics_state('edit-metrics')['counts']['model_attempts'] == 1
    assert floating.conversation.metrics_state(child['id'])['counts']['model_attempts'] == 0


def test_edit_retry_uses_a_new_attempt_and_rejects_old_late_events(draft):
    service, floating, source = draft
    first = begin(service, floating)
    first.collector.count('model_attempts')
    floating.conversation.update(first.conversation_id, 'failed', '本次未完成', status='failed')
    floating.conversation.update(first.conversation_id, 'retrying', '主动重试')
    second = for_edit(floating, first.conversation_id)
    assert second is not first and not first.current()
    second.collector.count('model_attempts')
    floating.conversation.event('draft_published', {'conversation_id': first.conversation_id, 'metric_attempt_id': first.id,
        'receipt': {'id': 'late', 'kind': 'paged_candidate', 'output': str(source), 'changes': []}})
    assert second.collector.snapshot()['marks_seconds']['completed'] is None
    report = floating.conversation.metrics_state(first.conversation_id)
    assert report['attempt_count'] == 2 and report['counts']['model_attempts'] == 2
    assert len(report['previous']) == 1


def test_normal_reset_cancels_edit_spans_and_detaches_the_trace(draft):
    service, floating, _ = draft
    trace = begin(service, floating)
    trace.start('editor')
    floating.conversation.reset()
    report = trace.collector.snapshot()
    assert report['marks_seconds']['completed'] is not None
    assert report['marks_seconds']['first_effective_answer'] is None
    assert report['outcomes']['cancelled'] == 1 and not trace.current()


def test_copy_reopen_preserves_the_live_collector_across_document_rebind(readonly, monkeypatch):
    service, floating, source = readonly
    trace = begin(service, floating)
    publisher = EditableCopyPublisher(floating.app.editor, source, service.source_version)
    publisher.metrics_trace = trace
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **k: {
        'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = publisher.plan_copy(document_id=service.document_id, draft_revision=0)
    receipt = publisher.execute(plan['id'], document_id=service.document_id, draft_revision=0, confirmed=True)
    receipt.update(conversation_id=trace.conversation_id, source_version=service.source_version)
    floating.outputs[receipt['id']] = receipt
    before = floating.document['id']
    assert floating.select_output(receipt['id'], True)['ok']
    assert floating.document['id'] != before
    assert for_edit(floating, trace.conversation_id) is trace and trace.current()
    assert trace.collector.snapshot()['marks_seconds']['completed'] is None
    assert any(span['stage'] == 'reopen' for span in trace.collector.snapshot()['spans'])
    assert 'metrics_trace' not in json.dumps(floating.conversation.state())


def test_candidate_publication_records_file_stages_without_serializing_metrics(draft, monkeypatch):
    service, floating, source = draft
    trace = begin(service, floating)
    candidate, _ = snapshot(service)
    import base64
    data = base64.b64decode(candidate)
    publisher = CandidatePublisher(floating.app.editor, source, service.source_version)
    publisher.metrics_trace = trace
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **k: {
        'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = publisher.plan(data, document_id=service.document_id, draft_revision=1)
    receipt = publisher.execute(plan['id'], document_id=service.document_id, draft_revision=1, confirmed=True)
    report = trace.collector.snapshot()
    assert {'validation', 'publication', 'reopen', 'render'} <= {span['stage'] for span in report['spans']}
    assert not any(span['active'] for span in report['spans'])
    assert report['marks_seconds']['completed'] is None
    log = (floating.app.editor.audit / receipt['id'] / 'audit.json').read_text('utf-8')
    assert 'metrics_trace' not in log and trace.id not in log


def test_renderer_rejection_keeps_the_failed_render_span_and_publishes_nothing(draft, monkeypatch):
    service, floating, source = draft
    trace = begin(service, floating)
    candidate, _ = snapshot(service)
    import base64
    publisher = CandidatePublisher(floating.app.editor, source, service.source_version)
    publisher.metrics_trace = trace
    plan = publisher.plan(base64.b64decode(candidate), document_id=service.document_id, draft_revision=1)
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **k: {
        'inputs_unchanged': True, 'render': {'status': 'error'}})
    with pytest.raises(GuardError):
        publisher.execute(plan['id'], document_id=service.document_id, draft_revision=1, confirmed=True)
    report = trace.collector.snapshot()
    assert any(span['stage'] == 'render' and span['outcome'] == 'error' for span in report['spans'])
    assert not any(span['active'] for span in report['spans'])
    assert not list(source.parent.glob('*.agent-*.docx'))


def test_metrics_failure_cannot_skip_or_replace_the_measured_operation():
    metrics = TurnMetrics()
    trace = EditMetrics(metrics, 'opaque', lambda: True)
    trace.collector.start = lambda *_: (_ for _ in ()).throw(ValueError('metrics failed'))
    ran = []
    with measure(trace, 'publication'):
        ran.append(True)
    assert ran == [True]
    with pytest.raises(GuardError):
        with measure(None, 'validation'):
            raise GuardError('source changed')
    assert for_edit(SimpleNamespace(conversation=None), 'opaque') is None


def test_finished_collector_rejects_late_span_even_if_the_owner_is_present():
    metrics = TurnMetrics()
    trace = EditMetrics(metrics, 'opaque', lambda: True)
    metrics.mark('completed')
    assert metrics.is_complete() and not trace.current()
    assert trace.start('publication') is None and not metrics.snapshot()['spans']


def test_nested_application_and_validation_are_not_added_as_wall_time():
    clock = [10.0]
    metrics = TurnMetrics(clock=lambda: clock[0])
    clock[0] = 11.0
    apply = metrics.start('apply')
    clock[0] = 13.0
    validation = metrics.start('validation')
    clock[0] = 15.0
    metrics.stop(validation)
    render = metrics.start('render')
    clock[0] = 17.0
    metrics.stop(render)
    clock[0] = 18.0
    metrics.stop(apply)
    clock[0] = 20.0
    metrics.mark('completed')
    report = metrics.snapshot()
    assert report['elapsed_seconds'] == 10
    assert report['observed_wall_seconds'] == 7
    assert report['unattributed_wall_seconds'] == 3
    assert report['cross_stage_overlap_seconds'] == 4
    assert sum(report['stage_wall_seconds'].values()) == 11
