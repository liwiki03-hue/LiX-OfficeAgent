import base64
import io
import json
import os
from pathlib import Path
import threading
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.edit_metrics import for_edit
from office_agent.instruction_runtime import InstructionRuntime
from office_agent.model_connections import ModelRoute
from office_agent.paged_documents import GuardError
from office_agent.writes import digest
from tests.test_editable_copies import readonly
from tests.test_paged_editor import draft


class InlineThread:
    def __init__(self, *, target, args=(), daemon=None):
        self.target, self.args = target, args

    def start(self):
        self.target(*self.args)


class Connections:
    def __init__(self, service, replacement=None):
        self.service = service
        self.replacement = replacement
        self.requests = 0
        self.route = ModelRoute('word-metric-test', 'test', 'simulated', 'agent', 'v1', 'openai',
                                'https://example.invalid/v1', 32768, 'test-only', 0)

    def capture(self, capability):
        return self.route

    def build_model(self, route):
        owner = self
        class Model:
            def bind_tools(self, tools):
                return self

            def invoke(self, messages):
                owner.requests += 1
                if owner.requests == 1:
                    reference = owner.service.references[owner.service.pending_reference_id]
                    return {'content': '', 'tool_calls': [{'name': 'propose_draft_edit', 'id': 'word-proposal',
                        'args': {'reference_id': reference['id'], 'old': reference['old'],
                                 'new': owner.replacement or reference['old'].replace('2025', '2026'), 'reason': '按明确要求调整年份，保留温度与单位。'}}],
                        'usage_metadata': {'input_tokens': 13, 'output_tokens': 8, 'total_tokens': 21}}
                return {'content': '已形成建议，等待采用与副本校验。',
                        'usage_metadata': {'input_tokens': 17, 'output_tokens': 6, 'total_tokens': 23}}
        return Model()


def exported(book):
    buffer = io.BytesIO()
    book.save(buffer)
    manifest = [{'id': f'metric-node-{index}', 'text': paragraph.text} for index, paragraph in enumerate(book.paragraphs)]
    return base64.b64encode(buffer.getvalue()).decode(), manifest


def inline(monkeypatch):
    monkeypatch.setattr('office_agent.webui.paged_editor.threading', SimpleNamespace(Thread=InlineThread, RLock=threading.RLock))


def start_round(draft, monkeypatch, *, replacement=None, followup=None):
    page, floating, source = draft
    inline(monkeypatch)
    encoded, manifest = exported(Document(source))
    page.sync(page.document_id, 0, encoded, manifest, True)
    text = '将年份改为2026年，保留温度与单位。'
    floating.conversation.receive('word-edit', text, page.document_id, 0)
    runtime = InstructionRuntime(page.source_version, persist=floating.persist, notify=floating.notice)
    floating.runtime = floating.app.instruction_runtime = runtime
    runtime.submit('word-edit', text)
    if followup:
        runtime.submit('word-append', followup)
    old = page.manifest[1]['text']
    reference = page.reference(page.document_id, 0, {'node_id': page.manifest[1]['id'], 'start': 0, 'end': len(old)})['reference']
    page.pending_reference_id = reference['id']
    connections = Connections(page, replacement)
    floating.api._connections = connections
    trace = for_edit(floating, 'word-edit')
    assert trace is not None and page.start()
    pending = (row for row in page.proposals.values() if row['status'] == 'pending')
    proposal = next(pending) if replacement is None else next(pending, None)
    return page, floating, source, trace, proposal, connections


def adopt(page, proposal):
    revision = page.revision
    grant = page.prepare_apply(proposal['id'], page.document_id, revision)
    book = Document(io.BytesIO(page.data))
    book.paragraphs[1].text = proposal['new']
    encoded, manifest = exported(book)
    page.changed(page.document_id, revision + 1, grant['token'])
    page.applied(grant['token'], page.document_id, revision, revision + 1)
    page.sync(page.document_id, revision + 1, encoded, manifest)


def mock_render(monkeypatch, calls):
    def render(*args, **kwargs):
        calls.append(True)
        return {'inputs_unchanged': True, 'render': {'status': 'different'}}
    monkeypatch.setattr('office_agent.fidelity.compare_documents', render)


@pytest.mark.parametrize('mode', ['default', 'cancel_during_render', 'stop_after_commit', 'bad_receipt'])
def test_word_pair_publication_uses_host_opt_in_and_preserves_commit_boundary(draft, monkeypatch, mode):
    page, floating, source, _, proposal, _ = start_round(draft, monkeypatch)
    original = digest(source)
    adopt(page, proposal)
    plan = page.plan(page.document_id, page.revision)['plan']
    seen = []
    stops = []
    if mode == 'default':
        monkeypatch.delenv('OFFICE_AGENT_DOCX_PAIR_PUBLISH', raising=False)
    else:
        monkeypatch.setenv('OFFICE_AGENT_DOCX_PAIR_PUBLISH', '1')

    def render(root, backup, candidate, **kwargs):
        seen.append(kwargs)
        result = {'inputs_unchanged': True, 'render': {'status': 'different'},
                  'source_sha256': digest(backup), 'output_sha256': digest(candidate),
                  'batch_conversion': {'used': True}, 'render_environment': {'complete': True},
                  'conversion_variant': 'docx-pair-candidate-v1'}
        if mode == 'cancel_during_render':
            stops.append(floating.stop()['publication_stop'])
            assert kwargs['cancelled']()
            result['render'] = {'status': 'error', 'code': 'worker_cancelled'}
        if mode == 'bad_receipt':
            result['output_sha256'] = original
        return result

    monkeypatch.setattr('office_agent.fidelity.compare_documents', render)
    if mode == 'stop_after_commit':
        link = os.link
        def stop_before_link(candidate, output):
            stops.append(floating.stop()['publication_stop'])
            assert not seen[0]['cancelled']()
            return link(candidate, output)
        monkeypatch.setattr('office_agent.paged_documents.os.link', stop_before_link)
    assert page.publish(plan['id'], page.document_id, page.revision, True)['publishing']
    assert len(seen) == 1 and seen[0].get('batch_conversion', False) == (mode != 'default')
    assert not page.publishing and not floating.app.busy and page._pair_publication is None
    assert digest(source) == original
    job = floating.app.editor.audit / plan['id']
    audit = json.loads((job / 'audit.json').read_text('utf8'))
    assert not (job / 'working.docx').exists()
    if mode in {'cancel_during_render', 'bad_receipt'}:
        assert audit['status'] == 'failed' and not Path(audit['output']).exists()
        assert not floating.outputs and page.published_revision == 0
        assert page.publication_result['status'] == ('stopped' if mode == 'cancel_during_render' else 'failed')
        assert stops == (['cancelling'] if mode == 'cancel_during_render' else [])
    else:
        assert audit['status'] == 'verified' and digest(Path(audit['output'])) == audit['candidate_sha256']
        assert stops == (['committing'] if mode == 'stop_after_commit' else [])
        assert floating.conversation.turns['word-edit']['status'] == 'complete'


def test_word_model_apply_publish_reopen_and_render_share_one_unfinished_attempt(draft, monkeypatch):
    page, floating, source, trace, proposal, connections = start_round(draft, monkeypatch)
    original = digest(source)
    before = trace.collector.snapshot()
    assert connections.requests == 2 and before['counts']['model_attempts'] == 2
    assert before['usage']['server']['input_tokens'] == 30
    assert before['marks_seconds']['completed'] is None
    assert before['marks_seconds']['proposal_ready'] is not None
    assert floating.conversation.turns['word-edit']['preview']['text'] == proposal['new']
    assert proposal['conversation_id'] == 'word-edit'
    adopt(page, proposal)
    assert trace.collector.snapshot()['marks_seconds']['completed'] is None
    render_calls = []
    mock_render(monkeypatch, render_calls)
    plan = page.plan(page.document_id, page.revision)['plan']
    assert page.publish(plan['id'], page.document_id, page.revision, True)['publishing']
    result = trace.collector.snapshot()
    assert result['turn_id'] == before['turn_id'] and result['marks_seconds']['completed'] is not None
    assert {'context', 'model', 'validation', 'tool', 'apply', 'reopen', 'render', 'publication'} <= {row['stage'] for row in result['spans']}
    assert not any(row['active'] for row in result['spans']) and render_calls == [True]
    assert floating.conversation.turns['word-edit']['status'] == 'complete'
    assert Document(floating.outputs[plan['id']]['output']).paragraphs[1].text == proposal['new']
    assert digest(source) == original
    encoded = json.dumps(result, ensure_ascii=False)
    assert source.name not in encoded and proposal['old'] not in encoded and 'test-only' not in encoded
    publication = [payload for kind, payload in list(floating.notifications.queue) if kind == 'draft_published'][-1]
    assert publication['conversation_id'] == 'word-edit' and publication['metric_attempt_id'] == trace.id
    assert 'metrics_trace' not in json.dumps(floating.outputs[plan['id']], default=str)


def test_number_order_conflict_stops_without_model_retry(draft, monkeypatch):
    page, floating, source, trace, proposal, connections = start_round(
        draft, monkeypatch, replacement='温度120 ℃，实验于2026年完成。',
        followup='追加要求：把温度放在句首。')
    assert floating.runtime.revision == 2 and connections.requests == 1
    assert proposal is None and not page.proposals and not floating.outputs
    assert page.snapshot_revision == page.revision == 0 and digest(source) == page.source_version
    assert page.start_blocked == 'number_conflict'
    row = floating.conversation.turns['word-edit']
    assert row['status'] == 'failed' and row['failure_code'] == 'number_conflict' and '调换了受保护数字的顺序' in row['text']
    with pytest.raises(GuardError, match='请先调整'):
        page.start_pending(page.document_id, 0, page.pending_reference_id, 2, floating.runtime.session_id)
    assert connections.requests == 1
    report = trace.collector.snapshot()
    assert report['counts']['tool_failures'] == 1 and report['counts']['model_attempts'] == 1
    assert report['marks_seconds']['first_effective_answer'] is None
    request = '请保留原数字顺序，只将年份改为2026年，其他内容不变。'
    reference_id = page.pending_reference_id
    corrected = Connections(page)
    floating.api._connections = corrected
    floating.conversation.receive('word-correction', request, page.document_id, 0)
    page.submit('word-correction', request, reference_id, page.document_id, page.source_version,
                floating.runtime.session_id, page.inbox_id)
    proposals = [value for value in page.proposals.values() if value['status'] == 'pending']
    assert corrected.requests == 2 and len(proposals) == 1 and not floating.outputs
    assert proposals[0]['instruction_revision'] == 3 and proposals[0]['new'] == '实验于2026年完成，温度120 ℃。'


def test_resending_the_same_number_conflict_does_not_call_model(draft, monkeypatch):
    repeated = '追加要求：把温度放在句首。'
    page, floating, source, _, _, connections = start_round(
        draft, monkeypatch, replacement='温度120 ℃，实验于2026年完成。', followup=repeated)
    reference_id = page.pending_reference_id
    floating.conversation.receive('word-repeated', repeated, page.document_id, 0)
    page.submit('word-repeated', repeated, reference_id, page.document_id, page.source_version,
                floating.runtime.session_id, page.inbox_id)
    assert connections.requests == 1 and page.start_blocked == 'number_conflict'
    assert floating.conversation.turns['word-repeated']['failure_code'] == 'number_conflict'
    assert not page.proposals and not floating.outputs and digest(source) == page.source_version


def test_manual_change_and_publish_do_not_attach_to_previous_ai_attempt(draft, monkeypatch):
    page, floating, _, trace, proposal, _ = start_round(draft, monkeypatch)
    adopt(page, proposal)
    book = Document(io.BytesIO(page.data))
    book.paragraphs[2].text = '用户随后手工编辑的内容。'
    encoded, manifest = exported(book)
    page.changed(page.document_id, 2)
    page.sync(page.document_id, 2, encoded, manifest)
    render_calls = []
    mock_render(monkeypatch, render_calls)
    plan = page.plan(page.document_id, 2)['plan']
    page.publish(plan['id'], page.document_id, 2, True)
    assert page.publisher.metrics_trace is None
    assert trace.collector.snapshot()['marks_seconds']['completed'] is None
    assert not any(row['stage'] in {'reopen', 'render', 'publication'} for row in trace.collector.snapshot()['spans'])
    notice = [payload for kind, payload in list(floating.notifications.queue) if kind == 'draft_published'][-1]
    assert notice['conversation_id'] is None and 'metric_attempt_id' not in notice
    assert floating.conversation.turns['word-edit']['status'] != 'complete'


@pytest.mark.parametrize('action,outcome', [('manual', 'stale'), ('stop', 'cancelled'), ('close', 'stale'), ('fresh', 'stale'), ('expired', 'rejected')])
def test_interrupted_apply_span_is_closed_without_marking_proposal_complete(draft, monkeypatch, action, outcome):
    page, _, _, trace, proposal, _ = start_round(draft, monkeypatch)
    grant = page.prepare_apply(proposal['id'], page.document_id, 0)
    if action == 'manual':
        page.changed(page.document_id, 1)
    elif action == 'stop':
        page.stop()
    elif action == 'close':
        page.close()
    elif action == 'fresh':
        page.start_again(True)
    else:
        page.applying[grant['token']]['deadline'] = 0
        with pytest.raises(GuardError):
            page.applied(grant['token'], page.document_id, 0, 1)
    spans = [row for row in trace.collector.snapshot()['spans'] if row['stage'] == 'apply']
    assert len(spans) == 1 and not spans[0]['active'] and spans[0]['outcome'] == outcome
    assert not page._metric_apply
    assert trace.collector.snapshot()['marks_seconds']['completed'] is None


def test_editable_copy_preparation_and_publish_remain_in_same_pending_edit_attempt(readonly, monkeypatch):
    page, floating, source = readonly
    inline(monkeypatch)
    text = '修改最后一段，保留全部数值与单位。'
    floating.conversation.receive('copy-edit', text, page.document_id, 0)
    runtime = InstructionRuntime(page.source_version, persist=floating.persist, notify=floating.notice)
    floating.runtime = floating.app.instruction_runtime = runtime
    runtime.submit('copy-edit', text)
    floating.access.for_instruction()
    trace = for_edit(floating, 'copy-edit')
    original = digest(source)
    render_calls = []
    mock_render(monkeypatch, render_calls)
    result = page.automatic_editable_copy(page.document_id, 0, page.inbox_id, runtime.session_id,
                                         runtime.revision, floating.access.state()['id'])
    assert result['preparing'] and page.editable_copy_id is not None
    snapshot = trace.collector.snapshot()
    assert snapshot['counts']['model_attempts'] == 0 and snapshot['marks_seconds']['completed'] is None
    assert {'context', 'validation', 'reopen', 'render', 'publication'} <= {span['stage'] for span in snapshot['spans']}
    assert digest(source) == original and render_calls == [True]
    notice = [payload for kind, payload in list(floating.notifications.queue) if kind == 'draft_published'][-1]
    assert notice['conversation_id'] == 'copy-edit' and notice['metric_attempt_id'] == trace.id
    assert floating.conversation.turns['copy-edit']['stage'] == 'copy_verified'
