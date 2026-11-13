import base64
from copy import deepcopy
import json
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.paged_documents import GuardError
from office_agent.webui.draft_suggestions import DraftModel, validate_suggestion
from tests.test_paged_editor import draft
from tests.test_paged_inbox import receive


def armed(service, floating, monkeypatch, text='把年份改为2026，保持温度和其他数字不变。'):
    book = Document(service.source)
    service.sync(service.document_id, 0, base64.b64encode(service.data).decode(),
                 [{'id': 'n'+str(i), 'text': p.text} for i, p in enumerate(book.paragraphs)], True)
    receive(service, text=text)
    row = service.reference(service.document_id, 0, {'node_id': 'n1', 'start': 0, 'end': len(book.paragraphs[1].text)})['reference']
    service.pending_reference_id = row['id']
    runtime = floating.runtime
    service.running_runtime = runtime
    floating.app.busy = floating.app._model_task_active = True
    route = SimpleNamespace(context_window=32768, public=lambda: {'connection_id': 'fixture', 'model_id': 'fixture', 'capability': 'agent'})
    restarts = []
    monkeypatch.setattr(service, 'start', lambda: restarts.append(True))
    return row, runtime, route, restarts


class Model:
    def __init__(self, reference, during=None, replacement='2026'):
        self.row = deepcopy(reference)
        self.calls = 0
        self.during = during
        self.replacement = replacement

    def bind_tools(self, tools):
        return self

    def invoke(self, messages):
        self.calls += 1
        if self.during:
            self.during()
        if self.calls == 1:
            return {'tool_calls': [{'name': 'propose_draft_edit', 'args': {'reference_id': self.row['id'],
                'old': self.row['old'], 'new': self.row['old'].replace('2025', self.replacement), 'reason': '按明确要求修改年份。'}, 'id': 'fixture'}]}
        return {'content': '建议待审阅，未写入。'}


def run(service, floating, monkeypatch, model, row, runtime, route):
    monkeypatch.setattr(floating.api._connections, 'build_model', lambda _: model, raising=False)
    service._run(runtime, row, route, 0, service.document_id, 0)


def test_explicit_year_change_reaches_real_proposal_without_changing_source(draft, monkeypatch):
    service, floating, source = draft
    original = source.read_bytes()
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    model = Model(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    proposal = next(iter(service.proposals.values()))
    assert proposal['new'] == '实验于2026年完成，温度120 ℃。'
    assert proposal['status'] == 'pending' and model.calls == 2
    assert proposal['model'] == route.public()
    assert source.read_bytes() == original and service.revision == 0 and not restarts


def test_model_receives_one_submittable_reference_and_bound_schema(draft, monkeypatch):
    service, floating, source = draft
    row, runtime, route, _ = armed(service, floating, monkeypatch)
    original = source.read_bytes()
    class InspectingModel(Model):
        def bind_tools(self, tools):
            parameters = tools[0]['function']['parameters']
            field = parameters['properties']['reference_id']
            assert field['enum'] == [row['id']]
            assert field['pattern'] == '^' + row['id'] + '$'
            assert parameters['additionalProperties'] is False
            return self
        def invoke(self, messages):
            text = str(messages)
            assert row['id'] in text and row['old'] in text
            assert row['draft_locator_id'] not in text
            assert 'source_host_locator_id' not in text and 'node_id' not in text
            return super().invoke(messages)
    model = InspectingModel(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    assert model.calls == 2 and len(service.proposals) == 1
    assert source.read_bytes() == original and service.revision == 0


@pytest.mark.parametrize('identifier', ['draft_locator_id', 'node_id', 'foreign'])
def test_wrong_reference_is_not_silently_repaired(draft, monkeypatch, identifier):
    service, floating, source = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    wrong = deepcopy(row)
    wrong['id'] = row.get(identifier, 'foreign-reference')
    original = source.read_bytes()
    model = Model(wrong)
    run(service, floating, monkeypatch, model, row, runtime, route)
    assert not service.proposals and not restarts
    assert source.read_bytes() == original and service.revision == 0
    assert not list(source.parent.glob('*.agent-*.docx'))


def test_native_acceptance_year_request_reaches_pending_proposal(draft, monkeypatch):
    service, floating, source = draft
    original = source.read_bytes()
    question = '只将选中句中的2025年改为2026年，保留温度120 ℃、标点及其他文字，不修改标题和表格。'
    row, runtime, route, restarts = armed(service, floating, monkeypatch, question)
    model = Model(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    proposal = next(iter(service.proposals.values()))
    assert proposal['status'] == 'pending'
    assert proposal['new'] == '实验于2026年完成，温度120 ℃。'
    assert source.read_bytes() == original and service.revision == 0 and not restarts


@pytest.mark.parametrize('question', [
    '只将选中句中的2025年改为2026年',
    '请仅把所选段落中的年份改为2026年',
    '将这句话中的2025年替换为2026年',
])
def test_scoped_year_command_preserves_other_numbers(question):
    before = '实验于2025年完成，温度120 ℃，耗时3.5秒。'
    allowed = before.replace('2025年', '2026年')
    assert validate_suggestion(before, allowed, '修改年份', [{'text': question}])
    for changed in [allowed.replace('120', '121'), allowed.replace('3.5', '3.6'), allowed.replace('℃', '℉')]:
        with pytest.raises(ValueError):
            validate_suggestion(before, changed, '修改年份', [{'text': question}])


@pytest.mark.parametrize('requirements', [
    ['不要只将选中句中的2025年改为2026年'],
    ['如果合适，把年份改为2026。'],
    ['例如：只将选中句中的2025年改为2026年'],
    ['将第二段中的2025年改为2026年'],
    ['只将选中句中的2025年改为2026年', '不要修改年份'],
    ['只将选中句中的2025年改为2026年；保持原年份不变'],
    ['只将选中句中的2025年改为2026年', '将年份改为2027年'],
])
def test_scoped_year_ambiguity_does_not_authorize_change(requirements):
    with pytest.raises(ValueError):
        validate_suggestion('实验于2025年完成，温度120 ℃。', '实验于2026年完成，温度120 ℃。',
                            '修改年份', [{'text': text} for text in requirements])


def test_year_protection_still_allows_non_numeric_rewording():
    assert validate_suggestion('实验于2025年完成，温度120 ℃。', '实验在2025年完成，温度120 ℃。',
                               '只调整措辞', [{'text': '调整措辞，不要修改年份'}]) == []


def test_source_change_before_send_makes_zero_requests(draft, monkeypatch):
    service, floating, source = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    source.write_bytes(b'changed before request')
    model = Model(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    assert model.calls == 0 and not service.proposals and not restarts


def test_source_change_on_response_never_emits_answer(draft, monkeypatch):
    service, floating, source = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    model = Model(row, lambda: source.write_bytes(b'changed during request'))
    run(service, floating, monkeypatch, model, row, runtime, route)
    assert model.calls == 1 and not service.proposals and not restarts
    assert not any(kind == 'draft_answer' for kind, _ in floating.notifications.queue)


def test_budget_exhaustion_never_schedules_another_worker(draft, monkeypatch):
    service, floating, _ = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    runtime.model_calls = runtime.model_limit
    model = Model(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    assert model.calls == 0 and not restarts
    assert service.state()['pending_reason'] == 'budget'


def test_auth_failure_is_visible_sanitized_and_not_retried(draft, monkeypatch):
    service, floating, _ = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    class AuthenticationError(Exception):
        status_code = 401
    def fail():
        raise AuthenticationError('PRIVATE_SERVICE_BODY')
    model = Model(row, fail)
    run(service, floating, monkeypatch, model, row, runtime, route)
    failures = [payload for kind, payload in floating.notifications.queue if kind == 'draft_model_error']
    assert failures and '鉴权' in failures[-1]['text']
    assert 'PRIVATE_SERVICE_BODY' not in json.dumps(list(floating.notifications.queue))
    assert model.calls == 1 and not restarts and not service.proposals


def test_stopped_batch_grant_cannot_be_acknowledged(draft, monkeypatch):
    service, floating, _ = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    operation = {'document_id': service.document_id, 'session_epoch': floating.epoch, 'source_version': service.source_version,
        'draft_revision': 0, 'instruction_revision': 1, 'target_id': 'n1', 'expected_before': row['old'],
        'operation_id': 'one', 'idempotency_key': 'one', 'action': 'replace_text', 'start': 0, 'end': len(row['old']), 'after': row['old']}
    grant = service.prepare_operations(service.document_id, 0, 1, [operation])
    floating.stop()
    with pytest.raises(GuardError):
        service.applied_operations(grant['token'], service.document_id, 0, 1)
    assert service.revision == 0


def test_batch_cannot_claim_a_fabricated_instruction_version(draft, monkeypatch):
    service, floating, _ = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    operation = {'document_id': service.document_id, 'session_epoch': floating.epoch, 'source_version': service.source_version,
        'draft_revision': 0, 'instruction_revision': 91, 'target_id': 'n1', 'expected_before': row['old'],
        'operation_id': 'one', 'idempotency_key': 'one', 'action': 'replace_text', 'start': 0, 'end': len(row['old']), 'after': row['old']}
    with pytest.raises(GuardError):
        service.prepare_operations(service.document_id, 0, 91, [operation])


@pytest.mark.parametrize('question,new', [
    ('精简文字，保持数字。', '实验于2026年完成，温度120℃。'),
    ('不要把年份改为2026。', '实验于2026年完成，温度120℃。'),
    ('如果合适，把年份改为2026。', '实验于2026年完成，温度121℃。'),
    ('把年份改为2026，保持温度。', '实验于2026年完成，温度120℉。'),
    ('把年份改为2026。', '实验于2026年完成，温度121℃。'),
    ('把年份改为2026。', '实验于2027年完成，温度120℃。'),
    ('将2024改为2026。', '实验于2026年完成，温度120℃。'),
    ('将2025改为2026；把年份改为2027。', '实验于2026年完成，温度120℃。'),
])
def test_year_authorization_never_unlocks_other_literals(question, new):
    with pytest.raises(ValueError):
        validate_suggestion('实验于2025年完成，温度120℃。', new, '修改年份', [{'text': question}])


@pytest.mark.parametrize('before', ['2025年和2024年', '[2025年]', '$2025年$'])
def test_ambiguous_year_citation_and_formula_still_refused(before):
    with pytest.raises(ValueError):
        validate_suggestion(before, before.replace('2025', '2026'), '修改年份', [{'text': '将2025改为2026'}])


@pytest.mark.parametrize('question', ['把年份改为2026，保持温度和数字不变。', '将2025改为2026', '请把2025年修改为2026年'])
def test_only_explicit_unique_year_is_allowed(question):
    changes = validate_suggestion('实验于2025年完成，温度120℃。', '实验于2026年完成，温度120℃。', '用户要求', [{'text': question}])
    assert changes == [{'kind': 'year', 'before': '2025', 'after': '2026'}]


@pytest.mark.parametrize('status,code', [(401,'authentication_failed'),(404,'model_not_found'),(429,'rate_limited'),(503,'service_failed')])
def test_provider_errors_are_redacted_in_invoke_and_stream(status, code):
    failures = []
    class ProviderError(Exception):
        status_code = status
    class Provider:
        def invoke(self, messages):
            raise ProviderError('PRIVATE_RESPONSE_HEADER')
        def stream(self, messages):
            yield {'content': '收到的部分内容'}
            raise ProviderError('PRIVATE_RESPONSE_HEADER')
    model = DraftModel(Provider(), failures)
    with pytest.raises(ValueError) as invocation:
        model.invoke([])
    stream = model.stream([])
    assert next(stream)['content'] == '收到的部分内容'
    with pytest.raises(ValueError) as streaming:
        next(stream)
    assert all(row[0] == code for row in failures) and len(failures) == 2
    assert 'PRIVATE_RESPONSE_HEADER' not in str(invocation.value) + str(streaming.value)


def test_stop_between_validation_and_proposal_commit(draft, monkeypatch):
    service, floating, _ = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    def validating(*args):
        floating.stop()
        return []
    monkeypatch.setattr('office_agent.webui.paged_editor.validate_suggestion', validating)
    model = Model(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    assert model.calls == 1 and not service.proposals and not restarts


def test_step_checks_are_bounded_and_include_source_validation(draft, monkeypatch):
    import math
    import os
    from pathlib import Path
    import statistics
    service, floating, source = draft
    row, runtime, route, restarts = armed(service, floating, monkeypatch)
    model = Model(row)
    run(service, floating, monkeypatch, model, row, runtime, route)
    measurements = service.last_model_checks
    assert {item['phase'] for item in measurements} == {'before_model','model_start','after_model','before_tool','after_tool'}
    assert len(measurements) == 8 and all(item['milliseconds'] >= 0 for item in measurements)
    folder = os.getenv('OFFICE_AGENT_INBOX_EVIDENCE')
    if folder:
        destination = Path(folder)
        destination.mkdir(parents=True, exist_ok=True)
        values = sorted(item['milliseconds'] for item in measurements)
        (destination / 'BOUNDARY_COST.json').write_text(json.dumps({'scope': 'one synthetic DOCX; mocked model; full source guard retained',
            'source_bytes': source.stat().st_size, 'samples': measurements, 'median_ms': statistics.median(values),
            'p95_ms': values[math.ceil(len(values)*.95)-1], 'model_calls': model.calls}, indent=2), encoding='utf8')
