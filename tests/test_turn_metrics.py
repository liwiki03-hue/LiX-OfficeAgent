import json
import re
from types import SimpleNamespace

import pytest

from office_agent.turn_metrics import TurnMetrics, response_usage


class Clock:
    def __init__(self):
        self.value = 10.0

    def __call__(self):
        return self.value


def test_overlapping_tools_report_wall_time_and_each_duration():
    clock = Clock()
    metrics = TurnMetrics(clock=clock, tool_names=('read_sheet_range',))
    clock.value = 11
    first = metrics.start('tool', tool='read_sheet_range')
    clock.value = 12
    second = metrics.start('tool', tool='read_sheet_range')
    clock.value = 14
    metrics.stop(first)
    clock.value = 16
    metrics.stop(second)
    metrics.count('tool_requests', 2)
    metrics.count('tool_executions', 2)
    metrics.mark('first_effective_answer', at=15)
    metrics.mark('completed')
    clock.value = 100
    snapshot = metrics.snapshot()
    assert snapshot['elapsed_seconds'] == 6
    assert snapshot['stage_wall_seconds']['tool'] == 5
    assert [span['duration_seconds'] for span in snapshot['spans']] == [3, 4]
    assert snapshot['marks_seconds'] == {'completed': 6, 'first_effective_answer': 5}
    assert snapshot['counts']['tool_requests'] == snapshot['counts']['tool_executions'] == 2
    assert snapshot['outcomes']['ok'] == 2
    assert re.fullmatch('[0-9a-f]{32}', snapshot['turn_id'])
    assert first != second and all(re.fullmatch('[0-9a-f]{32}', span['id']) for span in snapshot['spans'])


def test_missing_usage_stays_unavailable_and_estimates_are_separate():
    metrics = TurnMetrics()
    metrics.count('model_attempts', 2)
    metrics.record_usage({'content': 'private question', 'path': 'C:/private/sample.docx'})
    metrics.record_usage({'usage_metadata': {'input_tokens': 12, 'output_tokens': 4, 'total_tokens': 16}})
    metrics.record_estimate(input_tokens=20, output_tokens=8, total_tokens=28)
    snapshot = metrics.snapshot()
    server = snapshot['usage']['server']
    assert server['input_tokens'] == 12 and server['output_tokens'] == 4
    assert server['reasoning_tokens'] is None
    assert server['reported_requests'] == server['unavailable_requests'] == 1
    assert server['complete'] is False
    assert server['field_reported_requests']['input_tokens'] == 1
    assert snapshot['usage']['local_estimate']['total_tokens'] == 28
    assert 'private' not in json.dumps(snapshot)
    empty = TurnMetrics().snapshot()['usage']['server']
    assert empty['input_tokens'] is None and empty['output_tokens'] is None
    assert empty['requests'] == 0 and empty['complete'] is False


@pytest.mark.parametrize('reply,expected', [
    ({'usage_metadata': {'input_tokens': 2, 'output_tokens': 0, 'total_tokens': 2, 'output_token_details': {'reasoning': 0}}},
     {'input_tokens': 2, 'output_tokens': 0, 'total_tokens': 2, 'reasoning_tokens': 0}),
    (SimpleNamespace(response_metadata={'token_usage': {'prompt_tokens': 10, 'completion_tokens': 4,
                                                     'completion_tokens_details': {'reasoning_tokens': 3}}}),
     {'input_tokens': 10, 'output_tokens': 4, 'reasoning_tokens': 3}),
    ({'response_metadata': {'prompt_eval_count': 8, 'eval_count': 2}}, {'input_tokens': 8, 'output_tokens': 2}),
    ({'usage': {'input_tokens': 5, 'output_tokens': 3}}, {'input_tokens': 5, 'output_tokens': 3}),
    ({'usage_metadata': {'input_tokens': True, 'output_tokens': -1, 'total_tokens': 1.5, 'reasoning_tokens': 'secret'}}, {}),
    ({'usage_metadata': {'input_tokens': 1000000001}, 'response_metadata': {'token_usage': 'private'}}, {}),
])
def test_numeric_usage_normalization(reply, expected):
    assert response_usage(reply) == expected


def test_usage_does_not_double_count_parallel_metadata_or_infer_reasoning():
    metrics = TurnMetrics()
    metrics.record_usage({'usage_metadata': {'input_tokens': 5, 'output_tokens': 2},
                          'response_metadata': {'token_usage': {'prompt_tokens': 5, 'completion_tokens': 2},
                                                'reasoning': 'hidden deliberation', 'key': 'secret'}})
    server = metrics.snapshot()['usage']['server']
    assert server['input_tokens'] == 5 and server['output_tokens'] == 2
    assert server['total_tokens'] is None and server['reasoning_tokens'] is None
    assert server['complete'] is True
    assert 'deliberation' not in json.dumps(metrics.snapshot())


def test_span_context_marks_errors_without_recording_error_text():
    clock = Clock()
    metrics = TurnMetrics(clock=clock)
    with pytest.raises(RuntimeError, match='private'):
        with metrics.span('model'):
            clock.value = 12
            raise RuntimeError('private key C:/private.docx')
    snapshot = metrics.snapshot()
    assert snapshot['outcomes']['error'] == 1
    assert snapshot['stage_wall_seconds']['model'] == 2
    assert 'private' not in json.dumps(snapshot)


def test_open_spans_and_late_confirmation_are_measured_without_changing_first_answer():
    clock = Clock()
    metrics = TurnMetrics(clock=clock)
    identifier = metrics.start('model')
    clock.value = 12
    assert metrics.snapshot()['spans'][0]['active'] is True
    assert metrics.snapshot()['stage_wall_seconds']['model'] == 2
    clock.value = 14
    metrics.mark('first_effective_answer', at=11)
    metrics.mark('first_effective_answer', at=13)
    metrics.stop(identifier, outcome='cancelled')
    assert metrics.snapshot()['marks_seconds']['first_effective_answer'] == 1
    assert metrics.snapshot()['outcomes']['cancelled'] == 1


def test_fixed_labels_and_registered_tool_names_reject_uncontrolled_text():
    metrics = TurnMetrics(tool_names=('inspect_document',))
    for stage, tool in [('private body', None), ('tool', 'C:/private.docx'), ('tool', 'unregistered'), ('model', 'inspect_document')]:
        with pytest.raises(ValueError, match='span_invalid'):
            metrics.start(stage, tool=tool)
    with pytest.raises(ValueError, match='count_invalid'):
        metrics.count('private key')
    with pytest.raises(ValueError, match='count_invalid'):
        metrics.count('decisions', True)
    with pytest.raises(ValueError, match='mark_invalid'):
        metrics.mark('private question')
    with pytest.raises(ValueError, match='tool_name_invalid'):
        TurnMetrics(tool_names=('C:/private.docx',))
    assert 'private' not in json.dumps(metrics.snapshot())


def test_invalid_times_estimates_and_duplicate_stops_are_rejected():
    clock = Clock()
    metrics = TurnMetrics(clock=clock)
    with pytest.raises(ValueError, match='timestamp_invalid'):
        metrics.start('model', at=9)
    with pytest.raises(ValueError, match='timestamp_invalid'):
        metrics.mark('first_effective_answer', at=11)
    identifier = metrics.start('model')
    metrics.stop(identifier)
    with pytest.raises(ValueError, match='span_finished'):
        metrics.stop(identifier)
    with pytest.raises(ValueError, match='estimate_invalid'):
        metrics.record_estimate(input_tokens=True)
    with pytest.raises(ValueError, match='clock_invalid'):
        TurnMetrics(clock=lambda: float('nan'))


def test_snapshots_are_detached_from_internal_state():
    metrics = TurnMetrics()
    metrics.record_usage({'usage': {'input_tokens': 1, 'output_tokens': 2}})
    snapshot = metrics.snapshot()
    snapshot['counts']['model_attempts'] = 500
    snapshot['usage']['server']['input_tokens'] = 500
    assert metrics.snapshot()['counts']['model_attempts'] == 0
    assert metrics.snapshot()['usage']['server']['input_tokens'] == 1


def test_tool_registration_extends_only_valid_complete_batches():
    metrics = TurnMetrics(tool_names=('initial_read',))
    metrics.register_tools(name for name in ('capability_read', 'propose_write'))
    for name in ('initial_read', 'capability_read', 'propose_write'):
        with metrics.span('tool', tool=name):
            pass
    with pytest.raises(ValueError, match='tool_name_invalid'):
        metrics.register_tools(('uncommitted_read', 'C:/private.docx'))
    with pytest.raises(ValueError, match='span_invalid'):
        metrics.start('tool', tool='uncommitted_read')
    with pytest.raises(ValueError, match='tool_name_invalid'):
        metrics.register_tools('unregistered')
    assert {span['tool'] for span in metrics.snapshot()['spans']} == {'initial_read', 'capability_read', 'propose_write'}


def test_raw_reply_reasoning_content_and_think_text_never_enter_metrics():
    metrics = TurnMetrics()
    raw = SimpleNamespace(content='<think>private-thought-871</think>visible answer',
                          additional_kwargs={'reasoning_content': 'private-provider-reasoning-362', 'api_key': 'secret-key-437'},
                          response_metadata={'token_usage': {'prompt_tokens': 11, 'completion_tokens': 6,
                                                            'completion_tokens_details': {'reasoning_tokens': 4}},
                                             'diagnostic': 'C:/private-document-927.docx'})
    metrics.record_usage(raw)
    encoded = json.dumps(metrics.snapshot())
    assert metrics.snapshot()['usage']['server']['reasoning_tokens'] == 4
    assert metrics.snapshot()['usage']['server']['input_tokens'] == 11
    assert all(value not in encoded for value in ('private-thought-871', 'private-provider-reasoning-362',
                                                 'secret-key-437', 'private-document-927', 'visible answer', '<think>'))


def test_public_monotonic_time_can_be_confirmed_later_without_overwriting_first_answer():
    clock = Clock()
    metrics = TurnMetrics(clock=clock)
    assert metrics.now() == 10
    clock.value = 12
    first_visible = metrics.now()
    clock.value = 20
    metrics.mark('first_effective_answer', at=first_visible)
    metrics.mark('first_effective_answer', at=metrics.now())
    assert metrics.snapshot()['marks_seconds']['first_effective_answer'] == 2
