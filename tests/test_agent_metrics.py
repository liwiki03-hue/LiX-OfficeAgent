import json

from pydantic import BaseModel, ConfigDict

from office_agent.agent import OfficeAgent, ToolSpec
from office_agent.turn_metrics import TurnMetrics


class Query(BaseModel):
    model_config = ConfigDict(extra='forbid', strict=True)
    keyword: str


def measured_agent(tmp_path, monkeypatch, model, *, tools=(), dispatcher=lambda *_: None):
    import office_agent.turn_metrics as module
    clock = [10.0]
    metrics = TurnMetrics(clock=lambda: clock[0], tool_names=[spec.name for spec in tools])
    monkeypatch.setattr(module, 'TurnMetrics', lambda **_: metrics)
    events = []
    agent = OfficeAgent(model, dispatcher, list(tools), tmp_path, persist=False,
                        token_estimator=lambda _: 10, on_answer_event=events.append)
    return agent, clock, events


def test_first_visible_answer_excludes_hidden_stream_and_keeps_usage(tmp_path, monkeypatch):
    class Model:
        def bind_tools(self, _):
            return self
        def stream(self, _):
            clock[0] = 11.0
            yield {'content': '<think>private hidden material'}
            clock[0] = 12.0
            yield {'content': '</think>平均延迟 120 ms。'}
            clock[0] = 15.0
            yield {'content': '依据选区。', 'response_metadata': {'finish_reason': 'stop'},
                   'usage_metadata': {'input_tokens': 27, 'output_tokens': 11,
                                      'output_token_details': {'reasoning': 4}}}
    agent, clock, events = measured_agent(tmp_path, monkeypatch, Model())
    state = agent.run(agent.new_state('private question'))
    report = agent.metrics.snapshot()
    assert state.answer == '平均延迟 120 ms。依据选区。'
    assert report['marks_seconds'] == {'first_effective_answer': 2.0, 'completed': 5.0}
    assert report['stage_wall_seconds']['model'] == 5.0
    assert report['usage']['server']['input_tokens'] == 27
    assert report['usage']['server']['reasoning_tokens'] == 4
    assert report['counts']['model_attempts'] == 1 and report['counts']['tool_requests'] == 0
    assert 'private' not in json.dumps(report) and 'hidden' not in json.dumps(events)


def test_tool_preface_is_cancelled_before_first_effective_answer(tmp_path, monkeypatch):
    class Model:
        calls = 0
        def bind_tools(self, _):
            return self
        def stream(self, _):
            self.calls += 1
            if self.calls == 1:
                clock[0] = 11.0
                yield {'content': '准备读取'}
                clock[0] = 13.0
                yield {'content': '', 'tool_calls': [{'name': 'lookup', 'args': {'keyword': 'fact'}, 'id': 'read'}],
                       'response_metadata': {'finish_reason': 'tool_calls'}}
            else:
                clock[0] = 17.0
                yield {'content': '已核对：120 ms。'}
                clock[0] = 19.0
                yield {'content': '', 'response_metadata': {'finish_reason': 'stop'}}
    spec = ToolSpec(name='lookup', description='Find a fact', schema_model=Query)
    agent, clock, events = measured_agent(tmp_path, monkeypatch, Model(), tools=[spec],
                                        dispatcher=lambda *_: {'value': 120, 'unit': 'ms'})
    state = agent.run(agent.new_state('查出平均延迟'))
    report = agent.metrics.snapshot()
    assert state.status == 'done' and state.answer == '已核对：120 ms。'
    assert any(event['type'] == 'answer_draft_cancel' for event in events)
    assert report['marks_seconds'] == {'first_effective_answer': 7.0, 'completed': 9.0}
    assert report['counts']['model_attempts'] == 2 and report['counts']['tool_executions'] == 1
    assert report['usage']['server']['unavailable_requests'] == 2
    assert report['usage']['server']['input_tokens'] is None


def test_failed_model_finishes_metrics_without_claiming_an_answer(tmp_path, monkeypatch):
    class Model:
        def bind_tools(self, _):
            return self
        def stream(self, _):
            clock[0] = 12.0
            raise RuntimeError('provider unavailable')
            yield
    agent, clock, events = measured_agent(tmp_path, monkeypatch, Model())
    result = agent.run(agent.new_state('question'))
    report = agent.metrics.snapshot()
    assert result.answer is None and report['counts']['model_failures'] == 1
    assert report['marks_seconds']['completed'] == 2.0
    assert report['marks_seconds']['first_effective_answer'] is None
    assert report['usage']['server']['unavailable_requests'] == 1


def test_metrics_labels_do_not_change_approved_tool_name_semantics(tmp_path):
    class Model:
        replies = iter([{'tool_calls': [{'name': '查询证据', 'args': {'keyword': 'fact'}, 'id': 'local'}]},
                        {'content': '120 ms'}])
        def bind_tools(self, _):
            return self
        def invoke(self, _):
            return next(self.replies)
    spec = ToolSpec(name='查询证据', description='Read an approved fact', schema_model=Query)
    agent = OfficeAgent(Model(), lambda *_: {'value': 120}, [spec], tmp_path,
                        persist=False, token_estimator=lambda _: 10)
    result = agent.run(agent.new_state('查出延迟'))
    report = agent.metrics.snapshot()
    assert result.answer == '120 ms' and report['counts']['tool_executions'] == 1
    assert '查询证据' not in json.dumps(report, ensure_ascii=False)
