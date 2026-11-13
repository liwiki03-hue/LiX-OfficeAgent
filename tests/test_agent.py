"""Offline safety and budget tests: no API credentials or Office installation needed."""
import json
from pathlib import Path

import pytest
from pydantic import BaseModel, ConfigDict

from office_agent.agent import (
    _DEGRADE_AFTER_FAILURES,
    _MAX_COMPLETED,
    _MAX_ERRORS,
    _SYSTEM,
    AgentState,
    Evidence,
    OfficeAgent,
    RequiredState,
    ToolSpec,
    _bound_completed,
    _bound_errors,
    _consecutive_failures,
    _degradation_hint,
    _location_summary,
    configured_decision_limit,
)



class Query(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    keyword: str


class FileRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    path: str


class MockModel:
    def __init__(self, replies):
        self.replies = iter(replies)
        self.messages = []
        self.schemas = []

    def bind_tools(self, schemas):
        self.schemas = schemas
        return self

    def invoke(self, messages):
        self.messages.append(messages)
        return next(self.replies)


def build(tmp_path, replies, *, estimate=None, confirm=None, tools=None, dispatcher=None,
          context=1000, reply=100, decision_limit=15):
    model = MockModel(replies)
    calls = []

    def default_dispatch(name, args):
        calls.append((name, args))
        return {"huge": "secret" * 10000}

    agent = OfficeAgent(model, dispatcher or default_dispatch,
                        tools if tools is not None else [ToolSpec(name="lookup", description="Find facts",
                                                                  schema_model=Query)], tmp_path,
                        confirm=confirm, context_tokens=context, reply_tokens=reply,
                        decision_limit=decision_limit,
                         token_estimator=estimate or (lambda value: len(value) // 3))
    return agent, model, calls


def test_progress_events_follow_durable_checkpoints(tmp_path):
    model = MockModel([{"content": "done"}])
    observed = []
    agent = OfficeAgent(model, lambda name, args: None, [], tmp_path,
                        token_estimator=lambda text: 10,
                        on_progress=lambda state: observed.append((state.decisions,
                            agent.checkpoint.exists(), state.status)))
    result = agent.run(agent.new_state("Goal"))
    assert result.status == "done" and observed
    assert observed[0][1] and observed[-1] == (1, True, "done")


def test_compresses_at_80_percent_before_call(tmp_path):
    # The supplied estimator deliberately maps the history-bearing request to 80%
    # and the compressed request below 90%; all required-state fields must survive.
    estimator = lambda text: 700 if 'obsolete transcript' in text else 200
    agent, model, _ = build(tmp_path, [{"content": "done"}], estimate=estimator)
    state = agent.new_state("Make a report", constraints=["keep formulas"],
                            file_versions={"report.xlsx": "v2"}, pending=["chart"])
    state.required.completed.append("loaded")
    state.required.confirmations.append("previous approval")
    state.required.errors.append("network timeout")
    state.history.append({"role": "assistant", "content": "obsolete transcript"})
    result = agent.run(state)
    assert result.status == "done" and result.compressions == 1
    assert result.history == []
    assert result.required.model_dump() == state.required.model_dump()
    assert all("obsolete transcript" not in str(m) for m in model.messages[0])
    assert "keep formulas" in str(model.messages[0])


def test_compresses_retrieval_even_without_history(tmp_path):
    agent, model, _ = build(tmp_path, [{"content": "done"}],
        estimate=lambda text: 700 if "big retrieval" in text else 100)
    state = agent.run(agent.new_state("Goal"), snippets=["big retrieval"])
    assert state.status == "done" and state.compressions == 1
    assert "big retrieval" not in str(model.messages)


def test_hard_gate_90_percent_does_not_call_model(tmp_path):
    agent, model, _ = build(tmp_path, [{"content": "should not happen"}],
                            estimate=lambda text: 810)
    result = agent.run(agent.new_state("Goal"))  # 810 + reserved 100 >= 900
    assert result.status == "budget" and result.decisions == 0
    assert model.messages == []
    assert agent.load().status == "budget"


def test_budget_stop_states_the_fix_and_wrote_nothing(tmp_path):
    agent, model, calls = build(tmp_path, [{"content": "should not happen"}],
                                estimate=lambda text: 950)
    result = agent.run(agent.new_state("Goal"))
    assert result.status == "budget" and calls == [] and model.messages == []
    assert any("OFFICE_CONTEXT_WINDOW" in error and "No write was performed" in error
               for error in result.required.errors)


def test_tool_schema_drops_only_redundant_titles():
    class Titled(BaseModel):
        model_config = ConfigDict(extra="forbid", strict=True)
        title: str
        body: str

    schema = ToolSpec(name="lookup", description="Find facts", schema_model=Titled).model_schema()
    params = schema["function"]["parameters"]
    # Titles duplicate the field/model names already present, so a property schema is
    # just its type and the property/required names themselves stay intact.
    assert "title" not in params
    assert params["properties"] == {"title": {"type": "string"}, "body": {"type": "string"}}
    assert params["required"] == ["title", "body"]
    assert schema["function"]["description"] == "Find facts"
    assert schema["type"] == "function" and schema["function"]["name"] == "lookup"


def test_compression_sheds_old_excerpts_but_keeps_locations_and_required_facts(tmp_path):
    agent, _, _ = build(tmp_path, [{"content": "done"}], context=1_000_000, reply=100)
    state = agent.new_state("Fill the B column formula", constraints=["never edit the source file"],
                            file_versions={"book.xlsx": "sha256:abc"}, pending=["confirm plan p-1"])
    state.required.confirmations.append("read:inspect:{}")
    for index in range(10):
        state.required.evidence.append(Evidence(
            tool="read_excel_region", result_path=str(tmp_path / "results" / f"{index:02d}.json"),
            sha256=f"{index:064d}", bytes=512,
            summary=f"read_excel_region (fresh result, stored at {index:02d}.json): " + "cellvalue" * 25))
    payload = "cellvalue" * 25
    before = agent.estimate(state)
    # A window this size puts the excerpt-bearing state over 80% but leaves room
    # for the location-only form, which is exactly what compression must achieve.
    agent.context_tokens = int(before / .8) - 10
    snippets = agent._compress(state, ["snippet " * 200])
    after = agent.estimate(state, snippets)

    assert state.compressions == 1 and after < .8 * agent.context_tokens
    assert snippets == []
    # Required facts are never traded away for space.
    assert state.required.goal == "Fill the B column formula"
    assert state.required.constraints == ["never edit the source file"]
    assert state.required.file_versions == {"book.xlsx": "sha256:abc"}
    assert state.required.pending == ["confirm plan p-1"]
    assert state.required.confirmations == ["read:inspect:{}"]
    # Every evidence location survives; only repeated payload excerpts are dropped.
    assert len(state.required.evidence) == 10
    assert [e.result_path for e in state.required.evidence] == [
        str(tmp_path / "results" / f"{index:02d}.json") for index in range(10)]
    assert [e.sha256 for e in state.required.evidence] == [f"{index:064d}" for index in range(10)]
    assert payload in state.required.evidence[-1].summary  # newest excerpt is still inline
    assert all(payload not in e.summary for e in state.required.evidence[:-1])
    assert all(len(e.summary) < len(state.required.evidence[-1].summary)
               for e in state.required.evidence[:-1])


def test_compression_triggers_at_exactly_80_percent_and_not_one_token_below(tmp_path):
    """The documented trigger is ">= 80% of the window", checked before the request.

    The reserved reply is part of the estimate, so a 1000-token window with a 100-token
    reply reaches the trigger at 800: 699 + 100 = 799 must not shed anything, while
    700 + 100 = 800 must compress.
    """
    observed = {}
    for estimate_value in (699, 700):
        agent, _, _ = build(tmp_path / str(estimate_value), [{"content": "done"}],
                            estimate=lambda text, value=estimate_value: value, context=1000)
        state = agent.new_state("Goal")
        state.history.append({"role": "assistant", "content": "obsolete transcript"})
        observed[estimate_value] = agent.run(state)

    below = observed[699]
    assert below.status == "done" and below.compressions == 0
    assert below.history  # nothing was shed below the trigger
    at_trigger = observed[700]
    assert at_trigger.status == "done" and at_trigger.compressions == 1
    assert at_trigger.history == []
    assert at_trigger.required.goal == "Goal"


def test_budget_gate_stops_at_exactly_90_percent_before_any_model_call(tmp_path):
    agent, model, calls = build(tmp_path, [{"content": "must never be requested"}],
                                estimate=lambda text: 800)  # 800 + 100 reserved = exactly 90%
    result = agent.run(agent.new_state("Goal"))
    assert result.status == "budget" and result.decisions == 0
    assert model.messages == [] and calls == []
    assert any("90%" in error and "No write was performed" in error
               for error in result.required.errors)
    assert agent.load().status == "budget"  # the stop is durable and resumable


def test_budget_gate_raises_budget_error_instead_of_returning_a_payload(tmp_path):
    from office_agent.agent import BudgetError

    agent, model, _ = build(tmp_path, [{"content": "done"}], estimate=lambda text: 950)
    with pytest.raises(BudgetError):
        agent._budget(agent.new_state("Goal"), [])
    assert model.messages == []


def test_a_run_over_90_percent_proceeds_only_after_compression_frees_space(tmp_path):
    """The gate is "still above 90% *after* compression", as documented in OFFICE_AGENT.md.

    A payload that cannot be compressed must stop; one that can must still make progress
    rather than dying with a budget status that a resume would repeat forever.
    """
    estimator = lambda text: 950 if "obsolete transcript" in text else 200
    agent, model, _ = build(tmp_path, [{"content": "done"}], estimate=estimator)
    state = agent.new_state("Goal")
    state.history.append({"role": "assistant", "content": "obsolete transcript"})
    result = agent.run(state)
    assert result.status == "done" and result.compressions == 1
    assert len(model.messages) == 1
    assert all("obsolete transcript" not in str(message) for message in model.messages[0])


def test_default_agent_can_finish_after_more_than_15_decisions(tmp_path):
    replies = [{"tool_calls": [{"name": "lookup", "args": {"keyword": str(index)}, "id": str(index)}]}
               for index in range(20)] + [{"content": "done"}]
    agent, model, calls = build(tmp_path, replies, estimate=lambda text: 10, decision_limit=None)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "done" and result.decisions == 21
    assert len(model.messages) == 21 and len(calls) == 20
    assert result.answer == "done"


def test_stop_request_prevents_next_model_call(tmp_path):
    stop = False
    calls = []
    def dispatch(name, args):
        nonlocal stop
        calls.append(name)
        stop = True
        return {"result": "read"}
    model = MockModel([{"tool_calls": [{"name": "lookup", "args": {"keyword": "x"}}]},
                       {"content": "should never be requested"}])
    agent = OfficeAgent(model, dispatch, [ToolSpec(name="lookup", description="Find facts",
                       schema_model=Query)], tmp_path, token_estimator=lambda text: 10,
                       stop_requested=lambda: stop)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "limit" and result.decisions == 1
    assert len(model.messages) == 1 and calls == ["lookup"]
    assert "Stopped by user" in result.required.errors[-1]


def test_repeated_invalid_tool_calls_stop_without_an_iteration_cap(tmp_path):
    replies = [{"tool_calls": [{"name": "unknown", "args": {}, "id": str(index)}]}
               for index in range(20)]
    agent, model, calls = build(tmp_path, replies, estimate=lambda text: 10, decision_limit=None)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "limit" and result.decisions == 6
    assert len(model.messages) == 6 and calls == []
    assert "consecutive failed tool calls" in result.required.errors[-1]


def test_the_15th_decision_is_the_last_request_and_stops_with_limit(tmp_path):
    replies = [{"tool_calls": [{"name": "lookup", "args": {"keyword": str(index)}, "id": str(index)}]}
               for index in range(20)]
    agent, model, calls = build(tmp_path, replies, estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "limit" and result.decisions == 15
    assert len(model.messages) == 15          # the 16th reply was never requested
    assert len(calls) == 15                   # the 15th decision's tool did run, not dropped
    assert result.required.errors[-1] == "Stopped after 15 model decisions"
    assert agent.load().status == "limit"
    # Resuming a session that already hit the cap must not call the model again.
    requested = len(model.messages)
    assert agent.resume().status == "limit"
    assert len(model.messages) == requested


def test_configured_decision_limit_is_bounded_and_explicit(monkeypatch):
    monkeypatch.delenv("OFFICE_DECISION_LIMIT", raising=False)
    assert configured_decision_limit() is None
    for value in ("1", "32", "64"):
        monkeypatch.setenv("OFFICE_DECISION_LIMIT", value)
        assert configured_decision_limit() == int(value)
    monkeypatch.setenv("OFFICE_DECISION_LIMIT", "0")
    assert configured_decision_limit() is None
    for value in ("65", "2.5", " 32", "٣٢", "no"):
        monkeypatch.setenv("OFFICE_DECISION_LIMIT", value)
        with pytest.raises(ValueError, match="OFFICE_DECISION_LIMIT"):
            configured_decision_limit()


def test_extended_limit_stops_and_resume_does_not_request_more(tmp_path):
    replies = [{"tool_calls": [{"name": "lookup", "args": {"keyword": str(n)}, "id": str(n)}]}
               for n in range(40)]
    agent, model, calls = build(tmp_path, replies, estimate=lambda text: 10, decision_limit=32)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "limit" and result.decisions == 32
    assert len(model.messages) == len(calls) == 32
    assert result.required.errors[-1] == "Stopped after 32 model decisions"
    assert agent.resume().decisions == 32 and len(model.messages) == 32
    for limit in (False, 0, 65, "32"):
        with pytest.raises(ValueError, match="Decision limit"):
            OfficeAgent(MockModel([]), lambda name, args: None, [], tmp_path,
                        decision_limit=limit)


def test_system_prompt_requires_complete_cell_references():
    """Bare addresses were a real failure mode with the local model, so the rule is explicit."""
    assert "A1:B2" in _SYSTEM and "Sheet1!A1" in _SYSTEM
    assert "never a bare index" in _SYSTEM
    assert "read-only" in _SYSTEM


def test_degradation_hint_is_deterministic_and_only_after_the_threshold():
    assert _degradation_hint("read_excel_region", 1) == ""
    assert _degradation_hint("read_excel_region", _DEGRADE_AFTER_FAILURES - 1) == ""
    hint = _degradation_hint("read_excel_region", _DEGRADE_AFTER_FAILURES)
    assert "read_excel_region" in hint
    assert str(_DEGRADE_AFTER_FAILURES) in hint
    assert ".env" in hint and "MODEL_PROVIDER" in hint and "*_MODEL" in hint
    assert "DEEPSEEK_MODEL=deepseek-chat" in hint
    # Deterministic: the same inputs always produce the same sentence.
    assert hint == _degradation_hint("read_excel_region", _DEGRADE_AFTER_FAILURES)


def test_repeated_failures_earn_the_hint_and_never_retry_forever(tmp_path):
    failing = {"tool_calls": [{"name": "lookup", "args": {"keyword": 5}, "id": "x"}]}
    agent, model, _ = build(tmp_path, [failing] * 16, estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "limit" and result.decisions == 15      # capped, never unbounded
    assert len(model.messages) == 15
    errors = result.required.errors
    assert not any(".env" in error for error in errors[:2])          # below threshold
    assert any(f"failed {_DEGRADE_AFTER_FAILURES} times in a row" in error for error in errors)
    assert len([error for error in errors if "MODEL_PROVIDER" in error]) >= 1


def test_the_failure_streak_resets_after_a_successful_call(tmp_path):
    failing = {"tool_calls": [{"name": "lookup", "args": {"keyword": 5}, "id": "x"}]}
    working = {"tool_calls": [{"name": "lookup", "args": {"keyword": "ok"}, "id": "y"}]}
    agent, _, _ = build(tmp_path, [failing, failing, working] + [failing] * 3 + [{"content": "done"}],
                        estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "done"
    hints = [error for error in result.required.errors if "in a row" in error]
    # Two failures before the success do not count; only the three after it do.
    assert len(hints) == 1
    assert f"failed {_DEGRADE_AFTER_FAILURES} times in a row" in hints[0]


def test_consecutive_failures_are_derived_from_the_persisted_history():
    state = AgentState(required=RequiredState(goal="g"))
    assert _consecutive_failures(state, "lookup") == 0
    for index in range(2):
        state.history.append({"role": "assistant", "content": "",
                              "tool_calls": [{"name": "lookup", "args": {}, "id": str(index)}]})
        state.history.append({"role": "tool", "tool_call_id": str(index),
                              "content": "Tool rejected or failed; consult errors in required_state."})
    assert _consecutive_failures(state, "lookup") == 2
    assert _consecutive_failures(state, "other_tool") == 0
    state.history.append({"role": "assistant", "content": "", "tool_calls": [{"name": "lookup",
                                                                             "args": {}, "id": "3"}]})
    assert _consecutive_failures(state, "lookup") == 2               # the unpaired call is not counted
    state.history.append({"role": "tool", "tool_call_id": "ok", "content": "fresh result"})
    assert _consecutive_failures(state, "lookup") == 0               # a success breaks the streak


def test_batch_results_stay_inside_the_compression_gate(tmp_path):
    """A batch over 20 files must not push the run past the documented 80%/90% gates.

    The batch tool returns one bounded summary per file; this test measures the *catalogue*
    with that tool present and the live context after several batch calls, so a batch
    feature cannot silently eat the window that the gates protect.
    """
    from office_agent.app import DEFAULT_CONTEXT_WINDOW, DEFAULT_REPLY_BUDGET, build_tools

    target = tmp_path / "book.xlsx"
    target.write_bytes(b"placeholder")
    tools = build_tools((target,), ocr_enabled=True)

    def dispatch(name, args):
        if name != "batch_documents":
            return {"sheet": "Data", "rows": []}
        return {"directory": str(tmp_path), "operation": "summary", "wrote_nothing": True,
                "files": [{"file": f"doc-{index:02d}.docx", "paragraph_count": 120,
                           "character_count": 12000, "preview": "f" * 120} for index in range(20)],
                "skipped": [{"file": "old.doc", "reason": "legacy format"}], "errors": []}

    calls = [{"tool_calls": [{"name": "batch_documents", "args": {"path": str(target)},
                              "id": str(index)}]} for index in range(6)]
    calls.append({"content": "20 份文档已汇总。"})
    agent, model, _ = build(tmp_path, calls, tools=tools, dispatcher=dispatch,
                            confirm=lambda plan: True, context=DEFAULT_CONTEXT_WINDOW,
                            reply=DEFAULT_REPLY_BUDGET,
                            estimate=lambda text: (len(text.encode("utf-8")) + 2) // 3)
    result = agent.run(agent.new_state("汇总这个目录里的文档"))

    assert result.status == "done" and result.answer == "20 份文档已汇总。"
    assert len(result.required.evidence) == 1        # six identical results are recorded once
    assert agent.estimate(result) < .8 * agent.context_tokens
    assert all(len(evidence.summary) < 400 for evidence in result.required.evidence)
    assert all("f" * 500 not in str(message) for messages in model.messages for message in messages)
    assert agent.fixed_overhead() < .4 * agent.context_tokens


def test_a_batch_run_stays_inside_the_window_gate(tmp_path):
    """Distinct batch summaries must stay inside the context window.

    Compression sheds older excerpts and bounds the completed list, so 30 distinct batch
    summaries never exceed the window: the run stops at the 15-decision cap with the estimate
    still under the window, and no edit is ever proposed.
    """
    from office_agent.app import build_tools

    target = tmp_path / "book.docx"
    target.write_bytes(b"placeholder")
    tools = build_tools((target,), ocr_enabled=True)
    served = []

    def dispatch(name, args):
        served.append(name)
        return {"directory": str(tmp_path), "operation": "search", "wrote_nothing": True,
                "files": [{"file": f"doc-{len(served)}.docx", "hit_count": 20,
                           "hits": [{"match": "needle " * 10}] * 20}]}

    calls = [{"tool_calls": [{"name": "batch_documents",
                              "args": {"path": str(target), "query": "needle"},
                              "id": str(index)}]} for index in range(30)]
    agent, model, _ = build(tmp_path, calls, tools=tools, dispatcher=dispatch,
                            confirm=lambda plan: True, context=1000, reply=768,
                            estimate=lambda text: (len(text.encode("utf-8")) + 2) // 3)
    # Room to start, but not enough for twenty distinct batch summaries piling up.
    agent.context_tokens = int(agent.fixed_overhead() / 0.6)
    result = agent.run(agent.new_state("在所有文档里搜索 needle"))

    assert result.status == "limit"                  # 30 次调用超过 15 次决策上限
    assert result.decisions == 15
    assert model.messages
    assert agent.estimate(result) < agent.context_tokens  # 压缩保证不超窗口
    assert not any(name.startswith("plan_") for name in served)   # 无写入
    assert agent.load().status == "limit"


def test_repeated_identical_tool_result_is_recorded_once(tmp_path):
    call = {"tool_calls": [{"name": "lookup", "args": {"keyword": "a"}}]}
    agent, _, calls = build(tmp_path, [call, call, {"content": "done"}], estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert len(calls) == 2 and result.decisions == 3
    assert len(result.required.evidence) == 1
    assert len([line for line in result.required.completed if "Executed" in line]) == 1


def test_changed_tool_result_is_recorded_separately(tmp_path):
    payloads = iter([{"value": 1}, {"value": 2}])
    served = []

    def dispatch(name, args):
        served.append(name)
        return next(payloads)

    agent, _, _ = build(tmp_path, [
        {"tool_calls": [{"name": "lookup", "args": {"keyword": "a"}}]},
        {"tool_calls": [{"name": "lookup", "args": {"keyword": "a"}}]},
        {"content": "done"}], dispatcher=dispatch, estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert len(served) == 2 and len(result.required.evidence) == 2
    assert '"value": 1' in result.required.evidence[0].summary
    assert '"value": 2' in result.required.evidence[1].summary
    assert all(Path(e.result_path).is_file() for e in result.required.evidence)


def test_widest_tool_catalog_completes_a_multi_step_task_inside_the_default_budget(tmp_path):
    """Regression for the observed GUI run that died at 7 decisions with status 'budget'.

    Tool descriptions are serialized into every call, so the fixed schema overhead plus
    per-call evidence growth has to leave room for a real task. Large results must stay
    on disk with only a bounded excerpt in context.
    """
    from office_agent.app import DEFAULT_CONTEXT_WINDOW, DEFAULT_REPLY_BUDGET, build_tools

    target = tmp_path / "book.xlsx"
    target.write_bytes(b"placeholder")
    tools = build_tools((target,), ocr_enabled=True)  # every registered tool, worst case
    served = []

    def dispatch(name, args):
        served.append(name)
        if name == "plan_excel_formula_fill":
            return {"plan_id": "p-1", "preview": {"before": {"B2": None}, "after": {"B2": "=A2*3"}}}
        # Each read returns a distinct oversized payload: results belong on disk.
        return {"sheet": "Data", "secret": "x" * 4000, "read": len(served),
                "rows": [{"cell": f"B{row}", "value": row} for row in range(200)]}

    calls = [{"tool_calls": [{"name": "read_excel_region",
                              "args": {"path": str(target), "sheet": "Data",
                                       "row": index + 1}, "id": str(index)}]}
             for index in range(8)]
    calls.append({"tool_calls": [{"name": "plan_excel_formula_fill",
                                  "args": {"kind": "excel_formula_fill", "path": str(target),
                                           "sheet": "Data", "source_cell": "B1",
                                           "target_range": "B2:B2",
                                           "expected_source_formula": "=A1*3"}, "id": "plan"}]})
    calls.append({"content": "预览已完成，等待你在界面确认写入。"})

    agent, model, _ = build(tmp_path, calls, tools=tools, dispatcher=dispatch,
                            confirm=lambda plan: True, context=DEFAULT_CONTEXT_WINDOW,
                            reply=DEFAULT_REPLY_BUDGET,
                            estimate=lambda text: (len(text.encode("utf-8")) + 2) // 3)
    result = agent.run(agent.new_state("Fill the B column formula down one row"))

    assert result.status == "done" and result.answer == "预览已完成，等待你在界面确认写入。"
    # 8 reads + 1 plan + the final answer decision; nothing was rejected.
    assert result.decisions == 10 and not result.required.errors
    assert len(result.required.evidence) == 9
    # The exhaustive payload never enters the model context; only bounded excerpts do.
    assert all("x" * 500 not in str(message) for messages in model.messages for message in messages)
    assert all(len(evidence.summary) < 400 for evidence in result.required.evidence)
    # The catalogue is re-sent every call, so the window must leave room for real work.
    assert agent.fixed_overhead() < .4 * agent.context_tokens
    assert agent.estimate(result) < .8 * agent.context_tokens
    assert all(evidence.sha256 and Path(evidence.result_path).is_file()
               for evidence in result.required.evidence)
    assert any("p-1" in item for item in result.required.pending)


def test_a_window_too_small_for_the_catalogue_fails_safely_before_any_write(tmp_path):
    """A small window that still fits the catalogue must do useful work and then stop safely."""
    from office_agent.app import DEFAULT_CONTEXT_WINDOW, build_tools

    assert DEFAULT_CONTEXT_WINDOW >= 16384  # must accommodate the measured catalogue overhead
    target = tmp_path / "book.xlsx"
    target.write_bytes(b"placeholder")
    tools = build_tools((target,), ocr_enabled=True)
    calls = [{"tool_calls": [{"name": "read_excel_region",
                              "args": {"path": str(target), "sheet": "Data"}, "id": str(index)}]}
             for index in range(30)]
    calls.append({"content": "never reached"})
    served = []

    def dispatch(name, args):
        served.append(name)
        return {"sheet": "Data", "read": len(served), "payload": "y" * 2000}

    # The catalogue has grown past 8192 tokens of fixed overhead, so this uses a window that
    # still fits it but is small enough to run out mid-task; the property under test is
    # unchanged: use the window up, then stop on the budget without ever proposing an edit.
    agent, model, _ = build(tmp_path, calls, tools=tools, dispatcher=dispatch,
                            confirm=lambda plan: True, context=12288, reply=768,
                            estimate=lambda text: (len(text.encode("utf-8")) + 2) // 3)
    result = agent.run(agent.new_state("Read a lot"))
    assert result.status == "budget"
    assert model.messages  # it did as much useful work as the window allowed
    assert any("OFFICE_CONTEXT_WINDOW" in error for error in result.required.errors)
    assert not any(name.startswith("plan_") for name in served)  # no edit was ever proposed
    assert agent.load().status == "budget"  # resumable progress stays on disk


def test_a_window_smaller_than_the_catalogue_makes_no_model_call_at_all(tmp_path):
    """If the tool catalogue alone cannot fit, refuse before spending a single call."""
    from office_agent.app import build_tools

    target = tmp_path / "book.xlsx"
    target.write_bytes(b"placeholder")
    tools = build_tools((target,), ocr_enabled=True)
    served = []

    def dispatch(name, args):
        served.append(name)
        return {}

    # 8192 is below the measured fixed overhead of the full catalogue, so the very first
    # budget check must fire: a misconfigured window costs nothing and writes nothing.
    agent, model, _ = build(tmp_path, [{"content": "never reached"}], tools=tools,
                            dispatcher=dispatch, confirm=lambda plan: True,
                            context=8192, reply=768,
                            estimate=lambda text: (len(text.encode("utf-8")) + 2) // 3)
    result = agent.run(agent.new_state("Read a lot"))
    assert result.status == "budget"
    assert not model.messages  # no call was attempted
    assert not served  # no tool ran
    assert any("OFFICE_CONTEXT_WINDOW" in error for error in result.required.errors)

def test_stops_after_15_model_decisions_and_resumes_count(tmp_path):
    replies = [{"tool_calls": [{"name": "lookup", "args": {"keyword": "a"}, "id": str(i)}]}
               for i in range(16)]
    agent, model, calls = build(tmp_path, replies, estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "limit" and result.decisions == 15
    assert len(model.messages) == len(calls) == 15
    restarted, second_model, _ = build(tmp_path, [{"content": "must not run"}],
                                        estimate=lambda text: 10)
    assert restarted.resume().decisions == 15
    assert second_model.messages == []
    # The dispatcher returns one fixed payload, so 15 byte-identical results are
    # retained once: required state must not grow with repeated identical reads.
    assert len(restarted.load().required.evidence) == 1
    assert all("secret" * 100 not in str(m) for messages in model.messages for m in messages)
    assert "secret" in result.required.evidence[0].summary  # bounded useful excerpt
    assert all(Path(e.result_path).is_file() for e in result.required.evidence)


def test_resume_requires_new_confirmation_for_pending_read(tmp_path):
    target = tmp_path / "approved.docx"
    spec = ToolSpec(name="read_document", description="Read document",
                    schema_model=FileRequest, access="read", allowed_paths=(target,))
    reply = {"tool_calls": [{"name": "read_document", "args": {"path": str(target)}, "id": "1"}]}
    agent, _, calls = build(tmp_path, [reply], tools=[spec], estimate=lambda text: 10,
                            confirm=lambda plan: False)
    blocked = agent.run(agent.new_state("Read", file_versions={str(target): "sha256:abc"}))
    assert blocked.status == "confirmation_required"
    assert blocked.pending_plan.file_versions[str(target)] == "sha256:abc"
    assert calls == []
    restarted, _, dispatched = build(tmp_path, [{"content": "done"}], tools=[spec],
                                      estimate=lambda text: 10, confirm=lambda plan: True)
    result = restarted.resume()
    assert result.status == "done" and len(dispatched) == 1
    assert len(result.required.confirmations) == 1
    assert result.pending_plan is None


def test_stop_during_resumed_read_confirmation_does_not_dispatch(tmp_path):
    target = tmp_path / "approved.docx"
    spec = ToolSpec(name="read_document", description="Read document",
                    schema_model=FileRequest, access="read", allowed_paths=(target,))
    reply = {"tool_calls": [{"name": "read_document", "args": {"path": str(target)}, "id": "1"}]}
    agent, _, _ = build(tmp_path, [reply], tools=[spec], estimate=lambda text: 10,
                        confirm=lambda plan: False)
    blocked = agent.run(agent.new_state("Read", file_versions={str(target): "sha256:abc"}))
    assert blocked.status == "confirmation_required"
    stop = False
    def confirm(plan):
        nonlocal stop
        stop = True
        return True
    restarted = OfficeAgent(MockModel([{"content": "must not run"}]),
                            lambda name, args: pytest.fail("read dispatched after stop"),
                            [spec], tmp_path, token_estimator=lambda text: 10,
                            confirm=confirm, stop_requested=lambda: stop)
    result = restarted.resume()
    assert result.status == "limit" and result.decisions == 1
    assert "Stopped by user" in result.required.errors[-1]


def test_write_tool_only_returns_plan_for_external_ui_confirmation(tmp_path):
    target = tmp_path / "approved.docx"
    spec = ToolSpec(name="propose_edit", description="Prepare diff only",
                    schema_model=FileRequest, access="write", allowed_paths=(target,))
    calls = []

    def plan_only(name, args):
        calls.append((name, args))
        return {"plan_id": "p-1", "preview": "replace heading"}

    agent, _, _ = build(tmp_path, [{"tool_calls": [{"name": "propose_edit",
        "args": {"path": str(target)}}]}, {"content": "Plan prepared"}],
        tools=[spec], dispatcher=plan_only, estimate=lambda text: 10)
    result = agent.run(agent.new_state("Edit"))
    assert result.status == "done" and len(calls) == 1
    assert result.required.confirmations == []
    assert any("p-1" in item for item in result.required.pending)
    assert "p-1" in result.required.evidence[0].summary


def test_repeated_compression_notes_do_not_accumulate(tmp_path):
    agent, _, _ = build(tmp_path, [{"content": "done"}], estimate=lambda text: 10)
    state = agent.new_state("Goal")
    state.required.completed.extend(f"Executed tool {index}" for index in range(40))
    state.required.errors.extend(f"error {index}" for index in range(40))
    for _ in range(3):
        _bound_completed(state.required)
        _bound_errors(state.required)
    assert len(state.required.completed) == _MAX_COMPLETED + 1
    assert len(state.required.errors) == _MAX_ERRORS + 1
    assert sum("summarised" in line for line in state.required.completed) == 1
    assert sum("suppressed" in line for line in state.required.errors) == 1
    # The newest entries survive; the note stays a single summary line.
    assert state.required.completed[-1] == "Executed tool 39"
    assert state.required.errors[-1] == "error 39"
    assert _location_summary(Evidence(tool="t", result_path="p", sha256="a" * 64, bytes=1,
                                      summary="t: payload")) == "t: stored result, 1 bytes, sha256=" + "a" * 16


def test_resume_after_budget_stop_makes_progress_instead_of_looping(tmp_path):
    """The observed GUI run ended in status 'budget' with a resume checkpoint. Resume
    must free space by shedding excerpts and finish, not fail the same way forever."""
    estimate = lambda text: (len(text.encode("utf-8")) + 2) // 3
    agent, model, _ = build(tmp_path, [{"content": "recovered"}], context=1_000_000, reply=100,
                            estimate=estimate)
    state = agent.new_state("Fill the column", constraints=["never edit the source"])
    for index in range(20):
        state.required.evidence.append(Evidence(
            tool="read_excel_region", result_path=str(tmp_path / f"{index:02d}.json"),
            sha256=f"{index:064d}", bytes=900, summary="read_excel_region: " + "payload" * 40))
    # Put the window just below the un-compressed size so the first call must be refused.
    agent.context_tokens = int(agent.estimate(state) / .95)
    state.status = "budget"
    agent.save(state)

    result = agent.resume()
    assert result.status == "done" and result.answer == "recovered"
    assert result.compressions >= 1 and len(model.messages) == 1
    assert len(result.required.evidence) == 20
    assert all(Path(e.result_path).name.endswith(".json") for e in result.required.evidence)
    assert result.required.goal == "Fill the column"
    assert result.required.constraints == ["never edit the source"]


def test_tool_error_describes_each_offending_field_so_a_retry_can_succeed(tmp_path):
    """A local 7B model was observed failing on a discriminator it could not guess. The
    recorded error must name the field and reason completely instead of being cut off."""
    agent, model, _ = build(tmp_path, [
        {"tool_calls": [{"name": "lookup", "args": {"keyword": 5, "extra": "nope"}}]},
        {"content": "done"}], estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "done"
    reported = " | ".join(result.required.errors)
    assert "Tool failed: ValidationError:" in reported
    assert "keyword:" in reported and "extra:" in reported
    # Complete field-level reasons, not a truncated pydantic dump.
    assert "errors.pydantic.dev" not in reported
    assert len(result.required.errors[0]) < 460


def test_tool_error_without_validation_details_still_reports_cause(tmp_path):
    def explode(name, args):
        raise RuntimeError("dispatcher exploded")

    agent, _, _ = build(tmp_path, [
        {"tool_calls": [{"name": "lookup", "args": {"keyword": "a"}}]},
        {"content": "done"}], dispatcher=explode, estimate=lambda text: 10)
    result = agent.run(agent.new_state("Query"))
    assert result.status == "done"
    assert result.required.errors == ["Tool failed: RuntimeError: dispatcher exploded"]


def test_fixed_overhead_counts_tool_descriptions_and_reply_reserve(tmp_path):
    class Item(BaseModel):
        model_config = ConfigDict(extra="forbid", strict=True)
        path: str
        keyword: str

    small, _, _ = build(tmp_path, [], estimate=lambda text: len(text) // 3, context=8192, reply=768)
    wide, _, _ = build(tmp_path, [], context=8192, reply=768,
                       estimate=lambda text: (len(text.encode("utf-8")) + 2) // 3,
                       tools=[ToolSpec(name=f"tool_{index}", access="read",
                                       description="Read an approved document region "
                                                   f"number {index} with a bounded result",
                                       schema_model=Item, allowed_paths=(tmp_path / "doc.docx",))
                              for index in range(16)])
    assert small.fixed_overhead() >= 768  # the reply reserve is always included
    # The catalog is re-sent on every call, so more tools must cost strictly more.
    assert wide.fixed_overhead() > small.fixed_overhead()
    assert wide.fixed_overhead() < wide.context_tokens


def test_rejects_unapproved_tool_paths_and_schema_extras(tmp_path):
    target = tmp_path / "allowed.docx"
    spec = ToolSpec(name="read_document", description="Read allowed document",
                    schema_model=FileRequest, access="read", allowed_paths=(target,))
    agent, _, calls = build(tmp_path, [{"tool_calls": [{"name": "read_document",
        "args": {"path": str(tmp_path / 'other.docx')}}]}, {"content": "done"}],
        tools=[spec], estimate=lambda text: 10, confirm=lambda plan: True)
    result = agent.run(agent.new_state("Read"))
    assert result.status == "done" and calls == []
    assert any("outside approved" in error for error in result.required.errors)
    with pytest.raises(ValueError, match="Dangerous"):
        build(tmp_path, [], tools=[ToolSpec(name="python", description="unsafe", schema_model=Query)])

    class Permissive(BaseModel):
        keyword: str

    with pytest.raises(ValueError, match="strict"):
        build(tmp_path, [], tools=[ToolSpec(name="lookup", description="unsafe", schema_model=Permissive)])


def test_select_tools_covers_every_catalogue_tool_by_kind():
    """Every tool in the full catalogue must belong to the tier(s) that can serve it.

    This is the regression guard for catalogue tiering: if a new tool is registered in
    ``build_tools`` but never added to a tier, the union check fails and the omission
    cannot silently hide a capability.
    """
    from office_agent.app import _CROSS_TOOLS, _IMAGE_TOOLS, _KIND_TOOLS, _SHARED_TOOLS, build_tools

    tools = build_tools((Path("unused.xlsx"),), ocr_enabled=True)
    names = {tool.name for tool in tools}
    categorised = _SHARED_TOOLS | set().union(*_KIND_TOOLS.values())
    assert categorised == names
    # The word and spreadsheet tiers overlap on the shared read tools, the cross-type read
    # tools, and the image_gen planner tool (which only needs a DOCX/XLSX session anchor);
    # the legacy tier shares nothing with them beyond the always-present read tools.
    assert (_KIND_TOOLS["word"] & _KIND_TOOLS["spreadsheet"]) == (
        _SHARED_TOOLS | _CROSS_TOOLS | _IMAGE_TOOLS)
    assert (_KIND_TOOLS["word"] & _KIND_TOOLS["legacy"]) == _SHARED_TOOLS
    assert (_KIND_TOOLS["spreadsheet"] & _KIND_TOOLS["legacy"]) == _SHARED_TOOLS


def test_select_tools_drops_the_wrong_file_type_and_keeps_same_specs(tmp_path):
    from office_agent.app import build_tools, select_tools

    target = tmp_path / "book.xlsx"
    target.write_bytes(b"x")
    tools = build_tools((target,), ocr_enabled=True)
    by_name = {tool.name: tool for tool in tools}

    sheet = select_tools(tools, "spreadsheet")
    sheet_names = {tool.name for tool in sheet}
    assert "plan_excel_chart" in sheet_names and "read_excel_region" in sheet_names
    assert {"inspect", "com_read"} <= sheet_names
    assert not (sheet_names & {"plan_word", "read_word_text", "plan_docx_merge",
                               "plan_legacy_word_replace", "plan_legacy_cell_set", "ocr_docx_image"})

    word = select_tools(tools, "word")
    word_names = {tool.name for tool in word}
    assert "plan_word" in word_names and "plan_docx_merge" in word_names
    assert "ocr_docx_image" in word_names
    assert not (word_names & {"plan_excel", "read_excel_region", "plan_legacy_cell_set"})

    legacy = select_tools(tools, "legacy")
    assert {tool.name for tool in legacy} == {"inspect", "com_read",
                                               "plan_legacy_word_replace", "plan_legacy_cell_set"}

    unknown = select_tools(tools, "unknown")
    assert {tool.name for tool in unknown} == {"inspect", "com_read"}

    # Filtering never adds tools and returns the exact same spec objects, so the access
    # level, path allowlist and strict schema the agent already validated are unchanged.
    assert len(sheet) < len(tools)
    for spec in sheet:
        assert spec is by_name[spec.name]


def test_tiered_catalogue_meaningfully_reduces_fixed_overhead(tmp_path):
    from office_agent.app import DEFAULT_CONTEXT_WINDOW, build_tools, select_tools

    target = tmp_path / "doc.docx"
    target.write_bytes(b"x")
    full = build_tools((target,), ocr_enabled=True)
    tiered = select_tools(full, "word")
    estimate = lambda text: (len(text.encode("utf-8")) + 2) // 3
    full_agent, _, _ = build(tmp_path, [], tools=full, context=DEFAULT_CONTEXT_WINDOW,
                             reply=768, estimate=estimate)
    tiered_agent, _, _ = build(tmp_path, [], tools=tiered, context=DEFAULT_CONTEXT_WINDOW,
                               reply=768, estimate=estimate)
    assert len(tiered) < len(full)
    # The wrong-file-type tier is gone, so the fixed overhead drops substantially.
    assert tiered_agent.fixed_overhead() < full_agent.fixed_overhead() * 0.7


def test_write_tools_are_gated_until_the_model_requests_them(tmp_path):
    from office_agent.app import _WRITE_GATE, build_tools, select_tools, split_read_write

    target = tmp_path / "book.xlsx"
    target.write_bytes(b"x")
    tiered = select_tools(build_tools((target,), ocr_enabled=True), "spreadsheet")
    read_tools, write_tools = split_read_write(tiered)

    read_names = {tool.name for tool in read_tools}
    assert _WRITE_GATE.name in read_names
    assert not any(name.startswith("plan_") for name in read_names)
    write_names = {tool.name for tool in write_tools}
    # The gated write tier is exactly the plan_* previews plus image_gen, which is also a
    # plan-only tool: it never calls a provider and never writes a file.
    assert write_names == {"plan_excel", "plan_excel_trim", "plan_excel_range_trim",
                           "plan_excel_range_set", "plan_excel_sort", "plan_excel_chart",
                           "plan_excel_bar_chart", "plan_excel_formula_fill", "image_gen"}
    assert not (read_names & write_names)

    class RecordingModel:
        def __init__(self, replies):
            self.replies = iter(replies)
            self.messages = []
            self.bindings = []

        def bind_tools(self, schemas):
            self.bindings.append([schema["function"]["name"] for schema in schemas])
            return self

        def invoke(self, messages):
            self.messages.append(messages)
            return next(self.replies)

    model = RecordingModel([
        {"tool_calls": [{"name": _WRITE_GATE.name, "args": {}, "id": "g1"}]},
        {"content": "done"},
    ])
    agent = OfficeAgent(model, lambda name, args: {"ok": True}, read_tools, tmp_path,
                        write_tools=write_tools, write_gate=_WRITE_GATE.name,
                        context_tokens=32768, reply_tokens=768,
                        token_estimator=lambda text: (len(text.encode("utf-8")) + 2) // 3)
    result = agent.run(agent.new_state("把 A1 改成 7"))

    # First decision sees only the read tier plus the gate: no write tool leaked.
    assert _WRITE_GATE.name in model.bindings[0]
    assert not any(name.startswith("plan_") for name in model.bindings[0])
    # After the gate fired, the write tools joined the catalogue and the gate disappeared.
    assert _WRITE_GATE.name not in agent.tools
    assert write_names <= set(agent.tools)
    assert len(model.bindings) == 2
    assert any(name.startswith("plan_") for name in model.bindings[1])
    assert result.status == "done"


def test_write_gate_requires_a_present_gate_and_disjoint_write_tools(tmp_path):
    from office_agent.app import _WRITE_GATE, build_tools, select_tools, split_read_write

    target = tmp_path / "book.xlsx"
    target.write_bytes(b"x")
    read_tools, write_tools = split_read_write(
        select_tools(build_tools((target,), ocr_enabled=True), "spreadsheet"))
    # write_gate must name a tool that is actually in the initial catalogue.
    with pytest.raises(ValueError, match="write_gate"):
        OfficeAgent(object(), lambda name, args: {}, read_tools, tmp_path,
                    write_tools=write_tools, write_gate="not_present")
    # Write tools may not already be in the initial catalogue (no silent double-send).
    with pytest.raises(ValueError, match="initial catalogue"):
        OfficeAgent(object(), lambda name, args: {}, read_tools, tmp_path,
                    write_tools=read_tools, write_gate=_WRITE_GATE.name)
