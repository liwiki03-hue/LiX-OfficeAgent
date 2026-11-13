"""Focused B5 visibility/recovery tests; scripted replies are NOT real model traces."""
from pathlib import Path

import pytest
from pydantic import ValidationError

from office_agent.agent import BudgetError, OfficeAgent, ToolCall
from office_agent.app import build_tools, select_tools, split_read_write


class ScriptedModel:
    def __init__(self, replies):
        self.replies = iter(replies)
        self.visible = []

    def bind_tools(self, schemas):
        self.visible.append({schema["function"]["name"] for schema in schemas})
        return self

    def invoke(self, messages):
        return next(self.replies)


def call(name, **args):
    return {"tool_calls": [{"name": name, "args": args}]}


def make_agent(tmp_path, replies=(), kind="word", estimator=None):
    path = tmp_path / "approved.docx"
    tools = build_tools((path,))
    read, write = split_read_write(select_tools(tools, kind))
    model = ScriptedModel(replies)
    dispatches = []
    agent = OfficeAgent(model, lambda name, args: dispatches.append((name, args)),
                        read, tmp_path / "state", persist=False, confirm=lambda _: True,
                        write_tools=write, write_gate="request_write_tools",
                        capability_tools=tools, context_tokens=32768,
                        token_estimator=estimator)
    return agent, model, dispatches, path


WORD_READ = {"inspect", "com_read", "search", "batch_documents", "summarize_word_document",
             "read_word_text", "read_word_region", "read_word_table_region",
             "read_word_table_structure", "read_word_structure", "request_write_tools",
             "request_read_tools", "read_document_image", "review_document"}
EXCEL_READ = {"inspect", "com_read", "search", "batch_documents", "read_excel_region",
              "summarize_excel_column", "filter_sort_excel_view", "read_excel_structure",
              "request_write_tools", "request_read_tools", "read_document_image",
              "review_document"}


def test_b5_fixed_visible_names_and_wrong_pack_next_round_recovery(tmp_path):
    agent, model, dispatched, _ = make_agent(tmp_path, [
        call("request_read_tools", pack="spreadsheet"), {"content": "done"}])
    state = agent.run(agent.new_state("read a spreadsheet; initial routing was wrong"))
    assert state.status == "done"
    assert model.visible == [WORD_READ, EXCEL_READ]
    assert dispatched == []
    assert state.tool_pack == "spreadsheet" and not state.write_tools_requested
    assert all(not name.startswith("plan_") for names in model.visible for name in names)


def test_b5_explicit_write_only_current_pack_and_read_switch_relocks(tmp_path):
    agent, model, dispatched, _ = make_agent(tmp_path, [
        call("request_read_tools", pack="spreadsheet"), call("request_write_tools"),
        call("request_read_tools", pack="word"), {"content": "done"}])
    state = agent.run(agent.new_state("preview spreadsheet edit, then read Word"))
    excel_write = {"plan_excel", "plan_excel_trim", "plan_excel_range_trim", "plan_excel_range_set",
                   "plan_excel_sort", "plan_excel_chart", "plan_excel_bar_chart", "plan_excel_formula_fill"}
    assert model.visible[2] == (EXCEL_READ - {"request_write_tools"}) | excel_write | {"image_gen"}
    assert model.visible[3] == WORD_READ
    assert not state.write_tools_requested and dispatched == []


def test_b5_all_recovery_never_exposes_all_write_schemas(tmp_path):
    agent, model, _, _ = make_agent(tmp_path, [call("request_read_tools", pack="all"),
                                             call("request_write_tools"), {"content": "done"}])
    agent.run(agent.new_state("need other read capabilities"))
    assert all(not name.startswith("plan_") for names in model.visible for name in names)


def test_b5_restores_routing_but_retains_strict_schema_allowlist_and_read_confirmation(tmp_path):
    agent, _, _, path = make_agent(tmp_path)
    state = agent.new_state("read")
    state.tool_pack = "spreadsheet"
    state.write_tools_requested = True
    # Simulates decoded checkpoint metadata; no saved ToolSpec or approval is trusted.
    agent._restore_tool_visibility(state)
    assert "plan_excel" in agent.tools and "plan_word" not in agent.tools
    with pytest.raises(ValidationError):
        agent._validate_plan(ToolCall(name="read_excel_region", args={"path": str(path), "extra": True}), state)
    with pytest.raises(ValueError, match="outside approved"):
        agent._validate_plan(ToolCall(name="inspect", args={"path": str(tmp_path / "outside.xlsx")}), state)
    agent.confirm = lambda _: False
    agent._execute(ToolCall(name="inspect", args={"path": str(path)}), state)
    assert state.status == "confirmation_required"


def test_b5_routed_schemas_keep_80_90_and_prewrite_budget_gates(tmp_path):
    agent, model, dispatched, path = make_agent(tmp_path, [{"content": "done"}],
        estimator=lambda text: 27000 if "large history" in text else 100)
    state = agent.new_state("keep goal", constraints=["keep safety"])
    state.history = [{"role": "user", "content": "large history"}]
    state.tool_pack = "spreadsheet"
    result = agent.run(state)
    assert result.compressions == 1 and result.required.constraints == ["keep safety"]
    assert model.visible == [EXCEL_READ]
    blocked, model, dispatched, path = make_agent(tmp_path, [{"content": "unreachable"}],
                                                  estimator=lambda _: 30000)
    state = blocked.new_state("preview")
    state.tool_pack = "spreadsheet"
    state.write_tools_requested = True
    assert blocked.run(state).status == "budget"
    assert model.visible == [] and dispatched == []
    with pytest.raises(BudgetError):
        blocked._execute(ToolCall(name="plan_excel", args={"path": str(path), "sheet": "Sheet1",
                                                            "cell": "A1", "value": "new"}), state)
    assert dispatched == []
