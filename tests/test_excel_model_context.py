"""Excel result plumbing checks: no workbook, model service, GUI or Office needed."""
import json

from office_agent.agent import OfficeAgent, ToolCall, _MODEL_READ_FRAGMENT, _location_summary, _read_context
from office_agent.app import build_tools


def test_model_receives_exact_sheet_title_then_values_beyond_old_excerpt(tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"path fixture only")
    tools = [tool for tool in build_tools((path,))
             if tool.name in {"read_excel_structure", "read_excel_region"}]
    structure = {"classification": {"path": "long-classification-" * 100},
                 "sheet_titles": ["Data123"], "sheet_count": 1,
                 "sheets": [{"title": "Data123", "state": "visible"}], "readable": True}
    cells = [{"cell": f"A{row}", "value": f"ROW_{row}", "data_type": "s"} for row in range(1, 20)]
    cells.append({"cell": "A20", "value": "LAST_ROW_20", "data_type": "s"})

    class Model:
        step = 0

        def bind_tools(self, schemas):
            return self

        def invoke(self, messages):
            step, self.step = self.step, self.step + 1
            if step == 0:
                return {"tool_calls": [{"name": "read_excel_structure", "args": {"path": str(path)}}]}
            tool_message = [item["content"] for item in messages if item["role"] == "tool"][-1]
            if step == 1:
                assert "Data123" in tool_message
                return {"tool_calls": [{"name": "read_excel_region", "args": {
                    "path": str(path), "sheet": "Data123", "row": 1, "column": 1, "rows": 20, "columns": 1}}]}
            assert "LAST_ROW_20" in tool_message
            return {"content": "Observed all 20 returned cells."}

    model = Model()
    agent = OfficeAgent(model, lambda name, args: structure if name == "read_excel_structure" else {
        "sheet": "Data123", "cells": cells}, tools, tmp_path / "sessions",
        confirm=lambda plan: True, context_tokens=32768, reply_tokens=768)
    state = agent.run(agent.new_state("Summarize the selected workbook"))
    assert state.status == "done" and not state.required.errors
    assert "Data123" in state.required.evidence[0].summary
    assert all(len(item.summary) < 400 for item in state.required.evidence)


def test_large_atomic_cell_is_omitted_with_exact_continuation_and_formulas_preserved():
    formula = {"cell": "B1", "formula_text": "=A1*3", "cached_value": 12,
               "cached_value_unverified": True}
    result = {"sheet": "Data123", "cells": [formula, {"cell": "B2", "value": "large" * 10000}]}
    text = _read_context("read_excel_region", result, {"row": 1, "rows": 20, "columns": 1})
    fragment = json.loads(text)
    assert len(text) <= _MODEL_READ_FRAGMENT
    assert fragment["cells"] == [formula]
    assert fragment["context_truncated"] is True and fragment["next_cell"] == "B2"
    assert fragment["cells_total_in_tool_result"] == 2
    assert "large" * 100 not in text


def test_inspector_prioritizes_titles_and_dimensions_not_exhaustive_cell_dump():
    result = {"classification": {"path": "classification-first" * 1000}, "readable": True,
              "sheets": [{"title": "Data123", "max_row": 124, "max_column": 5,
                          "cells": [{"cell": "A1", "value": "secret" * 10000}]}]}
    text = _read_context("inspect", result, {})
    fragment = json.loads(text)
    assert fragment["sheet_titles"] == ["Data123"]
    assert fragment["sheets"][0]["max_row"] == 124
    assert "Data123" in text[:200]
    assert "secret" not in text and len(text) <= _MODEL_READ_FRAGMENT
    assert "classification" in fragment["context_omitted_fields"]


def test_repeated_read_rebuilds_context_even_after_evidence_excerpt_compression(tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"path fixture only")
    tools = [tool for tool in build_tools((path,)) if tool.name == "read_excel_structure"]
    approvals = []
    result = {"classification": {"path": "padding" * 1000}, "sheet_titles": ["Data123"],
              "sheet_count": 1, "sheets": [{"title": "Data123"}]}
    agent = OfficeAgent(None, lambda name, args: result, tools, tmp_path / "sessions",
                        confirm=lambda plan: approvals.append(plan) or True)
    state = agent.new_state("read twice")
    call = ToolCall(name="read_excel_structure", args={"path": str(path)}, id="first")
    agent._execute(call, state)
    evidence = state.required.evidence[0]
    digest, locator = evidence.sha256, evidence.result_path
    evidence.summary = _location_summary(evidence)
    agent._execute(call, state)
    assert len(approvals) == 2 and len(state.required.evidence) == 1
    assert evidence.sha256 == digest and evidence.result_path == locator
    assert "Data123" in evidence.summary
    assert "Data123" in state.history[-1]["content"]


def test_fact_request_receives_precise_prompt_and_complete_long_cell(tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"synthetic dispatch fixture only")
    tools = [tool for tool in build_tools((path,))
             if tool.name in {"read_excel_structure", "read_excel_region"}]
    value = "合成填充" * 260 + "尾部证据=7319"
    structure = {"sheet_titles": ["季度 汇总"], "sheet_count": 1,
                 "sheets": [{"title": "季度 汇总"}]}

    class Model:
        step = 0

        def bind_tools(self, schemas):
            return self

        def invoke(self, messages):
            prompt = messages[0]["content"]
            assert "never assume Sheet1" in prompt
            assert "1-based row/column" in prompt
            assert "only schema-declared argument names" in prompt
            assert "do not quote an entire long cell" in prompt
            step, self.step = self.step, self.step + 1
            if step == 0:
                return {"tool_calls": [{"name": "read_excel_structure", "args": {"path": str(path)}}]}
            if step == 1:
                return {"tool_calls": [{"name": "read_excel_region", "args": {
                    "path": str(path), "sheet": "季度 汇总", "row": 1,
                    "column": 2, "rows": 1, "columns": 1}}]}
            context = [item["content"] for item in messages if item["role"] == "tool"][-1]
            assert value in context and "7319" in context
            return {"content": "尾部证据编号为7319。"}

    agent = OfficeAgent(Model(), lambda name, args: structure if name == "read_excel_structure" else {
        "sheet": "季度 汇总", "cells": [{"cell": "B1", "value": value}]}, tools,
        tmp_path / "sessions", confirm=lambda plan: True, context_tokens=32768, reply_tokens=768)
    state = agent.run(agent.new_state("只回答B1末尾证据编号，不复述完整单元格"))
    assert state.status == "done" and not state.required.errors
    assert state.answer == "尾部证据编号为7319。"
