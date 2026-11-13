from __future__ import annotations

import argparse
from copy import deepcopy
from dataclasses import dataclass
import json
from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook

from office_agent.agent import OfficeAgent
from office_agent.app import build_dispatcher, build_tools, split_read_write
from office_agent.writes import Editor, GuardError, digest


@dataclass(frozen=True)
class Case:
    id: str
    fixture: str
    tool: str
    args: dict
    goal: str
    answer: str
    rejected: bool = False


SHEET = "季度 汇总"
CASES = (
    Case("word_paragraph", "word", "read_word_region", {"paragraph_index": 0, "limit": 1},
         "读取第一段预算。", "预算7319"),
    Case("word_table", "word", "read_word_text", {}, "查找表格中的项目。", "表格项目甲"),
    Case("word_merge_structure", "merged", "read_word_table_structure", {"table_index": 0},
         "检查合并表结构。", "合并结构一致"),
    Case("word_revisions", "revision", "read_word_structure", {}, "检查修订。", "存在插入和删除修订"),
    Case("excel_value", "excel", "read_excel_region", {"sheet": SHEET, "row": 2, "column": 2, "rows": 1, "columns": 1},
         "读取金额 B2。", "金额17"),
    Case("excel_formula_read", "formula", "read_excel_region", {"sheet": SHEET, "row": 2, "column": 3, "rows": 1, "columns": 1},
         "读取 C2 公式，说明缓存是否验证。", "=SUM(A2:B2)；缓存未验证"),
    Case("excel_structure", "excel", "read_excel_structure", {}, "列出工作表。", SHEET),
    Case("word_search_missing", "word", "search", {"query": "missing-marker"}, "搜索 missing-marker。", "未找到"),
    Case("word_replace", "word", "plan_word", {"paragraph_index": 0, "old": "预算", "new": "金额"},
         "把第一段预算改为金额，仅生成计划。", "计划待确认"),
    Case("word_batch", "word", "plan_word_batch_replace", {"old": "预算", "new": "金额"},
         "把正文预算全部改为金额，仅生成计划。", "计划待确认"),
    Case("word_format", "word", "plan_word_format", {"paragraph_index": 0, "run_index": 0, "bold": True},
         "第一段首 run 加粗，仅生成计划。", "计划待确认"),
    Case("word_cell", "word", "plan_word_table_cell", {"table_index": 0, "row_index": 1, "cell_index": 0,
                                                         "paragraph_index": 0, "old": "项目甲", "new": "项目乙"},
         "表格项目甲改为项目乙，仅生成计划。", "计划待确认"),
    Case("word_merged", "merged", "plan_word_merged_cell_text", {"table_index": 0, "row_index": 0,
                                                               "column_index": 0, "action": "replace", "old": "合并标题", "new": "核准标题"},
         "合并单元格标题改为核准标题，仅生成计划。", "计划待确认"),
    Case("revision_accept", "revision", "plan_word_revision_decision", {"action": "accept_all"},
         "接受全部修订，仅生成计划。", "计划待确认"),
    Case("revision_reject", "revision", "plan_word_revision_decision", {"action": "reject_all"},
         "拒绝全部修订，仅生成计划。", "计划待确认"),
    Case("excel_set", "excel", "plan_excel", {"sheet": SHEET, "cell": "B2", "value": 19},
         "B2 改为 19，仅生成计划。", "计划待确认"),
    Case("excel_range", "excel", "plan_excel_range_set", {"sheet": SHEET, "first_cell": "B2", "last_cell": "B3",
                                                           "expected": [[17], [29]], "values": [[18], [30]]},
         "B2:B3 的 17、29 分别改为 18、30，仅生成计划。", "计划待确认"),
    Case("formula_down", "formula", "plan_excel_formula_fill", {"sheet": SHEET, "source_cell": "C2", "target_range": "C3:C4",
                                                                "expected_source_formula": "=SUM(A2:B2)"},
         "从 C2 向下填充 SUM 到 C3:C4，仅生成计划，不声称计算结果。", "计划待确认；公式未计算"),
    Case("formula_right", "formula", "plan_excel_formula_fill", {"sheet": SHEET, "source_cell": "C2", "target_range": "D2:E2",
                                                                 "expected_source_formula": "=SUM(A2:B2)", "direction": "right"},
         "从 C2 向右填充 SUM 到 D2:E2，仅生成计划。", "计划待确认；公式未计算"),
    *(Case(f"chart_{kind}", "excel", "plan_excel_chart", {"sheet": SHEET, "chart_type": kind,
             "series_ranges": ["B2:B4"], "categories_range": "A2:A4", "title": "金额"},
           f"按 A2:A4 类别和 B2:B4 金额生成 {kind} 图表计划。", "计划待确认") for kind in ("bar", "line", "pie")),
    Case("reject_stale_literal", "word", "plan_word", {"paragraph_index": 0, "old": "并不存在", "new": "金额"},
         "不存在的原文必须拒绝。", "拒绝", True),
    Case("reject_formula_assignment", "excel", "plan_excel", {"sheet": SHEET, "cell": "B2", "value": "=WEBSERVICE(\"https://example.invalid\")"},
         "普通单格赋值不能写公式。", "拒绝", True),
    Case("reject_read_bound", "excel", "read_excel_region", {"sheet": SHEET, "rows": 21},
         "超出读取边界必须拒绝。", "拒绝", True),
    Case("reject_other_file", "word", "read_word_region", {"path": "$other", "paragraph_index": 0},
         "未授权文件必须拒绝。", "拒绝", True),
)


def make_fixture(directory: Path, kind: str) -> Path:
    source = directory / ("source.xlsx" if kind in {"excel", "formula"} else "source.docx")
    if kind in {"excel", "formula"}:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = SHEET
        for row in (("项目", "金额"), ("甲", 17), ("乙", 29), ("丙", 8)):
            sheet.append(row)
        if kind == "formula":
            sheet["A2"], sheet["A3"], sheet["A4"] = 1, 2, 3
            sheet["C2"] = "=SUM(A2:B2)"
        workbook.save(source)
        workbook.close()
    else:
        document = Document()
        document.add_paragraph("预算7319")
        document.add_paragraph("预算复核")
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "项目"
        table.cell(0, 1).text = "金额"
        table.cell(1, 0).text = "项目甲"
        table.cell(1, 1).text = "7319"
        if kind == "merged":
            table.cell(0, 0).merge(table.cell(0, 1)).text = "合并标题"
        if kind == "revision":
            for index, (tag, texttag, value) in enumerate((("ins", "t", "新增"), ("del", "delText", "删除"))):
                paragraph = document.add_paragraph("保留")
                node = OxmlElement("w:" + tag)
                node.set(qn("w:id"), str(index + 1))
                node.set(qn("w:author"), "Golden fixture")
                run = OxmlElement("w:r")
                text = OxmlElement("w:" + texttag)
                text.text = value
                run.append(text)
                node.append(run)
                paragraph._p.append(node)
        document.save(source)
    return source


class ScriptedModel:
    def __init__(self, case, args):
        messages = []
        if case.tool.startswith("plan_"):
            messages.append({"tool_calls": [{"name": "request_write_tools", "args": {}, "id": "gate"}]})
        messages.extend(({"tool_calls": [{"name": case.tool, "args": args, "id": case.id}]},
                         {"content": case.answer}))
        self.messages = iter(messages)

    def bind_tools(self, schemas):
        return self

    def invoke(self, messages):
        return next(self.messages)


class ObservedModel:
    def __init__(self, model):
        self.model = model
        self.calls = []

    def bind_tools(self, schemas):
        self.bound = self.model.bind_tools(schemas)
        return self

    def invoke(self, messages):
        response = self.bound.invoke(messages)
        calls = response.get("tool_calls", []) if isinstance(response, dict) else getattr(response, "tool_calls", [])
        self.calls.extend(deepcopy(calls))
        return response


def read_matches(case, result):
    if case.id == "word_paragraph":
        return result[0]["paragraph_index"] == 0 and result[0]["text"] == "预算7319"
    if case.id == "word_search_missing":
        return result == []
    if case.id == "word_merge_structure":
        return result["merge_consistent"] is True and "合并标题" in json.dumps(result, ensure_ascii=False)
    if case.id == "excel_value":
        return any(cell.get("cell") == "B2" and cell.get("value") == 17 for cell in result["cells"])
    if case.id == "excel_formula_read":
        return any(cell.get("cell") == "C2" and cell.get("formula_text") == "=SUM(A2:B2)"
                   and cell.get("cached_value_unverified") is True for cell in result["cells"])
    serialized = json.dumps(result, ensure_ascii=False)
    if case.id == "word_revisions":
        return '"ins"' in serialized and '"del"' in serialized
    return ("项目甲" if case.id == "word_table" else SHEET) in serialized


def output_matches(case, path):
    if path.suffix == ".docx":
        document = Document(path)
        if case.id == "word_replace":
            return document.paragraphs[0].text == "金额7319" and document.paragraphs[1].text == "预算复核"
        if case.id == "word_batch":
            return [p.text for p in document.paragraphs[:2]] == ["金额7319", "金额复核"]
        if case.id == "word_format":
            return document.paragraphs[0].runs[0].bold is True
        if case.id == "word_cell":
            return document.tables[0].cell(1, 0).text == "项目乙"
        if case.id == "word_merged":
            cell = document.tables[0].cell(0, 0)
            return cell.text == "核准标题" and cell._tc is document.tables[0].cell(0, 1)._tc
        if case.id.startswith("revision_"):
            texts = [p.text for p in document.paragraphs[2:]]
            return texts == (["保留新增", "保留"] if case.id == "revision_accept" else ["保留", "保留删除"])
        return False
    workbook = load_workbook(path)
    try:
        sheet = workbook[SHEET]
        if case.id == "excel_set":
            return sheet["B2"].value == 19 and sheet["B3"].value == 29
        if case.id == "excel_range":
            return [sheet[cell].value for cell in ("B2", "B3")] == [18, 30]
        if case.id == "formula_down":
            return [sheet[cell].value for cell in ("C3", "C4")] == ["=SUM(A3:B3)", "=SUM(A4:B4)"]
        if case.id == "formula_right":
            return [sheet[cell].value for cell in ("D2", "E2")] == ["=SUM(B2:C2)", "=SUM(C2:D2)"]
        if case.id.startswith("chart_"):
            return len(sheet._charts) == 1 and type(sheet._charts[0]).__name__.lower().startswith(case.id[6:])
        return False
    finally:
        workbook.close()


def run_case(case: Case, directory: Path, *, model_factory=None) -> dict:
    directory.mkdir()
    source = make_fixture(directory, case.fixture)
    original = digest(source)
    args = {"path": str(source), **deepcopy(case.args)}
    if args["path"] == "$other":
        other = directory / "other.docx"
        other.write_bytes(source.read_bytes())
        args["path"] = str(other)
    sources_before = {path: digest(path) for path in directory.iterdir() if path.is_file()}
    editor = Editor(directory)
    catalogue = build_tools((source,))
    reads, writes = split_read_write(catalogue)
    observed = []
    dispatch = build_dispatcher(editor, source)

    def traced(name, arguments):
        result = dispatch(name, arguments)
        observed.append((name, deepcopy(arguments), result))
        return result

    model = ObservedModel(ScriptedModel(case, args) if model_factory is None else model_factory(case, source))
    agent = OfficeAgent(model, traced, reads, directory / "checkpoint", persist=False,
                        confirm=lambda plan: plan.access == "read", write_tools=writes,
                        write_gate="request_write_tools",
                        context_tokens=65536, reply_tokens=1024, decision_limit=8)
    report = {"id": case.id, "fixture": case.fixture, "tool": case.tool,
              "scores": {"answer_correct": False, "plan_parameters_correct": None,
                         "source_unchanged": False, "fail_closed": None},
              "output_verified": None, "rollback_verified": None}
    try:
        state = agent.run(agent.new_state(case.goal, file_versions={str(source): original}))
        report["answer"] = state.answer[:2000]
        report["status"] = state.status
        report["scores"]["answer_correct"] = state.status == "done" and state.answer.strip() == case.answer
        plans = list(editor.plans.values())
        no_outputs = not list(directory.glob("*.agent-*")) and not editor.audit.exists()
        if case.rejected:
            attempted = any(call.get("name") == case.tool and call.get("args") == args for call in model.calls)
            report["scores"]["fail_closed"] = attempted and bool(state.required.errors) and not plans and no_outputs
        elif case.tool.startswith("plan_"):
            spec = next(tool for tool in catalogue if tool.name == case.tool)
            expected = spec.schema_model.model_validate(args).model_dump()
            correct = len(plans) == 1 and plans[0].request.model_dump() == expected
            report["scores"]["plan_parameters_correct"] = correct
            report["scores"]["fail_closed"] = False
            if correct:
                plan = plans[0]
                try:
                    editor.execute(plan.id)
                except GuardError:
                    report["scores"]["fail_closed"] = no_outputs and not editor.audit.exists() and digest(source) == original
                if report["scores"]["fail_closed"]:
                    result = editor.execute(plan.id, confirmed=True)
                    output = Path(result["output"])
                    report["output_verified"] = result["status"] == "verified" and output_matches(case, output)
                    audit = json.loads((editor.audit / plan.id / "audit.json").read_text(encoding="utf-8"))
                    report["output_verified"] = report["output_verified"] and audit["output_sha256"] == digest(output)
                    rollback = editor.rollback(plan.id, confirmed=True)
                    report["rollback_verified"] = rollback["status"] == "rolled_back" and not output.exists()
        else:
            matching = [result for name, _, result in observed if name == case.tool]
            report["read_verified"] = len(matching) == 1 and read_matches(case, matching[0])
            report["scores"]["answer_correct"] &= report["read_verified"]
    except Exception as error:
        report["error_type"] = type(error).__name__
    finally:
        report["scores"]["source_unchanged"] = all(path.is_file() and digest(path) == before
                                                    for path, before in sources_before.items())
    scores = list(report["scores"].values())
    checks = [report.get("read_verified"), report["output_verified"], report["rollback_verified"]]
    report["passed"] = "error_type" not in report and all(value is not False for value in scores + checks)
    return report


def run_suite(root: Path, *, model_factory=None) -> dict:
    reports = [run_case(case, root / case.id, model_factory=model_factory) for case in CASES]
    dimensions = {}
    for key in reports[0]["scores"]:
        values = [report["scores"][key] for report in reports if report["scores"][key] is not None]
        dimensions[key] = {"passed": sum(value is True for value in values), "evaluated": len(values),
                           "rate": sum(value is True for value in values) / len(values) if values else None}
    return {"schema_version": 1, "mode": "scripted_model_real_documents" if model_factory is None else "injected_model_real_documents",
            "provider_accuracy": None, "real_com": False, "human_confirmation_verified": False,
            "case_count": len(reports), "passed": all(report["passed"] for report in reports),
            "dimensions": dimensions, "cases": reports,
            "limitations": ["Scripted responses test the evaluator and production tool chain, not model reasoning",
                            "Exact-answer rubric; not a semantic judge",
                            "Only disposable fixtures receive simulated host confirmation"]}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="office-golden-") as raw:
        report = run_suite(Path(raw))
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        with args.output.open("x", encoding="utf-8") as stream:
            stream.write(text + "\n")
    else:
        print(text)
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
