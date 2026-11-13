"""Day-one synthetic document scenarios: real readers, scripted model, no provider or COM."""
import json

from docx import Document
from openpyxl import Workbook
import pytest

from office_agent.agent import OfficeAgent
from office_agent.app import build_dispatcher, build_tools, select_tools
from office_agent.writes import Editor


class ScriptedModel:
    def __init__(self, name, arguments):
        self.name = name
        self.arguments = arguments
        self.calls = 0
        self.tool_result_seen = False

    def bind_tools(self, schemas):
        assert self.name in {item["function"]["name"] for item in schemas}
        return self

    def invoke(self, messages):
        self.calls += 1
        if self.calls == 1:
            return {"tool_calls": [{"name": self.name, "args": self.arguments, "id": "synthetic-read"}]}
        self.tool_result_seen = any(item["role"] == "tool" and "fresh result" in item["content"]
                                    for item in messages)
        return {"content": "工具调用结束；请根据结构化结果判断能否读取正文。"}


@pytest.fixture
def sample_documents(tmp_path):
    word = tmp_path / "sample.docx"
    document = Document()
    document.add_paragraph("项目说明：仅供离线验收。")
    document.add_paragraph("核准预算是 7319 元。")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "名称"
    table.cell(0, 1).text = "金额"
    table.cell(1, 0).text = "甲"
    table.cell(1, 1).text = "7319"
    document.add_paragraph("synthetic padding " * 300 + "end-marker")
    document.save(word)

    excel = tmp_path / "sample.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "季度 汇总"
    for row in (("部门", "金额", "备注"), ("甲", 7319, None),
                ("乙", None, "=B2*2"), ("丙", 8, "普通文本")):
        sheet.append(row)
    book.save(excel)
    return {"word": word, "excel": excel}


CASES = [
    ("W01", "word", "read_word_region", {"paragraph_index": 1, "limit": 1},
     lambda result: result[0]["paragraph_index"] == 1 and "7319" in result[0]["text"]),
    ("W02", "word", "read_word_text", {},
     lambda result: "7319" in json.dumps(result, ensure_ascii=False) and
                    "名称" in json.dumps(result, ensure_ascii=False)),
    ("W03", "word", "read_word_region", {"paragraph_index": 2, "limit": 1},
     lambda result: result[0]["text_truncated"] is True and len(result[0]["text"]) == 4096),
    ("W05", "word", "search", {"query": "不存在的事项"},
     lambda result: result == []),
    ("X01", "excel", "read_excel_structure", {},
     lambda result: "季度 汇总" in json.dumps(result, ensure_ascii=False)),
    ("X02", "excel", "read_excel_region", {"sheet": "季度 汇总", "row": 2,
                                             "column": 1, "rows": 1, "columns": 2},
     lambda result: any(cell["cell"] == "B2" and cell["value"] == 7319
                        for cell in result["cells"])),
    ("X03", "excel", "summarize_excel_column", {"sheet": "季度 汇总", "column": 2,
                                                  "start_row": 2, "end_row": 4},
     lambda result: "7327" in json.dumps(result, ensure_ascii=False)),
    ("X04", "excel", "read_excel_region", {"sheet": "季度 汇总", "row": 3,
                                             "column": 2, "rows": 2, "columns": 2},
     lambda result: any(cell["cell"] == "C3" and cell["formula_text"] == "=B2*2" and
                        cell["cached_value_unverified"] is True for cell in result["cells"])
                    and not any(cell["cell"] == "B3" for cell in result["cells"])),
]


@pytest.mark.parametrize("case_id,kind,tool,parameters,expected", CASES, ids=[case[0] for case in CASES])
def test_day1_scripted_read_only(case_id, kind, tool, parameters, expected, sample_documents, tmp_path):
    source = sample_documents[kind]
    original = source.read_bytes()
    editor = Editor(tmp_path)
    model = ScriptedModel(tool, {"path": str(source), **parameters})
    allowed = [item for item in build_tools((source,)) if item.name == tool]
    reads = []

    def approve_read(plan):
        reads.append(plan)
        return True

    agent = OfficeAgent(model, build_dispatcher(editor, source), allowed, tmp_path / f"state-{case_id}",
                        confirm=approve_read, persist=False, context_tokens=32768,
                        reply_tokens=768)
    state = agent.run(agent.new_state(f"离线验收 {case_id}"))
    assert state.status == "done" and state.answer and not state.required.errors
    assert model.calls == 2 and model.tool_result_seen
    assert len(reads) == 1 and reads[0].access == "read"
    assert reads[0].arguments["path"] == str(source)
    assert len(state.required.evidence) == 1
    locator = state.required.evidence[0].result_path
    assert locator.startswith("memory:")
    result = json.loads(agent.results[locator.removeprefix("memory:")])
    assert expected(result), f"{case_id}: unexpected bounded read result {result!r}"
    assert source.read_bytes() == original
    assert not list(tmp_path.glob("*.agent-*.docx"))
    assert not list(tmp_path.glob("*.agent-*.xlsx"))


def test_w04_legacy_metadata_read_excludes_docx_tools(tmp_path):
    source = tmp_path / "legacy.doc"
    source.write_bytes(b"synthetic legacy file; no OOXML content")
    original = source.read_bytes()
    editor = Editor(tmp_path)
    allowed = select_tools(build_tools((source,)), "legacy")
    names = {tool.name for tool in allowed}
    assert "inspect" in names and "read_word_text" not in names
    assert "read_word_region" not in names and "image_gen" not in names

    model = ScriptedModel("inspect", {"path": str(source)})
    reads = []

    def approve_read(plan):
        reads.append(plan)
        return True

    agent = OfficeAgent(model, build_dispatcher(editor, source), allowed, tmp_path / "state-W04",
                        confirm=approve_read, persist=False, context_tokens=32768,
                        reply_tokens=768)
    state = agent.run(agent.new_state("离线验收 W04 旧版文件仅检查元信息"))
    assert state.status == "done" and not state.required.errors
    assert len(reads) == 1 and model.tool_result_seen
    locator = state.required.evidence[0].result_path
    report = json.loads(agent.results[locator.removeprefix("memory:")])
    assert report["readable"] is False
    assert "legacy_format_unsupported" in report["risks"]
    assert source.read_bytes() == original
