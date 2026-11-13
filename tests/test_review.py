"""Tests for the deterministic, read-only document review engine."""
from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from office_agent.review import ERROR, INFO, WARN, review_document, review_docx, review_xlsx

docx = pytest.importorskip("docx")
openpyxl = pytest.importorskip("openpyxl")

from docx import Document as DocxDocument  # noqa: E402
from docx.shared import Pt  # noqa: E402
from openpyxl import Workbook  # noqa: E402


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _codes(report: dict) -> set[str]:
    return {f["code"] for f in report["findings"]}


# --------------------------------------------------------------------------- Word

def _write_docx(path: Path, build) -> Path:
    doc = DocxDocument()
    build(doc)
    doc.save(path)
    return path


def test_placeholder_and_sensitive_are_detected_and_masked(tmp_path):
    def build(doc):
        doc.add_paragraph("这份报告的负责人是 TODO 待填")
        doc.add_paragraph("联系电话 13800138000，身份证 110101199003071234")

    path = _write_docx(tmp_path / "a.docx", build)
    report = review_docx(path)
    assert "word.placeholder.unfilled" in _codes(report)
    assert "word.sensitive.exposure" in _codes(report)

    sensitive = [f for f in report["findings"] if f["code"] == "word.sensitive.exposure"]
    assert sensitive, "expected a sensitive-data finding"
    # The full value must never be reproduced in the evidence.
    for finding in sensitive:
        assert "13800138000" not in finding["evidence"]
        assert "110101199003071234" not in finding["evidence"]
        assert "*" in finding["evidence"]
    assert report["counts"][ERROR] >= 1


def test_heading_level_skip_is_reported(tmp_path):
    def build(doc):
        doc.add_heading("第一章", level=1)
        doc.add_heading("1.1.1 跳级小节", level=3)

    path = _write_docx(tmp_path / "b.docx", build)
    assert "word.heading.level_skip" in _codes(review_docx(path))


def test_manual_numbering_gap_is_reported(tmp_path):
    def build(doc):
        for text in ("1. 准备工作", "2. 需求分析", "4. 结论"):
            doc.add_paragraph(text)

    path = _write_docx(tmp_path / "c.docx", build)
    assert "word.numbering.manual_gap" in _codes(review_docx(path))


def test_format_inconsistency_needs_a_group_norm(tmp_path):
    def build(doc):
        for i in range(4):
            run = doc.add_paragraph(f"正文段落 {i}").runs[0]
            run.font.name = "Arial"
            run.font.size = Pt(12)
        deviant = doc.add_paragraph("格式不同的段落")
        deviant.runs[0].font.name = "Times New Roman"
        deviant.runs[0].font.size = Pt(9)

    path = _write_docx(tmp_path / "d.docx", build)
    report = review_docx(path)
    assert "word.format.inconsistent" in _codes(report)
    # A paragraph must not be reported twice: the indent check uses its own signature.
    inconsistent = [f for f in report["findings"] if f["location"] == "第 5 段"]
    assert len([f for f in inconsistent if f["code"] == "word.format.inconsistent"]) == 1


def test_a_two_way_split_claims_no_norm(tmp_path):
    def build(doc):
        for i in range(3):
            doc.add_paragraph(f"甲 {i}").runs[0].font.name = "Arial"
        for i in range(3):
            doc.add_paragraph(f"乙 {i}").runs[0].font.name = "SimSun"

    path = _write_docx(tmp_path / "e.docx", build)
    # 3/6 is not a minority: no group norm exists, so nothing is claimed.
    assert "word.format.inconsistent" not in _codes(review_docx(path))


def test_empty_table_is_reported(tmp_path):
    def build(doc):
        doc.add_paragraph("正文")
        doc.add_table(rows=2, cols=2)

    path = _write_docx(tmp_path / "f.docx", build)
    report = review_docx(path)
    assert "word.table.empty" in _codes(report)
    assert report["counts"][INFO] >= 1


def test_inherited_format_is_not_treated_as_a_spelling(tmp_path):
    """Well-formatted paragraphs must not be flagged against unformatted ones."""
    def build(doc):
        for i in range(4):
            run = doc.add_paragraph(f"规范正文 {i}").runs[0]
            run.font.name = "SimSun"
            run.font.size = Pt(12)
        for i in range(6):
            doc.add_paragraph(f"未显式设格式 {i}")

    report = review_docx(_write_docx(tmp_path / "i.docx", build))
    # "Inherited" is unknown, not a rival spelling: only explicit formats form a norm.
    assert "word.format.inconsistent" not in _codes(report)


def test_id_number_is_not_also_reported_as_a_bank_card(tmp_path):
    def build(doc):
        doc.add_paragraph("身份证 110101199003071234")

    report = review_docx(_write_docx(tmp_path / "j.docx", build))
    titles = [f["title"] for f in report["findings"] if f["code"] == "word.sensitive.exposure"]
    assert any("身份证号" in t for t in titles)
    assert not any("银行卡号" in t for t in titles)


def test_multiple_placeholder_marks_in_one_paragraph_are_one_finding(tmp_path):
    def build(doc):
        doc.add_paragraph("此处待填 XXX")

    report = review_docx(_write_docx(tmp_path / "k.docx", build))
    got = [f for f in report["findings"] if f["code"] == "word.placeholder.unfilled"]
    assert len(got) == 1
def test_review_never_modifies_the_source(tmp_path):
    def build(doc):
        doc.add_paragraph("TODO 未完成")
        doc.add_table(rows=1, cols=1)

    path = _write_docx(tmp_path / "g.docx", build)
    before = _sha(path)
    review_docx(path)
    assert _sha(path) == before


def test_scan_scope_is_reported(tmp_path):
    def build(doc):
        doc.add_paragraph("正文一")
        doc.add_paragraph("正文二")

    report = review_docx(_write_docx(tmp_path / "h.docx", build))
    assert report["scanned"]["paragraphs"] == 2
    assert report["truncated"] is False
    assert set(report["counts"]) == {ERROR, WARN, INFO}


# -------------------------------------------------------------------------- Excel

def test_excel_empty_column_mixed_type_and_formula_error(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    ws.append(["名称", "空列", "混合"])
    ws.append(["甲", None, 1])
    ws.append(["乙", None, 2])
    ws.append(["丙", None, "N/A"])
    ws.append(["丁", None, 4])
    ws.append(["戊", None, "#DIV/0!"])
    path = tmp_path / "a.xlsx"
    wb.save(path)

    report = review_xlsx(path)
    codes = _codes(report)
    assert "excel.column.empty" in codes
    assert "excel.column.type_mixed" in codes
    assert "excel.formula.error" in codes
    assert report["counts"][ERROR] >= 1


def test_excel_duplicate_row_and_outlier(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["编号", "金额"])
    for i, value in enumerate([10, 11, 12, 13, 14, 1000]):
        ws.append([i + 1, value])
    ws.append([6, 1000])  # identical to the last data row above
    path = tmp_path / "b.xlsx"
    wb.save(path)

    report = review_xlsx(path)
    codes = _codes(report)
    assert "excel.row.duplicate" in codes
    assert "excel.value.outlier" in codes


def test_excel_review_is_read_only(tmp_path):
    wb = Workbook()
    wb.active.append(["a", 1])
    path = tmp_path / "c.xlsx"
    wb.save(path)
    before = _sha(path)
    review_xlsx(path)
    assert _sha(path) == before


def test_dispatch_rejects_unsupported_type(tmp_path):
    from office_agent.documents import UnsupportedDocumentError

    bogus = tmp_path / "x.txt"
    bogus.write_text("hi", encoding="utf-8")
    with pytest.raises(UnsupportedDocumentError):
        review_document(bogus)


# ------------------------------------------------------------- tool registration

def test_review_tool_is_read_only_and_cross_tier(tmp_path):
    from office_agent.app import build_tools, select_tools

    path = _write_docx(tmp_path / "t.docx", lambda d: d.add_paragraph("正文"))
    tools = build_tools((path,))
    tool = next((t for t in tools if t.name == "review_document"), None)
    assert tool is not None, "review_document must be registered"
    assert tool.access == "read"
    assert tool.allowed_paths == (path,)
    for kind in ("word", "spreadsheet"):
        assert tool in select_tools(tools, kind)
    assert tool not in select_tools(tools, "legacy")


def test_dispatcher_reviews_and_guards_stale_version(tmp_path):
    from office_agent.app import build_dispatcher
    from office_agent.writes import Editor

    path = _write_docx(tmp_path / "t.docx", lambda d: d.add_paragraph("TODO 未完成"))
    dispatch = build_dispatcher(Editor(tmp_path), path)
    report = dispatch("review_document", {"path": str(path)})
    assert report["kind"] == "word"
    assert "word.placeholder.unfilled" in _codes(report)

    # A file changed after the session began must fail closed, not be reviewed stale.
    path.write_bytes(path.read_bytes() + b"changed")
    with pytest.raises(ValueError, match="changed"):
        dispatch("review_document", {"path": str(path)})


def test_dispatcher_requires_docx_or_xlsx(tmp_path):
    from office_agent.app import build_dispatcher
    from office_agent.writes import Editor

    legacy = tmp_path / "old.doc"
    legacy.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
    dispatch = build_dispatcher(Editor(tmp_path), legacy)
    with pytest.raises(ValueError, match="DOCX or XLSX"):
        dispatch("review_document", {"path": str(legacy)})


# ------------------------------------------------------------------- rendering

def _sample_report() -> dict:
    return {
        "kind": "word", "file": "报告.docx", "truncated": False,
        "scanned": {"paragraphs": 12, "tables": 1},
        "counts": {"error": 1, "warn": 1, "info": 0},
        "findings": [
            {"code": "word.placeholder.unfilled", "severity": "error", "title": "疑似未填写的占位内容",
             "location": "第 3 段", "evidence": "匹配「TODO」", "suggestion": "替换为真实内容"},
            {"code": "word.format.inconsistent", "severity": "warn", "title": "段落格式不一致",
             "location": "第 5 段", "evidence": "字号=9pt", "suggestion": "统一字号"},
        ],
    }


def test_bridge_renders_review_as_readable_table():
    from office_agent.webui.bridge import _table_from_result

    table = _table_from_result("review_document", _sample_report())
    assert table is not None
    assert table["headers"] == ["严重度", "问题", "位置", "证据", "建议"]
    assert len(table["rows"]) == 2
    assert table["rows"][0][0] == "错误"
    assert "错误 1" in table["caption"] and "扫描范围" in table["caption"]


def test_bridge_states_scope_explicitly_when_nothing_is_found():
    from office_agent.webui.bridge import _table_from_result

    report = _sample_report()
    report["findings"] = []
    report["counts"] = {"error": 0, "warn": 0, "info": 0}
    table = _table_from_result("review_document", report)
    assert table is not None and len(table["rows"]) == 1
    assert "未在扫描范围内发现问题" in table["rows"][0][1]


def test_tk_render_is_human_readable():
    from office_agent.app import render_tool_result

    text = render_tool_result("review_document", _sample_report())
    assert text is not None
    assert "审阅结果" in text and "扫描范围" in text
    assert "[错误]" in text and "[警告]" in text and "建议：" in text


def test_model_context_digest_is_bounded_but_keeps_counts():
    from office_agent.agent import _read_context

    report = _sample_report()
    report["findings"] = [
        {"code": "word.format.inconsistent", "severity": "warn", "title": "格式不一致",
         "location": f"第 {i} 段", "evidence": "证据" * 60, "suggestion": "统一格式"}
        for i in range(200)
    ]
    report["counts"] = {"error": 0, "warn": 200, "info": 0}
    digest = _read_context("review_document", report, {})
    assert digest is not None
    import json as _json
    payload = _json.loads(digest)
    # The bounded fragment is capped, but the totals and scope survive.
    assert payload["counts"]["warn"] == 200
    assert payload["findings_total"] == 200
    assert payload["context_findings_truncated"] is True
    assert len(digest) < 7000
