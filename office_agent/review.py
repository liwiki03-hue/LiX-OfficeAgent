"""Deterministic, read-only document review.

This module never writes to, and never mutates, the source document. It produces a
bounded, reproducible list of findings from a DOCX or XLSX so the model can explain
them in plain language. Every finding carries a stable ``code``, a severity, a
location and bounded evidence; nothing here depends on a model or the network.

Scope is deliberately reported: a scan that examined 120 paragraphs and stopped is
never reported as "clean". Absence of a finding means "not detected in what was
scanned", which the caller must state honestly.
"""
from __future__ import annotations

import os
import re
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from statistics import median
from typing import Any, Sequence

from .documents import UnsupportedDocumentError, classify_document

try:  # optional until a DOCX is reviewed
    from docx import Document as _DocxDocument
except ImportError:  # pragma: no cover - exercised only without python-docx
    _DocxDocument = None

try:  # optional until an XLSX is reviewed
    import openpyxl
except ImportError:  # pragma: no cover - exercised only without openpyxl
    openpyxl = None


ERROR, WARN, INFO = "error", "warn", "info"
_SEVERITY_RANK = {ERROR: 0, WARN: 1, INFO: 2}

# Bounds. A review is a report, not a full pass over an unbounded file. What was not
# scanned is always reported so "no finding" is never confused with "not checked".
MAX_PARAGRAPHS = 3000
MAX_TABLES = 50
MAX_SHEETS = 20
MAX_CELLS = 20_000
MAX_FINDINGS = 200
MAX_EVIDENCE = 120
MAX_HEADING_DEPTH = 9

# A format group smaller than this cannot establish a norm, so inconsistency is not
# claimed: two paragraphs differing is a coincidence, not a deviation.
_MIN_GROUP_FOR_NORM = 4
# A deviating spelling must be a clear minority before it is reported as inconsistent.
_MINORITY_FRACTION = 0.5
# The dominant spelling must be a clear majority before anything else is called a
# deviation. Below this, a 50/50 split is a document with two styles, not a fault.
_MIN_DOMINANT_SHARE = 0.6


@dataclass(frozen=True)
class Finding:
    code: str
    severity: str
    title: str
    location: str
    evidence: str
    suggestion: str

    def public(self) -> dict[str, Any]:
        return asdict(self)


def _clip(text: Any, limit: int = MAX_EVIDENCE) -> tuple[str, bool]:
    value = "" if text is None else str(text)
    value = value.replace("\n", " ").replace("\r", " ").strip()
    return (value[:limit], True) if len(value) > limit else (value, False)


def _mask(value: str, keep_head: int = 3, keep_tail: int = 4) -> str:
    """Mask a matched sensitive value; the review never reproduces it in full."""
    if len(value) <= keep_head + keep_tail:
        return "*" * len(value)
    return value[:keep_head] + "*" * (len(value) - keep_head - keep_tail) + value[-keep_tail:]


# --------------------------------------------------------------------------- Word

# Deliberately narrow: a bare "XXX" inside ordinary prose (e.g. "XXXL size") is rare in
# the document types this agent handles, and a false placeholder report is worse than none.
_PLACEHOLDER_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("three_or_more_x", re.compile(r"[XxｘＸ]{3,}")),
    ("todo_marker", re.compile(r"\bTODO\b|\bTBD\b", re.IGNORECASE)),
    ("cn_pending", re.compile(r"待填|待补充|待完善|待确认|待定|请填写|此处填写")),
    ("bracket_blank", re.compile(r"【[^】]{0,30}】|\{\{[^}]{0,30}\}\}|\{[^}\n]{0,30}\}")),
    ("question_run", re.compile(r"[?？]{3,}")),
    ("underscore_blank", re.compile(r"_{4,}|＿{3,}")),
)

_ID_CARD = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")
_MOBILE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
_BANK_CARD = re.compile(r"(?<!\d)\d{16,19}(?!\d)")

_MANUAL_NUMBER = re.compile(r"^\s*(\d{1,3})\s*[.、．]\s*\S")
_HEADING_LEVEL = re.compile(r"heading\s*([1-9])", re.IGNORECASE)


@dataclass(frozen=True)
class _ParaView:
    index: int
    style: str
    level: int | None
    font: str | None
    size_pt: float | None
    bold: bool | None
    indent_twips: int | None
    align: str | None
    text: str


def _effective_font_name(run: Any, paragraph: Any) -> str | None:
    if run.font is not None and run.font.name:
        return str(run.font.name)
    style = getattr(paragraph, "style", None)
    font = getattr(style, "font", None)
    if font is not None and font.name:
        return str(font.name)
    return None


def _effective_size_pt(run: Any, paragraph: Any) -> float | None:
    for source in (getattr(run, "font", None), getattr(getattr(paragraph, "style", None), "font", None)):
        size = getattr(source, "size", None) if source is not None else None
        if size is not None:
            return round(float(size.pt), 2)
    return None


def _heading_level(style_name: str) -> int | None:
    match = _HEADING_LEVEL.search(style_name or "")
    if not match:
        return None
    level = int(match.group(1))
    return level if 1 <= level <= MAX_HEADING_DEPTH else None


def _indent_twips(paragraph: Any) -> int | None:
    try:
        value = paragraph.paragraph_format.first_line_indent
    except Exception:
        return None
    if value is None:
        return None
    try:
        return int(round(float(value.twips)))
    except Exception:
        return None


def _align_name(paragraph: Any) -> str | None:
    try:
        value = paragraph.paragraph_format.alignment
    except Exception:
        return None
    return None if value is None else str(value).rsplit(" ", 1)[-1].upper()


def _paragraph_view(paragraph: Any, index: int) -> _ParaView:
    style = paragraph.style.name if paragraph.style else "Normal"
    text = (paragraph.text or "").strip()
    runs = [r for r in paragraph.runs if (r.text or "").strip()]
    first = runs[0] if runs else None
    return _ParaView(
        index=index,
        style=style,
        level=_heading_level(style),
        font=_effective_font_name(first, paragraph) if first is not None else None,
        size_pt=_effective_size_pt(first, paragraph) if first is not None else None,
        bold=(first.font.bold if first is not None and first.font is not None else None),
        indent_twips=_indent_twips(paragraph),
        align=_align_name(paragraph),
        text=text,
    )


def _format_signature(view: _ParaView) -> tuple[Any, ...]:
    """Font-level spelling only; indentation has its own signature and finding code."""
    return (view.font, view.size_pt, view.bold, view.align)


def _indent_signature(view: _ParaView) -> tuple[Any, ...]:
    return (view.indent_twips,)


def _describe_format(sig: tuple[Any, ...]) -> str:
    font, size, bold, align = sig
    parts = [
        f"字体={font or '继承'}",
        f"字号={size if size is not None else '继承'}pt",
        f"加粗={'是' if bold else ('否' if bold is False else '继承')}",
        f"对齐={align or '继承'}",
    ]
    return "，".join(parts)


def _describe_indent(sig: tuple[Any, ...]) -> str:
    (indent,) = sig
    return f"首行缩进={indent if indent is not None else '无'}twips"


def _norm_deviations(views: Sequence[_ParaView], *, group_label: str, code: str, noun: str,
                     signature_fn: Any, describe_fn: Any) -> list[Finding]:
    """Report paragraphs whose format deviates from the group's dominant spelling."""
    groups: dict[str, list[_ParaView]] = {}
    for view in views:
        groups.setdefault(view.style, []).append(view)
    findings: list[Finding] = []
    for style, members in groups.items():
        # "Inherited" (None) is *unknown*, not a spelling. python-docx does not resolve
        # docDefaults, so a paragraph with no explicit font is not evidence of any
        # particular format. Comparing against it would flag the well-formatted
        # paragraphs that do set a font, so only explicit formats establish a norm.
        explicit = [m for m in members if any(v is not None for v in signature_fn(m))]
        if len(explicit) < _MIN_GROUP_FOR_NORM:
            continue
        counts = Counter(signature_fn(m) for m in explicit)
        if len(counts) < 2:
            continue
        dominant, dominant_count = counts.most_common(1)[0]
        if dominant_count / len(explicit) < _MIN_DOMINANT_SHARE:
            continue  # no clear norm: two styles coexist, nothing to call a deviation
        for view in explicit:
            sig = signature_fn(view)
            if sig == dominant:
                continue
            if counts[sig] / len(explicit) >= _MINORITY_FRACTION:
                continue  # a near-equal split: no norm can be claimed
            evidence, _ = _clip(view.text)
            findings.append(Finding(
                code=code, severity=WARN,
                title=f"{group_label}「{style}」中{noun}与同组多数段落不一致",
                location=f"第 {view.index + 1} 段",
                evidence=f"实际：{describe_fn(sig)}｜同组多数：{describe_fn(dominant)}｜文本：{evidence}",
                suggestion=f"将本段{noun}改为与同组一致的格式，或确认此处确为有意例外",
            ))
            if len(findings) >= MAX_FINDINGS:
                return findings
    return findings


def _scan_placeholders(views: Sequence[_ParaView]) -> list[Finding]:
    findings: list[Finding] = []
    for view in views:
        marks: list[str] = []
        for _name, pattern in _PLACEHOLDER_PATTERNS:
            match = pattern.search(view.text)
            if match and match.group(0)[:20] not in marks:
                marks.append(match.group(0)[:20])
        if not marks:
            continue
        evidence, _ = _clip(view.text)
        findings.append(Finding(
            code="word.placeholder.unfilled", severity=ERROR,
            title="疑似未填写的占位内容",
            location=f"第 {view.index + 1} 段",
            evidence=f"匹配：{'、'.join(marks)}｜文本：{evidence}",
            suggestion="确认并替换为真实内容；提交前应清除全部占位标记",
        ))
        if len(findings) >= MAX_FINDINGS:
            return findings
    return findings


def _scan_sensitive(views: Sequence[_ParaView]) -> list[Finding]:
    findings: list[Finding] = []
    patterns = (("身份证号", _ID_CARD), ("手机号", _MOBILE), ("银行卡号", _BANK_CARD))
    for view in views:
        spans: list[tuple[int, int]] = []
        for label, pattern in patterns:
            match = pattern.search(view.text)
            if not match:
                continue
            # An 18-digit ID is also a valid 16–19 digit "card number". Once a more
            # specific rule has claimed those characters, a looser rule must not
            # re-report the same digits under a second name.
            if any(match.start() < end and start < match.end() for start, end in spans):
                continue
            spans.append((match.start(), match.end()))
            findings.append(Finding(
                code="word.sensitive.exposure", severity=WARN,
                title=f"文中出现疑似{label}",
                location=f"第 {view.index + 1} 段",
                evidence=f"掩码：{_mask(match.group(0))}",  # never reproduce in full
                suggestion=f"如非必要，建议脱敏或移除该{label}；如为示例数据请标注",
            ))
            if len(findings) >= MAX_FINDINGS:
                return findings
    return findings


def _scan_heading_levels(views: Sequence[_ParaView]) -> list[Finding]:
    findings: list[Finding] = []
    outline = [(v.index, v.level) for v in views if v.level is not None]
    previous: int | None = None
    for index, level in outline:
        if previous is not None and level > previous + 1:
            findings.append(Finding(
                code="word.heading.level_skip", severity=WARN,
                title="标题层级出现跳级",
                location=f"第 {index + 1} 段（{level} 级标题）",
                evidence=f"上一标题为 {previous} 级，本标题直接到 {level} 级",
                suggestion="补齐缺失的中间层级，或调整大纲级别使其连续",
            ))
        previous = level
        if len(findings) >= MAX_FINDINGS:
            break
    return findings


def _scan_manual_numbering(views: Sequence[_ParaView]) -> list[Finding]:
    """Detect obvious gaps in hand-typed numbering (1. 2. 4.) within one style."""
    findings: list[Finding] = []
    per_style: dict[str, list[tuple[int, int, str]]] = {}
    for view in views:
        if view.level is not None:
            continue  # auto-numbered headings are not hand-typed numbers
        match = _MANUAL_NUMBER.match(view.text)
        if match:
            per_style.setdefault(view.style, []).append(
                (view.index, int(match.group(1)), view.text))
    for style, items in per_style.items():
        if len(items) < 2:
            continue
        for prev, cur in zip(items, items[1:]):
            if cur[1] == prev[1] + 1:
                continue
            # A restart at 1 is a new list, not a gap.
            if cur[1] <= prev[1]:
                continue
            text, _ = _clip(cur[2])
            findings.append(Finding(
                code="word.numbering.manual_gap", severity=WARN,
                title="手动编号疑似断号",
                location=f"第 {cur[0] + 1} 段",
                evidence=f"上一项为 {prev[1]}，本项为 {cur[1]}｜文本：{text}",
                suggestion="检查是否漏项；如需自动编号请改用列表样式",
            ))
            if len(findings) >= MAX_FINDINGS:
                return findings
    return findings


def _scan_empty_tables(tables: Sequence[Any]) -> list[Finding]:
    findings: list[Finding] = []
    for position, table in enumerate(tables):
        try:
            cells = [c.text.strip() for row in table.rows for c in row.cells]
        except Exception:
            continue
        if cells and not any(cells):
            findings.append(Finding(
                code="word.table.empty", severity=INFO,
                title="表格为空",
                location=f"第 {position + 1} 个表格",
                evidence=f"共 {len(table.rows)} 行，无任何非空单元格",
                suggestion="确认该空表是否为模板残留，可否删除",
            ))
        if len(findings) >= MAX_FINDINGS:
            break
    return findings


def review_docx(path: str | os.PathLike[str]) -> dict[str, Any]:
    if _DocxDocument is None:
        raise UnsupportedDocumentError("python-docx is not installed")
    c = classify_document(path)
    if c.kind != "word" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only readable DOCX can be reviewed")
    doc = _DocxDocument(path)

    findings: list[Finding] = []
    paragraphs = [p for p in doc.paragraphs]
    truncated = len(paragraphs) > MAX_PARAGRAPHS
    views = [_paragraph_view(p, i) for i, p in enumerate(paragraphs[:MAX_PARAGRAPHS])]
    non_empty = [v for v in views if v.text]

    findings += _scan_placeholders(non_empty)
    findings += _scan_sensitive(non_empty)
    findings += _scan_heading_levels(views)
    findings += _scan_manual_numbering(non_empty)
    findings += _norm_deviations(non_empty, group_label="段落格式", code="word.format.inconsistent",
                                 noun="格式", signature_fn=_format_signature,
                                 describe_fn=_describe_format)

    # A body paragraph's first-line indent is the single most common format fault in
    # reports; check it only where a group is large enough to have a norm.
    headings = [v for v in non_empty if v.level is not None]
    body = [v for v in non_empty if v.level is None]
    findings += _norm_deviations(body, group_label="正文缩进", code="word.indent.inconsistent",
                                 noun="缩进", signature_fn=_indent_signature,
                                 describe_fn=_describe_indent)

    tables = list(doc.tables)
    findings += _scan_empty_tables(tables[:MAX_TABLES])

    findings = findings[:MAX_FINDINGS]
    return _report("word", path, findings, scanned={
        "paragraphs": len(views),
        "tables": min(len(tables), MAX_TABLES),
        "headings": len(headings),
        "body_paragraphs": len(body),
    }, truncated=truncated or len(tables) > MAX_TABLES)


# -------------------------------------------------------------------------- Excel

_FORMULA_ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!")
_DATE_HINT = re.compile(r"^\s*\d{1,4}\s*[-/年.]\s*\d{1,2}\s*[-/月.]\s*\d{1,4}")


def _cell_kind(value: Any) -> str:
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)):
        return "number"
    return "text"


def _scan_sheet(name: str, grid: list[list[Any]]) -> list[Finding]:
    findings: list[Finding] = []
    if not grid:
        return findings

    # Column-wise: empty columns and mixed types (data rows only, header row excluded).
    width = max((len(r) for r in grid), default=0)
    for col in range(width):
        column = [row[col] if col < len(row) else None for row in grid]
        data = column[1:] if len(column) > 1 else column
        kinds = [_cell_kind(v) for v in data]
        non_empty = [k for k in kinds if k != "empty"]
        if data and not non_empty:
            findings.append(Finding(
                code="excel.column.empty", severity=INFO,
                title=f"工作表「{name}」存在整列为空",
                location=f"第 {col + 1} 列",
                evidence=f"该列 {len(data)} 个数据格全部为空",
                suggestion="确认该列可否删除，或补全遗漏的数据",
            ))
            continue
        distinct = set(non_empty)
        if len(non_empty) >= 3 and "number" in distinct and "text" in distinct:
            numbers = sum(1 for k in non_empty if k == "number")
            texts = sum(1 for k in non_empty if k == "text")
            findings.append(Finding(
                code="excel.column.type_mixed", severity=WARN,
                title=f"工作表「{name}」某列数据类型不一致",
                location=f"第 {col + 1} 列",
                evidence=f"数字 {numbers} 个、文本 {texts} 个混用",
                suggestion="统一该列类型；文本形式的数字会导致求和与排序出错",
            ))

    # Formula error values surfaced as cached strings.
    for r, row in enumerate(grid):
        for cidx, value in enumerate(row):
            if isinstance(value, str):
                token = value.strip().upper()
                if token in _FORMULA_ERRORS:
                    findings.append(Finding(
                        code="excel.formula.error", severity=ERROR,
                        title=f"工作表「{name}」存在公式错误值",
                        location=f"{_col_letter(cidx)}{r + 1}",
                        evidence=f"值为 {value}",
                        suggestion="定位该公式引用的单元格，修复断裂的引用或除数",
                    ))

    # Duplicate whole rows (non-empty, compared on their normalised values).
    seen: dict[tuple[Any, ...], int] = {}
    for r, row in enumerate(grid):
        values = tuple(v for v in row)
        if not any(v not in (None, "") for v in values):
            continue
        key = tuple(str(v) if v is not None else "" for v in values)
        if key in seen:
            findings.append(Finding(
                code="excel.row.duplicate", severity=WARN,
                title=f"工作表「{name}」存在重复行",
                location=f"第 {r + 1} 行",
                evidence=f"与第 {seen[key] + 1} 行内容完全相同",
                suggestion="确认是否为重复录入，可否合并或删除",
            ))
        else:
            seen[key] = r

    # Robust outlier check per numeric column (median ± 3·MAD).
    for col in range(width):
        numbers = [(r, row[col]) for r, row in enumerate(grid[1:], start=1)
                   if col < len(row) and isinstance(row[col], (int, float))
                   and not isinstance(row[col], bool)]
        if len(numbers) < 5:
            continue
        values = [v for _, v in numbers]
        med = median(values)
        deviations = [abs(v - med) for v in values]
        mad = median(deviations)
        if mad <= 0:
            continue
        for r, value in numbers:
            if abs(value - med) > 3 * mad:
                findings.append(Finding(
                    code="excel.value.outlier", severity=INFO,
                    title=f"工作表「{name}」某数值明显偏离同列水平",
                    location=f"{_col_letter(col)}{r + 1}",
                    evidence=f"值 {value}，同列中位数 {med}",
                    suggestion="确认为真实数据还是录入错误（如多打/少打一位）",
                ))
        if len(findings) >= MAX_FINDINGS:
            return findings
    return findings


def _col_letter(index: int) -> str:
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def review_xlsx(path: str | os.PathLike[str]) -> dict[str, Any]:
    if openpyxl is None:
        raise UnsupportedDocumentError("openpyxl is not installed")
    c = classify_document(path)
    if c.kind != "spreadsheet" or c.protected or not c.readable:
        raise UnsupportedDocumentError("only readable XLSX can be reviewed")
    # data_only=True surfaces cached formula results, which is where error values live.
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)

    findings: list[Finding] = []
    sheets_scanned = 0
    cells_scanned = 0
    truncated = False
    try:
        for sheet_index, sheet in enumerate(workbook.worksheets):
            if sheets_scanned >= MAX_SHEETS:
                truncated = True
                break
            grid: list[list[Any]] = []
            for row in sheet.iter_rows(values_only=True):
                if cells_scanned + len(row) > MAX_CELLS:
                    truncated = True
                    break
                cells_scanned += len(row)
                grid.append(list(row))
            findings += _scan_sheet(sheet.title, grid)
            sheets_scanned += 1
            if len(findings) >= MAX_FINDINGS:
                break
    finally:
        workbook.close()

    findings = findings[:MAX_FINDINGS]
    return _report("spreadsheet", path, findings, scanned={
        "sheets": sheets_scanned,
        "cells": cells_scanned,
    }, truncated=truncated)


# ------------------------------------------------------------------------- shared

def _report(kind: str, path: str | os.PathLike[str], findings: Sequence[Finding], *,
            scanned: dict[str, int], truncated: bool) -> dict[str, Any]:
    ordered = sorted(findings, key=lambda f: (_SEVERITY_RANK.get(f.severity, 9), f.code))
    counts = Counter(f.severity for f in ordered)
    return {
        "kind": kind,
        "file": Path(path).name,
        "scanned": scanned,
        "truncated": truncated,
        "counts": {ERROR: counts.get(ERROR, 0), WARN: counts.get(WARN, 0), INFO: counts.get(INFO, 0)},
        "findings": [f.public() for f in ordered],
    }


def review_document(path: str | os.PathLike[str]) -> dict[str, Any]:
    """Review one DOCX or XLSX. Read-only: the source file is never modified."""
    c = classify_document(path)
    if c.kind == "word":
        return review_docx(path)
    if c.kind == "spreadsheet":
        return review_xlsx(path)
    raise UnsupportedDocumentError("only DOCX and XLSX can be reviewed")


__all__ = ["Finding", "review_document", "review_docx", "review_xlsx",
           "ERROR", "WARN", "INFO"]
