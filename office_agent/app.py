"""Tk Windows desktop MVP. Human file selection grants exactly one document per task."""
from __future__ import annotations

import json
import os
import queue
import shutil
import threading
import unicodedata
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, font as tkfont, messagebox, ttk
except ImportError:  # A headless Python can still run non-GUI tests and tools.
    tk = None
    filedialog = messagebox = ttk = tkfont = None

from pydantic import Field, StrictInt, StrictStr, field_validator

from typing import Literal

from .branding import APP_ICON
from .agent import OfficeAgent, StrictModel, ToolSpec, configured_decision_limit, DEFAULT_DECISION_LIMIT
from .documents import OCRProvider, classify_document, inspect_file, ocr_docx_image, search_document
from .documents import batch_document_report
from .imagegen import ImageGenError, ImageRequest
from .imagetransactions import MAX_PLANS
from .writes import (DocxMerge, DocxSplit, Editor, ExcelBarChart, ExcelChart, ExcelFormulaFill, ExcelRangeSet, ExcelRangeTrim, ExcelSet, ExcelSort, ExcelTrim,
                     LegacyCellSet, LegacyWordReplace, WordBatchReplace, WordMergedCellText, WordReplace, WordFormat,
                     WordRevisionDecision, WordTableCellReplace, WordTableRow, digest)

# The full tool catalogue is re-sent on every model call. At an 8192 or 16384 window the tool
# descriptions alone would consume most of the space available for document work, and the
# configured providers offer less context than the default below. This default
# matches what the configured providers actually offer (deepseek-chat 64k, qwen-plus 32k, local
# qwen2.5:7b 32k). It must never exceed the configured model's real window: the GUI prints the
# measured overhead and the window at the start of every task so a mismatch is visible, and
# tests/test_agent.py measures the full catalogue against this default.
DEFAULT_CONTEXT_WINDOW = 32768
DEFAULT_REPLY_BUDGET = 768


def _writable_directory(path: Path) -> bool:
    """True only if this process can really create and remove entries under ``path``.

    Permission bits and ACL listings lie in confined environments, so the check performs
    an actual create/write/delete cycle. Any directory it had to create for the test is
    removed again, so merely selecting a file or declining a confirmation leaves the
    document folder exactly as it was.
    """
    missing = []
    for candidate in [path, *path.parents]:
        if candidate.exists():
            break
        missing.append(candidate)
    probe = path / f".write-probe-{uuid.uuid4().hex}"
    try:
        probe.mkdir(parents=True, exist_ok=True)
        (probe / "probe.txt").write_text("ok", encoding="utf-8")
    except OSError:
        usable = False
    else:
        usable = True
    shutil.rmtree(probe, ignore_errors=True)
    for candidate in missing:  # deepest first; stops at the first directory still in use
        try:
            candidate.rmdir()
        except OSError:
            break
    return usable


class InspectArgs(StrictModel):
    path: str


class PptxStructureArgs(StrictModel):
    path: str
    offset: int = Field(default=0, ge=0, le=1000)
    limit: int = Field(default=20, ge=1, le=20)


class PptxSlideArgs(StrictModel):
    path: str
    slide_index: int = Field(default=0, ge=0, le=999)
    offset: int = Field(default=0, ge=0, le=20000)
    limit: int = Field(default=20, ge=1, le=50)
    text_offset: int = Field(default=0, ge=0, le=33554432)
    max_chars: int = Field(default=4000, ge=200, le=20000)


class PdfReadArgs(StrictModel):
    path: str
    page: int = Field(default=0, ge=0, le=1999)
    limit: int = Field(default=3, ge=1, le=5)
    char_offset: int = Field(default=0, ge=0, le=2000000)
    max_chars: int = Field(default=4000, ge=200, le=20000)
    layout: bool = True


class DirectorySearchArgs(StrictModel):
    path: str
    query: str = Field(min_length=1, max_length=200)
    top_k: int = Field(default=5, ge=1, le=10)


class ReviewArgs(StrictModel):
    path: str


class OCRImageArgs(StrictModel):
    path: str
    relationship: str = Field(pattern=r"^rId[0-9]{1,8}$")


class DocumentImageArgs(StrictModel):
    path: str
    relationship: str = Field(pattern=r"^rId[0-9]{1,8}$")
    part: str = Field(default="word/document.xml", max_length=120,
                      pattern=r"^(word/(document|header[0-9]{1,8}|footer[0-9]{1,8})|xl/drawings/drawing[0-9]{1,8})\.xml$")


class ImageGenArgs(StrictModel):
    """Planner-only schema. Mirrors ImageRequest so no invalid request ever reaches a plan."""

    path: str
    prompt: StrictStr = Field(min_length=1, max_length=4000)
    size: Literal["1024x1024", "1024x1536", "1536x1024"] = "1024x1024"
    style: Literal["natural", "vivid"] = "natural"
    n: StrictInt = Field(default=1, ge=1, le=1)
    prompts: list[StrictStr] | None = Field(default=None, min_length=1, max_length=8)

    @field_validator("prompt", "prompts")
    @classmethod
    def validate_prompt(cls, value: str | list[str] | None) -> str | list[str] | None:
        if value is None:
            return None
        for prompt in value if isinstance(value, list) else [value]:
            if not prompt.strip() or len(prompt) > 4000 or any(0xD800 <= ord(c) <= 0xDFFF for c in prompt):
                raise ValueError("invalid_image_prompt")
        return value


class SearchArgs(StrictModel):
    path: str
    query: str = Field(min_length=1, max_length=200)


class WordRegionArgs(StrictModel):
    path: str
    paragraph_index: int = Field(default=0, ge=0)
    limit: int = Field(default=8, ge=1, le=20)


class WordTableRegionArgs(StrictModel):
    path: str
    table_index: int = Field(ge=0)
    start_row: int = Field(default=0, ge=0)
    start_column: int = Field(default=0, ge=0)
    rows: int = Field(default=5, ge=1, le=20)
    columns: int = Field(default=5, ge=1, le=20)


class WordStructureArgs(StrictModel):
    path: str
    limit: int = Field(default=20, ge=1, le=50)


class WordSummaryArgs(StrictModel):
    path: str
    chunk_size: int = Field(default=20, ge=1, le=200)
    max_chunks: int = Field(default=25, ge=1, le=100)
    preview_chars: int = Field(default=200, ge=1, le=1000)


class WordTextArgs(StrictModel):
    path: str
    offset: int = Field(default=0, ge=0)
    limit: int = Field(default=20, ge=1, le=50)
    max_chars: int = Field(default=4000, ge=200, le=20000)


class ExcelRegionArgs(StrictModel):
    path: str
    sheet: str = Field(min_length=1)
    row: int = Field(default=1, ge=1, le=1048576)
    column: int = Field(default=1, ge=1, le=16384)
    rows: int = Field(default=8, ge=1, le=20)
    columns: int = Field(default=5, ge=1, le=20)


class ExcelViewArgs(StrictModel):
    path: str
    sheet: str = Field(min_length=1)
    start_row: int = Field(ge=1, le=1048576)
    end_row: int = Field(ge=1, le=1048576)
    start_column: int = Field(ge=1, le=16384)
    columns: int = Field(ge=1, le=20)
    key_column: int = Field(ge=1, le=16384)
    equals: str | int | float | bool | None = None
    descending: bool = False


class COMReadArgs(StrictModel):
    path: str
    include_tables: bool = False
    region: str | None = None


class ExcelSummaryArgs(StrictModel):
    path: str
    sheet: str = Field(min_length=1)
    column: int = Field(ge=1, le=16384)
    start_row: int = Field(ge=1, le=1048576)
    end_row: int = Field(ge=1, le=1048576)


class ExcelStructureArgs(StrictModel):
    path: str
    limit: int = Field(default=50, ge=1, le=100)


class BatchArgs(StrictModel):
    path: str
    operation: Literal["inspect", "summary", "search"] = "summary"
    query: str | None = None
    limit: int = Field(default=20, ge=1, le=20)
    per_file: int = Field(default=20, ge=1, le=20)


class WordTableStructureArgs(StrictModel):
    path: str
    table_index: int = Field(default=0, ge=0)
    limit: int = Field(default=50, ge=1, le=100)


class DocxEvidenceArgs(StrictModel):
    path: str
    offset: int = Field(default=0, ge=0, le=4000)
    limit: int = Field(default=3, ge=1, le=5)


# One request model per plan-only tool: the tool name is the discriminator of record.
from .advanced_writes import TOOL_MODELS
from .professional_edits import ConstrainedDocx

_PLAN_MODELS = {
    **TOOL_MODELS,
    'plan_constrained_docx': ConstrainedDocx,
    "plan_word": WordReplace,
    "plan_word_batch_replace": WordBatchReplace,
    "plan_word_format": WordFormat,
    "plan_word_table_cell": WordTableCellReplace,
    "plan_word_table_row": WordTableRow,
    "plan_word_merged_cell_text": WordMergedCellText,
    "plan_word_revision_decision": WordRevisionDecision,
    "plan_legacy_word_replace": LegacyWordReplace,
    "plan_legacy_cell_set": LegacyCellSet,
    "plan_docx_merge": DocxMerge,
    "plan_docx_split": DocxSplit,
    "plan_excel": ExcelSet,
    "plan_excel_trim": ExcelTrim,
    "plan_excel_range_trim": ExcelRangeTrim,
    "plan_excel_range_set": ExcelRangeSet,
    "plan_excel_sort": ExcelSort,
    "plan_excel_bar_chart": ExcelBarChart,
    "plan_excel_chart": ExcelChart,
    "plan_excel_formula_fill": ExcelFormulaFill,
}


def build_dispatcher(editor: Editor, path: Path, *, on_plan=None, ocr_provider: OCRProvider | None = None,
                     expected_version: str | None = None, on_image=None, image_service_factory=None, image_count: int = 1):
    """The single authorized-file dispatcher the model's tool calls go through.

    Every call re-checks that the target is the exact selected file and that its bytes
    are unchanged since the session started, so a stale or redirected path fails closed.
    Write tools only ever produce a preview plan; execution stays with the UI.
    ``image_service_factory`` is an explicit host injection: without it image_gen stays
    unavailable and explains how to configure it instead of calling any image service.
    """
    editor.require_confirmation = True
    if type(image_count) is not int or not 1 <= image_count <= MAX_PLANS:
        raise ValueError(f"图片张数须为 1–{MAX_PLANS}")
    path = editor.space.file(str(path))
    expected = digest(path) if expected_version is None else expected_version
    image_service = None
    image_batch_planned = False

    def _image_service():
        """Lazily create the P3 approval service; never call a provider from here."""
        nonlocal image_service
        if image_service is None:
            if image_service_factory is None:
                raise ValueError(
                    "云端生图未配置：请在 .env 设置 IMAGE_PROVIDER/IMAGE_API_KEY/IMAGE_BASE_URL/IMAGE_MODEL "
                    "后重启应用；本工具不会在没有配置与界面独立确认时调用任何生图服务")
            try:
                image_service = image_service_factory()
            except ImageGenError as error:
                raise ValueError(f"云端生图不可用（{error.code}）：检查 IMAGE_* 配置；"
                                 "未调用任何生图服务") from None
        return image_service


    def dispatch(name: str, args: dict):
        if name in {'read_docx_task_navigation', 'plan_docx_task'}:
            from .document_tasks import NavigationArgs, TaskSpec, navigation, compile_task, public_task
            checked = editor.space.file(str(path))
            if checked != path or digest(checked) != expected:
                raise ValueError('Selected file changed; refusing stale operation')
            if name == 'read_docx_task_navigation':
                return navigation(editor, checked, **NavigationArgs.model_validate(args).model_dump())
            task = compile_task(editor, checked, TaskSpec.model_validate(args).model_dump())
            editor._document_task_previews = {task['id']: task}
            return {**public_task(task), 'plan_id': task['id'], 'plan_kind': 'document_task_preview', 'requires_task_import': True,
                    'preview_only': True, 'note': '任务覆盖方案已登记；此ID不能直接写入，用户在完整修改任务中导入并逐组独立确认，不自动写入'}
        if name == 'plan_docx_blocks':
            from .block_edits import BlockIntent, BlockDocx
            from .range_edits import now
            checked = editor.space.file(str(path))
            if checked != path or digest(checked) != expected:
                raise ValueError('Selected file changed; refusing stale operation')
            intent = BlockIntent.model_validate(args)
            plan = editor.plan(BlockDocx(**intent.model_dump(), path=str(path), revision_date=now()))
            if on_plan is not None:
                on_plan(('plan', plan.public()))
            return {'plan_id': plan.id, 'preview': plan.public(), 'note': '普通结构净稿；非原生结构修订，必须独立确认'}
        if name in {'read_docx_ranges', 'plan_docx_ranges'}:
            from .range_edits import RangeIntent, RangeViewArgs, RangeDocx, now, view
            checked = editor.space.file(str(path))
            if checked != path or digest(checked) != expected:
                raise ValueError('Selected file changed; refusing stale operation')
            if name == 'read_docx_ranges':
                values = RangeViewArgs.model_validate(args)
                return view(editor, checked, values.offset)
            intent = RangeIntent.model_validate(args)
            plan = editor.plan(RangeDocx(**intent.model_dump(), path=str(path), revision_date=now()))
            if on_plan is not None:
                on_plan(('plan', plan.public()))
            return {'plan_id': plan.id, 'preview': plan.public(), 'note': '仅预览；独立人工确认后才写副本'}
        checked = editor.space.file(args["path"])
        if checked != path:
            raise ValueError("Only the explicitly selected file is authorized")
        if name == "read_document_image":
            from .documents import _image_source_preflight
            _image_source_preflight(checked)
        if digest(checked) != expected:
            raise ValueError("Selected file changed since this session started; refusing stale operation")
        if name == "inspect":
            return inspect_file(checked)
        if name == 'read_docx_evidence':
            from .professional_edits import snapshot
            values = DocxEvidenceArgs.model_validate(args).model_dump(exclude={'path'})
            return snapshot(editor, checked, **values)
        if name in {'read_pptx_structure', 'read_pptx_slide', 'read_pdf', 'search_directory_documents'}:
            schemas = {'read_pptx_structure': PptxStructureArgs, 'read_pptx_slide': PptxSlideArgs,
                       'read_pdf': PdfReadArgs, 'search_directory_documents': DirectorySearchArgs}
            validated = schemas[name].model_validate(args).model_dump(exclude={'path'})
            if name.startswith('read_pptx_'):
                from .presentations import read_pptx_slide, read_pptx_structure
                result = (read_pptx_structure if name == 'read_pptx_structure' else read_pptx_slide)(checked, **validated)
            elif name == 'read_pdf':
                from .pdf_reader import read_pdf
                result = read_pdf(checked, **validated)
            else:
                from .directory_search import DirectoryIndex
                result = DirectoryIndex(editor.space, checked).build().query(**validated)
            if editor.space.file(args['path']) != path or digest(checked) != expected:
                raise ValueError('Selected file changed during read')
            return result
        if name == "review_document":
            if checked.suffix.lower() not in {".docx", ".xlsx"}:
                raise ValueError("DOCX or XLSX required")
            from .review import review_document as _review_document
            return _review_document(checked)
        if name == "read_document_image":
            from .documents import extract_document_image
            image = extract_document_image(checked, args["relationship"], part=args.get("part", "word/document.xml"))
            if editor.space.file(args["path"]) != path or digest(checked) != expected:
                raise ValueError("Selected file changed during image extraction")
            if on_image is not None and image.metadata["preview_supported"]:
                on_image(image)
            return {**image.metadata, "preview_available": on_image is not None and image.metadata["preview_supported"]}
        if name in _LEGACY_TOOLS:
            raise ValueError(".doc/.xls COM 写预览通道已冻结：实现保留，但真机往返未验证；"
                             "本次不启动写探针、不生成计划。重新启用须独立授权。")
        classification = classify_document(checked)
        if not classification.readable and not (name == "com_read" and
                checked.suffix.lower() in {".doc", ".xls"}) and not (
                name in {"plan_legacy_word_replace", "plan_legacy_cell_set"} and
                checked.suffix.lower() in {".doc", ".xls"}):
            raise ValueError("File is unreadable or risky; only inspection or bounded legacy COM read allowed")
        if name == "image_gen":
            nonlocal image_batch_planned
            if checked.suffix.lower() not in {".docx", ".xlsx"}:
                raise ValueError("DOCX or XLSX required")
            if image_count > 1 and image_batch_planned:
                raise ValueError("本次图片张数已规划；请先逐张确认或发起新任务")
            # Revalidate even for direct callers bypassing the agent's tool schema.
            validated = ImageGenArgs.model_validate(args)
            prompts = validated.prompts or [validated.prompt] * image_count
            if len(prompts) != image_count:
                raise ValueError("逐张提示词数量必须等于用户指定的图片张数")
            service = _image_service()
            if image_count > 1:
                image_batch_planned = True  # Partial failure cannot trigger a second batch.
            plans = []
            try:
                for prompt_text in prompts:
                    request = ImageRequest(prompt=prompt_text, size=validated.size,
                                           style=validated.style, n=validated.n)
                    plans.append(service.plan(request))
            except ImageGenError as error:
                if error.code in {"image_configuration_missing", "image_configuration_invalid"}:
                    raise ValueError(
                        f"云端生图未配置或配置无效（{error.code}）：请在 .env 设置有效的 "
                        "IMAGE_PROVIDER/IMAGE_API_KEY/IMAGE_BASE_URL/IMAGE_MODEL；未调用任何生图服务"
                    ) from None
                raise ValueError(f"生图计划失败（{error.code}）；未发送云端请求，请重新开始任务") from None
            if len({p["plan_id"] for p in plans}) != len(plans):
                raise ValueError("生图计划标识重复；未发送云端请求，请重新开始任务")
            for index, public in enumerate(plans, start=1):
                if on_plan is not None:
                    on_plan(("image_plan", {**public, "batch_index": index, "batch_total": image_count}
                             if image_count > 1 else public))
            return {"plan_id": plans[0]["plan_id"], "preview": plans[0],
                    "plan_ids": [p["plan_id"] for p in plans],
                    "note": "仅生成逐张计划：未调用任何生图服务、未写文件；每张都必须由用户在界面独立确认费用后"
                            "才会请求云端，未保存图片只在内存中，逐张保存新 PNG 还需另一次确认"}
        if name == "com_read":
            from .com import capabilities, extract_office
            # Availability is checked before opening anything so an unusable COM path
            # reports the documented fallback instead of a raw automation error.
            office = capabilities().get("excel" if checked.suffix.lower() in {".xls", ".xlsx"} else "word", {})
            def legacy_fallback():
                if editor.space.file(args['path']) != checked or digest(checked) != expected:
                    raise ValueError('Selected file changed before legacy fallback')
                from .legacy_reader import read_legacy_document
                try:
                    result = read_legacy_document(checked, include_tables=args['include_tables'], region=args['region'])
                    if digest(checked) != expected:
                        raise ValueError('Selected file changed during legacy read')
                    return result
                except Exception as error:
                    raise ValueError('Office COM is unavailable; bounded LibreOffice legacy read refused (' +
                                     (getattr(error, 'code', None) or 'conversion_failed') + ')') from None
            if not office.get("available") and checked.suffix.lower() in {'.doc', '.xls'}:
                return legacy_fallback()
            if not office.get("available"):
                raise ValueError("Office COM is unavailable in this environment: "
                                 f"{office.get('note', 'detection failed')}. "
                                 "Use the Python read tools (inspect/read_word_region/read_excel_region) instead.")
            try:
                return extract_office(str(checked), include_tables=args["include_tables"], region=args["region"])
            except Exception:
                if checked.suffix.lower() in {'.doc', '.xls'}:
                    return legacy_fallback()
                raise
        if name == "ocr_docx_image":
            if ocr_provider is None:
                raise ValueError("No OCR provider configured")
            return ocr_docx_image(checked, args["relationship"], ocr_provider)
        if name == "read_word_text":
            if checked.suffix.lower() != ".docx":
                raise ValueError("DOCX required")
            from .documents import read_word_text
            return read_word_text(checked, offset=args["offset"], limit=args["limit"],
                                  max_chars=args["max_chars"])
        if name == "summarize_word_document":
            if checked.suffix.lower() != ".docx":
                raise ValueError("DOCX required")
            from .documents import summarize_word_document
            return summarize_word_document(checked, chunk_size=args["chunk_size"],
                                           max_chunks=args["max_chunks"],
                                           preview_chars=args["preview_chars"])
        if name == "search":
            return search_document(checked, args["query"])[:50]
        if name == "read_word_region":
            if checked.suffix.lower() != ".docx":
                raise ValueError("DOCX required")
            from docx import Document
            doc = Document(checked)
            start = args["paragraph_index"]
            return [{"paragraph_index": i, "text": doc.paragraphs[i].text[:4096],
                     "text_truncated": len(doc.paragraphs[i].text) > 4096,
                     "style": doc.paragraphs[i].style.name if doc.paragraphs[i].style else None}
                    for i in range(start, min(start + args["limit"], len(doc.paragraphs)))]
        if name == "read_word_structure":
            if checked.suffix.lower() != ".docx":
                raise ValueError("DOCX required")
            from .documents import read_word_structure
            return read_word_structure(checked, limit=args["limit"])
        if name == "read_word_table_region":
            if checked.suffix.lower() != ".docx":
                raise ValueError("DOCX required")
            from .analytics import read_word_table_region
            return read_word_table_region(checked, args["table_index"], args["start_row"],
                args["start_column"], args["rows"], args["columns"])
        if name == "read_excel_region":
            if checked.suffix.lower() != ".xlsx":
                raise ValueError("XLSX required")
            if args["row"] + args["rows"] - 1 > 1048576 or args["column"] + args["columns"] - 1 > 16384:
                raise ValueError("Excel region exceeds sheet bounds")
            from openpyxl import load_workbook
            formulas = load_workbook(checked, data_only=False, read_only=True)
            cached = None
            try:
                cached = load_workbook(checked, data_only=True, read_only=True)
            except Exception:
                formulas.close()
                raise
            try:
                if args["sheet"] not in formulas:
                    titles = [sheet.title for sheet in formulas.worksheets][:5]
                    raise ValueError(f"worksheet not found; available worksheet titles (first 5): "
                                     f"{titles!r}. Use an exact title, do not guess.")
                ws, values = formulas[args["sheet"]], cached[args["sheet"]]
                cells = []
                for row in ws.iter_rows(min_row=args["row"], max_row=args["row"] + args["rows"] - 1,
                         min_col=args["column"], max_col=args["column"] + args["columns"] - 1):
                    for cell in row:
                        if cell.value is None:
                            continue
                        entry = {"cell": cell.coordinate, "value": cell.value, "data_type": cell.data_type}
                        if cell.data_type == "f":
                            entry = {"cell": cell.coordinate, "formula_text": cell.value,
                                     "cached_value": values[cell.coordinate].value,
                                     "cached_value_unverified": True}
                        cells.append(entry)
                return {"sheet": args["sheet"], "cells": cells}
            finally:
                formulas.close()
                if cached is not None:
                    cached.close()
        if name == "filter_sort_excel_view":
            if checked.suffix.lower() != ".xlsx":
                raise ValueError("XLSX required")
            from .analytics import filter_sort_rows
            return filter_sort_rows(checked, args["sheet"], args["start_row"], args["end_row"],
                args["start_column"], args["columns"], args["key_column"],
                args["equals"], args["descending"])
        if name == "summarize_excel_column":
            if checked.suffix.lower() != ".xlsx":
                raise ValueError("XLSX required")
            from .analytics import summarize_column
            return summarize_column(checked, args["sheet"], args["column"],
                args["start_row"], args["end_row"])
        if name == "read_excel_structure":
            if checked.suffix.lower() != ".xlsx":
                raise ValueError("XLSX required")
            from .documents import read_excel_structure
            return read_excel_structure(checked, limit=args["limit"])
        if name == "read_word_table_structure":
            if checked.suffix.lower() != ".docx":
                raise ValueError("DOCX required")
            from .documents import read_word_table_structure
            return read_word_table_structure(checked, table_index=args["table_index"], limit=args["limit"])
        if name == "batch_documents":
            if checked.suffix.lower() not in {".docx", ".xlsx"}:
                raise ValueError("DOCX or XLSX required")
            from .documents import batch_document_report
            return batch_document_report(checked, validate=editor.space.file,
                                         operation=args["operation"], query=args["query"],
                                         limit=args["limit"], per_file=args["per_file"])
        plan_model = _PLAN_MODELS.get(name)
        if plan_model is None:
            raise ValueError("Tool not registered")
        # The tool name already determines the request kind. A model that echoes a
        # different discriminator (observed: kind="word_replace" on the batch tool) must
        # not turn an otherwise valid call into a rejection; every semantic field is
        # still strictly validated by the request model below.
        plan = editor.plan(plan_model.model_validate(
            {**args, "kind": plan_model.model_fields["kind"].default}))
        if on_plan is not None:
            on_plan(("plan", plan.public()))
        return {"plan_id": plan.id, "preview": plan.public(),
                "note": "仅预览；用户必须在 UI 独立确认后才能写入"}

    return dispatch


def build_tools(approved: tuple[Path, ...], *, ocr_enabled: bool = False) -> list[ToolSpec]:
    """The complete model-visible tool catalog for one authorized file.

    Kept as a named function because these schemas are serialized into every model
    call: the context-budget regression test measures them directly.
    """
    tools = [
        ToolSpec(name='read_pptx_structure', description='只读PPTX大纲、备注/图表/图片计数，按next_offset分页', schema_model=PptxStructureArgs, access='read', allowed_paths=approved),
        ToolSpec(name='read_pptx_slide', description='只读某页文本框、备注、表格格、图表标题和图片局部锚点；按next_offset及next_text_offset续读', schema_model=PptxSlideArgs, access='read', allowed_paths=approved),
        ToolSpec(name='read_pdf', description='只读PDF文本层及坐标；按next_page及next_char_offset续读，扫描页不做OCR', schema_model=PdfReadArgs, access='read', allowed_paths=approved),
        ToolSpec(name='search_directory_documents', description='只读所选文件父目录一层至多20个安全DOCX/XLSX的倒排检索；来源和位置随命中返回，不调用embedding服务', schema_model=DirectorySearchArgs, access='read', allowed_paths=approved),
        ToolSpec(name="inspect", description="识别所选文档的结构与风险，返回简短文件报告", schema_model=InspectArgs, access="read", allowed_paths=approved),
        ToolSpec(name="review_document", description="只读审阅所选 DOCX/XLSX 并给出问题清单：格式不一致、未填占位符、标题跳级、手动编号断号、疑似敏感信息（脱敏显示）、空表格、整列为空、列类型混用、公式错误值、重复行、数值离群；确定性规则、不写文件、不改源文件，报告含实际扫描范围与 error/warn/info 分级", schema_model=ReviewArgs, access="read", allowed_paths=approved),
        ToolSpec(name="com_read", description="只读所选doc/docx/xls/xlsx；旧doc/xls在COM不可用时尝试受限LibreOffice转换兜底，宏/外链/BIFF公式等拒绝；≤20段或400格，原件不变，不解锁写入", schema_model=COMReadArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_document_image", description="按结构报告的 relationship 和 source_part 提取所选 DOCX/XLSX 单张内部图片（≤10MB）；仅元数据进入模型，Web 界面只读预览、不写回源文件，图片可能保存到有界只读会话历史；DOCX 正文默认 part=word/document.xml，XLSX 必须给 drawing 部件", schema_model=DocumentImageArgs, access="read", allowed_paths=approved),
        ToolSpec(name="image_gen", description="按用户明确张数规划图片，未说默认一张；共用prompt，不同画面用prompts。每张n=1。只预览，不调用云端；逐张费用确认，保存再独立确认", schema_model=ImageGenArgs, access="write", allowed_paths=approved),
        ToolSpec(name="search", description="在所选文档中查找文本，结果保存到磁盘；参数 query 为关键字", schema_model=SearchArgs, access="read", allowed_paths=approved),
        ToolSpec(name="summarize_word_document", description="读取所选 DOCX 摘要脚手架：标题大纲、每张表行/列数、分块字符数与短预览；正文与表格另用专用工具取，摘要你自己写", schema_model=WordSummaryArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_word_text", description="按文档顺序读取所选 DOCX 正文段落与表格单元格（要展示文本内容时首选）；参数可省略；若 has_more 为真，按各块 remainder_hint 或 next_offset 续读", schema_model=WordTextArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_word_region", description="读取所选 DOCX 有界正文段落与样式；段落从 0 开始，limit 至多 20；**不含表格内容**，整篇文本请用 read_word_text", schema_model=WordRegionArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_word_table_region", description="读取所选 DOCX 指定表格至多 20x20 单元格文本，不修改文件", schema_model=WordTableRegionArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_word_table_structure", description="只读解析所选 DOCX 指定表格的合并与嵌套结构：每个单元格的逻辑/物理坐标、gridSpan、vMerge 起点与延续、承载文本的 anchor，以及嵌套表格层级；不自洽时报告 merge_consistent=false，不编造合并", schema_model=WordTableStructureArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_word_structure", description="读取所选 DOCX 的页眉页脚、批注、图片位置、修订与内容控件（各至多 limit 条），不修改文件", schema_model=WordStructureArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_excel_region", description="读取所选 XLSX 指定工作表至多 20x20 单元格，公式文本与未验证缓存分开标记", schema_model=ExcelRegionArgs, access="read", allowed_paths=approved),
        ToolSpec(name="summarize_excel_column", description="汇总 XLSX 某列最多1000行数值字面量，返回总和/平均/极值；公式值不参与统计", schema_model=ExcelSummaryArgs, access="read", allowed_paths=approved),
        ToolSpec(name="filter_sort_excel_view", description="只读筛选/排序 XLSX 数据视图（≤100行x20列，公式行排除），不更改工作表", schema_model=ExcelViewArgs, access="read", allowed_paths=approved),
        ToolSpec(name="read_excel_structure", description="只读识别所选 XLSX 结构：命名区域、条件格式、隐藏行列、合并区域、图表清单、图片锚点；只列出不求值，每项带总数与截断标志", schema_model=ExcelStructureArgs, access="read", allowed_paths=approved),
        ToolSpec(name="plan_word", description="仅预览 DOCX 正文段落索引的局部替换，不直接写文件；按原格式保留 runs", schema_model=WordReplace, access="write", allowed_paths=approved),
        ToolSpec(name="plan_word_batch_replace", description="仅预览 DOCX 正文中某个字面量的全文替换（超上限则整体拒绝，不做部分替换）；默认最多 20 处，必须由用户另行独立确认，不直接写文件", schema_model=WordBatchReplace, access="write", allowed_paths=approved),
        ToolSpec(name="plan_word_format", description="仅预览 DOCX 正文单个纯文本 run 的加粗、斜体、下划线修改，不直接写文件", schema_model=WordFormat, access="write", allowed_paths=approved),
        ToolSpec(name="plan_word_table_cell", description="仅预览 DOCX 简单表格指定单元格段落的一处文本替换，不直接写文件", schema_model=WordTableCellReplace, access="write", allowed_paths=approved),
        ToolSpec(name="plan_word_table_row", description="仅预览对普通表格插入一个空行或删除一行（含嵌套表格或合并单元格的表格一律拒绝）；删除行必须由用户另行独立确认，不直接写文件", schema_model=WordTableRow, access="write", allowed_paths=approved),
        ToolSpec(name="plan_word_merged_cell_text", description="预览合并表格承载文本单元格的replace/append；逻辑行列，覆盖格或延续格拒绝；保持合并与样式，独立确认", schema_model=WordMergedCellText, access="write", allowed_paths=approved),
        ToolSpec(name="plan_word_revision_decision", description="仅预览接受或拒绝文档正文中的全部修订（accept_all/reject_all）：接受=插入内容保留、被删内容消失；拒绝=插入内容移除、被删内容还原为正文；含移动/行级/段落标记/表格修订时整体拒绝并说明，需独立确认，不直接写文件", schema_model=WordRevisionDecision, access="write", allowed_paths=approved),
        ToolSpec(name="plan_legacy_word_replace", description="冻结：遗留 .doc 写预览实现保留，但真机往返未验证；当前应用拒绝此调用且不启动 COM 写探针；重新启用须独立授权，不修改源文件", schema_model=LegacyWordReplace, access="write", allowed_paths=approved),
        ToolSpec(name="plan_legacy_cell_set", description="冻结：遗留 .xls 单格写预览实现保留，但真机往返未验证；当前应用拒绝此调用且不启动 COM 写探针；重新启用须独立授权，不修改源文件", schema_model=LegacyCellSet, access="write", allowed_paths=approved),
        ToolSpec(name="batch_documents", description="只读批量：对所选文件所在目录（仅下一层，≤20 个）的 DOCX/XLSX 做统一的识别/摘要/搜索，结果按文件汇总；越界候选、遗留格式与报风险文件一律跳过并列出原因；不写入任何文件", schema_model=BatchArgs, access="read", allowed_paths=approved),
        ToolSpec(name="plan_docx_merge", description="仅预览把所选 DOCX 同目录中的多个 DOCX（sources 为文件名，按顺序）合并为一个新 DOCX；各源保留自己的节设置与图片/链接关系，源文件不被修改，需独立确认", schema_model=DocxMerge, access="write", allowed_paths=approved),
        ToolSpec(name="plan_docx_split", description="仅预览按 Heading 层级（level≤3）把所选 DOCX 拆成多个新 DOCX；每个标题一个部分，源文件不被修改，需独立确认", schema_model=DocxSplit, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel", description="仅预览 XLSX 单个单元格的新值，不直接写文件，公式不会计算", schema_model=ExcelSet, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_trim", description="仅预览 XLSX 指定文本单元格的首尾空格清理，要求精确原值，不直接写文件", schema_model=ExcelTrim, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_range_trim", description="仅预览 XLSX 矩形区域内文本字面量的首尾空格清理（≤200格；公式/数字/已整洁格不动并列出跳过坐标，无改动则整体拒绝），需独立确认", schema_model=ExcelRangeTrim, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_range_set", description="仅预览 XLSX 矩形区域字面量块写入（≤100格）：先用 read_excel_region 读出该区域，按实际形状逐格提供 expected（空白用 null）与 values；值+类型须精确匹配，公式/合并格拒绝，整块全改或全拒，需独立确认", schema_model=ExcelRangeSet, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_sort", description="仅预览普通单表 XLSX 有界区域的稳定排序写入（不含表头，≤100格、≤20列）；先用 read_excel_region 读取完整区域并逐格提供 expected，key_cell 必须是首行完整 A1 地址；仅允许字面量、无公式/合并/超链接，格式按坐标保留，需独立确认", schema_model=ExcelSort, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_chart", description="仅预览普通单表 XLSX 图表：柱状图/折线图/饼图，≤8 列同行的有限数值字面量（饼图仅1列），可带分类列、标题、轴标题、图例位置；A1 锚点，需独立确认，不计算公式", schema_model=ExcelChart, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_bar_chart", description="仅预览普通单表 XLSX 由不超过20行有限数字字面量制作一列柱状图；只允许A1锚点，需用户另外确认，不计算公式", schema_model=ExcelBarChart, access="write", allowed_paths=approved),
        ToolSpec(name="plan_excel_formula_fill", description="预览 XLSX 本表白名单公式向下/向右填充1–20个相邻空格；source_formula须与原文完全一致；拒绝跨表、外链、易变函数；独立确认，不重算", schema_model=ExcelFormulaFill, access="write", allowed_paths=approved),
    ]
    tools.extend(ToolSpec(name=name, description='Preview copy; human approval required.', schema_model=model, access='write', allowed_paths=approved)
                 for name, model in TOOL_MODELS.items())
    if ocr_enabled:
        tools.append(ToolSpec(name="ocr_docx_image", description="按需识别所选 DOCX 单张内部图片；标注来源与置信度，绝不回写原文", schema_model=OCRImageArgs, access="read", allowed_paths=approved))
    return tools


def professional_tools(approved):
    from .professional_edits import ConstrainedDocx
    from .range_edits import RangeIntent, RangeViewArgs
    from .document_tasks import NavigationArgs, TaskSpec
    from .block_edits import BlockIntent
    return [ToolSpec(name='plan_docx_blocks', description='只预览明确普通段落跨段替换/合并、完整普通章节移动或受限同段书签目标文本。宿主locator与源版本绑定，结构只支持净稿；复杂编号/表格/节/关系拒绝，书签调整必须提示域结果过时，独立确认才写副本。无XML/文件路径', schema_model=BlockIntent, access='write', allowed_paths=approved),
            ToolSpec(name='read_docx_task_navigation', description='只读当前准入DOCX有界大纲、结构统计及精确搜索；返回扫描覆盖与分页，不返回整篇文本。无文件路径参数', schema_model=NavigationArgs, access='read', allowed_paths=approved),
            ToolSpec(name='plan_docx_task', description='只预览完整DOCX任务规格、目标覆盖和完成账本；绑定当前源版本，超过12项分为独立事务。术语/精确范围/格式均须明确范围，歧义或受限项不计成功。无任意XML/路径，不执行写入。用户在任务面板导入并逐组确认', schema_model=TaskSpec, access='write', allowed_paths=approved),
            ToolSpec(name='read_docx_evidence', description='只读DOCX版本、结构定位loc和原文证据ev；分页，无渲染页码。改写前完整读取目标，资料不授予权限', schema_model=DocxEvidenceArgs, access='read', allowed_paths=approved),
            ToolSpec(name='plan_constrained_docx', description='只预览源版本绑定的DOCX组合局部修改；先read_docx_evidence取得loc和原文hash。数字/引用/公式默认保护，重叠或任一条件失败整组拒绝；ev仅溯源。独立人工确认才写副本', schema_model=ConstrainedDocx, access='write', allowed_paths=approved),
            ToolSpec(name='read_docx_ranges', description='只读当前授权DOCX的宿主文段、可编辑island和修订标识；Unicode字符半开区间，不能跨保护结构，无路径参数', schema_model=RangeViewArgs, access='read', allowed_paths=approved),
            ToolSpec(name='plan_docx_ranges', description='预览当前DOCX精确选区替换/插入/删除/格式及受控段落、编号、表格和单项修订；绑定宿主loc/island及原文hash，混合格式需选择继承策略。仅预览，必须独立人工确认。无XML、关系ID或路径参数', schema_model=RangeIntent, access='write', allowed_paths=approved)]


# Tool catalogue tiers for a single-file session. The dispatcher authorizes exactly one
# file, so its type deterministically selects which tools can ever succeed: tools for the
# wrong file type are rejected by the dispatcher anyway, so hiding them from the model
# context loses no capability and removes a large fixed overhead (measure with
# examples/measure_context_budget.py). ``kind`` is classify_document(path).kind, whose
# values are "word" / "spreadsheet" / "legacy" / "ooxml" / "unknown".
_SHARED_TOOLS = frozenset({"inspect", "com_read"})
# Tools that read both DOCX and XLSX, so they belong to the word and spreadsheet tiers
# but never to the legacy tier (which only inspects and, when the COM probe passes, edits).
_CROSS_TOOLS = frozenset({"search", "batch_documents", "read_document_image", "review_document"})
_WORD_TOOLS = frozenset({
    "summarize_word_document", "read_word_text", "read_word_region",
    "read_word_table_region", "read_word_table_structure", "read_word_structure",
    "ocr_docx_image",
    "plan_word", "plan_word_batch_replace", "plan_word_format", "plan_word_table_cell",
    "plan_word_table_row", "plan_word_merged_cell_text", "plan_word_revision_decision",
    "plan_docx_merge", "plan_docx_split",
})
_SPREADSHEET_TOOLS = frozenset({
    "read_excel_region", "summarize_excel_column", "filter_sort_excel_view",
    "read_excel_structure",
    "plan_excel", "plan_excel_trim", "plan_excel_range_trim", "plan_excel_range_set",
    "plan_excel_chart", "plan_excel_bar_chart", "plan_excel_formula_fill", "plan_excel_sort",
})
_LEGACY_TOOLS = frozenset({"plan_legacy_word_replace", "plan_legacy_cell_set"})
# P3 generation plans a brand-new image and only needs a selected DOCX/XLSX as the audited
# session anchor; it writes nothing until generation and then saving are separately approved.
_IMAGE_TOOLS = frozenset({"image_gen"})

_KIND_TOOLS = {
    "word": _SHARED_TOOLS | _CROSS_TOOLS | _WORD_TOOLS | _IMAGE_TOOLS,
    "spreadsheet": _SHARED_TOOLS | _CROSS_TOOLS | _SPREADSHEET_TOOLS | _IMAGE_TOOLS,
    "legacy": _SHARED_TOOLS | _LEGACY_TOOLS,
    'presentation': _SHARED_TOOLS | {'read_pptx_structure', 'read_pptx_slide'},
    'presentation_edit': _SHARED_TOOLS | {'read_pptx_structure', 'read_pptx_slide', 'plan_pptx_replace', 'plan_pptx_slide'},
    'pdf': _SHARED_TOOLS | {'read_pdf'},
    'directory': _SHARED_TOOLS | {'search_directory_documents'},
    'advanced': _SHARED_TOOLS | set(TOOL_MODELS),
}


def select_tools(tools: list[ToolSpec], kind: str) -> list[ToolSpec]:
    """Filter the full catalogue down to the tier that can act on one file type.

    ``kind`` is ``classify_document(path).kind``. Unknown kinds keep only the shared read
    tools (inspect / com_read). Nothing is ever added beyond ``tools``; each returned spec
    is the same object with the same access, allowlist and strict schema.
    """
    keep = _KIND_TOOLS.get(kind, _SHARED_TOOLS)
    if kind in {'word', 'advanced'}:
        keep = keep | {'read_docx_evidence', 'plan_constrained_docx', 'read_docx_ranges', 'plan_docx_ranges',
                       'read_docx_task_navigation', 'plan_docx_task', 'plan_docx_blocks'}
    return [tool for tool in tools if tool.name in keep]


class EmptyArgs(StrictModel):
    """No-argument schema for the write-gate tool."""


# The model never sees the plan_* tools until it asks for them: calling this gate unlocks
# the write tier for the remaining decisions. Its schema is empty, so its fixed cost is a
# handful of tokens while it keeps 18 plan_* schemas out of every read-only task.
_WRITE_GATE = ToolSpec(
    name="request_write_tools",
    description="当任务需要创建或修改文档内容（替换/格式化/表格增删改/合并单元格/图表/公式填充/区域写入/合并或拆分文档等）或需要预览云端生成新图片（image_gen，仍需费用与保存两次独立确认）时，先调用本工具解锁 plan_* 预览工具；解锁后仍必须由用户在界面单独确认才能真正写入，本工具本身不写任何文件",
    schema_model=EmptyArgs,
    access="none",
)


def split_read_write(tools: list[ToolSpec]) -> tuple[list[ToolSpec], list[ToolSpec]]:
    """Split a tier into the read catalogue (plus the write gate) and the write tools.

    Write tools enter context only after the model calls the gate, so their large schemas
    stay out of every read-only task; a read request never pays for the write catalogue.
    The gate is added only when there is at least one write tool to unlock, so an
    inspect-only tier (unknown file type) never advertises write tools it does not have.
    """
    read = [tool for tool in tools if tool.access != "write"]
    write = [tool for tool in tools if tool.access == "write"]
    if write:
        read = read + [_WRITE_GATE]
    return read, write


_STATUS_HELP = {
    "done": "任务完成。",
    "limit": "任务已安全停止；未执行任何未经你确认的写入。", 
    "budget": "上下文预算在压缩后仍超过 90%，已在调用模型前安全停止：未执行任何写入。",
    "confirmation_required": "等待你在界面上单独确认读取；代理不会自行继续。",
    "running": "本轮尚未结束，可从检查点继续。",
}


def describe_state(payload: dict) -> str:
    """Plain-language outcome shown ahead of the raw JSON checkpoint payload."""
    status = payload.get("status")
    limit = payload.get("decision_limit", DEFAULT_DECISION_LIMIT)
    status_help = (f"已达到 {limit} 次模型决策上限并安全停止：没有继续循环，也没有执行任何未经你确认的写入。"
                   if status == "limit" and limit is not None and
                   any(f"Stopped after {limit} model decisions" in str(e) for e in payload.get("errors") or [])
                   else _STATUS_HELP.get(status, status))
    steps = str(payload.get('iterations')) + (f"/{limit}" if limit is not None else "")
    lines = [f"结果：{status_help}",
             f"迭代 {steps} · 估算 token {payload.get('token_estimate')}/"
             f"{payload.get('context_capacity')} · 压缩 {payload.get('compressions')} 次"]
    if payload.get("answer"):
        lines.append("模型回复：" + str(payload["answer"]))
    errors = payload.get("errors") or []
    for error in errors[-2:]:
        lines.append("最近错误：" + str(error))
    if status in {"budget", "limit", "running", "confirmation_required"}:
        if payload.get("resume_checkpoint"):
            lines.append("可续跑检查点：" + str(payload["resume_checkpoint"]))
        lines.append("上方任何已预览的修改计划仍需你另行确认才会写入；未确认则原文件不变。")
    if status == "budget":
        lines.append("处理办法：提高 OFFICE_CONTEXT_WINDOW（当前工具描述本身约占固定开销），"
                     "或把任务拆成更小的步骤后重新提交。")
    return "\n".join(lines)


_MAX_RESULT_PREVIEW = 6000
_MAX_SEARCH_PREVIEW = 2000
_GENERIC_RENDER_LINES = 2000


def render_generic_result(value: Any, *, limit: int = _GENERIC_RENDER_LINES) -> str:
    """Loss-free fallback rendering for a read tool that has no bespoke renderer yet.

    Read results are the only place the user can see document content, so an unhandled
    read tool must still show everything it returned instead of an empty panel. The line
    budget is reported explicitly rather than dropped silently.
    """
    lines: list[str] = []
    truncated = False

    def walk(item: Any, indent: int) -> None:
        nonlocal truncated
        pad = "  " * indent
        if isinstance(item, dict):
            for key, entry in item.items():
                if len(lines) >= limit:
                    truncated = True
                    return
                if isinstance(entry, (dict, list)) and entry:
                    lines.append(f"{pad}{key}:")
                    walk(entry, indent + 1)
                else:
                    lines.append(f"{pad}{key}: {entry!r}")
        elif isinstance(item, (list, tuple)):
            for index, entry in enumerate(item):
                if len(lines) >= limit:
                    truncated = True
                    return
                if isinstance(entry, (dict, list, tuple)) and entry:
                    lines.append(f"{pad}[{index}]")
                    walk(entry, indent + 1)
                else:
                    lines.append(f"{pad}[{index}] {entry!r}")
        else:
            lines.append(f"{pad}{item!r}")

    walk(value, 0)
    if truncated:
        lines.append(f"…（本条结果超过通用渲染的 {limit} 行上限，其余内容未在此展开；"
                     "请让我缩小读取范围或分段继续读取）")
    return "\n".join(lines)


def _paragraph_line(entry: dict) -> str:
    """One paragraph line, keeping the tool's own truncation flag visible."""
    line = f"[段 {entry.get('paragraph_index')}] {entry.get('text', '')}"
    if entry.get("text_truncated"):
        line += "（本段更长，已被工具截断）"
    return line


def render_tool_result(tool: str, result: Any) -> str | None:
    """Human-readable rendering of a read result, for the window log.

    The model only ever receives a short excerpt (tool results are summarised so the
    context stays small), so it cannot echo a document back to the user. When the user
    asks to *see* content, the retrieved text has to be shown here instead.
    """
    if tool == "read_word_region" and isinstance(result, list):
        return "\n".join(_paragraph_line(p) for p in result if isinstance(p, dict))
    if tool == "search" and isinstance(result, list):
        hits = [f"[{hit.get('location', hit.get('sheet', ''))}] {hit.get('match', '')}"
                for hit in result if isinstance(hit, dict)]
        text = "\n".join(hits)
        if len(text) > _MAX_SEARCH_PREVIEW:
            # Never drop the tail silently: say what was held back and how to get it.
            return (text[:_MAX_SEARCH_PREVIEW] +
                    f"\n…（命中内容较长，此处只显示前 {_MAX_SEARCH_PREVIEW} 字符；"
                    f"共 {len(hits)} 处命中、{len(text)} 字符，请缩小关键字或让我继续查看）")
        return text
    if not isinstance(result, dict):
        return None
    if tool == "read_document_image":
        lines = [f"文档内部图片：{result.get('part')} · {result.get('relationship')}",
                 f"原始字节：{result.get('byte_count')}；仅元数据进入模型，不修改源文件。"]
        if result.get("bytes_decoded") is True:
            lines.append(f"元数据已解码：{result.get('format')}，{result.get('width')} × {result.get('height')}；非内容识别。")
        else:
            lines.append("元数据未解码（bytes_decoded=false），不推断宽高或格式。")
        if not result.get("preview_available"):
            lines.append("当前界面/此格式不能提供安全图片预览。")
        lines.append(str(result.get("note", "")))
        return "\n".join(lines)
    if tool == "review_document":
        labels = {"error": "错误", "warn": "警告", "info": "提示"}
        counts = result.get("counts") or {}
        scanned = result.get("scanned") or {}
        scope = "、".join(f"{key} {value}" for key, value in scanned.items())
        lines = [f"审阅结果：错误 {counts.get('error', 0)} / 警告 {counts.get('warn', 0)}"
                 f" / 提示 {counts.get('info', 0)}",
                 f"扫描范围：{scope}" + ("（已截断）" if result.get("truncated") else "")]
        findings = result.get("findings") or []
        if not findings:
            lines.append("未在扫描范围内发现问题。")
        for finding in findings:
            if not isinstance(finding, dict):
                lines.append(str(finding))
                continue
            lines.append(
                f"· [{labels.get(finding.get('severity'), finding.get('severity'))}] "
                f"{finding.get('title', '')}（{finding.get('location', '')}）\n"
                f"    证据：{finding.get('evidence', '')}\n"
                f"    建议：{finding.get('suggestion', '')}")
        return "\n".join(lines)
    if tool == "read_word_text":
        lines = []
        for block in result.get("blocks", []):
            if not isinstance(block, dict):
                lines.append(str(block))
                continue
            if block.get("kind") == "table":
                lines.append(f"[表格 {block.get('index')}: {block.get('row_count')} 行 × "
                             f"{block.get('column_count')} 列]")
                for row in block.get("rows", []):
                    lines.append("  | " + " | ".join(row) + " |")
                if block.get("rows_truncated"):
                    lines.append(f"  （本表还有 {block.get('rows_omitted', '若干')} 行未显示："
                                 f"{block.get('remainder_hint', '')}）")
            else:
                lines.append(block.get("text", ""))
                if block.get("text_truncated"):
                    lines.append("  （本段文本被截断：" + str(block.get(
                        "remainder_hint") or "可用 read_word_region 按段索引继续读取") + "）")
        text = "\n".join(lines)
        if result.get("has_more"):
            text += (f"\n（文档共 {result.get('total_blocks')} 个块，本次到第 "
                     f"{result.get('next_offset')} 块；要继续请让我用 next_offset="
                     f"{result.get('next_offset')} 再读一次）")
        return text
    if tool == "read_word_region":
        return "\n".join(_paragraph_line(p) for p in result if isinstance(p, dict))
    if tool == "read_word_table_region":
        lines = [f"[表格 {result.get('table_index')}: {result.get('total_rows')} 行 × "
                 f"{result.get('total_columns')} 列]"]
        for row in result.get("rows", []):
            lines.append("  | " + " | ".join(cell.get("text", "") for cell in row.get("cells", [])) + " |")
        return "\n".join(lines)
    if tool == "read_excel_region":
        return "\n".join(f"{cell.get('cell')}: "
                         + (f"{cell.get('formula_text')}（公式文本；缓存值 "
                            f"{cell.get('cached_value')!r} 未经验证，未被重新计算）"
                            if "formula_text" in cell else repr(cell.get("value")))
                         for cell in result.get("cells", []))
    if tool == "read_word_structure":
        lines = []
        for index, section in enumerate(result.get("sections", [])):
            for kind in ("header", "footer"):
                for item in section.get(kind, []):
                    lines.append(f"[{kind} {index}] {item.get('text', '')}")
        for comment in result.get("comments", []):
            lines.append(f"[批注 {comment.get('id')} by {comment.get('author')}] {comment.get('text', '')}")
        for image in result.get("images", []):
            lines.append(f"[图片] {image.get('placement')} 位置 {image.get('position')}")
        # Tracked revisions and content controls are safety-relevant: a document with
        # pending revisions must not look clean just because the panel omitted them.
        changes = result.get("tracked_changes", {})
        for entry in changes.get("insertions", []):
            lines.append(f"[修订·插入 {entry.get('author')}] {entry.get('text', '')}")
        for entry in changes.get("deletions", []):
            lines.append(f"[修订·删除 {entry.get('author')}] {entry.get('text', '')}")
        if changes.get("insertions_truncated") or changes.get("deletions_truncated"):
            lines.append(f"（修订共 {changes.get('insertion_count')} 处插入 / "
                         f"{changes.get('deletion_count')} 处删除，此处只显示前若干条）")
        if changes.get("parts_unscanned"):
            lines.append(f"（未扫描部件：{changes['parts_unscanned']}，不能据此认为没有修订）")
        for control in result.get("content_controls", []):
            lines.append(f"[内容控件 {control.get('control_type')} tag={control.get('tag')}] "
                         f"{control.get('text', '')}")
        return "\n".join(lines) or "（无页眉页脚、批注、图片、修订或内容控件）"
    if tool == "summarize_word_document":
        lines = [f"正文 {result.get('paragraph_count')} 段、表格 {result.get('table_count')} 张"]
        for table in result.get("tables", []):
            lines.append(f"  表格[{table.get('index')}] {table.get('rows')} 行 × {table.get('columns')} 列")
        for heading in result.get("headings", []):
            lines.append(f"  标题[{heading.get('index')}] {heading.get('text')}")
        return "\n".join(lines)
    if tool == "summarize_excel_column":
        return (f"[只读统计] 工作表 {result.get('sheet')} 第 {result.get('column')} 列"
                f"（行 {result.get('range_rows')}）：数值字面量 {result.get('count')} 个 · "
                f"合计 {result.get('sum')} · 平均 {result.get('average')} · "
                f"最小 {result.get('minimum')} · 最大 {result.get('maximum')}\n"
                f"公式格 {result.get('formula_cells_excluded')} 个未参与统计（公式不会被重新计算）· "
                f"空白 {result.get('empty')} · 非数值 {result.get('other')}")
    if tool == "filter_sort_excel_view":
        lines = [f"[只读视图] 工作表 {result.get('sheet')} · 数据行 {result.get('source_rows')} · "
                 f"列 {result.get('source_columns')} · 键列 {result.get('key_column')} · "
                 f"{'降序' if result.get('descending') else '升序'} · "
                 f"排除公式行 {result.get('formula_rows_excluded')} 行（未修改工作表）"]
        for row in result.get("rows", []):
            values = " | ".join(str(value) for value in row.get("values", []))
            lines.append(f"  第 {row.get('row')} 行 · 键 {row.get('key')!r} | {values}")
        return "\n".join(lines)
    if tool == "read_word_table_structure":
        # Merged cells are a safety fact: a coordinate covered by a gridSpan or carried on by
        # a vMerge continuation must never look like an ordinary writable cell.
        lines = []
        if result.get("merge_consistent") is False:
            lines.append("！合并结构不自洽（merge_consistent=false），不要据此定位坐标：")
            for issue in result.get("merge_issues", []):
                lines.append(f"  - {issue}")
        lines.append(f"[表格 {result.get('table_index', 0)}："
                     f"{result.get('row_count', '?')} 行 × {result.get('column_count', '?')} 列"
                     f"（grid_columns={result.get('grid_columns', '?')}）]")
        for cell in result.get("cells", []):
            span = cell.get("gridSpan", 1) or 1
            merge = "vMerge 延续" if cell.get("vMerge") == "continue" else (
                "vMerge 起点" if cell.get("vMerge") else "")
            notes = []
            if span > 1:
                span_text = f"跨 {span} 列"
                if cell.get("last_column") is not None:
                    span_text += f"（至第 {cell.get('last_column')} 列）"
                notes.append(span_text)
            if merge:
                notes.append(merge)
            if not cell.get("anchor"):
                notes.append("非承载文本")
            text = (cell.get("text") or "").replace("\n", " ")
            suffix = ("（" + "、".join(notes) + "）") if notes else ""
            lines.append(f"  第{cell.get('row')}行 第{cell.get('column')}列: {text}{suffix}")
        for continuation in result.get("continuations", []):
            lines.append(f"  [延续] 第{continuation.get('row')}行 第{continuation.get('column')}列"
                         f" 归属起点 第{continuation.get('origin_row', continuation.get('row'))}行"
                         f" 第{continuation.get('origin_column', continuation.get('column'))}列")
        for nested in result.get("nested_tables", []):
            lines.append(f"  [嵌套表格] 路径 {nested.get('path')} 深度 {nested.get('depth')}")
        return "\n".join(lines)
    if tool == "batch_documents":
        # A directory-level read must never look like "the folder contains only these files":
        # the skip list and its reasons are part of the answer, and the read-only guarantee is
        # stated outright.
        lines = [f"[批量只读 · {result.get('operation')}] 目录 {result.get('directory')}"
                 f"：处理 {result.get('file_count', 0)} 个文件，跳过 {result.get('skipped_count', 0)} 项"
                 f"{'，读取错误 ' + str(result.get('error_count')) + ' 项' if result.get('error_count') else ''}"
                 "；**未写入任何文件**（wrote_nothing="
                 f"{str(result.get('wrote_nothing')).lower()}）"]
        if result.get("note"):
            lines.append(f"  说明：{result.get('note')}")
        for entry in result.get("files", []):
            name = entry.get("file") or entry.get("name") or "?"
            detail = entry.get("summary") or entry.get("report") or entry.get("excerpt") or entry
            lines.append(f"  · {name}")
            lines.append(_render_generic_block(detail, indent=2))
        for skipped in result.get("skipped", []):
            lines.append(f"  × 跳过 {skipped.get('file', '?')}：{skipped.get('reason', '未说明原因')}")
        for failure in result.get("errors", []):
            lines.append(f"  × 读取失败 {failure.get('file', '?')}：{failure.get('error', '未说明原因')}")
        if result.get("truncated"):
            lines.append(f"  （目录内候选项超过 limit={result.get('limit')}，"
                         "仅扫描了前若干个；其余文件未被读取，也不在本清单里）")
        return "\n".join(lines)
    if tool == "com_read" or tool.startswith(("read_", "list_")):
        # A read-shaped tool with no bespoke renderer must still show the user its content.
        return render_generic_result(result)
    return None


# ---------------------------------------------------------------------------
# Desktop presentation layer
#
# Everything from here to the end of Desktop decides only *how* existing state and events
# are displayed. It never decides what may be read, previewed or written: authorization
# stays in writes.py, the preview/confirm gate stays in execute()/drain(), and the event
# kinds stay exactly as agent.py sends them. Widgets are created by Desktop.__init__ only,
# so tests and headless checks that build Desktop.__new__(Desktop) still drive the same
# logic through the guards below.
# ---------------------------------------------------------------------------

# Families are probed at run time, so no machine-specific font has to exist. The monospace
# candidates are CJK-capable on purpose: document content is mostly Chinese and a
# Latin-only mono font would render it as boxes.
_UI_FONT_CANDIDATES = ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "Noto Sans CJK SC",
                       "Source Han Sans SC", "PingFang SC", "DejaVu Sans", "Arial")
_MONO_FONT_CANDIDATES = ("Sarasa Mono SC", "Noto Sans Mono CJK SC", "Microsoft YaHei Mono",
                         "Noto Sans Mono", "DejaVu Sans Mono")

# --- design tokens ----------------------------------------------------------
# Every colour, spacing step, radius, type size and border width lives here; the layout code
# references tokens instead of raw numbers, so switching a theme is a palette swap and the
# spacing scale cannot drift between cards.
_SPACING = {1: 4, 2: 8, 3: 12, 4: 16, 5: 24, 6: 32}          # space-1 .. space-6
_RADIUS = {"sm": 3, "md": 6, "lg": 10}                        # radius-sm/md/lg
_TYPE_SCALE = {"caption": 9, "body": 10, "body_bold": 10, "heading": 11, "title": 15,
               "status": 12, "badge": 9, "mono": 10}          # caption .. title ladder
_BORDER_WIDTH = 1
_FOCUS_WIDTH = 2
_THEME_ENV = "OFFICE_AGENT_THEME"
_THEME_MODES = ("auto", "light", "dark")

# Layout tokens: Tk measures panes in text lines/characters and progress bars in percent, so
# those metrics get names too instead of raw numbers inside the build methods.
_WINDOW_SIZE = (1120, 800)
_WINDOW_MIN = (820, 620)
_COMPACT_HEIGHT = 780          # below this the rest of the fixed chrome compacts as well
_DENSE_HEIGHT = 880            # below this the pane footnotes give their space to the panes
_MIN_PANE_HEIGHT = 60          # every reading pane keeps at least this much height
_PANE_LINES = {"log": 16, "plan": 9, "result": 12}     # pane heights, in text lines
_PANE_CHARS = {"log": 34, "plan": 26, "result": 26}    # pane widths, in characters
_PANE_WEIGHTS = {"log": 3, "panels": 2, "plan": 2, "result": 3}
_WRAP_MARGIN = 56          # space kept around a label that wraps to the window width
_MIN_WRAP_WIDTH = 240      # never wrap status text narrower than this
_BUDGET_PERCENT_MAX = 100.0


@dataclass(frozen=True)
class Palette:
    """One complete colour scheme.

    Field names keep the semantics of the historical ``_PALETTE`` mapping
    (window/card/border/ink/muted/accent/ok/warn/danger plus their soft/dark variants), so a
    reader of that mapping cannot misread a colour. ``soft`` values are the card backgrounds
    of status chips, ``dark`` values are pressed/hover variants of the same hue, and ``faint``
    is only ever used on a ``card`` surface (it does not reach AA against ``window``).
    """

    name: str
    window: str
    window_alt: str
    card: str
    card_alt: str
    border: str
    border_strong: str
    ink: str
    muted: str
    faint: str
    accent: str
    accent_dark: str
    accent_soft: str
    on_solid: str
    ok: str
    ok_soft: str
    ok_dark: str
    warn: str
    warn_soft: str
    warn_dark: str
    danger: str
    danger_soft: str
    danger_dark: str
    focus: str
    selection: str

    def as_dict(self) -> dict:
        return asdict(self)


# Contrast of every text-on-surface pair used by the window is asserted in
# tests/test_app_resume.py (WCAG AA, >= 4.5:1) for both themes.
LIGHT_THEME = Palette(
    name="light",
    window="#eceff4", window_alt="#e3e9f0", card="#ffffff", card_alt="#f7f9fc",
    border="#c6cfda", border_strong="#99a7b7",
    ink="#1b2530", muted="#5b6a7a", faint="#6b7887",
    accent="#14508c", accent_dark="#0e3c6b", accent_soft="#e2ecf8", on_solid="#ffffff",
    ok="#1a6b42", ok_soft="#e2f3e8", ok_dark="#12512f",
    warn="#8a5300", warn_soft="#fdf1da", warn_dark="#6d4100",
    danger="#a3271d", danger_soft="#fbe4e1", danger_dark="#7d1a13",
    focus="#1a73c8", selection="#cfe3f8",
)

DARK_THEME = Palette(
    name="dark",
    window="#1d232b", window_alt="#232b34", card="#262e38", card_alt="#2c3540",
    border="#3c4653", border_strong="#57657a",
    ink="#e7edf4", muted="#a7b4c3", faint="#909eae",
    accent="#63a6ea", accent_dark="#4b8fdc", accent_soft="#24384d", on_solid="#10161d",
    ok="#5fc98d", ok_soft="#1f3a2c", ok_dark="#7fd7a4",
    warn="#e3b04b", warn_soft="#3d3526", warn_dark="#f0c46b",
    danger="#ef8a80", danger_soft="#43282a", danger_dark="#f2a49b",
    focus="#7cc0ff", selection="#2f4a66",
)

THEMES = {"light": LIGHT_THEME, "dark": DARK_THEME}
# Compatibility mapping: downstream readers of the historical key names still resolve them
# (to the light theme); the window itself always goes through the active Palette.
_PALETTE = LIGHT_THEME.as_dict()

# One Text tag per severity keeps a long transcript scannable. Tag names are stable; the
# colours come from the active theme in log_tag_specs(). "safety" and "budget" exist so that
# refusals and context compressions carry a visual marker instead of blending into the stream.
_LOG_TAG_NAMES = ("body", "meta", "head", "user", "plan", "ok", "warn", "error", "result",
                  "safety", "budget")

# Symbol in front of a badge label: colour plus shape plus words, so the state survives
# colour-blindness and monochrome output. All four exist in GB2312-era CJK fonts.
_BADGE_MARKS = {"info": "●", "ok": "√", "warn": "!", "danger": "×"}

# Agent status -> presentation level. The vocabulary belongs to office_agent.agent; this
# mapping only picks a colour for the status card.
_STATUS_LEVELS = {"done": "ok", "limit": "warn", "budget": "warn",
                  "confirmation_required": "warn", "running": "info"}

_MAX_STATUS_ERROR = 600       # the status card is bounded; the transcript keeps the full text
_MAX_RESULT_PANEL = 400_000   # pathological result size; reported explicitly when it bites
_DPI_MIN, _DPI_MAX = 1.0, 2.5  # bounds for DPI-derived pixel scaling
_BUDGET_REMINDER = 0.8        # share of the context window that triggers the on-screen reminder


def budget_advice(fraction: float | None) -> str:
    """Warning line shown once the context window is largely used, or "" when it is fine.

    Presentation only, and derived from the same ``budget_fraction`` the meter uses, so the
    bar, the sentence and the reminder can never disagree.
    """
    if fraction is None or fraction < _BUDGET_REMINDER:
        return ""
    level = "接近上限" if fraction < 0.9 else "已达上限区"
    return (f"注意：上下文已用 {fraction:.0%}（{level}）。继续读取会先触发压缩，"
            "必要时会提前安全停止；可减少一次任务的读取范围，或提高 OFFICE_CONTEXT_WINDOW。")


def theme_for(name: Any) -> Palette:
    """Palette for a theme name; an unknown name falls back to the light theme."""
    return THEMES.get(str(name or "").lower(), LIGHT_THEME)


def default_theme_mode() -> str:
    """Initial theme mode: OFFICE_AGENT_THEME when it names a mode, else follow the OS."""
    override = os.getenv(_THEME_ENV, "").strip().lower()
    return override if override in _THEME_MODES else "auto"


def detect_system_theme() -> str:
    """'dark' or 'light' from the OS preference; light when it cannot be determined.

    Windows is this window's supported host, so the preference comes from the personalise
    registry value. A missing key, a missing winreg or an unexpected value means "light":
    the window never guesses a dark scheme it cannot verify.
    """
    try:
        import winreg
    except ImportError:
        return "light"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            value, _kind = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "light" if int(value) else "dark"
    except Exception:
        return "light"


def capability_detection_notice(reason: str) -> str:
    """Neutral wording for the startup capability *detection*.

    Startup is not a failure: nothing was attempted and no file is at risk, so this must not
    borrow the "失败，原文件保留" language that belongs to a real refusal. The backend's reason
    is embedded verbatim, and the text says explicitly that only detection ran.
    """
    return ("遗留 .doc/.xls 编辑需要本机可用且允许 COM 自动化的 Microsoft Office；"
            f"当前检测结果：{reason}"
            "（仅做启动检测，未运行写入探针；.docx/.xlsx 的读取与写入不受影响。）")


# The backend's own vocabulary for a closed legacy channel; the window's only text-shape
# dependence, kept in one place so it can be checked against com.describe_write_refusal().
_LEGACY_REFUSAL_MARKERS = ("COM", ".doc/.xls")


def legacy_refusal_presentation(text: Any) -> str:
    """Error text as shown when a legacy edit is refused.

    A refused legacy edit is the one error that must state both the backend's reason *and*
    that nothing was written. The backend's sentence is kept untouched and the guarantee is
    appended — never paraphrased, never reordered — so the user still reads the difference
    between "automation is blocked here" and "Office is not installed". Every other error is
    passed through unchanged.
    """
    raw = str(text)
    if all(marker in raw for marker in _LEGACY_REFUSAL_MARKERS):
        return raw + "（本次未做任何写入，原文件保持不变。）"
    return raw


def log_tag_specs(theme: Palette) -> dict:
    """Transcript tag colours for the active theme (tag names stay stable)."""
    return {
        "body": {"foreground": theme.ink},
        "meta": {"foreground": theme.muted},
        "head": {"foreground": theme.accent, "font": "heading"},
        "user": {"foreground": theme.accent, "font": "body_bold"},
        "plan": {"foreground": theme.warn, "font": "body_bold", "background": theme.warn_soft},
        "ok": {"foreground": theme.ok, "font": "body_bold"},
        "warn": {"foreground": theme.warn, "background": theme.warn_soft},
        "error": {"foreground": theme.danger, "font": "body_bold"},
        "result": {"foreground": theme.ink},
        # Refusals/guard errors and context compressions get their own marks.
        "safety": {"foreground": theme.danger, "font": "body_bold",
                   "background": theme.danger_soft},
        "budget": {"foreground": theme.warn, "background": theme.warn_soft},
    }


def status_level(status: Any) -> str:
    """Presentation level (info/ok/warn/danger) for an agent status value."""
    return _STATUS_LEVELS.get(str(status or ""), "info")


# Agent status -> (level, short label) for the status badge. The wording is deliberately
# explicit text, not colour alone, so the state survives a colour-blind or monochrome view.
_STATUS_BADGES = {"running": ("info", "进行中"),
                  "done": ("ok", "已完成"),
                  "limit": ("warn", "已安全停止"),
                  "budget": ("warn", "上下文预算不足"),
                  "confirmation_required": ("warn", "等待你的确认")}


def status_badge(status: Any) -> tuple[str, str]:
    """(level, label) for an agent status value; an unknown value is shown verbatim."""
    key = str(status or "")
    if key in _STATUS_BADGES:
        return _STATUS_BADGES[key]
    return ("info", key or "尚未开始")


def _display_width(text: str) -> int:
    """Display cells of a string: East-Asian wide characters count as two."""
    return sum(2 if unicodedata.east_asian_width(char) in "WF" else 1 for char in text)


def align_table_text(text: str) -> str:
    """Pad ``|``-separated table rows so columns line up in a display font.

    Display-only: ``render_tool_result`` stays the single source of the table content, and
    this never changes words or row order — it only inserts spaces, and only in rows that look
    like the renderer's table lines (``| a | b |``). Rows with no table shape are returned
    untouched.
    """
    rows = text.split("\n")
    parsed: list[tuple[int, list[str], str, str]] = []
    for index, row in enumerate(rows):
        stripped = row.strip()
        if "|" not in stripped or not (stripped.startswith("|") or stripped.startswith("  |")):
            continue
        indent = row[:len(row) - len(row.lstrip())]
        cells = stripped.strip("|").split("|")
        parsed.append((index, [cell.strip() for cell in cells], indent,
                       "|" if stripped.endswith("|") else ""))
    if len(parsed) < 2:
        return text
    widths: dict[int, int] = {}
    for _index, cells, _indent, _tail in parsed:
        for position, cell in enumerate(cells):
            widths[position] = max(widths.get(position, 0), _display_width(cell))
    for index, cells, indent, _tail in parsed:
        padded = [cell + " " * (widths[position] - _display_width(cell))
                  for position, cell in enumerate(cells)]
        rows[index] = indent + "| " + " | ".join(padded) + " |"
    return "\n".join(rows)


class StatusBadge:
    """A rounded status chip: colour plus shape plus words.

    ttk has no chip widget and clam draws no corner radius, so the chip is a canvas rounded
    rectangle — this is where the ``_RADIUS`` tokens are actually consumed. The badge follows
    a StringVar, so the existing code paths keep setting text exactly as before.
    """

    _COLOURS = {"info": ("accent", "accent_soft"), "ok": ("ok", "ok_soft"),
                "warn": ("warn", "warn_soft"), "danger": ("danger", "danger_soft")}

    def __init__(self, parent: tk.Misc, *, textvariable, fonts: dict, theme: Palette):
        self.canvas = tk.Canvas(parent, highlightthickness=0, borderwidth=0, height=1)
        self._textvariable = textvariable
        self._fonts = fonts
        self._theme = theme
        self.level = "info"
        self._pad_x, self._pad_y = 8, 3
        self._trace = textvariable.trace_add("write", lambda *_args: self._redraw())

    def set(self, level: str, label: str) -> None:
        """Set the chip's level (colour) and label (words)."""
        self.level = level if level in self._COLOURS else "info"
        self._textvariable.set(label)
        self._redraw()

    def apply_theme(self, theme: Palette) -> None:
        self._theme = theme
        self._redraw()

    def _redraw(self) -> None:
        theme = self._theme
        foreground_name, background_name = self._COLOURS[self.level]
        foreground, background = getattr(theme, foreground_name), getattr(theme, background_name)
        mark = _BADGE_MARKS.get(self.level, "●")
        label = self._textvariable.get()
        text = f"{mark} {label}" if label else mark
        font = self._fonts["badge"]
        width = tkfont.Font(root=self.canvas, font=font).measure(text) + 2 * self._pad_x
        height = tkfont.Font(root=self.canvas, font=font).metrics("linespace") + 2 * self._pad_y
        self.canvas.configure(width=width, height=height, background=theme.card)
        self.canvas.delete("all")
        radius = min(_RADIUS["md"], height // 2)
        self.canvas.create_polygon(self._round_points(0, 0, width, height, radius),
                                   smooth=True, fill=background, outline=background)
        self.canvas.create_text(width / 2, height / 2, text=text, fill=foreground, font=font)

    @staticmethod
    def _round_points(x1: float, y1: float, x2: float, y2: float, r: int) -> list[float]:
        """Corner points for a smoothed polygon (Tk's stand-in for a rounded rectangle)."""
        return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
                x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]


def next_step_hint(*, has_file: bool, busy: bool, pending: bool, verified: bool,
                   level: str = "info") -> str:
    """The one line that answers "what do I click next?", from window state alone.

    Presentation only: it names the step the UI already offers, it never changes what the
    window permits. Order matters: a running task outranks everything, an unwritten plan
    outranks the file prompt and a previous output, and an error outranks the idle prompt.
    """
    if busy:
        return "下一步：等待本轮结束（按钮已临时禁用；不会执行任何未经你确认的写入）。"
    if pending:
        return ("下一步：核对上方『修改计划 / 差异预览（未写入）』，确认无误后点击"
                "『确认并保存输出副本』，并在弹窗中再次确认。")
    if not has_file:
        return "下一步：点击『选择授权文件』选择要处理的文档；该文件所在目录即授权目录。"
    if verified:
        return ("下一步：可查看输出副本与审计记录，或点击『撤销最近输出副本』（需再次确认）；"
                "也可以直接提交下一个任务。")
    if level == "danger":
        return ("下一步：查看会话记录中的完整错误；原文件保持不变。"
                "若是本机 Office 能力提示，可改用 .docx/.xlsx 的读取与写入工具继续。")
    return ("下一步：在『任务指令』输入自然语言任务后按 Enter（或点击『自然语言查询/规划』）；"
            "模型只做预览，写入必须由你单独确认。")


def _first_installed_family(candidates: tuple[str, ...], fallback: str) -> str:
    """First candidate family Tk reports as installed, else ``fallback`` (never a guess)."""
    try:
        installed = {name.casefold() for name in tkfont.families()}
    except Exception:
        return fallback
    for name in candidates:
        if name.casefold() in installed:
            return name
    return fallback


def budget_fraction(tokens: Any, capacity: Any) -> float | None:
    """Share of the context window in use (0..1), or None when it cannot be computed."""
    try:
        tokens, capacity = int(tokens), int(capacity)
    except (TypeError, ValueError):
        return None
    if capacity <= 0:
        return None
    return max(0.0, min(1.0, tokens / capacity))


def _budget_percent(tokens: Any, capacity: Any) -> str:
    """Readable token usage; an unusable capacity must never crash the window."""
    fraction = budget_fraction(tokens, capacity)
    return "未知" if fraction is None else f"{fraction:.0%}"


def format_progress_line(iterations: Any, tokens: Any, capacity: Any, compressions: Any,
                         decision_limit: int | None = DEFAULT_DECISION_LIMIT) -> str:
    """The progress sentence in the status card: iterations, tokens, compressions."""
    steps = str(iterations) + (f"/{decision_limit}" if decision_limit is not None else "")
    return (f"迭代 {steps} · 估算 token {tokens}/{capacity}"
            f"（{_budget_percent(tokens, capacity)}）· 压缩 {compressions} 次")


def _label_block(pairs: list[tuple[str, Any]]) -> list[str]:
    """Render ``label：value`` lines, keeping multi-line values on their own block."""
    lines: list[str] = []
    for label, value in pairs:
        text = str(value)
        if "\n" in text:
            lines.append(f"{label}：")
            lines.extend("    " + row for row in text.splitlines())
        else:
            lines.append(f"{label}：{text}")
    return lines


_PLAN_LABELS = (("sources", "将读取的其它文件（同目录，仅这些，按此顺序）"),
                ("id", "计划 ID"), ("kind", "操作类型"), ("source", "源文件"), ("target", "修改目标"),
                ("before", "修改前"), ("after", "修改后"), ("risks", "风险提示"),
                ("confirmation_required", "需要你单独确认"))

# Keys whose list value reads as an enumeration of names, not as a Python list.
_PLAN_LIST_KEYS = {"sources", "risks", "parts", "targets"}

_OUTPUT_LABELS = (("id", "任务 ID"), ("status", "状态"), ("kind", "操作类型"), ("source", "源文件"),
                  ("output", "输出副本"), ("backup", "备份"), ("audit", "审计记录"),
                  ("source_sha256", "源文件 SHA-256"), ("output_sha256", "输出副本 SHA-256"),
                  ("note", "说明"))


def format_plan_preview(plan: dict) -> str:
    """Readable view of an *unwritten* ``Plan.public()`` payload.

    Depends only on the documented plan keys (id/kind/source/target/before/after/risks/
    confirmation_required). Any other key is appended verbatim, so a change of plan shape
    can add information but can never hide it.
    """
    if not isinstance(plan, dict):
        return (f"【尚未写入】计划载荷不是字典，原样显示：{plan!r}\n"
                "（未确认前不会写入任何文件）")
    pairs: list[tuple[str, Any]] = []
    for key, label in _PLAN_LABELS:
        if key not in plan:
            continue
        value = plan[key]
        if key == "confirmation_required":
            value = "是（必须由你单独确认）" if value else "否"
        elif key in _PLAN_LIST_KEYS and isinstance(value, (list, tuple)):
            # Name lists read as an enumeration: a merge that will read A.docx/B.docx must be
            # legible before the user confirms, since that is the only gate on the wider read.
            value = "；".join(str(item) for item in value) if value else "（无）"
        pairs.append((label, value))
    lines = ["【尚未写入】以下修改计划等待你确认：在你点击『确认并保存输出副本』并在弹窗中确认之前，"
             "不会备份、不会生成输出副本，原文件保持不变。", ""]
    lines += _label_block(pairs)
    if plan.get("sources"):
        lines += ["", "除了上面列出的源文件外，本计划只会读取上面『将读取的其它文件』里的文件；"
                      "未列出的文件不会被打开，也不会被修改。"]
    extra = {key: value for key, value in plan.items() if key not in dict(_PLAN_LABELS)}
    if extra:
        lines.append("")
        lines.append("其他字段：" + json.dumps(extra, ensure_ascii=False, indent=2))
    if not pairs and not extra:
        lines.append("（计划内容为空；未确认前不会写入任何文件）")
    return "\n".join(lines)


def format_output_summary(payload: dict, *, undone: bool = False) -> str:
    """Readable view of a verified/undone payload; every key is shown, none is dropped."""
    header = ("【已撤销】只删除了审计哈希未变的输出副本；源文件与备份保留，仍可用备份人工恢复。"
              if undone else
              "【已写入副本】写入已完成并通过重开验证；源文件字节未改动，备份保留用于人工恢复。")
    pairs: list[tuple[str, Any]] = []
    if isinstance(payload, dict):
        for key, label in _OUTPUT_LABELS:
            if key not in payload:
                continue
            value = payload[key]
            if key == "status":
                value = {"verified": "verified（已重开验证）"}.get(str(value), value)
            pairs.append((label, value))
        diff = payload.get("diff")
        if isinstance(diff, dict):
            pairs.append(("差异", f"{diff.get('target')}：修改前 {diff.get('before')!r} → "
                                  f"修改后 {diff.get('after')!r}"))
    lines = [header, ""]
    lines += _label_block(pairs)
    extra = ({key: value for key, value in payload.items()
              if key not in dict(_OUTPUT_LABELS) and key != "diff"}
             if isinstance(payload, dict) else {})
    if extra:
        lines.append("")
        lines.append("其他字段：" + json.dumps(extra, ensure_ascii=False, indent=2))
    lines += ["", ("输出副本已不存在（撤销完成）；源文件与备份未修改。" if undone else
                   "撤销只删除哈希未变的输出副本；源文件与备份始终保留。")]
    return "\n".join(lines)


def _render_generic_block(value: Any, *, indent: int = 0, limit: int = 40) -> str:
    """Indented, bounded rendering of one per-file entry inside the batch report."""
    pad = "  " * indent
    if isinstance(value, dict):
        keys = [key for key in ("readable", "kind", "paragraph_count", "table_count",
                                "sheet_count", "match_count", "headings", "message", "note")
                if key in value]
        parts = [f"{key}={value[key]!r}" for key in keys]
        if parts:
            return pad + " | ".join(parts)
    text = render_generic_result(value)
    lines = text.split("\n")
    if len(lines) > limit:
        return "\n".join(pad + line for line in lines[:limit]) + \
            f"\n{pad}…（该文件的报告超过 {limit} 行，其余未展开）"
    return "\n".join(pad + line for line in lines)


def _cap_panel_text(text: str, limit: int = _MAX_RESULT_PANEL) -> tuple[str, str]:
    """Panel body plus an explicit note when a pathological result is capped.

    Normal results are never touched: this only bites far beyond what a document read
    returns, and the note says exactly what was held back.
    """
    if len(text) <= limit:
        return text, ""
    return (text[:limit],
            f"（本条结果过长：面板只显示前 {limit} 字符，共 {len(text)} 字符；"
            "会话记录中同样是有界预览，可让我缩小范围或分段继续读取）")


def _clip_repr(value: Any, limit: int = 2000) -> str:
    """repr() for diagnostics, bounded with an explicit marker instead of a silent cut."""
    text = repr(value)
    if len(text) <= limit:
        return text
    return text[:limit] + f"…（诊断信息只显示前 {limit} 字符，共 {len(text)} 字符）"


class Desktop:
    """Tk window: pick one file, describe a task, review the plan, confirm the write.

    Presentation only. Authorization, preview/confirm and audit logic stay in writes.py and
    the dispatcher; this class decides where each existing event is displayed.
    """

    # Widgets exist only after __init__. Tests drive the same logic through
    # Desktop.__new__(Desktop), so widget access is guarded by _panels_ready and every
    # widget has a class-level default.
    log = None
    _panels_ready = False
    _log_tags = frozenset(_LOG_TAG_NAMES)
    _theme = LIGHT_THEME
    _theme_mode = "auto"
    _theme_name = "light"
    _dpi_scale = 1.0
    _progress_state: tuple = (0, "—", "—", 0)
    _actions_narrow = False
    _compact = False
    _dense = False
    _pending_capability_notice: str | None = None
    _checkpoint_placeholder = True
    _status_level = "info"

    def __init__(self, root: tk.Tk, *, ocr_provider: OCRProvider | None = None):
        if tk is None:
            raise RuntimeError("The desktop UI needs a Windows Python with tkinter/Tcl-Tk installed")
        self.root = root
        self.ocr_provider = ocr_provider  # Explicit host injection; no provider imports or credentials in core.
        self.events: queue.Queue = queue.Queue()
        self.path: Path | None = None
        self.editor: Editor | None = None
        self.model_agent: OfficeAgent | None = None
        self.pending_id: str | None = None
        self.last_verified_id: str | None = None
        self.busy = False
        self._stop_task = threading.Event()
        self._model_task_active = False
        self.last_compressions = 0
        self.session_metadata: dict | None = None
        self._resume_session: tuple[Path, dict] | None = None
        self._fonts: dict[str, tuple] = {}
        self._action_buttons: tuple = ()
        self._theme_mode = default_theme_mode()
        self._theme_name = detect_system_theme() if self._theme_mode == "auto" else self._theme_mode
        self._theme = theme_for(self._theme_name)
        self._dpi_scale = self._detect_dpi_scale()   # before the layout is measured
        self._progress_state = (0, "—", "—", 0)
        self.style = ttk.Style(root)
        self._build_window()
        self._panels_ready = True
        self._apply_theme(self._theme_mode)   # styles + recolours + theme button label
        self._reset_panels()
        self._refresh_file_labels()
        self._set_busy_ui(False)
        self.prompt.focus_set()
        self.root.after(150, self.drain)
        self.root.after(500, self.probe)  # Startup check runs in a worker; UI stays responsive.

    # -- theme --------------------------------------------------------------
    def _sp(self, step: int) -> int:
        """A spacing token in device pixels, scaled for the current display."""
        return max(1, int(round(_SPACING[step] * self._dpi_scale)))

    def _detect_dpi_scale(self) -> float:
        """Pixel scale for this display, so paddings match a scaled font size.

        Tk sizes fonts in points (already DPI-aware) while our paddings are pixels; without
        this the layout looks cramped at 125%/150%. Fails closed to 1.0 when the display
        cannot be queried.
        """
        try:
            dpi = float(self.root.winfo_fpixels("1i"))
        except Exception:
            return _DPI_MIN
        return min(_DPI_MAX, max(_DPI_MIN, dpi / 96.0))

    def _resolve_theme(self, mode: str) -> tuple[str, str]:
        """Resolve a mode to (palette name, mode); "auto" follows the OS preference."""
        mode = mode if mode in _THEME_MODES else "auto"
        return (detect_system_theme() if mode == "auto" else mode), mode

    def _apply_theme(self, mode: str | None = None) -> None:
        """Re-style the live window for the resolved theme.

        No widget is created or destroyed here: the session log, the read-result pane, the
        pending plan and every checkpoint message survive a switch untouched.
        """
        name, mode = self._resolve_theme(mode or self._theme_mode)
        self._theme_mode, self._theme_name = mode, name
        self._theme = theme_for(name)
        self._configure_styles(self._theme)
        self._retheme_widgets(self._theme)
        self._set_progress(*self._progress_state)   # bar colours follow the theme too
        self._update_theme_button()

    def cycle_theme(self) -> None:
        """Low-key switch in the header: auto -> light -> dark -> auto."""
        order = _THEME_MODES
        self._apply_theme(order[(order.index(self._theme_mode) + 1) % len(order)])

    def _retheme_widgets(self, theme: Palette) -> None:
        """Recolour the widgets ttk styles cannot reach (plain Text/Tk surfaces)."""
        if not self._panels_ready:
            return
        self.root.configure(background=theme.window)
        self.status_badge.apply_theme(theme)
        for widget in (self.log, self.plan_text, self.result_text):
            widget.configure(background=theme.card, foreground=theme.ink,
                             insertbackground=theme.ink, selectbackground=theme.selection,
                             selectforeground=theme.ink, highlightbackground=theme.border,
                             highlightcolor=theme.focus)
        self._configure_log_tags()

    def _configure_log_tags(self) -> None:
        """Apply the active theme's colours to every transcript tag."""
        for name, spec in log_tag_specs(self._theme).items():
            self.log.tag_configure(name, **{
                key: (self._fonts[value] if key == "font" else value)
                for key, value in spec.items()})

    def _update_theme_button(self) -> None:
        if not self._panels_ready:
            return
        names = {"light": "浅色", "dark": "深色"}
        label = (f"跟随系统（{names.get(self._theme_name, self._theme_name)}）"
                 if self._theme_mode == "auto" else names.get(self._theme_mode, self._theme_mode))
        # Plain CJK text only: the label must render in every CJK font, so no symbol glyph.
        self.theme_var.set(f"主题：{label}（点击切换）")

    # -- window construction ------------------------------------------------
    def _make_fonts(self) -> dict[str, tuple]:
        """Resolve font families that really exist on this machine."""
        try:
            default_family = tkfont.nametofont("TkDefaultFont", root=self.root).actual("family")
        except Exception:
            default_family = "TkDefaultFont"
        ui = _first_installed_family(_UI_FONT_CANDIDATES, default_family)
        mono = _first_installed_family(_MONO_FONT_CANDIDATES, ui)
        size = _TYPE_SCALE
        return {"body": (ui, size["body"]), "body_bold": (ui, size["body_bold"], "bold"),
                "small": (ui, size["caption"]), "heading": (ui, size["heading"], "bold"),
                "title": (ui, size["title"], "bold"), "status": (ui, size["status"], "bold"),
                "badge": (ui, size["badge"], "bold"), "mono": (mono, size["mono"])}

    def _configure_styles(self, theme: Palette | None = None):
        """Colours, typography and metrics in one place; 'clam' is the theme that honours them."""
        theme = theme or self._theme
        style, fonts = self.style, self._fonts
        gap = self._sp
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=theme.window, foreground=theme.ink, font=fonts["body"])
        style.configure("TFrame", background=theme.window)
        style.configure("Card.TFrame", background=theme.card)
        style.configure("Banner.TFrame", background=theme.warn_soft)
        style.configure("TLabel", background=theme.window, foreground=theme.ink)
        style.configure("Card.TLabel", background=theme.card, foreground=theme.ink)
        style.configure("Title.TLabel", font=fonts["title"], background=theme.window)
        style.configure("Subtitle.TLabel", font=fonts["small"], foreground=theme.muted)
        style.configure("Task.TLabel", font=fonts["body_bold"], foreground=theme.accent)
        style.configure("FileName.TLabel", font=fonts["body_bold"])
        style.configure("Path.TLabel", font=fonts["small"], foreground=theme.muted)
        style.configure("Muted.TLabel", foreground=theme.muted)
        style.configure("Faint.TLabel", foreground=theme.faint, background=theme.card)
        style.configure("Heading.TLabel", font=fonts["heading"], foreground=theme.accent)
        style.configure("Warn.TLabel", font=fonts["body_bold"], foreground=theme.warn)
        style.configure("Safety.TLabel", font=fonts["body_bold"], foreground=theme.warn,
                        background=theme.warn_soft)
        style.configure("Danger.TLabel", font=fonts["body_bold"], foreground=theme.danger)
        style.configure("BannerWarn.TLabel", background=theme.warn_soft, foreground=theme.warn,
                        font=fonts["body_bold"])
        for level, colour in (("Info", theme.ink), ("Ok", theme.ok),
                              ("Warn", theme.warn), ("Danger", theme.danger)):
            style.configure(f"Status{level}.TLabel", background=theme.card, foreground=colour,
                            font=fonts["status"])
        style.configure("Badge.TLabel", font=fonts["badge"], foreground=theme.muted)
        for name, colour, background in (("Info", theme.accent, theme.accent_soft),
                                         ("Warn", theme.warn, theme.warn_soft),
                                         ("Ok", theme.ok, theme.ok_soft),
                                         ("Danger", theme.danger, theme.danger_soft)):
            style.configure(f"Badge{name}.TLabel", font=fonts["badge"], foreground=colour,
                            background=background, padding=(gap(2), gap(1)))
        # Three visual layers: the status card and a pending plan own the strong border; the
        # transcript/result panes stay quiet so they read as history.
        style.configure("Section.TLabelframe", background=theme.window, bordercolor=theme.border,
                        relief="solid", borderwidth=_BORDER_WIDTH)
        style.configure("Section.TLabelframe.Label", background=theme.window,
                        foreground=theme.accent, font=fonts["heading"])
        style.configure("Plan.TLabelframe", background=theme.window, bordercolor=theme.border_strong,
                        relief="solid", borderwidth=_BORDER_WIDTH)
        style.configure("Plan.TLabelframe.Label", background=theme.window,
                        foreground=theme.accent, font=fonts["heading"])
        style.configure("PlanPending.TLabelframe", background=theme.window,
                        bordercolor=theme.warn, relief="solid", borderwidth=_FOCUS_WIDTH)
        style.configure("PlanPending.TLabelframe.Label", background=theme.window,
                        foreground=theme.warn, font=fonts["heading"])
        style.configure("Quiet.TLabelframe", background=theme.window, bordercolor=theme.border,
                        relief="solid", borderwidth=_BORDER_WIDTH)
        style.configure("Quiet.TLabelframe.Label", background=theme.window,
                        foreground=theme.muted, font=fonts["heading"])
        style.configure("TNotebook", background=theme.window, bordercolor=theme.border,
                        tabmargins=(gap(1), gap(1), 0, 0), borderwidth=_BORDER_WIDTH)
        style.configure("TNotebook.Tab", background=theme.window_alt, foreground=theme.muted,
                        font=fonts["small"], padding=(gap(3), gap(1)))
        style.map("TNotebook.Tab",
                  background=[("selected", theme.card), ("active", theme.card_alt)],
                  foreground=[("selected", theme.accent), ("active", theme.ink)])
        style.configure("TButton", padding=(gap(3), gap(1)), font=fonts["body"],
                        focuscolor=theme.focus, focusthickness=_FOCUS_WIDTH)
        style.configure("Accent.TButton", padding=(gap(4), gap(2)), font=fonts["body_bold"],
                        foreground=theme.on_solid, background=theme.accent)
        style.map("Accent.TButton",
                  background=[("disabled", theme.border), ("pressed", theme.accent_dark),
                              ("active", theme.accent_dark)],
                  foreground=[("disabled", theme.muted)])
        style.configure("Confirm.TButton", padding=(gap(4), gap(2)), font=fonts["body_bold"],
                        focuscolor=theme.focus, focusthickness=_FOCUS_WIDTH)
        style.configure("Pending.TButton", padding=(gap(4), gap(2)), font=fonts["body_bold"],
                        foreground=theme.on_solid, background=theme.warn)
        style.map("Pending.TButton",
                  background=[("disabled", theme.border), ("pressed", theme.warn_dark),
                              ("active", theme.warn_dark)],
                  foreground=[("disabled", theme.muted)])
        style.configure("Danger.TButton", foreground=theme.danger)
        style.configure("Link.TButton", padding=(gap(2), gap(1)), font=fonts["small"],
                        foreground=theme.accent)
        style.configure("TEntry", padding=gap(2), fieldbackground=theme.card,
                        bordercolor=theme.border, focuscolor=theme.focus,
                        focusthickness=_FOCUS_WIDTH, insertcolor=theme.ink)
        for name, colour, trough in (("Budget", theme.accent, theme.accent_soft),
                                     ("BudgetWarn", theme.warn, theme.warn_soft),
                                     ("BudgetDanger", theme.danger, theme.danger_soft)):
            style.configure(f"{name}.Horizontal.TProgressbar", troughcolor=trough,
                            background=colour, bordercolor=theme.border, lightcolor=colour,
                            darkcolor=colour, thickness=gap(3))
        # The iteration meter is a separate bar: thin, accent-coloured, no status colours —
        # running out of decisions is not the same failure as running out of context.
        style.configure("Meter.Horizontal.TProgressbar", troughcolor=theme.window_alt,
                        background=theme.accent, bordercolor=theme.border,
                        lightcolor=theme.accent, darkcolor=theme.accent, thickness=gap(2))

    def _autowrap(self, label: ttk.Label, container: tk.Misc, margin: int | None = None) -> None:
        """Keep long status/notice text inside the window instead of forcing it wider."""
        margin = _WRAP_MARGIN if margin is None else margin
        # Set a bounded wrap immediately: a long error arriving before the first <Configure>
        # must not be able to widen the window on its own.
        label.configure(wraplength=max(_MIN_WRAP_WIDTH, _WINDOW_SIZE[0] - margin))

        def resize(event):
            label.configure(wraplength=max(_MIN_WRAP_WIDTH, int(event.width) - margin))
        container.bind("<Configure>", resize, add="+")

    def _card(self, parent: tk.Misc, title: str, *, style: str = "Section.TLabelframe",
              padding: tuple | None = None) -> ttk.Labelframe:
        """One card container: titled frame, single stretching column, token padding.

        Every card in the window comes from here, so the three visual layers differ only by
        style name (`Section`/`Plan`/`PlanPending`/`Quiet`) and never by ad-hoc padding.
        """
        gap = self._sp
        frame = ttk.Labelframe(parent, text=f" {title} ", style=style,
                               padding=padding or (gap(2), gap(2)))
        frame.columnconfigure(0, weight=1)
        return frame

    def _scrolled_text(self, parent: tk.Misc, *, height: int, font_key: str,
                       width: int = 30) -> tk.Text:
        """Read-only text area: long content is scrolled, never clipped.

        ``width`` is deliberately small (in characters): the panes are sized by the paned
        window and the grid weights, and a wide default would otherwise force the whole
        window wider than the 1120px it asks for.
        """
        theme = self._theme
        return tk.Text(parent, height=height, width=width, wrap="word", state="disabled",
                       relief="flat", background=theme.card, foreground=theme.ink,
                       font=self._fonts[font_key], padx=self._sp(2), pady=self._sp(2),
                       highlightthickness=_FOCUS_WIDTH, highlightbackground=theme.border,
                       highlightcolor=theme.focus, insertbackground=theme.ink,
                       selectbackground=theme.selection, selectforeground=theme.ink)

    def _build_window(self):
        """Header, file strip, notice banner, status card, panes, prompt and actions.

        Every measurement comes from the spacing/type tokens, so both themes and every DPI
        setting share one layout scale.
        """
        root = self.root
        gap = self._sp
        self._fonts = self._make_fonts()
        self._configure_styles()
        root.title("LiX — 审核后写入")
        if os.name == "nt":
            root.iconbitmap(str(APP_ICON))
        root.geometry(f"{_WINDOW_SIZE[0]}x{_WINDOW_SIZE[1]}")
        root.minsize(*_WINDOW_MIN)
        root.configure(background=self._theme.window)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)  # only the transcript/panels area grows

        header = ttk.Frame(root, padding=(gap(4), gap(4), gap(4), gap(1)))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="LiX · 审核后写入", style="Title.TLabel").grid(
            row=0, column=0, sticky="w")
        self.theme_var = tk.StringVar(value="◐ 主题：跟随系统")
        self.theme_button = ttk.Button(header, textvariable=self.theme_var, style="Link.TButton",
                                       command=self.cycle_theme)
        self.theme_button.grid(row=0, column=1, rowspan=2, sticky="ne")
        self.subtitle_label = ttk.Label(header,
                  text="选择文件 → 输入指令 → 查看计划 → 确认 → 写入 → 查看结果；"
                       "任何写入都要你单独确认。",
                  style="Subtitle.TLabel")
        self.subtitle_label.grid(row=1, column=0, sticky="w", pady=(gap(1), 0))
        self.task_var = tk.StringVar(value="当前任务：（尚未提交指令）")
        self.task_label = ttk.Label(header, textvariable=self.task_var, style="Task.TLabel",
                                    justify="left")
        self.task_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(gap(2), 0))
        self._autowrap(self.task_label, header)

        strip = ttk.Frame(root, padding=(gap(4), gap(2), gap(4), 0))
        strip.grid(row=1, column=0, sticky="ew")
        strip.columnconfigure(0, weight=1)
        strip_buttons = ttk.Frame(strip)
        strip_buttons.grid(row=0, column=0, sticky="w")
        self.select_button = ttk.Button(strip_buttons, text="选择授权文件", command=self.select,
                                        style="Accent.TButton")
        self.select_button.pack(side="left")
        self.probe_button = ttk.Button(strip_buttons, text="检测本机 Office", command=self.probe)
        self.probe_button.pack(side="left", padx=(gap(2), 0))
        self.file_var = tk.StringVar(value="当前文件：尚未选择")
        ttk.Label(strip, textvariable=self.file_var, style="FileName.TLabel").grid(
            row=0, column=1, sticky="e", padx=(gap(3), 0))
        self.dir_var = tk.StringVar(
            value="授权目录：选择文件后，该文件所在目录即授权目录；只有这一个文件会被授权读写。")
        self.dir_label = ttk.Label(strip, textvariable=self.dir_var, style="Path.TLabel",
                                   justify="left")
        self.dir_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(gap(2), 0))
        self._autowrap(self.dir_label, strip)

        self.notice_var = tk.StringVar(value="")
        self.notice_frame = ttk.Frame(root, style="Banner.TFrame", padding=(gap(4), gap(2)))
        self.notice_frame.grid(row=2, column=0, sticky="ew", padx=gap(4), pady=(gap(2), gap(2)))
        self.notice_frame.columnconfigure(0, weight=1)
        self.notice_label = ttk.Label(self.notice_frame, textvariable=self.notice_var,
                                      style="BannerWarn.TLabel", justify="left")
        self.notice_label.grid(row=0, column=0, sticky="ew")
        self._autowrap(self.notice_label, self.notice_frame)
        self.notice_frame.grid_remove()

        # Layer 1: the status card — what is happening, how far along, what to click next.
        self.status_card = card = self._card(root, "任务状态", padding=(gap(4), gap(3)))
        card.grid(row=3, column=0, sticky="ew", padx=gap(4), pady=(0, gap(2)))
        status_head = ttk.Frame(card)
        status_head.grid(row=0, column=0, sticky="ew")
        status_head.columnconfigure(1, weight=1)
        self.badge_var = tk.StringVar(value="尚未开始")
        self.status_badge = StatusBadge(status_head, textvariable=self.badge_var,
                                        fonts=self._fonts, theme=self._theme)
        self.status_badge.canvas.grid(row=0, column=0, sticky="w", padx=(0, gap(2)))
        self.badge_label = self.status_badge.canvas   # kept for callers/tests that resize it
        self.status = tk.StringVar(value="请选择一个文件，所在目录即授权目录；默认不覆盖原文件。")
        self.status_label = ttk.Label(status_head, textvariable=self.status, style="StatusInfo.TLabel",
                                      justify="left")
        self.status_label.grid(row=0, column=1, sticky="ew")
        self._autowrap(self.status_label, card, margin=_WRAP_MARGIN + gap(4))
        self.next_step_var = tk.StringVar(value="")
        self.next_step_label = ttk.Label(card, textvariable=self.next_step_var, style="Path.TLabel",
                                         justify="left")
        self.next_step_label.grid(row=1, column=0, sticky="ew", pady=(gap(1), 0))
        self._autowrap(self.next_step_label, card)
        # Progress sentence and budget bar share one container so compact mode can put them
        # side by side instead of stacking them.
        meter_row = ttk.Frame(card)
        meter_row.grid(row=2, column=0, sticky="ew", pady=(gap(3), 0))
        meter_row.columnconfigure(0, weight=1)
        self.meter_row = meter_row
        self.progress_var = tk.StringVar(value=format_progress_line(0, "—", "—", 0))
        self.progress_label = ttk.Label(meter_row, textvariable=self.progress_var, style="Path.TLabel")
        self.iteration_var = tk.StringVar(value="决策 0")
        self.iteration_label = ttk.Label(meter_row, textvariable=self.iteration_var,
                                         style="Path.TLabel")
        self.iteration_bar = ttk.Progressbar(meter_row, style="Meter.Horizontal.TProgressbar",
                                             maximum=1, value=0.0, length=gap(6) * 6)
        self.budget_bar = ttk.Progressbar(meter_row, style="Budget.Horizontal.TProgressbar",
                                          maximum=_BUDGET_PERCENT_MAX, value=0.0)
        self._layout_meter()
        self.checkpoint_var = tk.StringVar(value="续跑检查点：尚未运行任务")
        self.checkpoint_label = ttk.Label(card, textvariable=self.checkpoint_var,
                                          style="Path.TLabel", justify="left")
        self.checkpoint_label.grid(row=3, column=0, sticky="ew", pady=(gap(3), 0))
        self._autowrap(self.checkpoint_label, card)
        self.error_var = tk.StringVar(value="")
        self.error_label = ttk.Label(card, textvariable=self.error_var, style="Danger.TLabel",
                                     justify="left")
        self.error_label.grid(row=4, column=0, sticky="ew", pady=(gap(2), 0))
        self._autowrap(self.error_label, card)
        self.error_label.grid_remove()
        # Near-limit reminder and the indeterminate "working" bar: the two live signals that
        # keep a long model call from looking like a frozen window.
        self.limit_var = tk.StringVar(value="")
        self.limit_label = ttk.Label(card, textvariable=self.limit_var, style="Safety.TLabel",
                                     justify="left")
        self.limit_label.grid(row=5, column=0, sticky="ew", pady=(gap(2), 0))
        self._autowrap(self.limit_label, card)
        self.limit_label.grid_remove()
        self.busy_bar = ttk.Progressbar(card, style="Meter.Horizontal.TProgressbar",
                                        mode="indeterminate", maximum=100)
        self.busy_bar.grid(row=6, column=0, sticky="ew", pady=(gap(2), 0))
        self.busy_bar.grid_remove()

        body = ttk.PanedWindow(root, orient="horizontal")
        body.grid(row=4, column=0, sticky="nsew", padx=gap(4), pady=(0, gap(2)))
        self.body = body
        self.log_frame = log_frame = self._card(body, "会话记录（完整轨迹 · 可滚动）",
                                                style="Quiet.TLabelframe")
        right = ttk.Frame(body)
        body.add(log_frame, weight=_PANE_WEIGHTS["log"])
        body.add(right, weight=_PANE_WEIGHTS["panels"])

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = self._scrolled_text(log_frame, height=_PANE_LINES["log"], font_key="body",
                                      width=_PANE_CHARS["log"])
        self.log.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=log_scroll.set)
        self._configure_log_tags()
        ttk.Button(log_frame, text="复制全部记录", style="Link.TButton",
                   command=lambda: self._copy(self.log)).grid(row=1, column=0, columnspan=2,
                                                              sticky="e", pady=(gap(1), 0))

        # The right pane is a notebook, so each surface gets the full body height instead of
        # two squeezed halves — at the 820x620 minimum that is the difference between a
        # readable pane and a sliver. Which tab is selected is driven by the event stream.
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        self.notebook = ttk.Notebook(right)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Layer 2: the unwritten plan. It gets the stronger border (and an amber one while a
        # plan is waiting) because it is the safety-critical surface.
        self.plan_frame = plan_frame = self._card(self.notebook, "修改计划 / 差异预览（未写入）",
                                                  style="Plan.TLabelframe")
        plan_frame.rowconfigure(1, weight=1)
        plan_head = ttk.Frame(plan_frame)
        plan_head.grid(row=0, column=0, columnspan=2, sticky="ew")
        plan_head.columnconfigure(0, weight=1)
        self.plan_badge_var = tk.StringVar(value="○ 尚未生成计划")
        self.plan_badge = ttk.Label(plan_head, textvariable=self.plan_badge_var, style="Badge.TLabel")
        self.plan_badge.grid(row=0, column=0, sticky="w")
        self.plan_text = self._scrolled_text(plan_frame, height=_PANE_LINES["plan"],
                                             font_key="mono", width=_PANE_CHARS["plan"])
        ttk.Button(plan_head, text="复制", style="Link.TButton",
                   command=lambda: self._copy(self.plan_text)).grid(row=0, column=1, sticky="e")
        self.plan_text.grid(row=1, column=0, sticky="nsew", pady=(gap(2), 0))
        plan_scroll = ttk.Scrollbar(plan_frame, orient="vertical", command=self.plan_text.yview)
        plan_scroll.grid(row=1, column=1, sticky="ns", pady=(gap(2), 0))
        self.plan_text.configure(yscrollcommand=plan_scroll.set)
        self.plan_hint_var = tk.StringVar(
            value="计划在你点击『确认并保存输出副本』并在弹窗中确认之前不会写入任何文件。")
        self.plan_hint = ttk.Label(plan_frame, textvariable=self.plan_hint_var, style="Path.TLabel",
                                   justify="left")
        self.plan_hint.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(gap(1), 0))
        self._autowrap(self.plan_hint, plan_frame)

        # Layer 3: history and detail. Quiet border, muted section label.
        self.result_frame = result_frame = self._card(self.notebook, "读取结果（完整内容 · 可滚动）",
                                                      style="Quiet.TLabelframe")
        result_frame.rowconfigure(1, weight=1)
        result_head = ttk.Frame(result_frame)
        result_head.grid(row=0, column=0, columnspan=2, sticky="ew")
        result_head.columnconfigure(0, weight=1)
        self.result_badge_var = tk.StringVar(value="○ 暂无读取结果")
        self.result_badge = ttk.Label(result_head, textvariable=self.result_badge_var,
                                      style="Badge.TLabel")
        self.result_badge.grid(row=0, column=0, sticky="w")
        self.result_text = self._scrolled_text(result_frame, height=_PANE_LINES["result"],
                                               font_key="mono", width=_PANE_CHARS["result"])
        ttk.Button(result_head, text="复制", style="Link.TButton",
                   command=lambda: self._copy(self.result_text)).grid(row=0, column=1, sticky="e")
        self.result_text.grid(row=1, column=0, sticky="nsew", pady=(gap(2), 0))
        result_scroll = ttk.Scrollbar(result_frame, orient="vertical",
                                      command=self.result_text.yview)
        result_scroll.grid(row=1, column=1, sticky="ns", pady=(gap(2), 0))
        self.result_text.configure(yscrollcommand=result_scroll.set)
        self.result_hint_var = tk.StringVar(
            value="模型上下文里只有很短的摘录；用户要看的完整内容显示在本面板。"
                  "内容过长时面板内滚动查看，并在下方明确标注截断。")
        self.result_hint = ttk.Label(result_frame, textvariable=self.result_hint_var,
                                     style="Path.TLabel", justify="left")
        self.result_hint.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(gap(1), 0))
        self._autowrap(self.result_hint, result_frame)
        self.notebook.add(plan_frame, text=self._plan_tab_title(pending=False))
        self.notebook.add(result_frame, text="读取结果")
        self.notebook.select(plan_frame)   # first screen: the flow/plan guidance is the start

        prompt_bar = ttk.Frame(root, padding=(gap(4), gap(2), gap(4), 0))
        prompt_bar.grid(row=5, column=0, sticky="ew")
        self.prompt_bar = prompt_bar
        prompt_bar.columnconfigure(1, weight=1)
        ttk.Label(prompt_bar, text="任务指令", style="Heading.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, gap(3)))
        self.prompt = ttk.Entry(prompt_bar, font=self._fonts["body"])
        self.prompt.grid(row=0, column=1, sticky="ew")
        self.prompt.bind("<Return>", lambda _event: self.ask())
        self.prompt_hint = ttk.Label(prompt_bar, style="Path.TLabel",
                                     text="Enter 提交 · Ctrl+Enter 任意位置提交 · Esc 回到本输入框 · "
                                          "例：把第一段的 price 改成 amount，先预览，不修改原件")
        self.prompt_hint.grid(row=1, column=1, sticky="w", pady=(gap(1), 0))

        # Action row: primary actions keep the strong styles (accent = start a task, amber =
        # a plan is waiting for you); secondary/destructive ones stay quiet.
        actions = ttk.Frame(root, padding=(gap(4), gap(3), gap(4), gap(4)))
        actions.grid(row=6, column=0, sticky="ew")
        self.actions_frame = actions
        self.ask_button = ttk.Button(actions, text="自然语言查询/规划", command=self.ask,
                                     style="Accent.TButton")
        self.resume_button = ttk.Button(actions, text="继续当前检查点", command=self.resume)
        self.stop_button = ttk.Button(actions, text="停止", command=self.stop_task)
        self.confirm_button = ttk.Button(actions, text="确认并保存输出副本", command=self.execute,
                                         style="Confirm.TButton")
        self.undo_button = ttk.Button(actions, text="撤销最近输出副本", command=self.undo,
                                      style="Danger.TButton")
        self.clear_button = ttk.Button(actions, text="清空显示", command=self.clear_transcript,
                                       style="Link.TButton")
        self.busy_var = tk.StringVar(value="○ 就绪")
        self.busy_label = ttk.Label(actions, textvariable=self.busy_var, style="Muted.TLabel")
        self._action_buttons = (self.select_button, self.probe_button, self.ask_button,
                                self.resume_button, self.confirm_button, self.undo_button,
                                self.clear_button)
        self._layout_actions(narrow=False)
        self.body.bind("<Configure>", self._on_window_resize, add="+")
        root.bind("<Control-Return>", lambda _event: self.ask())
        root.bind("<Escape>", lambda _event: self.focus_prompt())

    def focus_prompt(self):
        """Esc: put the caret back in the instruction box (a safe, non-destructive cancel)."""
        if not self._panels_ready:
            return
        self.prompt.focus_set()

    def _layout_actions(self, narrow: bool):
        """Place the action buttons in one or two rows; both rows stay clickable.

        Narrow windows (down to the 820 px minimum) must never hide a button: instead of
        letting the row clip, the secondary actions move to a second row while the primary
        ones keep the top row.
        """
        self._actions_narrow = narrow
        gaps = self._sp(2)
        if narrow:
            placements = ((self.ask_button, 0, 0, "w"), (self.confirm_button, 0, 1, "w"),
                          (self.resume_button, 0, 2, "w"), (self.busy_label, 0, 3, "e"),
                          (self.undo_button, 1, 0, "w"), (self.clear_button, 1, 1, "w"),
                           (self.stop_button, 1, 2, "w"))
            spacer = 2
        else:
            placements = ((self.ask_button, 0, 0, "w"), (self.resume_button, 0, 1, "w"),
                          (self.confirm_button, 0, 2, "w"), (self.undo_button, 0, 3, "w"),
                          (self.clear_button, 0, 4, "w"), (self.stop_button, 0, 5, "w"),
                           (self.busy_label, 0, 6, "e"))
            spacer = 4
        for widget in (self.ask_button, self.resume_button, self.confirm_button,
                       self.undo_button, self.clear_button, self.stop_button, self.busy_label):
            widget.grid_forget()
        for column in range(7):
            self.actions_frame.columnconfigure(column, weight=0)
        for widget, row, column, sticky in placements:
            widget.grid(row=row, column=column, sticky=sticky,
                        padx=(0 if column == 0 else gaps, 0),
                        pady=(self._sp(1) if row else 0, 0))
        self.actions_frame.columnconfigure(spacer, weight=1)   # pushes the busy label right
        if not self._model_task_active:
            self.stop_button.grid_remove()

    def clear_transcript(self):
        """Clear the visible transcript after an explicit confirmation (audit files stay)."""
        if not self._panels_ready:
            return
        if not messagebox.askyesno("清空显示",
                "只清空窗口里的会话记录显示；磁盘上的审计记录与输出副本不受影响。\n"
                "之后新的事件仍会继续记录。是否清空？", default=messagebox.NO):
            self._log("已取消清空；会话记录保持不变。", "meta")
            return
        self._set_panel(self.log, "")
        self._log("会话记录显示已清空（磁盘审计不受影响）。", "meta")

    # -- small presentation helpers (all safe before the widgets exist) ------
    def show(self, message: str):
        """Append text to the session log; the plain entry point other code calls."""
        self._log(message, "body")

    def _log(self, message: str, tag: str = "body"):
        """Append one block to the transcript with a severity tag, if the widgets exist."""
        if self.log is None:
            return
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n", tag if tag in self._log_tags else "body")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_panel(self, widget, text: str):
        """Replace a read-only pane's content; panels scroll, so nothing is clipped."""
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")
        widget.yview_moveto(0)

    def _set_status(self, text: str, *, level: str = "info"):
        """Single entry point for the status line; ``level`` only chooses the colour."""
        self.status.set(text)
        self._status_level = level
        if self._panels_ready:
            self.status_label.configure(style=f"Status{level.title()}.TLabel")
        self._update_next_step()

    def _layout_meter(self):
        """Progress sentence on top; decision meter and budget meter side by side below.

        One layout for every window size: the two bars are thin, and keeping them together
        means the meters never disagree about which numbers they show.
        """
        gap = self._sp
        for widget in (self.progress_label, self.iteration_label, self.iteration_bar,
                       self.budget_bar):
            widget.grid_forget()
        self.meter_row.columnconfigure(1, weight=0)
        self.meter_row.columnconfigure(2, weight=1)
        self.progress_label.grid(row=0, column=0, columnspan=3, sticky="w")
        self.iteration_label.grid(row=1, column=0, sticky="w", pady=(gap(1), 0))
        self.iteration_bar.grid(row=1, column=1, sticky="ew", padx=(gap(1), gap(2)),
                                pady=(gap(1), 0))
        self.budget_bar.grid(row=1, column=2, sticky="ew", pady=(gap(1), 0))

    def _apply_density(self, *, dense: bool, compact: bool):
        """Adapt how much guidance the window shows to the height it actually has.

        Least-lossy step first: ``dense`` gives the pane footnotes' space to the panes (those
        footnotes repeat what the pane body, the badge and the status card already state).
        ``compact`` additionally withdraws the one-time flow line, the input example and the
        idle checkpoint placeholder, and puts the meter on one line. Every withdrawn line
        returns as soon as the window is tall enough again, and no status, safety or progress
        information is dropped at either step.
        """
        if not self._panels_ready:
            return
        if dense == self._dense and compact == self._compact:
            return
        self._dense, self._compact = dense, compact
        for widget, visible in ((self.plan_hint, not dense), (self.result_hint, not dense),
                                (self.subtitle_label, not compact),
                                (self.prompt_hint, not compact)):
            if visible:
                widget.grid()
            else:
                widget.grid_remove()
        self._layout_meter()
        self._update_checkpoint_visibility()

    def _set_actions_narrow(self, narrow: bool):
        """One source of truth for the action row shape (driven by the window width)."""
        if narrow != self._actions_narrow:
            self._layout_actions(narrow=narrow)

    def _actions_need_two_rows(self) -> bool:
        """True when the action buttons cannot fit one row at the current font/DPI scale.

        Measured from the buttons themselves rather than a fixed pixel threshold, so a
        larger font or a scaled display splits the row before anything gets clipped.
        """
        try:
            width = self.root.winfo_width()
        except tk.TclError:      # pragma: no cover - widgets already gone
            return False
        if width <= 1:
            return False
        buttons = (self.ask_button, self.resume_button, self.confirm_button,
                   self.undo_button, self.clear_button, self.busy_label)
        if self._model_task_active:
            buttons += (self.stop_button,)
        needed = sum(widget.winfo_reqwidth() + self._sp(2) for widget in buttons)
        return needed > width - 2 * self._sp(4)

    def _on_window_resize(self, event):
        """Adapt the chrome to the window size.

        Bound to this window's own body pane (not the toplevel), so the handler disappears
        together with the widgets and a destroyed window can never react to another window's
        resize. Sizes are read from the toplevel because the body is padded inside it.
        """
        if not self._panels_ready or event.widget is not self.body:
            return
        try:
            width, height = self.root.winfo_width(), self.root.winfo_height()
        except tk.TclError:      # pragma: no cover - widgets already gone
            return
        if width <= 1 or height <= 1:      # not realized yet
            return
        self._apply_density(dense=height < _DENSE_HEIGHT, compact=height < _COMPACT_HEIGHT)
        self._set_actions_narrow(self._actions_need_two_rows())

    def _update_checkpoint_visibility(self):
        """In compact mode the "no checkpoint yet" placeholder gives its space back."""
        if not self._panels_ready:
            return
        if self._checkpoint_placeholder and self._compact:
            self.checkpoint_label.grid_remove()
        else:
            self.checkpoint_label.grid()

    def _set_badge(self, level: str, label: str):
        """Status chip: colour, shape *and* words, so the state never rests on colour alone."""
        if not self._panels_ready:
            return
        self.status_badge.set(level, label)

    def _update_next_step(self):
        """Name the next available click, derived from the window's own state."""
        if not self._panels_ready:
            return
        self.next_step_var.set(next_step_hint(
            has_file=self.path is not None, busy=self.busy,
            pending=self.pending_id is not None, verified=self.last_verified_id is not None,
            level=self._status_level))

    def _plan_tab_title(self, *, pending: bool) -> str:
        """Tab label for the plan/output surface; a waiting plan says so in the tab itself."""
        return "计划待确认（未写入）" if pending else "修改计划 / 输出"

    def _select_tab(self, widget):
        """Bring a notebook tab forward (safety-critical events must not hide behind a tab)."""
        if not self._panels_ready:
            return
        try:
            self.notebook.select(widget)
        except tk.TclError:      # pragma: no cover - only if the widget is not a tab
            pass

    def _set_plan_emphasis(self, pending: bool):
        """Layer 2 gets the amber, thicker border exactly while a plan awaits confirmation."""
        if not self._panels_ready:
            return
        self.plan_frame.configure(style="PlanPending.TLabelframe" if pending
                                  else "Plan.TLabelframe")
        self.notebook.tab(self.plan_frame, text=self._plan_tab_title(pending=pending))
        if pending:
            self._select_tab(self.plan_frame)   # an unwritten plan must be seen, not tabbed away

    def _set_progress(self, iterations: Any, tokens: Any, capacity: Any, compressions: Any):
        """Iterations, estimated tokens and compressions, with a budget bar.

        The bar, the sentence and the near-limit reminder are derived from the same values in
        the same call, so the window can never disagree with the transcript about the budget.
        """
        self._progress_state = (iterations, tokens, capacity, compressions)
        if not self._panels_ready:
            return
        limit = getattr(getattr(self, "model_agent", None), "decision_limit", DEFAULT_DECISION_LIMIT)
        self.progress_var.set(format_progress_line(iterations, tokens, capacity, compressions, limit))
        self.iteration_var.set(f"决策 {iterations}" + (f"/{limit}" if limit is not None else ""))
        try:
            if limit is None:
                self.iteration_bar.grid_remove()
            else:
                self.iteration_bar.grid()
                self.iteration_bar.configure(maximum=limit, value=max(0.0, float(iterations)))
        except (TypeError, ValueError):
            self.iteration_bar.configure(value=0.0)
        fraction = budget_fraction(tokens, capacity)
        reminder = budget_advice(fraction)
        self.limit_var.set(reminder)
        if reminder:
            self.limit_label.grid()
        else:
            self.limit_label.grid_remove()
        self.budget_bar.configure(value=0.0 if fraction is None else fraction * 100.0)
        if fraction is not None and fraction >= 0.9:
            self.budget_bar.configure(style="BudgetDanger.Horizontal.TProgressbar")
        elif fraction is not None and fraction >= 0.8:
            self.budget_bar.configure(style="BudgetWarn.Horizontal.TProgressbar")
        else:
            self.budget_bar.configure(style="Budget.Horizontal.TProgressbar")

    def _set_task(self, question: str):
        """Show the instruction this window is currently working on."""
        if not self._panels_ready:
            return
        self.task_var.set("当前任务：" + (question.strip() or "（未提供指令）"))

    def _reset_progress(self):
        """Zero the progress card when a new task starts."""
        self._set_progress(0, "—", "—", 0)
        if self._panels_ready:
            self._set_error(None)

    def _set_checkpoint(self, checkpoint: Any):
        """Resume checkpoint, or an explicit statement that this run keeps none.

        A real checkpoint and the "nothing was persisted" warning are safety information and
        always stay visible; only the idle "no task yet" placeholder may give up its space.
        """
        if not self._panels_ready:
            return
        self._checkpoint_placeholder = False
        if checkpoint:
            self.checkpoint_var.set(f"续跑检查点：{checkpoint}"
                                    "（点击『继续当前检查点』会重新校验文件版本；不会恢复写入批准）")
        else:
            self.checkpoint_var.set("续跑检查点：本次运行未保存检查点"
                                    "（授权目录不可写时不落盘，重启后无法续跑）")
        self._update_checkpoint_visibility()

    def _set_error(self, message: Any):
        """Show the most recent error in the status card; the transcript keeps it in full."""
        if not self._panels_ready:
            return
        if not message:
            self.error_var.set("")
            self.error_label.grid_remove()
            return
        text = str(message)
        if len(text) > _MAX_STATUS_ERROR:
            text = (text[:_MAX_STATUS_ERROR] +
                    f"…（此处只显示前 {_MAX_STATUS_ERROR} 字符，错误全文共 {len(text)} 字符，"
                    "见下会话记录）")
        self.error_var.set("最近错误：" + text)
        self.error_label.grid()

    def _set_notice(self, text: str):
        """Persistent warning banner (e.g. this folder cannot hold session data)."""
        if not self._panels_ready:
            return
        # U+FF01 is present in every CJK font, unlike the ⚠ symbol some hosts render as tofu.
        self.notice_var.set("！" + text)
        self.notice_frame.grid()

    def _clear_notice(self):
        if not self._panels_ready:
            return
        self.notice_var.set("")
        self.notice_frame.grid_remove()

    def _set_busy_ui(self, busy: bool):
        """Disable the action buttons while a worker runs, and show that state.

        This is also the single writer of ``self.busy``: the flag and the indicator can never
        drift apart, and the "next step" line is recomputed from it here.
        """
        self.busy = busy
        if not busy:
            self._model_task_active = False
        if not self._panels_ready:
            return
        state = "disabled" if busy else "normal"
        for button in self._action_buttons:
            button.configure(state=state)
        self._set_actions_narrow(self._actions_need_two_rows())
        if busy and self._model_task_active:
            self.stop_button.grid()
        else:
            self.stop_button.grid_remove()
        self.stop_button.configure(state="normal" if busy and self._model_task_active
                                   and not self._stop_task.is_set() else "disabled")
        self.busy_var.set("● 正在处理模型任务，请等待结束" if busy else "○ 就绪")
        self.busy_label.configure(style="Warn.TLabel" if busy else "Muted.TLabel")
        # An indeterminate bar is the "still working" signal for a long model call.
        if busy:
            self.busy_bar.grid()
            try:
                self.busy_bar.start(60)
            except tk.TclError:      # pragma: no cover - widget already gone
                pass
        else:
            try:
                self.busy_bar.stop()
            except tk.TclError:      # pragma: no cover - widget already gone
                pass
            self.busy_bar.grid_remove()
        self._update_next_step()
        try:
            self.root.configure(cursor="watch" if busy else "")
        except tk.TclError:
            pass

    def _copy(self, widget):
        """Copy a pane to the clipboard so long content stays reusable.

        Reports to the transcript rather than the status card: the status line may be
        carrying a pending-plan or failure notice that must not be overwritten.
        """
        try:
            text = widget.get("1.0", "end-1c")
        except tk.TclError:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._log(f"已复制 {len(text)} 个字符到剪贴板。", "meta")

    def _refresh_file_labels(self):
        """Current file plus the authorized directory (the part that limits the tools)."""
        if not self._panels_ready:
            return
        if self.path is None:
            self.file_var.set("当前文件：尚未选择")
            self.dir_var.set("授权目录：选择文件后，该文件所在目录即授权目录；"
                             "只有这一个文件会被授权读写。")
            return
        root = self.editor.space.root if self.editor is not None else self.path.parent
        self.file_var.set(f"当前文件：{self.path.name}")
        self.dir_var.set(f"授权目录（只授权这一个文件）：{root}")

    def _reset_panels(self):
        """Empty-state text for the plan/result panes after a new file is selected."""
        if not self._panels_ready:
            return
        self.task_var.set("当前任务：（尚未提交指令）")
        self._set_panel(self.plan_text,
                        "尚未生成修改计划。\n\n"
                        "流程：输入任务指令 → 在这里查看修改前后的差异预览 → "
                        "点击『确认并保存输出副本』并在弹窗中确认 → 才会先备份、"
                        "再生成独立输出副本（原文件不动）。")
        self.plan_badge_var.set("○ 尚未生成计划")
        self.plan_badge.configure(style="Badge.TLabel")
        self.plan_hint_var.set("计划在你点击『确认并保存输出副本』并在弹窗中确认之前不会写入任何文件；"
                               "未确认则原文件保持不变。")
        self.confirm_button.configure(style="Confirm.TButton")
        self._set_panel(self.result_text,
                        "尚无读取结果。\n\n"
                        "当模型读取文档时，完整的读取内容会显示在这里："
                        "模型上下文里只有很短的摘录，用户要看的内容只能靠本面板。")
        self.result_badge_var.set("○ 暂无读取结果")
        self.result_badge.configure(style="Badge.TLabel")
        self.result_hint_var.set("长内容可在此面板内滚动查看；会话记录里保留同一结果的有界预览"
                                 "及明确的字符数说明。")
        self._set_progress(0, "—", "—", 0)   # one source for the bar, its style and the sentence
        self.checkpoint_var.set("续跑检查点：尚未运行任务")
        self._checkpoint_placeholder = True
        self._update_checkpoint_visibility()
        self._set_error(None)
        self._clear_notice()
        self._set_plan_emphasis(False)
        self._set_badge("info", "尚未开始")
        self._set_status("请选择一个文件，所在目录即授权目录；默认不覆盖原文件。", level="info")

    def _session_root(self, editor: Editor) -> Path:
        """Validate audit/session storage before any checkpoint path is used."""
        root = editor.space.root.resolve(strict=True)
        audit_raw = editor.audit
        if audit_raw.is_symlink():
            raise ValueError("Audit directory symlink is not permitted")
        audit = audit_raw.resolve(strict=False)
        if not audit.is_relative_to(root) or audit == root:
            raise ValueError("Audit directory is outside the authorized root")
        sessions_raw = audit / "sessions"
        if sessions_raw.is_symlink():
            raise ValueError("Session directory symlink is not permitted")
        sessions = sessions_raw.resolve(strict=False)
        if not sessions.is_relative_to(audit) or sessions == audit:
            raise ValueError("Session directory is outside the audit directory")
        return sessions

    def _prepare_session_base(self, editor: Editor) -> tuple[Path | None, str | None]:
        """Where this task's checkpoints and tool results go, or None to persist nothing.

        The only acceptable location is inside the authorized folder: document-derived
        writes must not escape it. When that folder is not writable for this process,
        the task runs **without persisting anything at all** rather than relocating the
        data somewhere unauthorized. Reads still work; checkpoints and the on-disk
        result files do not, and the returned message says so. Writes are refused
        separately, because a verified output copy cannot be published into a folder
        this process cannot write.
        """
        sessions = self._session_root(editor)          # validation only; creates nothing
        if _writable_directory(sessions):
            return sessions, None
        return None, (
            f"注意：授权目录 {editor.space.root} 对本进程不可写，无法在其中创建会话目录。"
            "本次任务改为不落盘运行：仍可正常读取与分析文档，但不会保存检查点或工具结果，"
            "因此不支持重启续跑；为遵守不越权写入的约定，也不会把文档内容写到其他目录。"
            "如需写入或保留记录，请把文档复制到可写目录（例如项目内的 office-smoke 目录）后重新选择。")

    def _session_metadata_file(self, session_dir: Path) -> Path:
        return session_dir / "session.json"

    def _write_session_metadata(self, session_dir: Path, path: Path, question: str) -> None:
        """Persist only restart-discovery data; approvals and allowlists are never persisted."""
        metadata = {
            "version": 1,
            "selected_path": str(path.resolve(strict=True)),
            "file_sha256": digest(path),
            "session_dir": str(session_dir.resolve(strict=False)),
            "question": question,
        }
        session_dir.mkdir(parents=True, exist_ok=True)
        temporary = self._session_metadata_file(session_dir).with_suffix(".tmp")
        temporary.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, self._session_metadata_file(session_dir))
        self.session_metadata = metadata

    def _discover_session(self, path: Path, editor: Editor) -> tuple[Path, dict] | None:
        """Return a checkpoint only when every persisted identity check matches the selection."""
        selected = path.resolve(strict=True)
        try:
            sessions = self._session_root(editor)
        except (OSError, ValueError):
            return None
        if not sessions.exists() or not sessions.is_dir():
            return None
        candidates = []
        for session_dir in sessions.iterdir():
            if not session_dir.is_dir() or session_dir.is_symlink():
                continue
            metadata_file = self._session_metadata_file(session_dir)
            checkpoint = session_dir / "agent-state.json"
            if (not metadata_file.is_file() or metadata_file.is_symlink() or
                    not checkpoint.is_file() or checkpoint.is_symlink()):
                continue
            try:
                metadata = json.loads(metadata_file.read_text(encoding="utf-8"))
                if metadata.get("version") != 1 or not isinstance(metadata.get("question"), str):
                    continue
                if not metadata["question"].strip() or len(metadata["question"]) > 2000:
                    continue
                stored_path = Path(metadata["selected_path"]).resolve(strict=False)
                stored_session = Path(metadata["session_dir"]).resolve(strict=False)
                if stored_path != selected or stored_session != session_dir.resolve(strict=False):
                    continue
                if not stored_session.is_relative_to(sessions) or stored_session == sessions:
                    continue
                if metadata.get("file_sha256") != digest(selected):
                    continue
                candidates.append((checkpoint.stat().st_mtime_ns, session_dir, metadata))
            except (OSError, KeyError, TypeError, ValueError):
                continue
        if not candidates:
            return None
        _, session_dir, metadata = max(candidates, key=lambda item: item[0])
        return session_dir, metadata

    def _build_model(self):
        """Build the configured provider lazily so restart resume uses the same factory."""
        from main import build_llm
        from dotenv import load_dotenv
        load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=True)
        return build_llm()

    def _confirm_read(self, plan):
        """Ask the UI thread for each read approval; time out closed windows safely."""
        event = threading.Event()
        decision = [False]
        self.events.put(("confirm_read", (plan, event, decision)))
        if not event.wait(120):
            event.set()  # A queued but expired request must never show a stale approval dialog.
            return False
        return decision[0]

    def select(self):
        if self.busy:
            return
        raw = filedialog.askopenfilename(filetypes=[("Office/PDF 文档", "*.doc *.docx *.xls *.xlsx *.pptx *.pdf"), ("全部文件", "*.*")])
        if not raw:
            return
        selected = Path(raw)
        if selected.is_symlink():
            messagebox.showerror("拒绝链接", "请直接选择真实文件，而不是快捷链接或符号链接。")
            return
        path = selected.resolve(strict=True)
        from .com import WindowsComBackend
        self.editor = Editor(path.parent, legacy_backend=WindowsComBackend(workdir=path.parent),
                             require_confirmation=True, render_validation=True)
        self.path = self.editor.space.file(str(path))
        self.model_agent = None
        self.pending_id = None
        self.last_verified_id = None
        self.session_metadata = None
        self.last_compressions = 0
        self._reset_panels()
        self._refresh_file_labels()
        discovered = self._discover_session(self.path, self.editor)
        if discovered:
            session_dir, metadata = discovered
            self.session_metadata = metadata
            self.show("发现与当前精确文件匹配的会话检查点；点击‘继续当前检查点’以重新验证并恢复。")
            self._log(f"可续跑检查点：{self._session_metadata_file(session_dir)}"
                      "（续跑会重新校验文件版本；不会恢复任何写入批准）", "ok")
        # Tell the user up front when this folder cannot hold the session store, so a task
        # never fails later with an unexplained access error.
        try:
            _, warning = self._prepare_session_base(self.editor)
        except Exception as exc:
            self.show(f"授权目录不可用：{exc}")
            self._set_notice(f"授权目录不可用：{exc}")
        else:
            if warning:
                self.show(warning)
                self._set_notice(warning)
        try:
            self.show(json.dumps(inspect_file(self.path), ensure_ascii=False, indent=2))
        except Exception as exc:
            self.show(f"文件初检失败：{type(exc).__name__}: {exc}；拒绝基于初检结果推测文件安全性。")

    def probe(self):
        """Report the Office capability *detection* without pretending a channel is open.

        Only the lightweight detection runs here. ``com.write_capabilities`` performs a real
        open/save probe, which must not happen merely because a window opened — that would be
        an unrequested Office instance. The write-channel verdict is reported when a legacy
        edit is actually requested. Startup is informational: it never borrows the
        failure/refusal wording, and the backend's own reason is passed through unchanged.
        """
        def task():
            try:
                from .com import capabilities
                report = capabilities()
                self.events.put(("text", "Office 检测：" + json.dumps(report, ensure_ascii=False)))
                unavailable = [(name, state) for name, state in report.items()
                               if isinstance(state, dict) and not state.get("available")]
                if unavailable:
                    reasons = "；".join(f"{name}: {state.get('note') or '未说明原因'}"
                                        for name, state in unavailable)
                    notice = capability_detection_notice(reasons)
                    self.events.put(("text", notice))
                    self._publish_capability_notice(notice)
            except Exception as exc:
                notice = capability_detection_notice(str(exc))
                self.events.put(("text", notice))
                self._publish_capability_notice(notice)
        threading.Thread(target=task, daemon=True).start()

    def _publish_capability_notice(self, text: str) -> None:
        """Worker -> UI mailbox for the capability banner.

        The event queue carries the kinds the agent contract owns, so this banner travels as a
        plain attribute the UI thread collects on its next drain tick. Assigning a string is
        atomic, and no Tk call ever happens off the UI thread.
        """
        self._pending_capability_notice = text

    def _apply_pending_capability_notice(self) -> None:
        """UI-thread side of the mailbox: show the warn banner once (never a failure state)."""
        text = getattr(self, "_pending_capability_notice", None)
        if not text:
            return
        self._pending_capability_notice = None
        self._set_notice(text)

    def stop_task(self):
        if not self.busy or not self._model_task_active:
            return
        self._stop_task.set()
        self.stop_button.configure(state="disabled")
        self._set_status("正在停止：当前模型请求返回后，不再执行后续步骤。", level="warn")

    def ask(self):
        if not self.path or not self.editor:
            self._set_status("请先点击‘选择授权文件’并选择要处理的文档。", level="warn")
            return
        if self.busy:
            self._set_status("上一个任务仍在处理中；请等待结束后再提交。", level="warn")
            return
        question = (self._resume_session[1].get("question", "").strip()
                    if self._resume_session else self.prompt.get().strip())
        if not question:
            self._set_status("请输入自然语言任务指令，然后点击‘自然语言查询/规划’。", level="warn")
            self.prompt.focus_set()
            return
        from .clarification import classify_intent, image_request_count
        try:
            image_count = image_request_count(question) if classify_intent(question).family == "image" else 1
        except ValueError as exc:
            self._set_status(str(exc), level="warn")
            return
        path, editor = self.path, self.editor
        resume_session = self._resume_session
        self._resume_session = None
        self.pending_id = None
        self._stop_task.clear()
        self._model_task_active = True
        self.busy = True
        self.last_compressions = 0
        self._set_status("处理中；可随时停止后续步骤。", level="info")
        self._set_badge("info", "进行中")
        self._set_task(question)
        self._reset_progress()
        self._set_busy_ui(True)
        self._log("用户: " + question, "user")

        def task():
            try:
                model = self._build_model()
                approved = (path,)
                # The catalogue must be built before the hook below reads it. It used to be
                # assigned after `accesses` was computed from it, which made `tools` a local
                # that was read before assignment: every task died with UnboundLocalError
                # before the first model call.
                tools = build_tools(approved, ocr_enabled=self.ocr_provider is not None)
                # One file per session: prune to the tier that can act on this file type,
                # then split it so the write tools stay out of context until the model
                # explicitly asks for them via the write gate.
                kind = classify_document(path).kind
                tiered = select_tools(tools, 'presentation_edit' if kind == 'presentation' else kind)
                read_tools, write_tools = split_read_write(tiered)
                accesses = {spec.name: spec.access for spec in tools}
                raw_dispatch = build_dispatcher(editor, path, on_plan=self.events.put,
                                                ocr_provider=self.ocr_provider, image_count=image_count)

                def dispatch(name, args):
                    result = raw_dispatch(name, args)
                    # Show read results in the window: the model itself only receives a
                    # bounded excerpt, so it cannot reproduce document content for the user.
                    if accesses.get(name) == "read":
                        preview = render_tool_result(name, result)
                        if preview:
                            self.events.put(("result", {"tool": name, "text": preview}))
                    return result
                sessions, session_warning = self._prepare_session_base(editor)
                if session_warning:
                    self.events.put(("text", session_warning))
                persist = sessions is not None
                # With persist=False the state directory is never created or written; the
                # authorized path is passed only so the value stays meaningful to readers.
                state_dir = (resume_session[0] if resume_session
                             else (sessions if persist else editor.audit / "sessions") / uuid.uuid4().hex)
                if persist and resume_session is None:
                    self._write_session_metadata(state_dir, path, question)
                agent = OfficeAgent(model, dispatch, read_tools, state_dir, confirm=self._confirm_read,
                    persist=persist, decision_limit=configured_decision_limit(),
                    stop_requested=self._stop_task.is_set,
                    context_tokens=int(os.getenv("OFFICE_CONTEXT_WINDOW", str(DEFAULT_CONTEXT_WINDOW))),
                    reply_tokens=int(os.getenv("OFFICE_REPLY_BUDGET", str(DEFAULT_REPLY_BUDGET))),
                    write_tools=write_tools, write_gate=_WRITE_GATE.name if write_tools else None,
                    capability_tools=tools if write_tools else [t for t in tools if t.access != "write"],
                     on_progress=lambda s: self.events.put(("progress", {
                         "iterations": s.decisions, "compressions": s.compressions, "status": s.status,
                          "decision_limit": agent.decision_limit,
                         "tokens": agent.estimate(s), "capacity": agent.context_tokens})))
                if resume_session is not None:
                    state = agent.load()
                    expected_version = {str(path): digest(path)}
                    if state.required.file_versions != expected_version:
                        raise ValueError("Selected file version differs from checkpoint; refusing resume")
                    # Persisted approvals and write plans are never executable after restart.
                    state.required.confirmations = []
                    state.required.pending = [item for item in state.required.pending
                        if not item.startswith("Preview and confirm edit plan ")]
                    if state.pending_plan is not None and state.pending_plan.access == "write":
                        state.required.errors.append("Discarded persisted write plan; fresh preview and confirmation required")
                        state.pending_plan = None
                else:
                    state = agent.new_state(question, constraints=["不得修改源文件", "任何写入先预览由用户单独确认"], file_versions={str(path): digest(path)})
                self.model_agent = agent
                overhead = agent.fixed_overhead()
                advice = ("；窗口偏小，建议提高 OFFICE_CONTEXT_WINDOW 或减少单次任务的读取次数"
                          if overhead > .6 * agent.context_tokens else "")
                self.events.put(("text", f"预算：固定开销约 {overhead} token（系统提示 + {len(agent.tools)} 个当前工具描述 + "
                    f"回复预留），本次窗口 {agent.context_tokens} token，可用于文件内容约 "
                    f"{max(0, agent.context_tokens - overhead)} token{advice}。"))
                state = agent.run(state)
                self.events.put(("state", {"status": state.status, "iterations": state.decisions,
                    "token_estimate": agent.estimate(state), "context_capacity": agent.context_tokens,
                    "compressions": state.compressions, "answer": state.answer,
                    "decision_limit": agent.decision_limit,
                    "errors": state.required.errors, "evidence": [e.model_dump() for e in state.required.evidence],
                    "resume_checkpoint": str(agent.checkpoint) if agent.checkpoint else None}))
            except (Exception, SystemExit) as exc:
                self.events.put(("error", f"{type(exc).__name__}: {exc}"))
            finally:
                self.events.put(("idle", None))
        threading.Thread(target=task, daemon=True).start()

    def resume(self):
        if self.busy or not self.path or not self.editor:
            return
        if self.model_agent is not None:
            agent = self.model_agent
            session = None
            if not agent.persist:
                self._set_status("该任务以不落盘方式运行，没有检查点可继续；请重新提交请求。", level="warn")
                self._log("上次任务未保存检查点（授权目录不可写），无法继续；请重新提交请求。", "warn")
                return
        else:
            session = self._discover_session(self.path, self.editor)
            if session is None:
                self._log("没有与当前精确文件和版本匹配的检查点；不会恢复其他文件的会话。", "warn")
                return
            session_dir, metadata = session
            if not messagebox.askyesno("确认继续检查点",
                    "将重新建立模型会话并读取该检查点中的文件操作记录。\n"
                    "不会恢复任何写入批准；后续写入仍需单独确认。\n是否继续？"):
                return
            self._resume_session = (session_dir, metadata)
            self.prompt.delete(0, "end")
            self.prompt.insert(0, metadata.get("question", ""))
            self.ask()
            return
        self._stop_task.clear()
        self._model_task_active = True
        self.busy = True
        self._set_status("从当前会话的持久化检查点继续；所有写入仍需单独确认。", level="info")
        self._set_badge("info", "从检查点继续")
        self._set_busy_ui(True)
        def task():
            try:
                current = digest(self.path)
                loaded = agent.load()
                if loaded.required.file_versions != {str(self.path): current}:
                    raise ValueError("Selected file version differs from checkpoint; refusing resume")
                state = agent.resume()
                self.events.put(("state", {"status": state.status, "iterations": state.decisions,
                    "token_estimate": agent.estimate(state), "context_capacity": agent.context_tokens,
                    "compressions": state.compressions, "answer": state.answer,
                    "decision_limit": agent.decision_limit,
                    "errors": state.required.errors, "evidence": [e.model_dump() for e in state.required.evidence],
                    "resume_checkpoint": str(agent.checkpoint) if agent.checkpoint else None}))
            except Exception as exc:
                self.events.put(("error", f"{type(exc).__name__}: {exc}"))
            finally:
                self.events.put(("idle", None))
        threading.Thread(target=task, daemon=True).start()

    def execute(self):
        if not self.editor:
            self._set_status("请先点击‘选择授权文件’并选择要处理的文档。", level="warn")
            return
        if self.busy:
            self._set_status("上一个任务仍在处理中；请等待结束后再确认写入。", level="warn")
            return
        if not self.pending_id:
            self._set_status("当前没有待确认的修改计划；请先提交一次规划并查看差异预览。", level="warn")
            return
        plan = self.editor.plans.get(self.pending_id)
        if not plan:
            self._log("计划已失效；请重新规划。", "warn")
            self._set_status("计划已失效；请重新提交任务生成新计划。", level="warn")
            self._set_plan_emphasis(False)
            self._set_badge("warn", "计划已失效")
            return
        # Publishing a verified output copy needs the authorized folder itself to be
        # writable; say so before asking for confirmation rather than failing afterwards.
        if not _writable_directory(self.editor.audit):
            messagebox.showerror("授权目录不可写",
                f"无法在授权目录创建审计目录：\n{self.editor.audit}\n\n"
                "写入需要先备份、再生成独立输出副本，因此授权目录必须可写。\n"
                "请把文档复制到一个可写目录（例如项目内的 office-smoke 目录）后重新选择。\n\n"
                "本次未做任何写入，原文件保持不变。")
            self.show("已取消：授权目录不可写，未做任何写入，原文件保持不变。")
            self._set_status("已取消：授权目录不可写，未做任何写入，原文件保持不变。", level="danger")
            self._set_badge("danger", "授权目录不可写")
            return
        if not messagebox.askyesno("确认真实写入", "将先备份并创建输出副本（原文件不动）。\n" +
                json.dumps(plan.public(), ensure_ascii=False, indent=2) + "\n是否确认执行？",
                default=messagebox.NO):      # the default button is Cancel: no accidental writes
            self.show("已取消：未做任何写入，原文件保持不变。")
            self._set_status("已取消：未做任何写入，原文件保持不变。", level="info")
            return
        self.busy = True
        self._set_status("正在备份、写入副本并重开验证...", level="info")
        self._set_badge("info", "正在写入输出副本")
        self._set_busy_ui(True)
        plan_id, editor = plan.id, self.editor
        def task():
            try:
                self.events.put(("verified", editor.execute(plan_id, confirmed=True)))
            except Exception as exc:
                self.events.put(("error", str(exc)))
            finally:
                self.events.put(("idle", None))
        self.pending_id = None
        threading.Thread(target=task, daemon=True).start()

    def undo(self):
        if not self.editor or not self.last_verified_id:
            self._set_status("没有可撤销的已验证输出副本。", level="warn")
            return
        if self.busy:
            self._set_status("上一个任务仍在处理中；请等待结束后再撤销。", level="warn")
            return
        job_id, editor = self.last_verified_id, self.editor
        if not messagebox.askyesno("确认撤销", "只删除上次生成且审计哈希未变的输出副本；源文件与备份保留。\n"
                f"任务 ID: {job_id}\n确认撤销？",
                default=messagebox.NO):      # Cancel is the default: nothing is deleted by reflex
            self.show("已取消撤销；输出副本保留。")
            self._set_status("已取消撤销；输出副本保留。", level="info")
            return
        self.busy = True
        self._set_status("正在删除审计哈希未变的输出副本（源文件与备份保留）...", level="info")
        self._set_badge("info", "正在撤销输出副本")
        self._set_busy_ui(True)
        def task():
            try:
                self.events.put(("undone", editor.rollback(job_id, confirmed=True)))
            except Exception as exc:
                self.events.put(("error", str(exc)))
            finally:
                self.events.put(("idle", None))
        threading.Thread(target=task, daemon=True).start()

    # -- event pump ---------------------------------------------------------
    def _answer_read_confirmation(self, payload):
        """Ask the user for one read approval; the worker's event is always set afterwards.

        A closed dialog (exception) or an expired request means *refused*: the original
        fail-closed behaviour is kept verbatim.
        """
        plan, event, decision = payload
        if event.is_set():
            return  # The worker timed out; do not resurrect a stale approval.
        try:
            decision[0] = messagebox.askyesno("确认读取所选文件",
                "仅允许本次工具读取已选择的文件。\n" +
                json.dumps({"tool": plan.tool, "arguments": plan.arguments},
                           ensure_ascii=False, indent=2) + "\n是否允许？")
        except Exception:
            decision[0] = False
        finally:
            event.set()

    def _progress_status(self, status: Any, iterations: Any, tokens: Any,
                         capacity: Any, compressions: Any) -> str:
        """The status line, stated once for both the progress and the state event."""
        limit = getattr(getattr(self, "model_agent", None), "decision_limit", DEFAULT_DECISION_LIMIT)
        steps = str(iterations) + (f"/{limit}" if limit is not None else "")
        return (f"状态 {status} | 迭代 {steps} | 估算 token {tokens}/{capacity} "
                f"({_budget_percent(tokens, capacity)}) | 压缩 {compressions} 次")

    def _render_plan_card(self, plan: dict, *, plan_id: str = ""):
        """PlanCard: the *unwritten* difference preview plus its badge, hint and confirm cue.

        The body text still comes from ``format_plan_preview``; only the widget updates live
        here, so the card can be driven straight from a payload in tests.
        """
        if not self._panels_ready:
            return
        self._set_panel(self.plan_text, align_table_text(format_plan_preview(plan)))
        self.plan_badge_var.set("● 未写入 · 等待你确认")
        self.plan_badge.configure(style="BadgeWarn.TLabel")
        self.plan_hint_var.set(
            f"计划 {plan_id} 尚未写入。点击下方『确认并保存输出副本』并在弹窗中确认后，"
            "才会先备份、再生成独立输出副本（原文件不动）；不确认则原文件保持不变。")
        self.confirm_button.configure(style="Pending.TButton")  # highlight the write gate

    def _render_output_card(self, payload: dict, *, undone: bool):
        """OutputCard: the verified output / undo summary, with source-and-backup wording."""
        if not self._panels_ready:
            return
        self._set_panel(self.plan_text, align_table_text(
            format_output_summary(payload, undone=undone)))
        if undone:
            self.plan_badge_var.set("● 已撤销输出副本 · 源文件与备份保留")
            self.plan_badge.configure(style="Badge.TLabel")
            self.plan_hint_var.set("只删除了审计哈希未变的输出副本；源文件与备份始终保留。"
                                   "需要再次写入请重新规划并单独确认。")
        else:
            self.plan_badge_var.set("● 已生成独立输出副本 · 已验证")
            self.plan_badge.configure(style="BadgeOk.TLabel")
            self.plan_hint_var.set("源文件字节未改动；备份保留用于人工恢复。"
                                   "撤销只删除哈希未变的输出副本（需再次确认）。")
        self.confirm_button.configure(style="Confirm.TButton")

    def _show_plan(self, plan: dict):
        """Make the *unwritten* difference preview and the confirmation entry unmissable."""
        self._log("操作计划 / 差异预览（未写入）:\n" + json.dumps(plan, ensure_ascii=False, indent=2),
                  "plan")
        plan_id = plan.get("id", "")
        self._set_status(f"已生成修改计划（未写入）：{plan_id}；请查看差异预览并点击"
                         "『确认并保存输出副本』单独确认。", level="warn")
        self._set_plan_emphasis(True)          # amber, thicker border while it waits
        self._set_badge("warn", "等待你确认计划")
        self._render_plan_card(plan, plan_id=plan_id)

    def _show_output(self, payload: dict, *, undone: bool):
        """Report a verified output copy, or a completed undo, without hiding any field."""
        title = "输出副本已撤销；源文件和备份未修改" if undone else "已验证输出与审计"
        self._log(title + ":\n" + json.dumps(payload, ensure_ascii=False, indent=2),
                  "warn" if undone else "ok")
        self._set_status("已撤销输出副本；源文件和备份未修改。" if undone else
                         "已生成独立输出副本并通过重开验证；源文件未改动，备份保留。",
                         level="warn" if undone else "ok")
        self._set_plan_emphasis(False)         # the plan is consumed either way
        self._set_badge("warn" if undone else "ok",
                        "已撤销输出副本" if undone else "已生成输出副本")
        self._render_output_card(payload, undone=undone)

    def _show_result(self, tool: str, text: Any):
        """Show a read result: the pane gets the full text, the transcript a bounded preview."""
        text = "" if text is None else str(text)
        clipped = len(text) > _MAX_RESULT_PREVIEW
        self._log(f"读取结果（{tool}）:\n" + text[:_MAX_RESULT_PREVIEW] +
                  (f"\n…（会话记录中的预览在此截断：完整内容共 {len(text)} 字符，"
                   "已完整显示在『读取结果』面板中）" if clipped else ""), "result")
        if not self._panels_ready:
            return
        shown, note = _cap_panel_text(text)
        # Display-only column alignment for the renderer's table lines; the transcript keeps
        # the exact text the renderer produced.
        self._set_panel(self.result_text, align_table_text(shown))
        if self.pending_id is None:      # never tab away from a plan that is still unwritten
            self._select_tab(self.result_frame)
        self.result_badge_var.set(f"● {tool} · {len(text)} 字符")
        self.result_badge.configure(style="BadgeOk.TLabel")
        self.result_hint_var.set(f"工具 {tool}：完整读取内容（{len(text)} 字符）显示在本面板，"
                                 "可滚动查看；会话记录里保留同一条结果的有界预览。" + note)

    def _handle_event(self, kind: str, payload):
        """Route one drained event; the kinds and their meanings are unchanged."""
        if kind == "plan":
            self.pending_id = payload["id"]
            self._show_plan(payload)
        elif kind == "verified":
            self.last_verified_id = payload["id"]
            # The plan has been consumed by Editor.execute (it pops it), so the window must
            # stop advertising a plan that is still awaiting confirmation.
            self.pending_id = None
            self._show_output(payload, undone=False)
        elif kind == "undone":
            self.last_verified_id = None
            self._show_output(payload, undone=True)
        elif kind == "progress":
            self._set_status(self._progress_status(payload["status"], payload["iterations"],
                                                   payload["tokens"], payload["capacity"],
                                                   payload["compressions"]),
                             level=status_level(payload["status"]))
            self._set_badge(*status_badge(payload["status"]))
            self._set_progress(payload["iterations"], payload["tokens"], payload["capacity"],
                               payload["compressions"])
            if payload["compressions"] > self.last_compressions:
                self._log(f"上下文压缩事件：第 {payload['compressions']} 次；"
                          "保留目标、约束、文件版本和证据定位。", "budget")
                self.last_compressions = payload["compressions"]
        elif kind == "state":
            status = payload["status"]
            self._set_status(self._progress_status(status, payload["iterations"],
                                                   payload["token_estimate"],
                                                   payload["context_capacity"],
                                                   payload["compressions"]),
                             level=status_level(status))
            self._set_badge(*status_badge(status))
            self._set_progress(payload["iterations"], payload["token_estimate"],
                               payload["context_capacity"], payload["compressions"])
            self._set_checkpoint(payload.get("resume_checkpoint"))
            errors = payload.get("errors") or []
            presented = legacy_refusal_presentation(errors[-1]) if errors else None
            self._set_error(presented)
            if presented:
                # A refusal reported through the state payload gets the same safety marker as
                # an error event, with the backend's sentence kept verbatim.
                self._log("错误: " + presented, "safety")
            self._log(describe_state(payload) + "\n" + json.dumps(payload, ensure_ascii=False, indent=2),
                      "ok" if status == "done" else "warn" if status in {"limit", "budget"} else "body")
        elif kind == "result":
            self._show_result(payload["tool"], payload["text"])
        elif kind == "idle":
            self.busy = False
            self._set_busy_ui(False)
        elif kind == "error":
            presented = legacy_refusal_presentation(payload)
            self._set_status("失败，原文件保留", level="danger")
            self._set_badge("danger", "失败，原文件保留")
            self._set_error(presented)
            # "safety" tag: a refusal/guard failure gets a marked block (danger colours), while
            # the wording stays the established "错误: …" that the transcript tests lock in.
            self._log("错误: " + presented, "safety")
        else:
            # "text": Office detection, budget line, session warnings — the raw line is kept.
            self._log(str(payload), "meta")

    def drain(self):
        """Pump worker events into the window; every kind keeps its original meaning."""
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "confirm_read":
                    self._answer_read_confirmation(payload)
                    continue
                try:
                    self._handle_event(kind, payload)
                except Exception as exc:
                    # A malformed payload must not kill the pump and silently freeze the UI.
                    self._log(f"界面处理事件 {kind} 时出错：{type(exc).__name__}: {exc}；"
                              f"原始事件：{_clip_repr(payload)}", "error")
                    self._set_error(f"界面处理事件 {kind} 失败：{type(exc).__name__}: {exc}")
        except queue.Empty:
            pass
        self._apply_pending_capability_notice()
        self.root.after(150, self.drain)


def _webview_available() -> bool:
    try:
        from importlib.util import find_spec
        return find_spec("webview") is not None
    except (ImportError, ValueError):
        return False


def _launch_web(ocr_provider: OCRProvider | None = None) -> None:
    """pywebview 前端；后端与 Tk 版完全同源（同一套 agent/writes/documents）。

    Windows 默认 Edge/netfx；其他平台仍默认 Qt。显式环境配置优先于 .env。
    初始化与渲染失败均报告原始错误，不自动安装依赖、重试或改后端。
    """
    from .webui.runtime import configure_runtime, startup_help
    gui = "未配置"
    try:
        gui = configure_runtime()
        import webview
        # CDP gives local processes complete control of this WebView; opt in only for acceptance.
        debug_port = os.getenv("OFFICE_WEBVIEW_DEBUG_PORT", "")
        port = int(debug_port) if (debug_port.isascii() and debug_port.isdecimal()
                                   and len(debug_port) <= 5) else 0
        # Also clear a prior opt-in if the same Python process starts another window.
        webview.settings["REMOTE_DEBUGGING_PORT"] = port if 1 <= port <= 65535 else None
        if os.name == 'nt' and gui == 'edgechromium' and os.getenv('OFFICE_FLOATING_ORB', '1') == '1':
            from .webui.floating_shell import launch_floating
            launch_floating(webview, gui, ocr_provider)
            return
        from .webui.bridge import Api, WebviewApp
        app = WebviewApp(ocr_provider=ocr_provider)
        static = Path(__file__).resolve().parent / "webui" / "static" / "index.html"
        window = webview.create_window(
            "LiX — 审核后写入",
            str(static),
            js_api=Api(app),
            width=1160, height=760, min_size=(820, 620),
            background_color="#0b0d10",
        )
        app.attach(window)
        webview.start(gui=gui, icon=str(APP_ICON))
    except Exception as exc:
        raise SystemExit(
            f"Web UI 启动失败（后端 {gui}）：{type(exc).__name__}: {exc}\n"
            + startup_help(gui)) from exc


def main():
    import argparse
    parser = argparse.ArgumentParser(description="LiX — 审核后写入")
    parser.add_argument("--ui", choices=["web", "tk", "auto"], default="auto",
                        help="界面渲染层：web=pywebview，tk=Tkinter，auto=有 pywebview 用 web 否则 tk")
    args = parser.parse_args()
    if args.ui in {"web", "auto"} and _webview_available():
        _launch_web()
        return
    if args.ui == "web":
        raise SystemExit("pywebview 未安装：运行 uv sync 安装依赖后再试。")
    if tk is None:
        raise SystemExit("Desktop UI unavailable: install/use a Windows Python with tkinter/Tcl-Tk, then rerun uv sync with that interpreter")
    root = tk.Tk()
    Desktop(root)
    root.mainloop()


if __name__ == "__main__":
    main()
