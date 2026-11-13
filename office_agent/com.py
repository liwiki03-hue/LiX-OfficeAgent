"""Opt-in read-only extraction plus a fail-closed write channel for Microsoft Office COM.

Nothing here opens a document of unknown origin: reading is limited to the one file the
caller selected, and the write probe uses only a blank document this module creates itself.
The write channel stays closed unless that probe succeeds end to end.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import threading
import uuid
from pathlib import Path

from openpyxl.utils.cell import range_boundaries
from contextlib import contextmanager
from typing import Any, Callable, Protocol


_MAX_PARAGRAPHS = 20
_MAX_CELLS = 20 * 20
_MAX_TEXT = 4096
_MAX_TOTAL_CELLS = _MAX_CELLS
_SUPPORTED = {".doc", ".docx", ".xls", ".xlsx"}
_APPLICATIONS = {"word": "Word.Application", "excel": "Excel.Application"}
# The write channel verifies whole bounded windows, not just the edited cell/paragraph.
_LEGACY_PARAGRAPH_WINDOW = 100
_LEGACY_REGION_ROWS = 20
_LEGACY_REGION_COLUMNS = 20


class OfficeUnavailableError(RuntimeError):
    """The Office write channel is closed; the message states why, without guessing."""


@contextmanager
def isolated_office(prog_id: str):
    """Yield a hidden, separate Office process and always tear it down."""
    if prog_id not in {"Word.Application", "Excel.Application"}:
        raise ValueError("unsupported Office application")
    import pythoncom
    from win32com.client import DispatchEx

    pythoncom.CoInitialize()
    instance = None
    try:
        # DispatchEx creates a separate process; Dispatch may attach to a user's Office.
        instance = DispatchEx(prog_id)
        instance.Visible = False
        instance.DisplayAlerts = 0
        # 3 is msoAutomationSecurityForceDisable. Set before opening any file.
        instance.AutomationSecurity = 3
        yield instance
    finally:
        if instance is not None:
            try:
                instance.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def extract_office(path: str, *, include_tables: bool = False,
                   region: str | None = None) -> dict[str, Any]:
    """Extract a small, read-only snapshot from one caller-selected Office file.

    This function deliberately accepts one path only. It does not search directories,
    follow links, or authorize a path on the caller's behalf.
    """
    if not isinstance(path, str) or not os.path.isabs(path):
        raise ValueError("path must be absolute")
    suffix = os.path.splitext(path)[1].lower()
    if suffix not in _SUPPORTED:
        raise ValueError("unsupported Office file type")
    if Path(path).is_symlink():
        raise ValueError("symbolic link refused")
    if not Path(path).is_file():
        raise ValueError("Office input must be an existing regular file")
    if region is not None:
        if not re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]{0,6}(?::[A-Za-z]{1,3}[1-9][0-9]{0,6})?", region):
            raise ValueError("region must be a bounded A1 rectangle")
        min_col, min_row, max_col, max_row = range_boundaries(region)
        if min_col > max_col or min_row > max_row or max_col > 16384 or max_row > 1048576 or max_col - min_col >= 20 or max_row - min_row >= 20:
            raise ValueError("COM region must be at most 20x20")
    if suffix in {".doc", ".docx"}:
        if region is not None:
            raise ValueError("cell region only applies to Excel")
        return _extract_word(path, include_tables)
    return _extract_excel(path, region)


def _clip(value: Any) -> str:
    return str(value if value is not None else "")[:_MAX_TEXT]


def _extract_word(path: str, include_tables: bool) -> dict[str, Any]:
    with isolated_office("Word.Application") as app:
        document = None
        try:
            document = app.Documents.Open(
                FileName=path,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                UpdateLinks=0,
                NoEncodingDialog=True,
            )
            paragraphs = []
            for index, paragraph in enumerate(document.Paragraphs, 1):
                if index > _MAX_PARAGRAPHS:
                    break
                paragraphs.append({"index": index, "text": _clip(
                    paragraph.Range.Text.rstrip("\r\a"))})
            result: dict[str, Any] = {
                "type": "word",
                "path": path,
                "paragraphs": paragraphs,
            }
            if include_tables:
                tables = []
                total_cells = 0
                for table_index, table in enumerate(document.Tables, 1):
                    if table_index > _MAX_PARAGRAPHS or total_cells >= _MAX_TOTAL_CELLS:
                        break
                    rows = []
                    for row_index, row in enumerate(table.Rows, 1):
                        if row_index > _MAX_PARAGRAPHS or total_cells >= _MAX_TOTAL_CELLS:
                            break
                        cells = []
                        for cell_index, cell in enumerate(row.Cells, 1):
                            if cell_index > _MAX_PARAGRAPHS or total_cells >= _MAX_TOTAL_CELLS:
                                break
                            total_cells += 1
                            cells.append({
                                "index": cell_index,
                                "text": _clip(cell.Range.Text.rstrip("\r\a")),
                            })
                        rows.append({"index": row_index, "cells": cells})
                    tables.append({"index": table_index, "rows": rows})
                result["tables"] = tables
            return result
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass


def _extract_excel(path: str, region: str | None) -> dict[str, Any]:
    with isolated_office("Excel.Application") as app:
        workbook = None
        try:
            workbook = app.Workbooks.Open(
                Filename=path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
            )
            sheets = []
            total_cells = 0
            for sheet_index, sheet in enumerate(workbook.Worksheets, 1):
                if total_cells >= _MAX_TOTAL_CELLS:
                    break
                if region is None:
                    source = sheet.UsedRange
                else:
                    source = sheet.Range(region)
                rows = []
                row_count = min(int(source.Rows.Count), 20)
                col_count = min(int(source.Columns.Count), 20)
                for row_index in range(1, row_count + 1):
                    if total_cells >= _MAX_TOTAL_CELLS:
                        break
                    cells = []
                    for col_index in range(1, col_count + 1):
                        if total_cells >= _MAX_TOTAL_CELLS:
                            break
                        cell = source.Cells(row_index, col_index)
                        cells.append({
                            "index": col_index,
                            # Formula and displayed text are distinct. Office may
                            # recalculate at open; do not claim that Text is a
                            # previously cached or freshly calculated result.
                            "formula": _clip(cell.Formula),
                            "display_text_calculation_status_unknown": _clip(cell.Text),
                        })
                        total_cells += 1
                    rows.append({"index": row_index, "cells": cells})
                sheets.append({"index": sheet_index, "name": _clip(sheet.Name),
                               "region": region or "UsedRange", "rows": rows})
            return {"type": "excel", "path": path, "sheets": sheets}
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass


def capabilities() -> dict:
    result = {}
    for name, prog_id in (("word", "Word.Application"), ("excel", "Excel.Application")):
        try:
            with isolated_office(prog_id):
                pass
            result[name] = {"available": True, "note": "独立实例启动检测成功；高级读写仍需专门验证"}
        except PermissionError:
            result[name] = {"available": False, "note": "独立 Office COM 实例启动遭权限拒绝；不能据此判断 Office 是否已安装。普通 docx/xlsx 采用 Python 降级；遗留/高级格式拒绝写入 (PermissionError)"}
        except Exception as exc:
            result[name] = {"available": False, "note": f"独立 Office COM 实例未能启动；不能据此判断 Office 是否已安装。普通 docx/xlsx 采用 Python 降级；遗留/高级格式拒绝写入 ({type(exc).__name__})"}
    return result


def run_in_office_thread(work: Callable[[], Any]) -> Any:
    """Run ``work`` on a dedicated thread that initialises and always releases COM there.

    COM is apartment-bound. Initialising it on whichever thread happens to call first would
    leak that apartment for the life of the process, and an exception would skip the
    matching ``CoUninitialize``. The work therefore runs on its own short-lived thread,
    where the pair is completed on every path, exceptions included, and the exception is
    re-raised on the caller's thread.
    """
    outcome: dict[str, Any] = {}

    def runner() -> None:
        initialized = False
        try:
            import pythoncom

            pythoncom.CoInitialize()
            initialized = True
            outcome["value"] = work()
        except BaseException as exc:  # includes import/startup failures on this thread
            outcome["error"] = exc
        finally:
            if initialized:
                try:
                    pythoncom.CoUninitialize()
                except BaseException as exc:
                    # Preserve the original work failure if cleanup also fails.
                    outcome.setdefault("error", exc)

    thread = threading.Thread(target=runner, name="office-com", daemon=True)
    thread.start()
    thread.join()
    if "error" in outcome:
        raise outcome["error"]
    return outcome.get("value")


def _describe_com_failure(exc: Exception) -> str:
    detail = str(exc)[:200]
    if isinstance(exc, PermissionError):
        return (f"COM 自动化被权限或策略拒绝（{detail}）；这不等于没有安装 Office，"
                f"普通 docx/xlsx 请使用 Python 通道")
    if isinstance(exc, ImportError):
        return f"缺少 pywin32/COM 支持（{detail}）"
    return (f"独立 Office 实例未能完成“新建空白文档 → 另存副本 → 重新打开”探针"
            f"（{type(exc).__name__}: {detail}）")


def _probe_in_instance(application: str, target: Path) -> dict[str, Any]:
    """Create, save and reopen a blank document inside one isolated instance."""
    with isolated_office(_APPLICATIONS[application]) as app:
        if application == "word":
            document = app.Documents.Add()
            try:
                document.Range().Text = "office write probe"
                document.SaveAs2(str(target), FileFormat=0)  # wdFormatDocument (.doc), not default OOXML
            finally:
                document.Close(SaveChanges=False)
            reopened = app.Documents.Open(FileName=str(target), ConfirmConversions=False,
                                          ReadOnly=True, AddToRecentFiles=False, UpdateLinks=0,
                                          NoEncodingDialog=True)
            try:
                text = str(reopened.Range().Text)
            finally:
                reopened.Close(SaveChanges=False)
        else:
            workbook = app.Workbooks.Add()
            try:
                workbook.Worksheets(1).Cells(1, 1).Value = "office write probe"
                workbook.SaveAs(str(target), FileFormat=56)  # xlExcel8 (.xls), not default OOXML
            finally:
                workbook.Close(SaveChanges=False)
            reopened = app.Workbooks.Open(Filename=str(target), UpdateLinks=0, ReadOnly=True,
                                          AddToMru=False, Notify=False)
            try:
                text = str(reopened.Worksheets(1).Cells(1, 1).Text)
            finally:
                reopened.Close(SaveChanges=False)
    if not target.is_file() or target.stat().st_size == 0:
        return {"application": application, "available": False, "stage": "verify",
                "probe": "blank_document_save_reopen",
                "reason": "探针新建的空白文档另存后的副本不存在或为空，写通道不可用"}
    return {"application": application, "available": True, "stage": "ok", "reason": "",
            "probe": "blank_document_save_reopen", "bytes": target.stat().st_size,
            "roundtrip_text": text[:_MAX_TEXT]}


def probe_write_channel(application: str, *, workdir: str | os.PathLike[str] | None = None) -> dict[str, Any]:
    """Can this machine really open **and save** a document through COM?

    Only a blank document created by this module is used: it is created inside the isolated
    instance, saved to a throwaway copy under ``workdir`` (or a fresh temporary directory),
    reopened, then deleted. No user file and no document of unknown origin is ever opened —
    that is a hard safety requirement, not a convenience. A failure keeps the write channel
    closed and is reported verbatim.
    """
    if application not in _APPLICATIONS:
        raise ValueError("unsupported Office application")
    try:
        if workdir is None:
            directory = Path(tempfile.mkdtemp(prefix="office-probe-"))
            owned_directory = True
        else:
            directory = Path(workdir)
            directory.mkdir(parents=True, exist_ok=True)
            owned_directory = False
    except OSError as exc:
        return {"application": application, "available": False, "stage": "workdir",
                "probe": "blank_document_save_reopen",
                "reason": f"探针无法创建自己的临时目录（{type(exc).__name__}: {str(exc)[:160]}）"}
    suffix = ".doc" if application == "word" else ".xls"
    target = directory / f"office-probe-{uuid.uuid4().hex[:12]}{suffix}"
    try:
        return run_in_office_thread(lambda: _probe_in_instance(application, target))
    except Exception as exc:
        return {"application": application, "available": False, "stage": "startup",
                "probe": "blank_document_save_reopen", "reason": _describe_com_failure(exc)}
    finally:
        try:
            target.unlink(missing_ok=True)
        finally:
            if owned_directory:
                shutil.rmtree(directory, ignore_errors=True)


def write_capabilities(*, workdir: str | os.PathLike[str] | None = None) -> dict[str, Any]:
    """Startup detection *and* the real open/save probe, per application.

    The two are different facts and must not be merged: "Office is not installed" and
    "Office is installed but automation is blocked" need different advice, and only a
    successful probe — never mere installation — may open the write channel.
    """
    startup = capabilities()
    result: dict[str, Any] = {}
    for application in ("word", "excel"):
        entry = dict(startup.get(application, {}))
        entry["startup_ok"] = bool(entry.get("available"))
        if not entry["startup_ok"]:
            entry["write_channel"] = False
            entry["reason"] = entry.get("note", "独立实例启动失败")
            entry["probe"] = None
            result[application] = entry
            continue
        probe = probe_write_channel(application, workdir=workdir)
        entry["probe"] = probe
        entry["write_channel"] = bool(probe.get("available"))
        entry["reason"] = probe.get("reason", "")
        result[application] = entry
    return result


def describe_write_refusal(application: str, probe: dict[str, Any] | None) -> str:
    """The user-facing sentence for a closed write channel; never guesses why."""
    name = "Word" if application == "word" else "Excel"
    reason = (probe or {}).get("reason") or "探针未通过"
    return (f"本机未检测到可用的 Microsoft {name} COM 写通道，.doc/.xls 编辑不可用：{reason}。"
            f"未做任何改动；普通 docx/xlsx 请使用 Python 通道")


class OfficeWriteBackend(Protocol):
    """The narrow Office surface the controlled legacy editors use; injectable for tests."""

    def probe(self, application: str) -> dict[str, Any]: ...
    def read_word_paragraphs(self, path: str, *, limit: int) -> list[str]: ...
    def replace_word_paragraph(self, path: str, *, index: int, old: str, new: str,
                               expected_count: int) -> str: ...
    def read_excel_cell(self, path: str, *, sheet: str, cell: str) -> Any: ...
    def read_excel_region(self, path: str, *, sheet: str, first_cell: str,
                          last_cell: str) -> dict[str, Any]: ...
    def set_excel_cell(self, path: str, *, sheet: str, cell: str, value: Any) -> Any: ...


def _paragraph_body(text: Any) -> str:
    """One Word paragraph's text without its trailing paragraph mark."""
    return str(text if text is not None else "").rstrip("\r\a")


def _word_paragraphs(path: str, limit: int) -> list[str]:
    with isolated_office("Word.Application") as app:
        document = None
        try:
            document = app.Documents.Open(FileName=path, ConfirmConversions=False, ReadOnly=True,
                                          AddToRecentFiles=False, UpdateLinks=0, NoEncodingDialog=True)
            paragraphs = []
            for index, paragraph in enumerate(document.Paragraphs, 1):
                if index > limit:
                    break
                paragraphs.append(_paragraph_body(paragraph.Range.Text))
            return paragraphs
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass


def _replace_word_paragraph(path: str, index: int, old: str, new: str, expected_count: int) -> str:
    with isolated_office("Word.Application") as app:
        document = None
        try:
            document = app.Documents.Open(FileName=path, ConfirmConversions=False, ReadOnly=False,
                                          AddToRecentFiles=False, UpdateLinks=0, NoEncodingDialog=True)
            paragraph = document.Paragraphs(index + 1)      # COM paragraph indexes are 1-based
            body = _paragraph_body(paragraph.Range.Text)
            found = body.count(old)
            if found != expected_count:
                raise ValueError(f"expected {expected_count} occurrence(s) of the literal, found {found}")
            # The paragraph mark is re-appended explicitly: assigning text without it would
            # merge this paragraph into the next one.
            paragraph.Range.Text = body.replace(old, new) + "\r"
            document.Save()
            return _paragraph_body(document.Paragraphs(index + 1).Range.Text)
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass


def _read_excel_cell(path: str, sheet: str, cell: str) -> Any:
    with isolated_office("Excel.Application") as app:
        workbook = None
        try:
            workbook = app.Workbooks.Open(Filename=path, UpdateLinks=0, ReadOnly=True,
                                          AddToMru=False, Notify=False)
            return workbook.Worksheets(sheet).Range(cell).Value
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass


def _read_excel_region(path: str, sheet: str, first_cell: str, last_cell: str) -> dict[str, Any]:
    min_col, min_row, max_col, max_row = range_boundaries(f"{first_cell}:{last_cell}")
    if min_col > max_col or min_row > max_row:
        raise ValueError("region must run from its top-left to its bottom-right cell")
    if (max_col - min_col + 1) > _LEGACY_REGION_COLUMNS or (max_row - min_row + 1) > _LEGACY_REGION_ROWS:
        raise ValueError(f"legacy region verification is bounded to "
                         f"{_LEGACY_REGION_ROWS}x{_LEGACY_REGION_COLUMNS} cells")
    with isolated_office("Excel.Application") as app:
        workbook = None
        try:
            workbook = app.Workbooks.Open(Filename=path, UpdateLinks=0, ReadOnly=True,
                                          AddToMru=False, Notify=False)
            worksheet = workbook.Worksheets(sheet)
            values: dict[str, Any] = {}
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    address = f"{worksheet.Cells(row, column).Address(False, False)}"
                    values[str(address).replace("$", "")] = worksheet.Cells(row, column).Value
            return values
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass


def _set_excel_cell(path: str, sheet: str, cell: str, value: Any) -> Any:
    with isolated_office("Excel.Application") as app:
        workbook = None
        try:
            workbook = app.Workbooks.Open(Filename=path, UpdateLinks=0, ReadOnly=False,
                                          AddToMru=False, Notify=False)
            worksheet = workbook.Worksheets(sheet)
            worksheet.Range(cell).Value = value
            workbook.Save()
            return worksheet.Range(cell).Value
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass


class WindowsComBackend:
    """The real COM write backend; every call runs in its own Office thread.

    Instances are created with ``DispatchEx`` inside :func:`isolated_office`, so ``Quit()``
    can only ever reach a process this application started — a user's open Office window is
    never attached to and never closed.
    """

    def __init__(self, *, workdir: str | os.PathLike[str] | None = None):
        self.workdir = str(workdir) if workdir is not None else None

    def probe(self, application: str) -> dict[str, Any]:
        return probe_write_channel(application, workdir=self.workdir)

    def read_word_paragraphs(self, path: str, *, limit: int = _LEGACY_PARAGRAPH_WINDOW) -> list[str]:
        _require_absolute_file(path)
        return run_in_office_thread(lambda: _word_paragraphs(path, limit))

    def replace_word_paragraph(self, path: str, *, index: int, old: str, new: str,
                               expected_count: int) -> str:
        _require_absolute_file(path)
        return run_in_office_thread(
            lambda: _replace_word_paragraph(path, index, old, new, expected_count))

    def read_excel_cell(self, path: str, *, sheet: str, cell: str) -> Any:
        _require_absolute_file(path)
        return run_in_office_thread(lambda: _read_excel_cell(path, sheet, cell))

    def read_excel_region(self, path: str, *, sheet: str, first_cell: str,
                          last_cell: str) -> dict[str, Any]:
        _require_absolute_file(path)
        return run_in_office_thread(lambda: _read_excel_region(path, sheet, first_cell, last_cell))

    def set_excel_cell(self, path: str, *, sheet: str, cell: str, value: Any) -> Any:
        _require_absolute_file(path)
        return run_in_office_thread(lambda: _set_excel_cell(path, sheet, cell, value))


def _require_absolute_file(path: str) -> None:
    """The write backend only ever accepts one absolute, regular, non-symlink file."""
    if not isinstance(path, str) or not os.path.isabs(path):
        raise ValueError("path must be absolute")
    if Path(path).is_symlink() or not Path(path).is_file():
        raise ValueError("Office target must be an existing regular file, not a link")
