import json
import os
import queue
import atexit
import tempfile
import threading
import time
import types
import unittest
from dataclasses import fields
from pathlib import Path
from unittest import mock
from unittest.mock import Mock

from docx import Document
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel, ConfigDict

from office_agent.agent import OfficeAgent, StrictModel, ToolSpec
from office_agent.app import (DARK_THEME, LIGHT_THEME, THEMES, Desktop, Palette, _cap_panel_text,
                               _display_width, _PALETTE, _writable_directory, align_table_text,
                               budget_advice, budget_fraction, build_dispatcher,
                               default_theme_mode, describe_state, detect_system_theme,
                               format_plan_preview, format_progress_line,
                               legacy_refusal_presentation, next_step_hint, render_tool_result,
                               status_badge, status_level, theme_for)
from office_agent.writes import Editor, digest


def _cell_value(path, sheet, cell):
    """Read one cell and close the workbook again.

    On Windows an open workbook handle blocks removal of the temporary directory, and this
    host's tempfile falls back to the project root — so a forgotten handle leaves ``tmp*``
    residue behind after the test run.
    """
    workbook = load_workbook(path)
    try:
        return workbook[sheet][cell].value
    finally:
        workbook.close()


class MockModel:
    """Minimal stand-in for the LangChain chat model used by main.build_llm."""

    def __init__(self, replies):
        self.replies = iter(replies)

    def bind_tools(self, _schemas):
        return self

    def invoke(self, _messages):
        return next(self.replies)


class SessionDiscoveryTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.path = self.root / "selected.docx"
        self.path.write_bytes(b"stable")
        self.editor = Editor(self.root)
        self.desktop = Desktop.__new__(Desktop)

    def tearDown(self):
        self.tmp.cleanup()

    def _checkpoint(self, sha=None):
        session = self.editor.audit / "sessions" / "session-1"
        session.mkdir(parents=True)
        (session / "agent-state.json").write_text("{}", encoding="utf-8")
        (session / "session.json").write_text(json.dumps({
            "version": 1,
            "selected_path": str(self.path.resolve()),
            "file_sha256": sha or digest(self.path),
            "session_dir": str(session.resolve()),
            "question": "inspect",
        }), encoding="utf-8")
        return session

    def test_discovers_exact_current_version(self):
        session = self._checkpoint()
        found = self.desktop._discover_session(self.path, self.editor)
        self.assertIsNotNone(found)
        self.assertEqual(found[0], session)

    def test_rejects_unknown_metadata_schema_and_unbounded_question(self):
        self._checkpoint()
        metadata_file = self.editor.audit / "sessions" / "session-1" / "session.json"
        metadata = json.loads(metadata_file.read_text(encoding="utf-8"))
        metadata["version"] = 99
        metadata["question"] = "x" * 2001
        metadata_file.write_text(json.dumps(metadata), encoding="utf-8")
        self.assertIsNone(self.desktop._discover_session(self.path, self.editor))

    def test_rejects_changed_source_version(self):
        self._checkpoint()
        self.path.write_bytes(b"changed")
        self.assertIsNone(self.desktop._discover_session(self.path, self.editor))

    def test_rejects_audit_symlink(self):
        audit = self.editor.audit
        audit.mkdir()
        outside = self.root / "outside"
        outside.mkdir()
        audit.rmdir()
        try:
            audit.symlink_to(outside, target_is_directory=True)
        except OSError as exc:
            self.skipTest(f"symlink unavailable: {exc}")
        self.assertIsNone(self.desktop._discover_session(self.path, self.editor))


class AskFeedbackTests(unittest.TestCase):
    def test_missing_selection_and_empty_prompt_explain_why_nothing_happened(self):
        desktop = Desktop.__new__(Desktop)
        desktop.path = None
        desktop.editor = None
        desktop.busy = False
        desktop.status = Mock()
        desktop.ask()
        desktop.status.set.assert_called_with("请先点击‘选择授权文件’并选择要处理的文档。")

        desktop.path = Path("sample.xlsx")
        desktop.editor = Mock()
        desktop._resume_session = None
        desktop.prompt = Mock()
        desktop.prompt.get.return_value = "   "
        desktop.ask()
        desktop.status.set.assert_called_with("请输入自然语言任务指令，然后点击‘自然语言查询/规划’。")
        desktop.prompt.focus_set.assert_called_once_with()

    def test_busy_task_reports_wait_instead_of_silent_return(self):
        desktop = Desktop.__new__(Desktop)
        desktop.path = Path("sample.xlsx")
        desktop.editor = Mock()
        desktop.busy = True
        desktop.status = Mock()
        desktop.ask()
        desktop.status.set.assert_called_once_with("上一个任务仍在处理中；请等待结束后再提交。")


class StateSummaryTests(unittest.TestCase):
    def _payload(self, **overrides):
        payload = {"status": "done", "iterations": 3, "token_estimate": 4000,
                   "context_capacity": 8192, "compressions": 1, "answer": None,
                   "errors": [], "resume_checkpoint": r"C:\work\book.xlsx.office_agent\state.json"}
        payload.update(overrides)
        return payload

    def test_budget_stop_tells_the_user_what_to_do_and_that_nothing_was_written(self):
        text = describe_state(self._payload(status="budget", answer=None,
            errors=["Pre-call estimate 7900 exceeds 90% of context (8192 tokens) even after "
                    "compression; raise OFFICE_CONTEXT_WINDOW or narrow the request, then retry. "
                    "No write was performed."]))
        self.assertIn("未执行任何写入", text)
        self.assertIn("OFFICE_CONTEXT_WINDOW", text)
        self.assertIn("可续跑检查点", text)
        self.assertIn("原文件不变", text)

    def test_decision_limit_stop_is_reported_as_safe_with_resume_path(self):
        text = describe_state(self._payload(status="limit", iterations=15, decision_limit=15,
                                            errors=["Stopped after 15 model decisions"]))
        self.assertIn("15 次模型决策上限", text)
        self.assertIn("15/15", text)
        self.assertIn("原文件不变", text)

    def test_completed_answer_is_shown_without_resume_noise(self):
        text = describe_state(self._payload(answer="已生成 B2 预览"))
        self.assertIn("任务完成", text)
        self.assertIn("已生成 B2 预览", text)
        self.assertNotIn("原文件不变", text)

    def test_missing_fields_do_not_break_the_summary(self):
        self.assertIn("结果：", describe_state({"status": "done"}))


class DispatcherGuardTests(unittest.TestCase):
    """The dispatcher is the single chokepoint for every model tool call."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.path = self.root / "sample.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 4; ws["B1"] = "=A1*3"
        wb.save(self.path)
        self.editor = Editor(self.root)

    def test_rejects_other_paths_and_stale_versions(self):
        dispatch = build_dispatcher(self.editor, self.path)
        other = self.root / "other.xlsx"
        Workbook().save(other)  # a real, readable file that is simply not the selected one
        with self.assertRaisesRegex(ValueError, "explicitly selected"):
            dispatch("inspect", {"path": str(other)})
        self.path.write_bytes(self.path.read_bytes() + b" ")
        with self.assertRaisesRegex(ValueError, "changed since this session started"):
            dispatch("inspect", {"path": str(self.path)})

    def test_write_tools_only_ever_return_a_preview(self):
        plans = []
        dispatch = build_dispatcher(self.editor, self.path, on_plan=plans.append)
        result = dispatch("plan_excel_formula_fill", {
            "path": str(self.path), "sheet": "Data", "source_cell": "B1",
            "target_range": "B2:B2", "expected_source_formula": "=A1*3"})
        self.assertEqual(result["preview"]["after"], "{'B2': '=A2*3'}")
        self.assertTrue(result["plan_id"])
        self.assertEqual(plans, [("plan", result["preview"])])
        # A preview is not a write: the source workbook is untouched.
        self.assertIsNone(_cell_value(self.path, "Data", "B2"))
        with self.assertRaisesRegex(ValueError, "Tool not registered"):
            dispatch("plan_drop_everything", {"path": str(self.path)})

    def test_the_tool_name_decides_the_kind_not_the_model_echo(self):
        """A local 7B model echoed kind='word_replace' at the batch tool and failed 7 times;
        the discriminator is redundant, so the tool name must win."""
        path = self.root / "doc.docx"
        document = Document()
        document.add_paragraph("price a")
        document.add_paragraph("price b")
        document.save(path)
        plans = []
        dispatch = build_dispatcher(self.editor, path, on_plan=plans.append)
        result = dispatch("plan_word_batch_replace", {
            "kind": "word_replace", "path": str(path), "old": "price", "new": "amount"})
        self.assertEqual(result["preview"]["kind"], "word_batch_replace")
        self.assertTrue(result["preview"]["confirmation_required"])
        self.assertIn("2 occurrence(s) in 2 paragraph(s)", result["preview"]["target"])
        # A wrong kind cannot smuggle a request into a different tool's rules either.
        other = dispatch("plan_word", {"kind": "excel_set", "path": str(path),
                                       "old": "price", "new": "amount", "paragraph_index": 0})
        self.assertEqual(other["preview"]["kind"], "word_replace")
        self.assertEqual(other["preview"]["target"], "paragraph[0]")

    def test_unregistered_plan_tools_are_refused(self):
        path = self.root / "doc.docx"
        Document().save(path)
        with self.assertRaisesRegex(ValueError, "Tool not registered"):
            build_dispatcher(self.editor, path)("plan_everything", {"path": str(path)})

    def test_read_word_structure_is_registered_and_bounded(self):
        path = self.root / "doc.docx"
        document = Document()
        document.add_paragraph("body")
        document.sections[0].header.paragraphs[0].text = "Header text"
        document.save(path)
        dispatch = build_dispatcher(self.editor, path)
        report = dispatch("read_word_structure", {"path": str(path), "limit": 5})
        self.assertEqual(report["sections"][0]["header"][0]["text"], "Header text")
        self.assertEqual(report["limit"], 5)
        xlsx = self.root / "book.xlsx"
        Workbook().save(xlsx)
        with self.assertRaisesRegex(ValueError, "DOCX required"):
            build_dispatcher(self.editor, xlsx)("read_word_structure",
                                                {"path": str(xlsx), "limit": 5})

    def test_com_read_is_refused_with_the_documented_fallback_when_office_is_absent(self):
        path = self.root / "legacy.doc"
        path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"0" * 64)
        dispatch = build_dispatcher(self.editor, path)
        with mock.patch("office_agent.com.capabilities",
                        return_value={"word": {"available": False, "note": "isolated start denied"}}):
            with self.assertRaisesRegex(ValueError, "COM is unavailable.*read_word_region|"
                                                    "COM is unavailable"):
                dispatch("com_read", {"path": str(path), "include_tables": False, "region": None})


class AskWorkerTests(unittest.TestCase):
    """The window's main flow: the worker must reach the model and feed the panels.

    Regression for a real defect: the worker read the local ``tools`` before assigning it
    (``accesses`` was computed above ``tools = build_tools(...)``), so every task submitted
    from the window died with UnboundLocalError before the first model call and the user
    only ever saw ``失败，原文件保留``. Driving ask() through the event queue catches that
    whole class of "worker references a later local" mistakes.
    """

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.path = self.root / "sample.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 4
        wb.save(self.path)
        self.editor = Editor(self.root)

    def _desktop(self):
        desktop = Desktop.__new__(Desktop)
        desktop.path = self.path
        desktop.editor = self.editor
        desktop.busy = False
        desktop.status = Mock()
        desktop.show = Mock()
        desktop.log = None           # headless: the panels are not built by __new__
        desktop.events = queue.Queue()
        desktop.prompt = Mock()
        desktop.prompt.get.return_value = "把 Data 工作表的 A1 读出来"
        desktop._resume_session = None
        desktop.ocr_provider = None
        desktop.pending_id = None
        desktop.last_verified_id = None
        desktop.last_compressions = 0
        desktop.model_agent = None
        desktop._stop_task = threading.Event()
        desktop._model_task_active = False
        desktop.session_metadata = None
        return desktop

    def _collect(self, desktop, timeout=20):
        """Answer read confirmations exactly like the window does, and collect the events."""
        events = []
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                item = desktop.events.get(timeout=0.2)
            except queue.Empty:
                continue
            events.append(item)
            kind, payload = item
            if kind == "confirm_read":
                _plan, event, decision = payload
                decision[0] = True          # the user approves this one read
                event.set()
            if kind == "idle":
                break
        return events

    def test_ask_runs_the_worker_and_reports_the_read_result(self):
        model = MockModel([
            {"tool_calls": [{"name": "read_excel_region",
                             "args": {"path": str(self.path), "sheet": "Data", "row": 1,
                                      "column": 1},
                             "id": "1"}]},
            {"content": "A1 是 4"},
        ])
        desktop = self._desktop()
        with mock.patch.object(Desktop, "_build_model", return_value=model), \
                mock.patch("office_agent.app.build_dispatcher",
                           return_value=lambda name, args: {"sheet": "Data",
                                                            "cells": [{"cell": "A1", "value": 4}]}):
            desktop.ask()
            events = self._collect(desktop)
        kinds = [kind for kind, _ in events]
        self.assertNotIn("error", kinds, [payload for kind, payload in events if kind == "error"])
        self.assertIn("text", kinds, kinds)              # budget/window line
        self.assertIn("progress", kinds, kinds)          # iterations/tokens/compressions
        self.assertIn("result", kinds, kinds)            # read-result panel feed
        self.assertIn("state", kinds, kinds)
        self.assertEqual(kinds[-1], "idle", kinds)
        state = next(payload for kind, payload in events if kind == "state")
        self.assertEqual(state["status"], "done")
        self.assertEqual(state["answer"], "A1 是 4")
        result = next(payload for kind, payload in events if kind == "result")
        self.assertEqual(result["tool"], "read_excel_region")
        self.assertIn("A1: 4", result["text"])
        # The worker leaves busy=True; only the window's idle handler clears it.
        self.assertTrue(desktop.busy)

    def test_ask_reports_why_nothing_happened_without_a_selection_or_task(self):
        desktop = self._desktop()
        desktop.path = None
        desktop.editor = None
        desktop.ask()
        desktop.status.set.assert_called_with("请先点击‘选择授权文件’并选择要处理的文档。")
        desktop.path, desktop.editor = self.path, self.editor
        desktop.prompt.get.return_value = "  "
        desktop.ask()
        desktop.status.set.assert_called_with("请输入自然语言任务指令，然后点击‘自然语言查询/规划’。")
        self.assertFalse(desktop.busy)


class DrainEventProtocolTests(unittest.TestCase):
    """drain() routes every documented event kind and survives a malformed payload."""

    def setUp(self):
        self.desktop = Desktop.__new__(Desktop)
        self.desktop.status = Mock()
        self.desktop.show = Mock()
        self.desktop.log = Mock()
        self.desktop.root = Mock()
        self.desktop.events = queue.Queue()
        self.desktop.busy = True
        self.desktop.last_compressions = 0
        self.desktop.pending_id = None
        self.desktop.last_verified_id = None

    def _drain(self):
        self.desktop.drain()
        self.desktop.root.after.assert_called_with(150, self.desktop.drain)  # pump stays alive

    def _logged(self):
        return "\n".join(str(call.args[1]) for call in self.desktop.log.insert.call_args_list)

    def _put_all(self):
        for item in [
            ("text", "预算：固定开销约 100 token，本次窗口 16384 token"),
            ("progress", {"iterations": 2, "compressions": 0, "status": "running",
                          "tokens": 100, "capacity": 1000}),
            ("progress", {"iterations": 3, "compressions": 2, "status": "running",
                          "tokens": 950, "capacity": 1000}),
            ("result", {"tool": "read_word_text", "text": "正文内容"}),
            ("plan", {"id": "p1", "kind": "excel_set", "target": "Data!A1", "before": "4",
                      "after": "7", "risks": ["公式不会计算"], "confirmation_required": True}),
            ("verified", {"id": "p1", "status": "verified", "output": "out.xlsx",
                          "backup": "backup.xlsx"}),
            ("undone", {"id": "p1", "status": "rolled_back"}),
            ("state", {"status": "budget", "iterations": 4, "token_estimate": 950,
                       "context_capacity": 1000, "compressions": 1, "answer": None,
                       "errors": ["Pre-call estimate 950 exceeds 90%"], "evidence": [],
                       "resume_checkpoint": None}),
            ("error", "GuardError: refused"),
            ("idle", None),
        ]:
            self.desktop.events.put(item)

    def test_every_event_kind_is_presented_with_its_safety_wording(self):
        self._put_all()
        self._drain()
        logged = self._logged()
        self.assertIn("预算：固定开销约 100 token", logged)
        self.assertIn("上下文压缩事件：第 2 次", logged)
        self.assertIn("正文内容", logged)                 # read result reaches the panel feed
        self.assertIn("差异预览（未写入）", logged)        # plan preview is flagged as unwritten
        self.assertIn("已验证输出与审计", logged)
        self.assertIn("输出副本已撤销；源文件和备份未修改", logged)
        self.assertIn("未执行任何写入", logged)            # budget stop keeps its explanation
        self.assertIn("错误: GuardError: refused", logged)
        statuses = [call.args[0] for call in self.desktop.status.set.call_args_list]
        self.assertTrue(any("未写入" in text and "单独确认" in text for text in statuses), statuses)
        self.assertTrue(any("源文件未改动" in text for text in statuses), statuses)
        self.assertTrue(any("源文件和备份未修改" in text for text in statuses), statuses)
        # A verified write consumes the plan, and an undo clears the output id: neither may
        # linger and make the window advertise work that no longer exists.
        self.assertIsNone(self.desktop.pending_id)
        self.assertIsNone(self.desktop.last_verified_id)
        self.assertFalse(self.desktop.busy)
        self.assertEqual(statuses[-1], "失败，原文件保留")

    def test_a_plan_waits_for_confirmation_until_a_verified_write_consumes_it(self):
        self.desktop.events.put(("plan", {"id": "p1", "kind": "excel_set", "target": "Data!A1",
                                          "before": "4", "after": "7", "risks": [],
                                          "confirmation_required": True}))
        self._drain()
        self.assertEqual(self.desktop.pending_id, "p1")
        self.desktop.events.put(("verified", {"id": "p1", "status": "verified"}))
        self._drain()
        self.assertIsNone(self.desktop.pending_id)
        self.assertEqual(self.desktop.last_verified_id, "p1")

    def test_a_malformed_payload_does_not_stop_the_pump(self):
        self.desktop.events.put(("progress", {"iterations": 1}))    # missing keys
        self.desktop.events.put(("idle", None))
        self._drain()
        self.assertIn("界面处理事件 progress 时出错", self._logged())
        self.assertFalse(self.desktop.busy, "the later idle event must still be processed")

    def test_read_confirmation_is_answered_and_a_closed_dialog_refuses(self):
        offered = types.SimpleNamespace(tool="inspect", arguments={"path": "x"})
        event, decision = threading.Event(), [False]
        self.desktop.events.put(("confirm_read", (offered, event, decision)))
        with mock.patch("office_agent.app.messagebox.askyesno", return_value=True):
            self._drain()
        self.assertTrue(event.is_set())
        self.assertTrue(decision[0])
        # A dialog that cannot be shown means refusal, and the worker is never left waiting.
        event2, decision2 = threading.Event(), [True]
        self.desktop.events.put(("confirm_read", (offered, event2, decision2)))
        with mock.patch("office_agent.app.messagebox.askyesno",
                        side_effect=RuntimeError("no display")):
            self._drain()
        self.assertTrue(event2.is_set())
        self.assertFalse(decision2[0])
        # An already-expired request is skipped instead of resurrecting a stale approval.
        event3, decision3 = threading.Event(), [False]
        event3.set()
        self.desktop.events.put(("confirm_read", (offered, event3, decision3)))
        with mock.patch("office_agent.app.messagebox.askyesno") as askyesno:
            self._drain()
        askyesno.assert_not_called()
        self.assertFalse(decision3[0])


class PresentationFormattingTests(unittest.TestCase):
    """Pure presentation helpers: they make safety facts louder, they never hide them."""

    def test_plan_preview_is_marked_unwritten_and_keeps_every_field(self):
        text = format_plan_preview({"id": "p1", "kind": "excel_set", "target": "Data!A1",
                                    "before": "4", "after": "7", "risks": ["公式不会计算"],
                                    "confirmation_required": True, "extra_field": {"kept": True}})
        self.assertIn("尚未写入", text)
        self.assertIn("原文件保持不变", text)
        self.assertIn("需要你单独确认：是", text)
        self.assertIn("公式不会计算", text)
        self.assertIn("extra_field", text)          # unknown plan keys are appended, not dropped
        self.assertIn('"kept": true', text)

    def test_plan_preview_handles_an_empty_or_odd_payload(self):
        self.assertIn("尚未写入", format_plan_preview({}))
        self.assertIn("尚未写入", format_plan_preview(None))

    def test_progress_line_reports_iterations_tokens_and_compressions(self):
        self.assertEqual(format_progress_line(3, 4000, 16384, 1, decision_limit=15),
                         "迭代 3/15 · 估算 token 4000/16384（24%）· 压缩 1 次")

    def test_progress_line_survives_an_unusable_capacity(self):
        self.assertIn("未知", format_progress_line(1, 10, 0, 0))
        self.assertIn("未知", format_progress_line(1, None, None, 0))
        self.assertIsNone(budget_fraction(10, 0))
        self.assertIsNone(budget_fraction("x", 10))
        self.assertEqual(budget_fraction(500, 1000), 0.5)

    def test_status_levels_follow_the_agent_status_field(self):
        self.assertEqual(status_level("done"), "ok")
        self.assertEqual(status_level("budget"), "warn")
        self.assertEqual(status_level("limit"), "warn")
        self.assertEqual(status_level(None), "info")

    def test_long_panel_content_is_capped_only_with_an_explicit_note(self):
        self.assertEqual(_cap_panel_text("内容", limit=10), ("内容", ""))
        capped, note = _cap_panel_text("x" * 10, limit=4)
        self.assertEqual(capped, "xxxx")
        self.assertIn("共 10 字符", note)

    def test_status_badge_always_carries_words_not_just_a_colour(self):
        self.assertEqual(status_badge("done"), ("ok", "已完成"))
        self.assertEqual(status_badge("limit")[0], "warn")
        self.assertEqual(status_badge("budget")[0], "warn")
        self.assertEqual(status_badge("confirmation_required")[0], "warn")
        self.assertEqual(status_badge("running"), ("info", "进行中"))
        self.assertEqual(status_badge(None), ("info", "尚未开始"))
        self.assertEqual(status_badge("something_new"), ("info", "something_new"))

    def test_table_rows_are_aligned_by_display_width_only(self):
        text = align_table_text("  表格 [2 行]\n  | 姓名 | 分 |\n  | 李炜基 | 9 |")
        rows = text.split("\n")
        self.assertEqual(rows[0], "  表格 [2 行]")          # non-table lines are untouched
        body = [row for row in rows if "|" in row]
        widths = {_display_width(row) for row in body}
        self.assertEqual(len(widths), 1, f"columns are not aligned: {body}")
        self.assertIn("李炜基", text)
        self.assertIn("姓名", text)

    def test_alignment_keeps_a_row_that_is_not_a_table(self):
        plain = "A1: 4\nB1: =A1*3（公式文本；未经验证）"
        self.assertEqual(align_table_text(plain), plain)

    def test_budget_reminder_comes_from_the_same_fraction_as_the_bar(self):
        self.assertEqual(budget_advice(None), "")
        self.assertEqual(budget_advice(0.5), "")
        self.assertIn("85%", budget_advice(0.85))
        self.assertIn("OFFICE_CONTEXT_WINDOW", budget_advice(0.85))
        self.assertIn("已达上限区", budget_advice(0.95))

    def test_merge_plan_lists_every_file_it_will_read(self):
        """The merge preview is the only gate on reading the model-named sibling files."""
        text = format_plan_preview({"id": "m1", "kind": "docx_merge", "source": r"C:\w\main.docx",
                                    "sources": ["A.docx", "B.docx"], "confirmation_required": True})
        self.assertIn("将读取的其它文件", text)
        self.assertIn("A.docx", text)
        self.assertIn("B.docx", text)
        self.assertIn("未列出的文件不会被打开", text)
        self.assertIn("需要你单独确认：是", text)

    def test_batch_report_shows_the_read_only_guarantee_and_every_skip(self):
        text = render_tool_result("batch_documents", {
            "directory": r"C:\w", "operation": "summary", "file_count": 1, "skipped_count": 2,
            "error_count": 1, "wrote_nothing": True, "truncated": True, "limit": 20,
            "files": [{"file": "a.docx", "paragraph_count": 3}],
            "skipped": [{"file": "old.doc", "reason": "legacy format needs COM"},
                        {"file": "../out.docx", "reason": "outside the authorized directory"}],
            "errors": [{"file": "b.xlsx", "error": "PermissionError: locked"}]})
        self.assertIn("未写入任何文件", text)
        self.assertIn("wrote_nothing=true", text)
        self.assertIn("a.docx", text)
        self.assertIn("old.doc", text)
        self.assertIn("legacy format needs COM", text)
        self.assertIn("outside the authorized directory", text)
        self.assertIn("PermissionError: locked", text)
        self.assertIn("其余文件未被读取", text)

    def test_next_step_hint_names_the_step_that_is_actually_available(self):
        idle = next_step_hint(has_file=False, busy=False, pending=False, verified=False)
        self.assertIn("选择授权文件", idle)
        running = next_step_hint(has_file=True, busy=True, pending=True, verified=False)
        self.assertIn("等待本轮结束", running)
        pending = next_step_hint(has_file=True, busy=False, pending=True, verified=False)
        self.assertIn("确认并保存输出副本", pending)
        self.assertIn("未写入", pending)
        verified = next_step_hint(has_file=True, busy=False, pending=False, verified=True)
        self.assertIn("撤销最近输出副本", verified)
        failed = next_step_hint(has_file=True, busy=False, pending=False, verified=False,
                                level="danger")
        self.assertIn("原文件保持不变", failed)


class ThemeTokenTests(unittest.TestCase):
    """The design tokens are complete, switchable, and readable in both themes.

    Stage 1 of UI_POLISH_PLANS: one token set drives every colour/spacing/size, and both
    palettes must keep WCAG AA contrast for every text-on-surface pair the window renders.
    """

    # Every pair of token roles the window actually paints as text on a surface.
    TEXT_PAIRS = (("ink", "card"), ("ink", "window"), ("ink", "card_alt"),
                  ("muted", "card"), ("muted", "window"), ("faint", "card"),
                  ("accent", "card"), ("accent", "window"), ("accent", "accent_soft"),
                  ("ok", "card"), ("warn", "card"), ("danger", "card"),
                  ("warn", "warn_soft"), ("ok", "ok_soft"), ("danger", "danger_soft"),
                  ("on_solid", "accent"), ("on_solid", "warn"), ("on_solid", "danger"))

    @staticmethod
    def _linear(channel):
        value = channel / 255
        return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4

    @classmethod
    def _luminance(cls, colour):
        raw = colour.lstrip("#")
        red, green, blue = (int(raw[index:index + 2], 16) for index in (0, 2, 4))
        return 0.2126 * cls._linear(red) + 0.7152 * cls._linear(green) + 0.0722 * cls._linear(blue)

    @classmethod
    def _contrast(cls, foreground, background):
        """WCAG 2.1 contrast ratio between two hex colours."""
        first, second = cls._luminance(foreground), cls._luminance(background)
        lighter, darker = max(first, second), min(first, second)
        return (lighter + 0.05) / (darker + 0.05)

    def test_both_themes_define_every_token(self):
        for name, theme in THEMES.items():
            with self.subTest(theme=name):
                self.assertEqual(theme.name, name)
                for field in fields(Palette):
                    self.assertTrue(getattr(theme, field.name), f"{name}.{field.name} is empty")

    def test_text_colours_meet_wcag_aa_in_both_themes(self):
        for name, theme in THEMES.items():
            for foreground, background in self.TEXT_PAIRS:
                with self.subTest(theme=name, pair=f"{foreground}/{background}"):
                    self.assertGreaterEqual(
                        self._contrast(getattr(theme, foreground), getattr(theme, background)),
                        4.5, f"{name}: {foreground} on {background}")

    def test_dark_theme_avoids_glare_and_keeps_the_expected_polarity(self):
        self.assertGreater(self._luminance(DARK_THEME.window), 0.005)   # not glaring pure black
        self.assertLess(self._luminance(LIGHT_THEME.window), 0.95)      # not glaring pure white
        self.assertLess(self._luminance(DARK_THEME.card), self._luminance(DARK_THEME.ink))
        self.assertGreater(self._luminance(LIGHT_THEME.card), self._luminance(LIGHT_THEME.ink))

    def test_theme_lookup_and_the_compatibility_mapping(self):
        self.assertIs(theme_for("dark"), DARK_THEME)
        self.assertIs(theme_for("LIGHT"), LIGHT_THEME)
        self.assertIs(theme_for("nonsense"), LIGHT_THEME)
        self.assertIn(detect_system_theme(), ("light", "dark"))     # never guesses anything else
        for key in ("window", "card", "border", "ink", "muted", "accent", "accent_dark",
                    "accent_soft", "ok", "ok_soft", "warn", "warn_soft", "danger", "danger_soft"):
            self.assertIn(key, _PALETTE)                            # historical names still resolve
        self.assertEqual(_PALETTE["card"], LIGHT_THEME.card)

    def test_theme_mode_reads_the_environment_when_set(self):
        with mock.patch.dict(os.environ, {"OFFICE_AGENT_THEME": "dark"}):
            self.assertEqual(default_theme_mode(), "dark")
        with mock.patch.dict(os.environ, {"OFFICE_AGENT_THEME": "Light"}):
            self.assertEqual(default_theme_mode(), "light")
        with mock.patch.dict(os.environ, {"OFFICE_AGENT_THEME": "nonsense"}):
            self.assertEqual(default_theme_mode(), "auto")


_TK_ROOT = None
_TK_ERROR = None


def _shared_tk_root():
    """One Tk root for every window test.

    Creating several Tcl interpreters in a row intermittently fails with
    ``invalid command name "tcl_findLibrary"`` / "Can't find a usable tk.tcl", so the root is
    created once, reused, and retried a couple of times before the tests give up.
    """
    global _TK_ROOT, _TK_ERROR
    if _TK_ROOT is not None:
        try:
            if _TK_ROOT.winfo_exists():
                return _TK_ROOT
        except Exception:                     # pragma: no cover - root already gone
            _TK_ROOT = None
    if _TK_ERROR is not None:
        raise unittest.SkipTest(f"no display for Tk: {_TK_ERROR}")
    import tkinter
    last = None
    for _attempt in range(3):
        try:
            root = tkinter.Tk()
        except Exception as exc:              # pragma: no cover - no display / Tcl race
            last = exc
            time.sleep(0.2)
            continue
        root.withdraw()                       # stay invisible; no mainloop is ever started
        atexit.register(root.destroy)
        _TK_ROOT = root
        return root
    _TK_ERROR = last
    raise unittest.SkipTest(f"no display for Tk: {last}")


class DesktopWindowTests(unittest.TestCase):
    """Build the real window when a display is available; skip on a headless host.

    Verifies the parts of UI_POLISH_PLANS stage 1 that only exist once widgets do: theme
    switching recolours everything, and it must not rebuild widgets or lose session content.
    """

    @classmethod
    def setUpClass(cls):
        cls.root = _shared_tk_root()

    def setUp(self):
        # The window schedules a startup capability probe. Inside a test it must not touch
        # real COM (slow, and on this host it reports an unavailable channel that would
        # overwrite whatever the test is asserting).
        patcher = mock.patch("office_agent.com.capabilities",
                             return_value={"word": {"available": True, "note": "test"},
                                           "excel": {"available": True, "note": "test"}})
        patcher.start()
        self.addCleanup(patcher.stop)
        self.desktop = Desktop(self.root)
        self.addCleanup(self.root.update_idletasks)
        self.addCleanup(self._drop_widgets)

    def _drop_widgets(self):
        """Remove this test's window so the next one starts from a clean root."""
        for child in list(self.root.winfo_children()):
            child.destroy()

    def _tag_colour(self, tag, option="foreground"):
        return str(self.desktop.log.tag_cget(tag, option))

    def test_window_builds_with_the_tokens_of_the_active_theme(self):
        desktop = self.desktop
        theme = desktop._theme
        self.assertEqual(desktop._theme_name, detect_system_theme())
        self.assertIn(desktop._theme_mode, ("auto", "light", "dark"))
        self.assertEqual(str(desktop.log.cget("background")), theme.card)
        self.assertEqual(str(desktop.root.cget("background")), theme.window)
        self.assertEqual(self._tag_colour("plan"), theme.warn)
        self.assertEqual(self._tag_colour("ok"), theme.ok)
        for style, option, expected in (("StatusInfo.TLabel", "background", theme.card),
                                        ("StatusWarn.TLabel", "foreground", theme.warn),
                                        ("StatusOk.TLabel", "foreground", theme.ok),
                                        ("StatusDanger.TLabel", "foreground", theme.danger),
                                        ("Accent.TButton", "background", theme.accent),
                                        ("TButton", "focuscolor", theme.focus),
                                        ("Section.TLabelframe", "bordercolor", theme.border)):
            with self.subTest(style=style):
                self.assertEqual(desktop.style.lookup(style, option), expected)

    def test_theme_cycle_recolours_without_rebuilding_or_losing_content(self):
        desktop = self.desktop
        self.addCleanup(desktop._apply_theme, "auto")
        desktop.show("会话内容必须保留")
        desktop._set_panel(desktop.result_text, "读取结果必须保留")
        desktop._set_panel(desktop.plan_text, "计划预览必须保留")
        widgets = (str(desktop.log), str(desktop.plan_text), str(desktop.result_text))
        seen = []
        for _ in range(3):                                # auto -> light -> dark -> auto
            desktop.cycle_theme()
            self.root.update_idletasks()
            theme = desktop._theme
            seen.append(theme.name)
            self.assertEqual(str(desktop.log.cget("background")), theme.card)
            self.assertEqual(str(desktop.root.cget("background")), theme.window)
            self.assertEqual(self._tag_colour("body"), theme.ink)
            self.assertEqual(desktop.style.lookup("Accent.TButton", "background"), theme.accent)
            self.assertIn("主题", desktop.theme_var.get())
            self.assertEqual((str(desktop.log), str(desktop.plan_text), str(desktop.result_text)),
                             widgets, "theme switch must not rebuild widgets")
            self.assertIn("会话内容必须保留", desktop.log.get("1.0", "end-1c"))
            self.assertEqual(desktop.result_text.get("1.0", "end-1c"), "读取结果必须保留")
            self.assertEqual(desktop.plan_text.get("1.0", "end-1c"), "计划预览必须保留")
        self.assertEqual(sorted(set(seen)), ["dark", "light"], seen)

    def test_progress_bar_and_sentence_share_one_source(self):
        desktop = self.desktop
        desktop._set_progress(3, 9000, 10000, 2)          # 90% -> danger band
        self.assertEqual(str(desktop.budget_bar.cget("style")),
                         "BudgetDanger.Horizontal.TProgressbar")
        self.assertIn("9000/10000", desktop.progress_var.get())
        desktop._set_progress(2, 500, 10000, 0)           # 5% -> normal band
        self.assertEqual(str(desktop.budget_bar.cget("style")),
                         "Budget.Horizontal.TProgressbar")
        desktop._apply_theme("dark")                      # a theme switch keeps both in step
        desktop._set_progress(3, 8500, 10000, 1)          # 85% -> warn band
        self.assertEqual(str(desktop.budget_bar.cget("style")),
                         "BudgetWarn.Horizontal.TProgressbar")
        self.assertEqual(desktop.style.lookup("BudgetWarn.Horizontal.TProgressbar", "background"),
                         DARK_THEME.warn)

    def test_first_screen_guides_instead_of_showing_blank_panels(self):
        desktop = self.desktop
        self.assertIn("选择授权文件", desktop.next_step_var.get())
        self.assertIn("尚未生成修改计划", desktop.plan_text.get("1.0", "end-1c"))
        self.assertIn("尚无读取结果", desktop.result_text.get("1.0", "end-1c"))
        self.assertEqual(desktop.badge_var.get(), "尚未开始")

    def test_extreme_single_line_and_long_error_do_not_widen_the_window(self):
        """20k characters with no space, plus a 5k-char error, must wrap instead of stretch."""
        desktop = self.desktop
        try:
            self.root.attributes("-alpha", 0.0)
        except Exception as exc:                # pragma: no cover - platform dependent
            self.skipTest(f"window alpha unavailable: {exc}")
        self.addCleanup(self.root.attributes, "-alpha", 1.0)
        self.addCleanup(self.root.withdraw)
        self.root.deiconify()
        self.root.geometry("1120x800")
        self.root.update()
        desktop._set_panel(desktop.result_text, "x" * 20000)
        desktop._set_error("E" * 5000)
        self.root.update()
        self.assertLessEqual(self.root.winfo_reqwidth(), 1120,
                            "long content stretched the window")
        self.assertLessEqual(desktop.result_text.winfo_reqwidth(), 400)
        self.assertGreaterEqual(desktop.status_card.winfo_width(), 400)
        self.assertIn("错误全文共 5000 字符", desktop.error_var.get())

    def test_narrow_geometry_keeps_every_control_inside_the_window(self):
        """800x600 must keep every key control mapped, on-screen, clickable and usable."""
        root = self.root
        desktop = self.desktop
        try:
            root.attributes("-alpha", 0.0)      # realize real geometry without showing it
        except Exception as exc:                # pragma: no cover - platform dependent
            self.skipTest(f"window alpha unavailable: {exc}")
        self.addCleanup(root.attributes, "-alpha", 1.0)
        self.addCleanup(root.withdraw)          # LIFO: hide first, then restore alpha
        root.deiconify()                        # a withdrawn window keeps its old geometry
        root.geometry("800x600")
        root.update()                           # real Configure + geometry pass
        self.assertGreaterEqual(root.winfo_width(), 780, "the window did not reach 800 px")
        self.assertLessEqual(root.winfo_width(), 820)
        right_edge = root.winfo_rootx() + root.winfo_width()
        bottom_edge = root.winfo_rooty() + root.winfo_height()
        widgets = (desktop.status_card, desktop.badge_label, desktop.next_step_label,
                   desktop.progress_label, desktop.budget_bar, desktop.log, desktop.prompt,
                   desktop.ask_button, desktop.resume_button, desktop.confirm_button,
                   desktop.undo_button, desktop.clear_button, desktop.theme_button)
        for widget in widgets:
            with self.subTest(widget=str(widget).rsplit(".", 1)[-1]):
                self.assertTrue(widget.winfo_ismapped(), "widget is not mapped at 800x600")
                self.assertLessEqual(widget.winfo_rootx() + widget.winfo_width(), right_edge + 2)
                self.assertLessEqual(widget.winfo_rooty() + widget.winfo_height(), bottom_edge + 2)
                self.assertGreaterEqual(widget.winfo_width(), 20, "widget collapsed to nothing")
        # Compact chrome gives the body real height, and each notebook surface gets all of it.
        self.assertTrue(desktop._compact, "short windows must use the compact chrome")
        self.assertGreaterEqual(desktop.log.winfo_height(), 40)
        # Whatever the row shape, no action may stick out of the window.
        for row in {int(widget.grid_info()["row"]) for widget in
                    (desktop.ask_button, desktop.confirm_button, desktop.resume_button,
                     desktop.undo_button, desktop.clear_button)}:
            needed = sum(widget.winfo_reqwidth() + desktop._sp(2) for widget in
                         (desktop.ask_button, desktop.confirm_button, desktop.resume_button,
                          desktop.undo_button, desktop.clear_button)
                         if int(widget.grid_info()["row"]) == row)
            self.assertLessEqual(needed, root.winfo_width(), f"action row {row} overflows")
        for frame, text in ((desktop.plan_frame, desktop.plan_text),
                            (desktop.result_frame, desktop.result_text)):
            desktop.notebook.select(frame)
            root.update()
            self.assertTrue(text.winfo_ismapped(), "the selected tab's pane must be readable")
            self.assertGreaterEqual(text.winfo_height(), 40, "pane has no readable height")
            self.assertGreaterEqual(text.winfo_width(), 100)
        root.geometry("1120x800")               # resize back: layout must follow the width
        root.update()
        self.assertFalse(desktop._compact)
        self.assertTrue(desktop.subtitle_label.winfo_ismapped())
        self.assertTrue(desktop.ask_button.winfo_ismapped())
        self.assertGreaterEqual(desktop.log.winfo_height(), 60)
        desktop.notebook.select(desktop.plan_frame)
        root.update()
        self.assertGreaterEqual(desktop.plan_text.winfo_height(), 80)
        # 800 px is the "dense" tier, not the compact one: the flow line stays, only the pane
        # footnote steps aside (it repeats what the pane body and status card already say).
        self.assertTrue(desktop._dense)
        self.assertTrue(desktop.subtitle_label.winfo_ismapped())
        self.assertTrue(desktop.prompt_hint.winfo_ismapped())
        self.assertFalse(desktop.plan_hint.winfo_ismapped())
        root.geometry("1400x900")               # tall enough: everything comes back
        root.update()
        self.assertFalse(desktop._dense)
        self.assertTrue(desktop.plan_hint.winfo_ismapped())
        self.assertTrue(desktop.prompt_hint.winfo_ismapped())
        self.assertGreaterEqual(desktop.plan_text.winfo_height(), 150)

    def test_narrow_window_keeps_every_action_clickable(self):
        desktop = self.desktop
        self.assertLessEqual(tuple(self.root.minsize()), (820, 620))
        desktop._layout_actions(narrow=True)   # forced: no update() in between, so it sticks
        widgets = (desktop.ask_button, desktop.resume_button, desktop.confirm_button,
                   desktop.undo_button, desktop.clear_button)
        rows = {}
        for widget in widgets:
            info = widget.grid_info()
            self.assertTrue(info, f"{widget.cget('text')} is not placed at all")
            rows.setdefault(int(info["row"]), []).append(widget)
        self.assertEqual(sorted(rows), [0, 1], "narrow mode must use two rows")
        # The row that holds the primary actions has to fit the minimum window width.
        gap = desktop._sp(2)
        primary = sum(widget.winfo_reqwidth() + gap for widget in rows[0]) + gap * 2
        self.assertLessEqual(primary, 820, f"primary actions need {primary}px")
        desktop._layout_actions(narrow=False)
        self.assertEqual(sorted({int(w.grid_info()["row"]) for w in widgets}), [0],
                         "the wide layout keeps every action on one row")

    def test_badge_and_plan_emphasis_follow_the_event_stream(self):
        desktop = self.desktop
        desktop.path = Path("sample.xlsx")     # as after select(): a plan implies a chosen file
        plan = {"id": "p9", "kind": "excel_set", "target": "Data!A1", "before": "4",
                "after": "7", "risks": [], "confirmation_required": True}
        desktop._handle_event("plan", plan)
        self.root.update_idletasks()
        self.assertEqual(desktop.badge_var.get(), "等待你确认计划")
        self.assertEqual(desktop.status_badge.level, "warn")
        self.assertEqual(str(desktop.plan_frame.cget("style")), "PlanPending.TLabelframe")
        self.assertEqual(str(desktop.confirm_button.cget("style")), "Pending.TButton")
        self.assertIn("确认并保存输出副本", desktop.next_step_var.get())

        desktop._handle_event("verified", {"id": "p9", "status": "verified", "output": "o.xlsx"})
        self.root.update_idletasks()
        self.assertEqual(desktop.badge_var.get(), "已生成输出副本")
        self.assertEqual(desktop.status_badge.level, "ok")
        self.assertEqual(str(desktop.plan_frame.cget("style")), "Plan.TLabelframe")
        self.assertIn("撤销最近输出副本", desktop.next_step_var.get())

        desktop._handle_event("undone", {"id": "p9", "status": "rolled_back"})
        self.assertEqual(desktop.badge_var.get(), "已撤销输出副本")

        desktop._handle_event("error", "GuardError: refused")
        self.assertEqual(desktop.badge_var.get(), "失败，原文件保留")
        self.assertEqual(desktop.status_badge.level, "danger")
        self.assertEqual(desktop.status.get(), "失败，原文件保留")
        self.assertIn("原文件保持不变", desktop.next_step_var.get())


class ComponentTests(unittest.TestCase):
    """Stage 3 components: the chip, the two meters, the cards and the safety markers."""

    @classmethod
    def setUpClass(cls):
        cls.root = _shared_tk_root()

    def setUp(self):
        patcher = mock.patch("office_agent.com.capabilities",
                             return_value={"word": {"available": True, "note": "test"},
                                           "excel": {"available": True, "note": "test"}})
        patcher.start()
        self.addCleanup(patcher.stop)
        self.desktop = Desktop(self.root)
        self.addCleanup(self.root.update_idletasks)
        self.addCleanup(self._drop_widgets)

    def _drop_widgets(self):
        for child in list(self.root.winfo_children()):
            child.destroy()

    def test_status_badge_is_a_chip_with_colour_shape_and_words(self):
        badge = self.desktop.status_badge
        for level, mark in (("info", "●"), ("ok", "√"), ("warn", "!"), ("danger", "×")):
            with self.subTest(level=level):
                badge.set(level, "示例状态")
                self.root.update_idletasks()
                self.assertEqual(badge.level, level)
                items = badge.canvas.find_all()
                self.assertEqual(len(items), 2, "the chip draws a background and one text item")
                text_item = [item for item in items
                             if badge.canvas.type(item) == "text"][0]
                self.assertEqual(badge.canvas.itemcget(text_item, "text"), f"{mark} 示例状态")
                self.assertGreaterEqual(badge.canvas.winfo_reqheight(), 16)
                self.assertGreaterEqual(badge.canvas.winfo_reqwidth(), 30)
        badge.set("nonsense", "未知级别落到 info")     # an unknown level must not crash the chip
        self.assertEqual(badge.level, "info")

    def test_the_chip_follows_the_theme(self):
        badge = self.desktop.status_badge
        badge.set("warn", "等待")
        self.desktop._apply_theme("dark")
        self.root.update_idletasks()
        fills = {badge.canvas.itemcget(item, "fill") for item in badge.canvas.find_all()}
        self.assertIn(DARK_THEME.warn_soft, fills)
        self.desktop._apply_theme("auto")

    def test_decision_and_budget_meters_share_the_progress_sentence(self):
        desktop = self.desktop
        desktop.model_agent = types.SimpleNamespace(decision_limit=15)
        desktop._set_progress(7, 9000, 10000, 2)
        self.assertIn("迭代 7/15", desktop.progress_var.get())
        self.assertIn("9000/10000", desktop.progress_var.get())
        self.assertEqual(desktop.iteration_var.get(), "决策 7/15")
        self.assertEqual(float(desktop.iteration_bar.cget("value")), 7.0)
        self.assertEqual(float(desktop.iteration_bar.cget("maximum")), 15.0)
        self.assertAlmostEqual(float(desktop.budget_bar.cget("value")), 90.0, places=3)
        self.assertEqual(str(desktop.budget_bar.cget("style")),
                         "BudgetDanger.Horizontal.TProgressbar")
        desktop._set_progress(0, "—", "—", 0)          # the reset path stays consistent
        self.assertEqual(desktop.iteration_var.get(), "决策 0/15")
        self.assertEqual(float(desktop.iteration_bar.cget("value")), 0.0)
        self.assertEqual(float(desktop.budget_bar.cget("value")), 0.0)

    def test_cards_come_from_one_container_helper_and_keep_their_layers(self):
        desktop = self.desktop
        for frame in (desktop.status_card, desktop.log_frame, desktop.plan_frame,
                      desktop.result_frame):
            self.assertEqual(frame.winfo_class(), "TLabelframe")
            self.assertGreaterEqual(frame.grid_columnconfigure(0)["weight"], 1)
        self.assertEqual(str(desktop.status_card.cget("style")), "Section.TLabelframe")
        self.assertEqual(str(desktop.plan_frame.cget("style")), "Plan.TLabelframe")
        self.assertEqual(str(desktop.log_frame.cget("style")), "Quiet.TLabelframe")

    def test_plan_and_output_cards_render_through_the_shared_formatters(self):
        desktop = self.desktop
        desktop.path = Path("sample.xlsx")
        plan = {"id": "p3", "kind": "excel_set", "target": "Data!A1", "before": "4",
                "after": "7", "risks": ["公式不会计算"], "confirmation_required": True}
        desktop._render_plan_card(plan, plan_id="p3")
        body = desktop.plan_text.get("1.0", "end-1c")
        self.assertIn("尚未写入", body)
        self.assertIn("公式不会计算", body)
        self.assertEqual(body, format_plan_preview(plan))     # same text as the pure formatter
        self.assertEqual(desktop.plan_badge_var.get(), "● 未写入 · 等待你确认")
        self.assertEqual(str(desktop.confirm_button.cget("style")), "Pending.TButton")
        desktop._render_output_card({"id": "p3", "status": "verified", "output": "o.xlsx"},
                                    undone=False)
        self.assertIn("已写入副本", desktop.plan_text.get("1.0", "end-1c"))
        self.assertEqual(str(desktop.confirm_button.cget("style")), "Confirm.TButton")

    def test_busy_state_disables_actions_and_shows_a_working_indicator(self):
        desktop = self.desktop
        desktop._set_busy_ui(True)
        self.root.update_idletasks()
        self.assertTrue(desktop.busy_var.get().startswith("●"), desktop.busy_var.get())
        self.assertTrue(desktop.busy_bar.winfo_manager(), "no working indicator while busy")
        self.assertIn("等待本轮结束", desktop.next_step_var.get())
        for button in desktop._action_buttons:
            self.assertTrue(button.instate(["disabled"]), f"{button.cget('text')} still enabled")
        desktop._set_busy_ui(False)
        self.root.update_idletasks()
        self.assertFalse(desktop.busy_bar.winfo_manager())
        for button in desktop._action_buttons:
            self.assertTrue(button.instate(["!disabled"]), f"{button.cget('text')} still disabled")

    def test_near_limit_progress_shows_and_hides_the_reminder(self):
        desktop = self.desktop
        desktop._set_progress(3, 8500, 10000, 1)
        self.root.update_idletasks()
        self.assertTrue(desktop.limit_label.winfo_manager())
        self.assertIn("85%", desktop.limit_var.get())
        desktop._set_progress(1, 500, 10000, 0)
        self.root.update_idletasks()
        self.assertFalse(desktop.limit_label.winfo_manager())

    def test_dpi_scale_drives_spacing_and_the_restyle(self):
        desktop = self.desktop
        self.assertGreaterEqual(desktop._dpi_scale, 1.0)
        self.assertEqual(desktop._sp(4), round(16 * desktop._dpi_scale))
        base_padding = desktop.style.lookup("TButton", "padding")
        desktop._dpi_scale = 1.5
        desktop._apply_theme()
        self.root.update_idletasks()
        self.assertEqual(desktop._sp(4), 24)
        self.assertNotEqual(desktop.style.lookup("TButton", "padding"), base_padding)
        desktop._dpi_scale = 1.0
        desktop._apply_theme("auto")

    def test_keyboard_shortcuts_and_focus_ring_are_registered(self):
        desktop = self.desktop
        bindings = desktop.root.bind()         # Tk reports normalised names such as <Key-Escape>
        self.assertTrue(any("Escape" in sequence for sequence in bindings), bindings)
        self.assertTrue(any("Control" in sequence and "Return" in sequence
                            for sequence in bindings), bindings)
        desktop.focus_prompt()                 # must not raise on a withdrawn window
        # Reading panes keep a visible focus ring colour from the theme.
        self.assertEqual(str(desktop.result_text.cget("highlightcolor")), desktop._theme.focus)
        self.assertGreaterEqual(int(desktop.result_text.cget("highlightthickness")), 2)
        self.assertEqual(desktop.style.lookup("TButton", "focuscolor"), desktop._theme.focus)

    def test_startup_capability_detection_is_informational_not_a_failure(self):
        """Startup reports a missing capability as information, never as a failed operation.

        Nothing was attempted and no file is at risk when the window opens, so borrowing the
        "失败，原文件保留" language here would be wrong and would blunt it for a real refusal.
        """
        desktop = self.desktop
        note = ("独立 Office COM 实例启动遭权限拒绝；不能据此判断 Office 是否已安装。"
                "普通 docx/xlsx 采用 Python 降级；遗留/高级格式拒绝写入 (PermissionError)")
        report = {"word": {"available": False, "note": note},
                  "excel": {"available": False, "note": f"[WinError 5] 拒绝访问；{note}"}}
        before = desktop.status.get()
        with mock.patch("office_agent.com.capabilities", return_value=report):
            desktop.probe()
            deadline = time.time() + 10
            while desktop.events.empty() and time.time() < deadline:
                time.sleep(0.05)
            self.assertFalse(desktop.events.empty(), "the probe produced no event at all")
            while not desktop.events.empty():
                desktop.drain()
            desktop.drain()                      # the banner travels via the UI mailbox
        self.root.update_idletasks()
        logged = desktop.log.get("1.0", "end-1c")
        # Informational: the status line is untouched, the chip stays informational.
        self.assertEqual(desktop.status.get(), before)
        self.assertNotEqual(desktop.status.get(), "失败，原文件保留")
        self.assertNotEqual(desktop.status_badge.level, "danger")
        self.assertNotEqual(str(desktop.status_label.cget("style")), "StatusDanger.TLabel")
        # A persistent warn banner explains the capability, with no failure wording.
        self.assertTrue(desktop.notice_frame.winfo_manager(), "the capability notice is not shown")
        banner = desktop.notice_var.get()
        self.assertIn("需要本机可用且允许 COM 自动化", banner)
        self.assertIn("仅做启动检测，未运行写入探针", banner)
        self.assertNotIn("失败", banner)
        self.assertNotIn("原文件保留", banner)
        # The neutral line and the untouched raw report are both in the transcript.
        self.assertIn("仅做启动检测，未运行写入探针", logged)
        self.assertIn('"available": false', logged)
        # And it is not silently swallowed: it never claims success either.
        self.assertIn("不能据此判断 Office 是否已安装", banner)

    def test_refused_legacy_edit_is_presented_at_danger_with_the_verbatim_reason(self):
        """A real refusal is the moment for danger + safety + "nothing was written"."""
        desktop = self.desktop
        note = ("独立 Office COM 实例启动遭权限拒绝；不能据此判断 Office 是否已安装。"
                "普通 docx/xlsx 采用 Python 降级；遗留/高级格式拒绝写入 (PermissionError)")
        desktop._handle_event("error",
                              "本机未检测到可用的 Microsoft Word COM 写通道，"
                              ".doc/.xls 编辑不可用：[WinError 5] 拒绝访问；"
                              f"{note}。未做任何改动；"
                              "普通 docx/xlsx 请使用 Python 通道")
        self.root.update_idletasks()
        logged = desktop.log.get("1.0", "end-1c")
        self.assertEqual(desktop.status.get(), "失败，原文件保留")
        self.assertEqual(str(desktop.status_label.cget("style")), "StatusDanger.TLabel")
        self.assertEqual(desktop.status_badge.level, "danger")
        self.assertIn("本次未做任何写入，原文件保持不变", logged)
        self.assertIn("[WinError 5]", logged)
        self.assertIn("不能据此判断 Office 是否已安装", logged)
        self.assertIn("PermissionError", logged)
        ranges = desktop.log.tag_ranges("safety")
        self.assertTrue(ranges, "the refusal was not marked as a safety block")
        marked = "".join(desktop.log.get(ranges[index], ranges[index + 1])
                         for index in range(0, len(ranges), 2))
        self.assertIn("[WinError 5]", marked)
        self.assertIn("本次未做任何写入", marked)
        # The danger hint reaches the next-step line once a file is in play (before that the
        # honest next step really is "choose a file").
        desktop.path = Path("legacy.doc")
        desktop._update_next_step()
        self.assertIn("原文件保持不变", desktop.next_step_var.get())
        desktop.path = None
        # A refusal carried inside a state payload gets exactly the same treatment.
        desktop._handle_event("state", {"status": "done", "iterations": 1, "token_estimate": 10,
                                        "context_capacity": 100, "compressions": 0, "answer": None,
                                        "errors": ["本机未检测到可用的 Microsoft Word COM 写通道，"
                                                   ".doc/.xls 编辑不可用：probe 未通过。未做任何改动；"
                                                   "普通 docx/xlsx 请使用 Python 通道"],
                                        "evidence": [], "resume_checkpoint": None})
        ranges = desktop.log.tag_ranges("safety")
        state_marked = "".join(desktop.log.get(ranges[index], ranges[index + 1])
                               for index in range(0, len(ranges), 2))
        self.assertIn("未检测到可用的 Microsoft Word COM 写通道", state_marked)
        self.assertIn("本次未做任何写入", state_marked)
        # An unrelated error must NOT gain the legacy guarantee.
        self.assertNotIn("本次未做任何写入", legacy_refusal_presentation("GuardError: refused"))

    def test_safety_and_budget_events_get_their_own_markers(self):
        desktop = self.desktop
        desktop._handle_event("progress", {"iterations": 2, "compressions": 1,
                                           "status": "running", "tokens": 100, "capacity": 1000})
        desktop._handle_event("error", "GuardError: explicit human confirmation required")
        self.root.update_idletasks()
        spans = []
        for tag in ("budget", "safety"):
            ranges = desktop.log.tag_ranges(tag)
            self.assertTrue(ranges, f"tag {tag} was never used")
            spans.append("".join(desktop.log.get(start, end) for start in ranges[::2]
                                 for end in [ranges[ranges.index(start) + 1]]))
        joined = " ".join(spans)
        self.assertIn("上下文压缩事件", joined)
        self.assertIn("GuardError", joined)
        self.assertEqual(str(desktop.log.tag_cget("safety", "background")),
                         desktop._theme.danger_soft)
        self.assertEqual(str(desktop.log.tag_cget("budget", "background")),
                         desktop._theme.warn_soft)


class ResultRenderingTests(unittest.TestCase):
    """The model only receives a bounded excerpt, so content the user asked to see must be
    rendered by the UI from the full tool result."""

    def test_word_text_renders_paragraphs_tables_and_the_continuation_hint(self):
        text = render_tool_result("read_word_text", {
            "blocks": [
                {"kind": "paragraph", "index": 0, "text": "湖南涉外经济学院"},
                {"kind": "table", "index": 1, "row_count": 6, "column_count": 9,
                 "rows": [["姓名", "李炜基"], ["学号", "234490133"]],
                 "rows_truncated": True, "rows_omitted": 4,
                 "remainder_hint": "其余行请用 read_word_table_region 继续读取"},
            ],
            "total_blocks": 5, "next_offset": 2, "has_more": True})
        self.assertIn("湖南涉外经济学院", text)
        self.assertIn("[表格 1: 6 行 × 9 列]", text)
        self.assertIn("| 姓名 | 李炜基 |", text)
        self.assertIn("还有 4 行未显示", text)
        self.assertIn("next_offset=2", text)

    def test_other_read_tools_render_their_content(self):
        region = render_tool_result("read_word_region", [
            {"paragraph_index": 0, "text": "第一段"}, {"paragraph_index": 1, "text": "第二段"}])
        self.assertIn("[段 0] 第一段", region)
        self.assertIn("[段 1] 第二段", region)

        table = render_tool_result("read_word_table_region", {
            "table_index": 0, "total_rows": 2, "total_columns": 2,
            "rows": [{"cells": [{"text": "a"}, {"text": "b"}]}]})
        self.assertIn("[表格 0: 2 行 × 2 列]", table)
        self.assertIn("| a | b |", table)

        excel = render_tool_result("read_excel_region", {"sheet": "Data", "cells": [
            {"cell": "A1", "value": 4},
            {"cell": "B1", "formula_text": "=A1*3", "cached_value": None,
             "cached_value_unverified": True}]})
        self.assertIn("A1: 4", excel)
        # A formula is never presented as a calculated value.
        self.assertIn("公式文本", excel)
        self.assertIn("未被重新计算", excel)

        structure = render_tool_result("read_word_structure", {
            "sections": [{"header": [{"text": "机密"}], "footer": []}],
            "comments": [{"id": "1", "author": "R", "text": "请核对"}],
            "images": [{"placement": "inline", "position": {"type": "paragraph", "index": 2}}]})
        self.assertIn("[header 0] 机密", structure)
        self.assertIn("[批注 1 by R] 请核对", structure)
        self.assertIn("[图片] inline", structure)

    def test_tools_without_user_facing_content_render_nothing(self):
        for tool, payload in (("inspect", {"readable": True}), ("read_word_text", "not a dict"),
                              ("plan_excel", {"plan_id": "x"})):
            with self.subTest(tool=tool):
                self.assertIsNone(render_tool_result(tool, payload))

    def test_truncated_content_is_marked_instead_of_dropped(self):
        text = render_tool_result("read_word_text", {"blocks": [
            {"kind": "paragraph", "index": 0, "text": "前半段", "text_truncated": True,
             "remainder_hint": "用 read_word_region 按段索引续读"}]})
        self.assertIn("前半段", text)
        self.assertIn("被截断", text)
        self.assertIn("用 read_word_region 按段索引续读", text)
        region = render_tool_result("read_word_region", [
            {"paragraph_index": 0, "text": "很长的一段", "text_truncated": True}])
        self.assertIn("很长的一段", region)
        self.assertIn("本段更长，已被工具截断", region)

    def test_search_preview_states_what_was_held_back(self):
        hits = [{"location": f"段{i}", "match": "m" * 100} for i in range(40)]
        text = render_tool_result("search", hits)
        self.assertIn("只显示前", text)
        self.assertIn("40 处命中", text)

    def test_an_unhandled_read_tool_still_shows_every_field(self):
        text = render_tool_result("read_excel_structure", {
            "defined_names": [{"name": "TaxRate"}],
            "sheets": [{"name": "Data", "hidden_rows": [3, 4]}]})
        self.assertIn("defined_names", text)
        self.assertIn("TaxRate", text)
        self.assertIn("hidden_rows", text)
        # Tools that carry no user-facing content still render nothing at all.
        self.assertIsNone(render_tool_result("inspect", {"readable": True}))
        self.assertIsNone(render_tool_result("plan_excel", {"plan_id": "x"}))

    def test_merged_table_structure_marks_spans_continuations_and_nesting(self):
        text = render_tool_result("read_word_table_structure", {
            "table_index": 1, "row_count": 3, "column_count": 4, "grid_columns": 5,
            "merge_consistent": True,
            "cells": [{"row": 0, "column": 0, "last_column": 2, "gridSpan": 3,
                       "vMerge": None, "anchor": True, "text": "标题"},
                      {"row": 1, "column": 0, "last_column": 0, "gridSpan": 1,
                       "vMerge": "continue", "anchor": False, "text": "标题"}],
            "continuations": [{"row": 1, "column": 0, "origin_row": 0, "origin_column": 0}],
            "nested_tables": [{"path": "0/0", "depth": 1}]})
        self.assertIn("grid_columns=5", text)
        self.assertIn("标题", text)
        self.assertIn("跨 3 列（至第 2 列）", text)
        self.assertIn("vMerge 延续", text)
        self.assertIn("非承载文本", text)
        self.assertIn("[延续]", text)
        self.assertIn("[嵌套表格] 路径 0/0 深度 1", text)
        inconsistent = render_tool_result("read_word_table_structure", {
            "merge_consistent": False, "merge_issues": ["row 0 gridSpan overflow"], "cells": []})
        self.assertIn("不自洽", inconsistent)
        self.assertIn("row 0 gridSpan overflow", inconsistent)


class WriteConfirmationGateTests(unittest.TestCase):
    """The UI gate that turns a preview into a write. Declining must change nothing."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)
        self.root = self.base / "authorized"          # the granted document folder
        self.root.mkdir()
        self.appstate = self.base / "appstate"        # stands in for the app-owned fallback
        self.appstate.mkdir()
        self.path = self.root / "sample.xlsx"
        wb = Workbook(); ws = wb.active; ws.title = "Data"
        ws["A1"] = 4; ws["B1"] = "=A1*3"
        wb.save(self.path)
        self.editor = Editor(self.root)
        self.desktop = Desktop.__new__(Desktop)
        self.desktop.editor = self.editor
        self.desktop.busy = False
        self.desktop.pending_id = None
        self.desktop.last_verified_id = None
        self.desktop.events = queue.Queue()
        self.desktop.status = Mock()
        self.desktop.show = Mock()

    def _plan(self):
        dispatch = build_dispatcher(self.editor, self.path)
        result = dispatch("plan_excel_formula_fill", {
            "path": str(self.path), "sheet": "Data", "source_cell": "B1",
            "target_range": "B2:B2", "expected_source_formula": "=A1*3"})
        self.desktop.pending_id = result["plan_id"]
        return result["plan_id"]

    def _settle(self):
        kind, payload = self.desktop.events.get(timeout=10)
        self.assertEqual(kind, "verified", payload)
        return payload

    def test_declining_confirmation_changes_nothing(self):
        self._plan()
        with mock.patch("office_agent.app.messagebox.askyesno", return_value=False):
            self.desktop.execute()
        self.assertIsNone(_cell_value(self.path, "Data", "B2"))
        self.assertEqual(list(self.root.glob("sample.agent-*.xlsx")), [])
        self.assertFalse(self.editor.audit.exists())
        self.desktop.show.assert_called_with("已取消：未做任何写入，原文件保持不变。")

    def test_confirming_writes_a_separate_output_and_leaves_the_source(self):
        plan_id = self._plan()
        before = self.path.read_bytes()
        with mock.patch("office_agent.app.messagebox.askyesno", return_value=True):
            self.desktop.execute()
        verified = self._settle()
        self.assertEqual(verified["id"], plan_id)
        self.assertEqual(_cell_value(verified["output"], "Data", "B2"), "=A2*3")
        self.assertIsNone(_cell_value(self.path, "Data", "B2"))
        self.assertEqual(self.path.read_bytes(), before)

    def test_execute_without_a_plan_explains_instead_of_doing_nothing(self):
        self.desktop.execute()
        self.desktop.status.set.assert_called_with(
            "当前没有待确认的修改计划；请先提交一次规划并查看差异预览。")

    def test_undo_without_a_verified_output_is_reported(self):
        self.desktop.undo()
        self.desktop.status.set.assert_called_with("没有可撤销的已验证输出副本。")

    def test_writability_probe_leaves_the_folder_exactly_as_it_was(self):
        """Selecting a file or declining a confirmation must not create .office_agent."""
        target = self.root / ".office_agent" / "sessions"
        self.assertFalse(target.parent.exists())
        self.assertTrue(_writable_directory(target))
        self.assertFalse(target.exists())
        self.assertFalse(target.parent.exists())
        # An existing directory is preserved, and a probe inside it still cleans up.
        target.mkdir(parents=True)
        self.assertTrue(_writable_directory(target))
        self.assertTrue(target.exists())
        self.assertEqual(list(target.iterdir()), [])

    def test_writability_probe_reports_failure_instead_of_raising(self):
        blocker = self.root / "not-a-directory.txt"
        blocker.write_text("x", encoding="utf-8")
        self.assertFalse(_writable_directory(blocker / "sub"))

    def test_write_confirmations_default_to_cancel(self):
        """A reflex Enter on the write dialogs must not write or delete anything."""
        self._plan()
        with mock.patch("office_agent.app.messagebox.askyesno", return_value=False) as askyesno:
            self.desktop.execute()
        self.assertEqual(askyesno.call_args.kwargs.get("default"), "no")
        self.assertIsNone(_cell_value(self.path, "Data", "B2"))
        self.desktop.last_verified_id = "deadbeef"
        with mock.patch("office_agent.app.messagebox.askyesno", return_value=False) as askyesno2:
            self.desktop.undo()
        self.assertEqual(askyesno2.call_args.kwargs.get("default"), "no")

    def test_unwritable_document_folder_runs_without_persisting_anything(self):
        """The reported failure: a Desktop document could not be read because the app
        insisted on creating .office_agent beside it. The authorized folder is the only
        acceptable place for session data, so the task must instead write nothing."""
        def document_folder_is_unwritable(path):
            return not (path == self.editor.audit or self.editor.audit in path.parents)

        with mock.patch("office_agent.app._writable_directory",
                        side_effect=document_folder_is_unwritable):
            sessions, warning = self.desktop._prepare_session_base(self.editor)
        self.assertIsNone(sessions, "no unauthorized location may be substituted")
        self.assertIsNotNone(warning, "the user must be told the run will not persist")
        self.assertIn("不落盘", warning)
        self.assertIn("不支持重启续跑", warning)
        self.assertIn(str(self.root), warning)
        self.assertFalse(self.editor.audit.exists(), "the document folder must stay untouched")

    def test_unwritable_folder_blocks_writes_with_an_explanation_not_a_traceback(self):
        self._plan()
        with mock.patch("office_agent.app._writable_directory", return_value=False), \
                mock.patch("office_agent.app.messagebox.showerror") as showerror:
            self.desktop.execute()
        showerror.assert_called_once()
        self.assertIn("授权目录不可写", showerror.call_args[0][0])
        self.assertIn("原文件保持不变", showerror.call_args[0][1])
        self.assertIsNone(_cell_value(self.path, "Data", "B2"))
        self.assertFalse(self.editor.audit.exists())

    def test_a_task_can_read_without_writing_anywhere(self):
        source = self.root / "doc.docx"
        document = Document()
        document.add_paragraph("Readable body text")
        document.save(source)

        class Reader(StrictModel):
            path: str

        agent = OfficeAgent(
            MockModel([{"tool_calls": [{"name": "read_word_region",
                                        "args": {"path": str(source)}, "id": "1"}]},
                       {"content": "已读出正文"}]),
            lambda name, args: {"paragraphs": ["Readable body text"]},
            [ToolSpec(name="read_word_region", description="Read paragraphs",
                      schema_model=Reader, access="read", allowed_paths=(source,))],
            self.base / "never-created", confirm=lambda plan: True, persist=False,
            token_estimator=lambda text: 10)
        before = sorted(str(p) for p in self.base.rglob("*"))
        result = agent.run(agent.new_state("展示文本内容"))
        self.assertEqual(result.status, "done")
        self.assertEqual(result.answer, "已读出正文")
        self.assertIsNone(agent.checkpoint)
        self.assertTrue(result.required.evidence[0].result_path.startswith("memory:"))
        self.assertIn(b"Readable body text", next(iter(agent.results.values())))
        # Nothing at all was written, and resume explains itself instead of crashing.
        self.assertEqual(sorted(str(p) for p in self.base.rglob("*")), before)
        with self.assertRaisesRegex(ValueError, "keeps no checkpoint"):
            agent.load()
        with self.assertRaisesRegex(ValueError, "keeps no checkpoint"):
            agent.resume()

    def test_a_persisting_run_still_writes_results_and_a_checkpoint(self):
        source = self.root / "doc.docx"
        Document().save(source)

        class Reader(StrictModel):
            path: str

        agent = OfficeAgent(
            MockModel([{"tool_calls": [{"name": "read_word_region",
                                        "args": {"path": str(source)}, "id": "1"}]},
                       {"content": "done"}]),
            lambda name, args: {"paragraphs": []},
            [ToolSpec(name="read_word_region", description="Read paragraphs",
                      schema_model=Reader, access="read", allowed_paths=(source,))],
            self.base / "state", confirm=lambda plan: True, token_estimator=lambda text: 10)
        result = agent.run(agent.new_state("goal"))
        self.assertEqual(result.status, "done")
        self.assertTrue(agent.checkpoint.is_file())
        self.assertEqual(agent.load().status, "done")
        self.assertTrue(Path(result.required.evidence[0].result_path).is_file())


if __name__ == "__main__":
    unittest.main()
