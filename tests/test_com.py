import os
import sys
import tempfile
import threading
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from office_agent.com import (WindowsComBackend, capabilities, describe_write_refusal, extract_office,
                              isolated_office, probe_write_channel, run_in_office_thread,
                              write_capabilities)


def _fake_com_modules(events, *, dispatch_factory=None):
    """Minimal stand-ins for pythoncom/win32com.client that record every call."""
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: events.append(("init", threading.current_thread().name))
    pythoncom.CoUninitialize = lambda: events.append(("uninit", threading.current_thread().name))
    win32com = types.ModuleType("win32com")
    client = types.ModuleType("win32com.client")
    if dispatch_factory is None:
        def dispatch_factory(name):
            class Instance:
                def Quit(self):
                    events.append("quit")
            events.append(name)
            return Instance()
    client.DispatchEx = dispatch_factory
    win32com.client = client
    return {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client}


class COMWriteProbeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def test_probe_only_touches_the_blank_document_it_creates_itself(self):
        """Safety red line: the probe never opens a user file or an untrusted document."""
        events = []
        target_holder = {}

        class Range:
            Text = "office write probe"

        class Document:
            def __init__(self):
                self.Range = lambda: Range()

            def SaveAs2(self, path, **kwargs):
                self_format = kwargs.get("FileFormat")
                if self_format != 0:
                    raise AssertionError("Word probe must save genuine .doc format")
                events.append(("saveas", path))
                target_holder["path"] = path
                Path(path).write_bytes(b"probe-copy")

            def Close(self, **kwargs):
                events.append(("close", kwargs))

        class Documents:
            def Add(self):
                events.append("add")
                return Document()

            def Open(self, **kwargs):
                events.append(("open", kwargs.get("FileName")))
                return Document()

        class Instance:
            def Quit(self):
                events.append("quit")

        Instance.Documents = Documents()
        modules = _fake_com_modules(events)
        modules["win32com.client"].DispatchEx = lambda name: Instance()
        with patch.dict(sys.modules, modules):
            report = probe_write_channel("word", workdir=str(self.root))

        self.assertTrue(report["available"], report)
        self.assertEqual(report["probe"], "blank_document_save_reopen")
        self.assertEqual(report["roundtrip_text"], "office write probe")
        # The only document opened or saved is the throwaway copy the probe made itself.
        opened = [entry[1] for entry in events if isinstance(entry, tuple) and entry[0] == "open"]
        saved = [entry[1] for entry in events if isinstance(entry, tuple) and entry[0] == "saveas"]
        self.assertEqual(opened, [target_holder["path"]])
        self.assertEqual(saved, [target_holder["path"]])
        self.assertTrue(str(target_holder["path"]).startswith(str(self.root)))
        # Created, verified, then removed: nothing is left behind for the user to trip over.
        self.assertFalse(Path(target_holder["path"]).exists())
        self.assertEqual([entry for entry in events if entry == "quit"], ["quit"])

    def test_probe_reports_permission_denial_without_claiming_office_is_missing(self):
        with patch("office_agent.com.isolated_office",
                   side_effect=PermissionError("[WinError 5] gen_py denied")):
            report = probe_write_channel("word", workdir=str(self.root))
        self.assertFalse(report["available"])
        self.assertEqual(report["stage"], "startup")
        self.assertIn("权限或策略拒绝", report["reason"])
        self.assertIn("不等于没有安装 Office", report["reason"])
        self.assertIn("gen_py denied", report["reason"])
        self.assertEqual(list(self.root.iterdir()), [])          # nothing written, nothing left

    def test_probe_reports_a_copy_that_could_not_be_saved(self):
        events = []

        class Cell:
            Value = None
            Text = ""

        class Worksheet:
            def Cells(self, row, column):
                return Cell()

        class Workbook:
            def Worksheets(self, index):
                return Worksheet()

            def SaveAs(self, path, **kwargs):
                if kwargs.get("FileFormat") != 56:
                    raise AssertionError("Excel probe must save genuine .xls format")
                pass                                              # silently saves nothing

            def Close(self, **kwargs):
                pass

        class Instance:
            def Quit(self):
                events.append("quit")

        Instance.Workbooks = types.SimpleNamespace(Add=lambda: Workbook(),
                                                   Open=lambda **kwargs: Workbook())
        modules = _fake_com_modules(events)
        modules["win32com.client"].DispatchEx = lambda name: Instance()
        with patch.dict(sys.modules, modules):
            report = probe_write_channel("excel", workdir=str(self.root))
        self.assertFalse(report["available"])
        self.assertEqual(report["stage"], "verify")
        self.assertIn("副本不存在或为空", report["reason"])
        self.assertEqual([entry for entry in events if entry == "quit"], ["quit"])

    def test_write_capabilities_separates_not_installed_from_automation_blocked(self):
        denied_startup = {"word": {"available": False, "note": "独立实例启动遭权限拒绝"},
                          "excel": {"available": False, "note": "独立实例启动遭权限拒绝"}}
        with patch("office_agent.com.capabilities", return_value=denied_startup), \
                patch("office_agent.com.probe_write_channel") as probe:
            result = write_capabilities(workdir=str(self.root))
        probe.assert_not_called()                                  # never probe what cannot start
        for entry in result.values():
            self.assertFalse(entry["startup_ok"])
            self.assertFalse(entry["write_channel"])
            self.assertIn("启动遭权限拒绝", entry["reason"])
        message = describe_write_refusal("word", result["word"])
        self.assertIn("未检测到可用的 Microsoft Word COM 写通道", message)
        self.assertIn("未做任何改动", message)

        started = {"word": {"available": True, "note": "独立实例启动检测成功"},
                   "excel": {"available": True, "note": "独立实例启动检测成功"}}
        blocked = {"available": False, "stage": "startup", "reason": "COM 自动化被权限或策略拒绝"}
        with patch("office_agent.com.capabilities", return_value=started), \
                patch("office_agent.com.probe_write_channel", return_value=blocked):
            result = write_capabilities(workdir=str(self.root))
        self.assertTrue(result["word"]["startup_ok"])
        self.assertFalse(result["word"]["write_channel"])          # installed is not writable
        self.assertIn("自动化被权限或策略拒绝", result["word"]["reason"])

    def test_office_work_runs_on_a_dedicated_thread_with_paired_lifecycle(self):
        events = []
        caller = threading.current_thread().name
        modules = _fake_com_modules(events)
        with patch.dict(sys.modules, modules):
            value = run_in_office_thread(lambda: (events.append(("work", threading.current_thread().name)), 7)[1])
        self.assertEqual(value, 7)
        threads = {name for _, name in events}
        self.assertNotIn(caller, threads)                          # never the caller's apartment
        self.assertEqual(len(events), 3)
        self.assertEqual(events[0][0], "init")
        self.assertEqual(events[-1][0], "uninit")
        self.assertLess(events.index(events[0]), events.index(events[-1]))

    def test_com_lifecycle_is_released_even_when_the_work_raises(self):
        events = []
        modules = _fake_com_modules(events)
        with patch.dict(sys.modules, modules):
            with self.assertRaisesRegex(RuntimeError, "boom"):
                run_in_office_thread(lambda: (_ for _ in ()).throw(RuntimeError("boom")))
        self.assertEqual([entry[0] for entry in events], ["init", "uninit"])

    def test_only_instances_started_by_dispatch_ex_are_quit(self):
        """A user's open Office window must never be attached to, let alone closed."""
        events = []

        class UserInstance:
            def __init__(self):
                self.quit_calls = 0

            def Quit(self):
                self.quit_calls += 1

        user_instance = UserInstance()
        created = []

        def dispatch(name):
            instance = types.SimpleNamespace(
                Quit=lambda: events.append("quit-created"),
                Visible=True, DisplayAlerts=1, AutomationSecurity=0)
            created.append(instance)
            events.append(name)
            return instance

        modules = _fake_com_modules(events, dispatch_factory=dispatch)
        with patch.dict(sys.modules, modules):
            with isolated_office("Word.Application"):
                pass
            with isolated_office("Excel.Application"):
                pass
        self.assertEqual(user_instance.quit_calls, 0)              # never touched
        self.assertEqual(events.count("quit-created"), 2)          # exactly our two instances
        self.assertTrue(all(instance.Visible is False for instance in created))
        self.assertTrue(all(instance.AutomationSecurity == 3 for instance in created))

    def test_word_write_opens_writable_saves_and_closes_without_saving(self):
        """The write helper must open writable, save once, and close without re-saving."""
        events = []

        class FakeRange:
            def __init__(self, document, index):
                self.document, self.index = document, index

            @property
            def Text(self):
                return self.document.texts[self.index]

            @Text.setter
            def Text(self, value):
                self.document.texts[self.index] = value

        class FakeParagraphs:
            def __init__(self, document):
                self.document = document

            def __iter__(self):
                return (types.SimpleNamespace(Range=FakeRange(self.document, index))
                        for index in range(len(self.document.texts)))

            def __call__(self, index):
                return types.SimpleNamespace(Range=FakeRange(self.document, index - 1))

        class Document:
            def __init__(self):
                self.texts = ["first\r", "price 5\r", "last\r"]
                self.Paragraphs = FakeParagraphs(self)

            def Save(self):
                events.append("save")

            def Close(self, **kwargs):
                events.append(("close", kwargs))

        document = Document()
        application = types.SimpleNamespace(
            Documents=types.SimpleNamespace(
                Open=lambda **kwargs: (events.append(("open", kwargs)), document)[1]))
        from contextlib import contextmanager

        @contextmanager
        def fake_instance(prog_id):
            yield application

        backend = WindowsComBackend()
        with patch("office_agent.com.isolated_office", fake_instance), \
                patch("office_agent.com.run_in_office_thread", lambda work: work()), \
                patch("office_agent.com.Path.is_file", return_value=True), \
                patch("office_agent.com.Path.is_symlink", return_value=False):
            written = backend.replace_word_paragraph("C:\\selected.doc", index=1, old="price",
                                                     new="amount", expected_count=1)
            read_back = backend.read_word_paragraphs("C:\\selected.doc", limit=5)
        self.assertEqual(written, "amount 5")
        self.assertEqual(read_back, ["first", "amount 5", "last"])
        self.assertEqual(document.texts[1], "amount 5\r")           # the paragraph mark survived
        self.assertEqual([entry for entry in events if entry == "save"], ["save"])
        self.assertFalse(events[0][1]["ReadOnly"])                  # an edit needs a writable open
        self.assertEqual(events[-1], ("close", {"SaveChanges": False}))
        self.assertEqual(events[3], ("open", {"FileName": "C:\\selected.doc",
                                              "ConfirmConversions": False, "ReadOnly": True,
                                              "AddToRecentFiles": False, "UpdateLinks": 0,
                                              "NoEncodingDialog": True}))

    def test_word_write_refuses_a_count_that_does_not_match(self):
        events = []

        class Document:
            def __init__(self):
                self.texts = ["price price\r"]

            def Paragraphs(self, index):
                return types.SimpleNamespace(Range=types.SimpleNamespace(
                    Text=self.texts[index - 1]))

            def Save(self):
                events.append("save")

            def Close(self, **kwargs):
                events.append(("close", kwargs))

        application = types.SimpleNamespace(
            Documents=types.SimpleNamespace(Open=lambda **kwargs: Document()))
        from contextlib import contextmanager

        @contextmanager
        def fake_instance(prog_id):
            yield application

        backend = WindowsComBackend()
        with patch("office_agent.com.isolated_office", fake_instance), \
                patch("office_agent.com.run_in_office_thread", lambda work: work()), \
                patch("office_agent.com.Path.is_file", return_value=True), \
                patch("office_agent.com.Path.is_symlink", return_value=False):
            with self.assertRaisesRegex(ValueError, "expected 1 occurrence"):
                backend.replace_word_paragraph("C:\\selected.doc", index=0, old="price",
                                               new="amount", expected_count=1)
        self.assertNotIn("save", events)                            # nothing was written
        self.assertEqual(events[-1], ("close", {"SaveChanges": False}))

    def test_backend_refuses_relative_paths_and_links(self):
        backend = WindowsComBackend()
        for call in (lambda: backend.read_word_paragraphs("relative.doc", limit=5),
                     lambda: backend.read_excel_cell("relative.xls", sheet="S", cell="A1")):
            with self.assertRaisesRegex(ValueError, "absolute"):
                call()
        with patch("office_agent.com.Path.is_symlink", return_value=True):
            with self.assertRaisesRegex(ValueError, "not a link"):
                backend.read_excel_cell("C:\\selected.xls", sheet="S", cell="A1")

    @unittest.skipUnless(os.environ.get("OFFICE_AGENT_REAL_COM") == "1",
                         "real Office disabled; explicitly set OFFICE_AGENT_REAL_COM=1")
    def test_real_office_roundtrip_is_explicitly_opt_in(self):
        """Two valid self-created legacy fixtures, each in a bounded worker.

        Probe failure is a failure when explicitly enabled, not a success/skip.
        Source hashes, outputs, unchanged neighbours, audit and rollback are checked.
        """
        from examples.office_com_diagnostics import run_real_case
        report = run_real_case(timeout=20)
        self.assertTrue(report["ok"])
        self.assertEqual([entry["application"] for entry in report["applications"]], ["word", "excel"])


class ControlledCOMWorkerTests(unittest.TestCase):
    def test_word_identity_uses_its_own_blank_window_and_closes_on_refusal(self):
        from examples.office_com_diagnostics import _identity_hwnd
        from unittest.mock import Mock
        from types import SimpleNamespace
        for refuse in (False, True):
            with self.subTest(refuse=refuse):
                document = SimpleNamespace(ActiveWindow=SimpleNamespace(Hwnd=456), Close=Mock())
                app = SimpleNamespace(AutomationSecurity=2,
                                      Documents=SimpleNamespace(Add=Mock(return_value=document)))
                def identify():
                    with _identity_hwnd(app, "Word.Application") as hwnd:
                        self.assertEqual(hwnd, 456)
                        self.assertEqual(app.AutomationSecurity, 3)
                        if refuse:
                            raise RuntimeError("parent refused")
                if refuse:
                    with self.assertRaisesRegex(RuntimeError, "parent refused"):
                        identify()
                else:
                    identify()
                self.assertEqual(app.AutomationSecurity, 2)
                document.Close.assert_called_once_with(SaveChanges=False)
                app.Documents.Add.assert_called_once_with(Visible=True)

    def test_loopback_json_transport_is_bounded_and_roundtrips_without_office(self):
        import socket
        import time
        from examples.office_com_diagnostics import _LoopbackConnection, ControlledCOMError
        left, right = socket.socketpair()
        with left, right:
            sender = _LoopbackConnection(left, time.monotonic() + 2)
            receiver = _LoopbackConnection(right, time.monotonic() + 2)
            sender.send(["stage", {"text": "自建夹具"}])
            self.assertTrue(receiver.poll(0.1))
            self.assertEqual(receiver.recv(), ["stage", {"text": "自建夹具"}])
            with self.assertRaisesRegex(ControlledCOMError, "too large"):
                sender.send("x" * (1024 * 1024 + 1))

    def test_each_quit_is_retired_before_starting_another_office_instance(self):
        from examples.office_com_diagnostics import _OwnedProcesses
        from unittest.mock import Mock
        for forced in (False, True):
            with self.subTest(forced=forced):
                owner = _OwnedProcesses.__new__(_OwnedProcesses)
                handle = Mock()
                owner.handles, owner.identities = [handle], {(123, "identity"): handle}
                owner.lifecycle, owner.api, owner.event = [], Mock(), Mock()
                owner.event.WaitForSingleObject.side_effect = [258, 0] if forced else [0]
                owner.retire(123, "identity")
                self.assertEqual(owner.handles, [])
                self.assertEqual(owner.identities, {})
                self.assertEqual(owner.api.TerminateProcess.call_count, int(forced))
                self.assertEqual(owner.lifecycle[-1]["event"],
                                 "terminated_after_quit" if forced else "quit_exited")
                handle.Close.assert_called_once()

    def test_timeout_stops_worker_and_cleans_only_registered_handles(self):
        from examples.office_com_diagnostics import ControlledCOMError, _supervise
        from unittest.mock import Mock
        worker = Mock(exitcode=0)
        worker.is_alive.side_effect = [True, True, False, False]
        connection, ownership = Mock(), Mock()
        ownership.cleanup.return_value = []
        with self.assertRaisesRegex(ControlledCOMError, "hard timeout.*NOT killed"):
            _supervise(worker, connection, ownership, 1, clock=iter([0, 2]).__next__)
        worker.terminate.assert_called_once()
        worker.join.assert_called_once_with(2)
        ownership.accept.assert_not_called()
        ownership.cleanup.assert_called_once()
        connection.close.assert_called_once()

    def test_confirmed_pid_is_acknowledged_before_result_and_cleaned(self):
        from examples.office_com_diagnostics import _supervise
        from unittest.mock import Mock
        worker = Mock(exitcode=0)
        worker.is_alive.side_effect = [True, False, False, False, False, False]
        connection, ownership = Mock(), Mock()
        connection.poll.return_value = True
        connection.recv.side_effect = [("owned", 123, "identity"), ("result", True, "")]
        ownership.cleanup.return_value = []
        _supervise(worker, connection, ownership, 1, clock=lambda: 0)
        ownership.accept.assert_called_once_with(123, "identity")
        connection.send.assert_called_once_with("accepted")
        worker.terminate.assert_not_called()
        ownership.cleanup.assert_called_once()

    def test_ownership_refusal_fails_closed_and_stops_worker(self):
        from examples.office_com_diagnostics import ControlledCOMError, _supervise
        from unittest.mock import Mock
        worker = Mock(exitcode=0)
        worker.is_alive.side_effect = [True, False, False]
        connection, ownership = Mock(), Mock()
        connection.poll.return_value = True
        connection.recv.return_value = ("owned", 16464, "preexisting")
        ownership.accept.side_effect = ControlledCOMError("pre-existing process")
        ownership.cleanup.return_value = []
        with self.assertRaisesRegex(ControlledCOMError, "failed closed.*pre-existing"):
            _supervise(worker, connection, ownership, 1, clock=lambda: 0)
        connection.send.assert_not_called()
        worker.terminate.assert_called_once()
        ownership.cleanup.assert_called_once()

    def test_preexisting_and_reused_pids_are_never_terminated(self):
        from examples.office_com_diagnostics import ControlledCOMError, _OwnedProcesses
        from unittest.mock import Mock
        ownership = _OwnedProcesses.__new__(_OwnedProcesses)
        ownership.preexisting, ownership.handles = {16464}, []
        ownership.api, ownership.process = Mock(), Mock()
        with self.assertRaisesRegex(ControlledCOMError, "pre-existing"):
            ownership.accept(16464, "old")
        ownership.api.OpenProcess.assert_not_called()
        ownership.process.GetProcessTimes.return_value = {"CreationTime": "reused"}
        with self.assertRaisesRegex(ControlledCOMError, "identity changed"):
            ownership.accept(123, "original")
        ownership.api.OpenProcess.return_value.Close.assert_called_once()
        ownership.api.TerminateProcess.assert_not_called()
        self.assertEqual(ownership.handles, [])

    def test_worker_failure_and_swallowed_quit_error_remain_failures(self):
        from examples.office_com_diagnostics import ControlledCOMError, _supervise
        from unittest.mock import Mock
        worker = Mock(exitcode=0)
        worker.is_alive.side_effect = [True, False, False, False, False, False]
        connection, ownership = Mock(), Mock()
        connection.poll.return_value = True
        connection.recv.side_effect = [("lifecycle_error", "Office Quit failed"),
                                       ("result", False, "original assertion failed")]
        ownership.cleanup.return_value = []
        with self.assertRaises(ControlledCOMError) as failure:
            _supervise(worker, connection, ownership, 1, clock=lambda: 0)
        self.assertIn("Office Quit failed", str(failure.exception))
        self.assertIn("original assertion failed", str(failure.exception))
        ownership.cleanup.assert_called_once()

    def test_readonly_diagnosis_does_not_probe_temp_or_activate_office(self):
        from examples.office_com_diagnostics import readonly_report
        with patch("tempfile.gettempdir", side_effect=AssertionError("write probe forbidden")), \
                patch("pathlib.Path.write_bytes", side_effect=AssertionError("write forbidden")):
            report = readonly_report()
        self.assertIn("not performed", report["com_validation"])
        self.assertTrue(all(entry["writable"].startswith("unknown")
                            for entry in report["cache_attributes"]))

    def test_cleanup_uses_retained_handle_not_pid_or_process_name(self):
        from examples.office_com_diagnostics import _OwnedProcesses
        from unittest.mock import Mock
        ownership = _OwnedProcesses.__new__(_OwnedProcesses)
        handle = Mock()
        ownership.handles, ownership.api, ownership.event = [handle], Mock(), Mock()
        ownership.event.WaitForSingleObject.side_effect = [258, 0]
        self.assertEqual(ownership.cleanup(), [])
        ownership.api.TerminateProcess.assert_called_once_with(handle, 1)
        handle.Close.assert_called_once()
        self.assertEqual(ownership.handles, [])


class COMLifecycleTests(unittest.TestCase):
    def test_permission_denied_detection_does_not_claim_office_missing(self):
        with patch("office_agent.com.isolated_office", side_effect=PermissionError("blocked by policy")):
            result = capabilities()
        self.assertEqual(set(result), {"word", "excel"})
        for entry in result.values():
            self.assertFalse(entry["available"])
            self.assertIn("权限拒绝", entry["note"])
            self.assertIn("不能据此判断 Office 是否已安装", entry["note"])
            self.assertIn("Python 降级", entry["note"])
            self.assertNotIn("blocked by policy", entry["note"])

    def test_only_created_instance_is_quit_even_on_error(self):
        events = []
        pythoncom = types.ModuleType("pythoncom")
        pythoncom.CoInitialize = lambda: events.append("init")
        pythoncom.CoUninitialize = lambda: events.append("uninit")
        win32com = types.ModuleType("win32com")
        client = types.ModuleType("win32com.client")
        class Instance:
            def Quit(self): events.append("quit")
        client.DispatchEx = lambda name: (events.append(name), Instance())[1]
        win32com.client = client
        with patch.dict(sys.modules, {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client}):
            with self.assertRaisesRegex(RuntimeError, "boom"):
                with isolated_office("Word.Application"):
                    raise RuntimeError("boom")
        self.assertEqual(events, ["init", "Word.Application", "quit", "uninit"])

    def test_word_extraction_is_read_only_and_closes_document_before_quit(self):
        events = []
        pythoncom = types.ModuleType("pythoncom")
        pythoncom.CoInitialize = lambda: events.append("init")
        pythoncom.CoUninitialize = lambda: events.append("uninit")
        win32com = types.ModuleType("win32com")
        client = types.ModuleType("win32com.client")

        class Paragraph:
            def __init__(self):
                self.Range = types.SimpleNamespace(Text="hello\r\a")

        class Document:
            Paragraphs = [Paragraph()]
            Tables = []
            def Close(self, **kwargs):
                events.append(("close", kwargs))

        class Documents:
            def Open(self, **kwargs):
                events.append(("open", kwargs))
                return Document()

        class Instance:
            def Quit(self):
                events.append("quit")

        Instance.Documents = Documents()
        client.DispatchEx = lambda name: (events.append(name), Instance())[1]
        win32com.client = client
        with patch.dict(sys.modules, {"pythoncom": pythoncom, "win32com": win32com,
                                      "win32com.client": client}), patch("office_agent.com.Path.is_file", return_value=True):
            result = extract_office("C:\\selected.doc")
        self.assertEqual(result["paragraphs"], [{"index": 1, "text": "hello"}])
        self.assertEqual(events[3], ("close", {"SaveChanges": False}))
        self.assertEqual(events[-2:], ["quit", "uninit"])
        self.assertTrue(events[2][1]["ReadOnly"])
        self.assertEqual(events[2][1]["UpdateLinks"], 0)

    def test_invalid_path_and_type_rejected_without_com(self):
        with self.assertRaises(ValueError):
            extract_office("relative.doc")
        with self.assertRaises(ValueError):
            extract_office("C:\\selected.pdf")

    def test_word_tables_have_global_cell_budget(self):
        import office_agent.com as adapter
        cell = types.SimpleNamespace(Range=types.SimpleNamespace(Text="x\\r\\a"))
        row = types.SimpleNamespace(Cells=[cell, cell])
        table = types.SimpleNamespace(Rows=[row, row])
        document = types.SimpleNamespace(Paragraphs=[], Tables=[table, table], Close=lambda **kw: None)
        application = types.SimpleNamespace(Documents=types.SimpleNamespace(Open=lambda **kw: document))
        from contextlib import contextmanager
        @contextmanager
        def fake_instance(prog_id):
            yield application
        with patch.object(adapter, "isolated_office", fake_instance), patch.object(adapter, "_MAX_TOTAL_CELLS", 1), patch("office_agent.com.Path.is_file", return_value=True):
            result = extract_office("C:\\selected.doc", include_tables=True)
        self.assertEqual(sum(len(row["cells"]) for table in result["tables"] for row in table["rows"]), 1)

    def test_com_region_limits_rejected_before_office_start(self):
        with patch("office_agent.com.Path.is_file", return_value=True):
            for region in ("A1:U1", "A1:A21", "A1:XFD1048576", "A1;B2"):
                with self.assertRaises(ValueError):
                    extract_office("C:\\selected.xls", region=region)
            with self.assertRaises(ValueError):
                extract_office("C:\\selected.doc", region="A1")

    def test_invalid_prog_id_rejected_before_com_import(self):
        with self.assertRaises(ValueError):
            with isolated_office("unknown.application"):
                pass


class NativeOfficeOwnershipTests(unittest.TestCase):
    def test_assignment_refusal_stops_suspended_child_without_resuming(self):
        from unittest.mock import Mock
        from examples.office_owned_launcher import NativeJob
        handle, thread = Mock(), Mock()
        api = types.SimpleNamespace(GetStdHandle=lambda value: None, TerminateProcess=Mock())
        process = types.SimpleNamespace(STARTUPINFO=lambda: types.SimpleNamespace(dwFlags=0),
                                        CreateProcess=Mock(return_value=(handle, thread, 4242, 1)),
                                        ResumeThread=Mock())
        constants = types.SimpleNamespace(STD_INPUT_HANDLE=-10, STD_OUTPUT_HANDLE=-11,
                                         STD_ERROR_HANDLE=-12, STARTF_USESTDHANDLES=256,
                                         CREATE_SUSPENDED=4, CREATE_NO_WINDOW=0x08000000)
        job = NativeJob.__new__(NativeJob)
        job.handle = Mock()
        job.api = types.SimpleNamespace(AssignProcessToJobObject=Mock(side_effect=RuntimeError('refused')))
        modules = {'win32api': api, 'win32process': process, 'win32con': constants,
                   'win32event': types.SimpleNamespace(WaitForSingleObject=lambda h, timeout: 0)}
        with patch.dict(sys.modules, modules), self.assertRaisesRegex(RuntimeError, 'refused'):
            job.spawn_worker(['owned-python.exe', '-c', 'pass'])
        self.assertTrue(process.CreateProcess.call_args.args[5] & constants.CREATE_SUSPENDED)
        process.ResumeThread.assert_not_called()
        api.TerminateProcess.assert_called_once_with(handle, 1)
        handle.Close.assert_called_once()
        thread.Close.assert_called_once()

    def test_job_cleanup_covers_unacknowledged_members_and_checks_zero(self):
        from unittest.mock import Mock
        from examples.office_owned_launcher import NativeJob
        for active in (0, 2):
            with self.subTest(active=active):
                job = NativeJob.__new__(NativeJob)
                job.handle, job.closed, job.empty_confirmed = Mock(), False, False
                job.api = types.SimpleNamespace(
                    JobObjectBasicAccountingInformation=1, TerminateJobObject=Mock(),
                    QueryInformationJobObject=Mock(side_effect=[{'ActiveProcesses': active},
                                                                {'ActiveProcesses': 0}]))
                events = job.close()
                self.assertTrue(job.empty_confirmed)
                self.assertEqual(events, [{'event': 'native_job_empty', 'forced': bool(active),
                                           'active_processes': 0}])
                if active:
                    job.api.TerminateJobObject.assert_called_once_with(job.handle, 1)
                else:
                    job.api.TerminateJobObject.assert_not_called()
                self.assertEqual(job.close(), [])
                job.handle.Close.assert_called_once()

    def test_native_window_selection_excludes_foreign_parent_and_child_pids(self):
        from unittest.mock import Mock
        from examples.office_owned_launcher import owned_document_windows
        parents = {10: 100, 20: 16464}
        pids = {**parents, 11: 100, 12: 16464}
        def children(parent, callback, value):
            for child in (11, 12):
                callback(child, value)
        gui = types.SimpleNamespace(
            EnumWindows=lambda callback, value: [callback(hwnd, value) for hwnd in parents],
            EnumChildWindows=Mock(side_effect=children), GetClassName=lambda hwnd: '_WwG')
        process = types.SimpleNamespace(GetWindowThreadProcessId=lambda hwnd: (1, pids[hwnd]))
        with patch.dict(sys.modules, {'win32gui': gui, 'win32process': process}):
            self.assertEqual(owned_document_windows(100, '_WwG'), [11])
        self.assertEqual(gui.EnumChildWindows.call_count, 1)
        self.assertEqual(gui.EnumChildWindows.call_args.args[0], 10)
