"""Read-only diagnostics; real COM is only reachable through the opt-in test worker.

Never kill by executable name. A DispatchEx call that hangs before returning an
Hwnd cannot be attributed safely: the worker is stopped, unknown Office stays.
"""
import importlib.util
from contextlib import contextmanager
import os
import re
import sys
from pathlib import Path
import tempfile
import time
import traceback
from unittest import TestCase
from unittest.mock import patch


class ControlledCOMError(RuntimeError):
    pass


class _LoopbackConnection:
    """Bounded JSON framing over a private loopback socket; no pickle/named pipes."""
    def __init__(self, sock, deadline):
        self.sock, self.deadline = sock, deadline

    def _remaining(self):
        self.sock.settimeout(max(0.001, self.deadline - time.monotonic()))

    def send(self, message):
        import json
        import struct
        data = json.dumps(message, ensure_ascii=True).encode("ascii")
        if len(data) > 1024 * 1024:
            raise ControlledCOMError("IPC message too large")
        self._remaining()
        self.sock.sendall(struct.pack("!I", len(data)) + data)

    def _read(self, size):
        chunks = bytearray()
        while len(chunks) < size:
            self._remaining()
            data = self.sock.recv(size - len(chunks))
            if not data:
                raise EOFError("COM worker connection closed")
            chunks.extend(data)
        return bytes(chunks)

    def recv(self):
        import json
        import struct
        size = struct.unpack("!I", self._read(4))[0]
        if size > 1024 * 1024:
            raise ControlledCOMError("IPC message too large")
        return json.loads(self._read(size))

    def poll(self, timeout=0):
        import select
        return bool(select.select([self.sock], [], [], timeout)[0])

    def close(self):
        self.sock.close()


class _SubprocessWorker:
    """Same supervisor interface; child inherits stdio rather than capturing pipes."""
    def __init__(self, process):
        self.process = process

    def is_alive(self):
        return self.process.poll() is None

    @property
    def exitcode(self):
        return self.process.poll()

    def terminate(self):
        self.process.terminate()

    def kill(self):
        self.process.kill()

    def join(self, timeout):
        import subprocess
        try:
            self.process.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            pass


class _OwnedProcesses:
    """Keep kernel handles alive, not just PIDs (prevents PID-reuse cleanup)."""
    def __init__(self):
        import win32api
        import win32event
        import win32process
        self.api, self.event, self.process = win32api, win32event, win32process
        self.preexisting = set(win32process.EnumProcesses())
        self.handles = []
        self.identities = {}
        self.lifecycle = []

    def retire(self, pid, created):
        """Wait for Quit, then stop only the exact acknowledged process if needed."""
        handle = self.identities[(pid, created)]
        forced = self.event.WaitForSingleObject(handle, 500) == 258
        if forced:
            self.api.TerminateProcess(handle, 1)
            if self.event.WaitForSingleObject(handle, 2000) == 258:
                raise ControlledCOMError("Owned Office process did not stop after Quit")
        self.lifecycle.append({"pid": pid, "created": created,
                               "event": "terminated_after_quit" if forced else "quit_exited"})
        self.handles.remove(handle)
        del self.identities[(pid, created)]
        handle.Close()

    def accept(self, pid, created):
        if pid in self.preexisting or pid == os.getpid():
            raise ControlledCOMError("Refusing ownership of a pre-existing process")
        handle = self.api.OpenProcess(0x100000 | 0x1000 | 1, False, pid)
        try:
            if str(self.process.GetProcessTimes(handle)["CreationTime"]) != created:
                raise ControlledCOMError("Process identity changed before ownership acknowledgement")
            job = getattr(self, "job", None)
            if job is not None and (not job.contains(handle)
                                    or self.event.WaitForSingleObject(handle, 0) != 258):
                raise ControlledCOMError("Office is not a live member of this diagnostic's private job")
            self.handles.append(handle)
            self.identities[(pid, created)] = handle
            self.lifecycle.append({"pid": pid, "created": created, "event": "claimed"})
        except BaseException:
            handle.Close()
            raise

    def cleanup(self):
        errors = []
        for handle in self.handles:
            try:
                identity = next((key for key, value in getattr(self, "identities", {}).items()
                                 if value == handle), None)
                forced = self.event.WaitForSingleObject(handle, 0) == 258
                if forced:
                    self.api.TerminateProcess(handle, 1)
                    if self.event.WaitForSingleObject(handle, 2000) == 258:
                        raise ControlledCOMError("Owned Office process did not stop")
                if identity:
                    self.lifecycle.append({"pid": identity[0], "created": identity[1],
                                           "event": "terminated_on_cleanup" if forced else "cleanup_exited"})
            except Exception as exc:
                errors.append(f"owned process cleanup failed: {type(exc).__name__}")
            finally:
                handle.Close()
        self.handles.clear()
        getattr(self, "identities", {}).clear()
        job = getattr(self, "job", None)
        if job is not None:
            try:
                self.lifecycle.extend(job.close())
            except Exception as exc:
                errors.append(f"native job cleanup failed: {exc}")
        return errors


@contextmanager
def _owned_directory(ownership):
    """Remove this fresh exact directory only after confirmed process cleanup."""
    import shutil
    directory = Path(tempfile.mkdtemp(prefix="office-controlled-")).resolve(strict=True)
    try:
        yield str(directory)
    finally:
        job = getattr(ownership, "job", None)
        if job is None or job.empty_confirmed:
            shutil.rmtree(directory)
        else:
            ownership.lifecycle.append({"event": "temporary_directory_preserved",
                                        "path": str(directory), "reason": "job cleanup unconfirmed"})


def _create_fixture(application, source):
    from office_agent.com import isolated_office
    progid = "Word.Application" if application == "word" else "Excel.Application"
    with isolated_office(progid) as app:
        if application == "word":
            document = app.Documents.Add()
            try:
                document.Range().Text = "First paragraph\rprice is 5\rLast paragraph"
                document.SaveAs2(str(source), FileFormat=0, AddToRecentFiles=False)
            finally:
                document.Close(SaveChanges=False)
        else:
            workbook = app.Workbooks.Add(-4167)  # xlWBATWorksheet: one blank worksheet
            try:
                sheet = workbook.Worksheets(1)
                sheet.Name = "Data"
                sheet.Range("A1").Value = 4
                sheet.Range("B1").Value = "keep"
                workbook.SaveAs(str(source), FileFormat=56, AddToMru=False)
            finally:
                workbook.Close(SaveChanges=False)


def _real_case(workdir, application, stage):
    """Only newly-created valid legacy documents; no user documents or model calls."""
    import hashlib
    import json
    from office_agent.com import WindowsComBackend, run_in_office_thread
    from office_agent.writes import Editor, GuardError, LegacyWordReplace, LegacyCellSet

    check = TestCase()
    root = Path(workdir).resolve(strict=True)
    backend = WindowsComBackend(workdir=str(root))
    stage("write_probe")
    probe = backend.probe(application)
    if not probe.get("available"):
        stage("probe_refused", probe=probe, output_count=0, successful_audit_count=0)
        raise ControlledCOMError(f"real {application} write probe failed: {probe.get('reason')}")
    source = root / ("fixture.doc" if application == "word" else "fixture.xls")
    stage("create_valid_fixture", probe=probe)
    run_in_office_thread(lambda: _create_fixture(application, source))
    check.assertTrue(source.is_file())
    check.assertTrue(source.read_bytes().startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"))
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    editor = Editor(root, legacy_backend=backend)
    if application == "word":
        paragraphs = backend.read_word_paragraphs(str(source), limit=5)
        check.assertIsInstance(paragraphs, list)  # Original type assertion retained.
        check.assertEqual(paragraphs, ["First paragraph", "price is 5", "Last paragraph"])
        request = LegacyWordReplace(path=str(source), paragraph_index=1, old="price", new="amount")
    else:
        check.assertEqual(backend.read_excel_cell(str(source), sheet="Data", cell="A1"), 4)
        request = LegacyCellSet(path=str(source), sheet="Data", cell="A1", value=7, expected_original=4)
    stage("plan", source_sha256=original_hash)
    plan = editor.plan(request)
    check.assertTrue(plan.confirmation_required)
    try:
        editor.execute(plan.id)
    except GuardError as exc:
        check.assertIn("explicit human confirmation required", str(exc))
    else:
        raise AssertionError("Unconfirmed fixture write must be refused")
    check.assertFalse(editor.audit.exists())
    stage("confirmed_fixture_execution", preview_before=plan.before, preview_after=plan.after)
    # Option A authorizes these fixed, self-created diagnostic fixture edits only.
    # This never restores persisted approval or grants authorization to user files.
    verified = editor.execute(plan.id, confirmed=True)
    check.assertEqual(verified["status"], "verified")
    output = Path(verified["output"])
    backup = Path(verified["backup"])
    check.assertTrue(output.is_file())
    check.assertTrue(output.read_bytes().startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"))
    output_hash = hashlib.sha256(output.read_bytes()).hexdigest()
    check.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), original_hash)
    check.assertEqual(hashlib.sha256(backup.read_bytes()).hexdigest(), original_hash)
    check.assertEqual(verified["output_sha256"], output_hash)
    if application == "word":
        check.assertEqual(backend.read_word_paragraphs(str(output), limit=5),
                          ["First paragraph", "amount is 5", "Last paragraph"])
    else:
        values = backend.read_excel_region(str(output), sheet="Data", first_cell="A1", last_cell="T20")
        check.assertEqual(values["A1"], 7)
        check.assertEqual(values["B1"], "keep")
    audit = json.loads((editor.audit / plan.id / "audit.json").read_text(encoding="utf-8"))
    check.assertEqual(audit["status"], "verified")
    stage("verified", source_sha256=original_hash, output_sha256=output_hash,
          verification=verified["legacy_verification"], audit_status=audit["status"])
    rolled_back = editor.rollback(plan.id, confirmed=True)
    check.assertEqual(rolled_back["status"], "rolled_back")
    check.assertFalse(output.exists())
    check.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), original_hash)
    return {"application": application, "ok": True, "probe": probe,
            "source_sha256": original_hash, "backup_sha256": original_hash,
            "output_sha256": output_hash, "source_unchanged": True,
            "unconfirmed_write_refused": True, "legacy_verification": verified["legacy_verification"],
            "audit_status_before_rollback": audit["status"], "rollback_status": rolled_back["status"],
            "fixture_and_outputs": "supervisor-owned temporary directory; removed after cleanup"}


@contextmanager
def _identity_hwnd(app, prog_id):
    """Word has Window.Hwnd, not Application.Hwnd. Never open/save user files.

    Word's identity bootstrap creates only a hidden unsaved blank, closes it even
    on refusal, and restores security. File work stays behind the ownership ACK.
    """
    if prog_id != "Word.Application":
        yield int(app.Hwnd)
        return
    previous_security = app.AutomationSecurity
    document = None
    try:
        app.AutomationSecurity = 3
        # A document window is required for Window.Hwnd; DispatchEx starts hidden.
        # Visible=False can create a windowless document on some Word versions.
        document = app.Documents.Add(Visible=True)
        yield int(document.ActiveWindow.Hwnd)
    finally:
        try:
            if document is not None:
                document.Close(SaveChanges=False)
        finally:
            app.AutomationSecurity = previous_security


def _worker(connection, workdir, application):
    handles = []
    try:
        import win32api
        import win32process
        import win32com.client
        original = win32com.client.DispatchEx

        def dispatch(*args, **kwargs):
            from examples.office_owned_launcher import launch_office
            expected = "Word.Application" if application == "word" else "Excel.Application"
            if not args or args[0] != expected:
                raise ControlledCOMError("Native diagnostic activation is restricted to this application")
            registrations = readonly_report()["office_registration"]
            executable = registrations[expected].get("server_executable")
            if not executable or not Path(executable).is_file():
                raise ControlledCOMError("Registered Office server executable is unavailable")

            def acknowledge(pid, created):
                connection.send(("owned", pid, created))
                if connection.recv() != "accepted":
                    raise ControlledCOMError("Parent refused Office ownership")

            app, child, pid, created = launch_office(application, executable, workdir, acknowledge)
            handles.append(child)

            class TrackedApplication:
                def __getattr__(self, name):
                    return getattr(app, name)

                def __setattr__(self, name, value):
                    setattr(app, name, value)

                def Quit(self, *args, **kwargs):
                    try:
                        result = app.Quit(*args, **kwargs)
                        connection.send(("released", pid, created))
                        if connection.recv() != "retired":
                            raise ControlledCOMError("Parent did not confirm Office exit")
                        return result
                    except BaseException:
                        connection.send(("lifecycle_error", "Office Quit failed"))
                        raise

            return TrackedApplication()

        def stage(name, **details):
            connection.send(("stage", {"application": application, "stage": name, **details}))

        original_mkdtemp = tempfile.mkdtemp

        def owned_mkdtemp(suffix=None, prefix=None, dir=None):
            # Production write probes also allocate temporary files: keep ALL of
            # them below the supervisor's root, including if the worker is killed.
            return original_mkdtemp(suffix=suffix, prefix=prefix, dir=workdir)

        with patch.object(win32com.client, "DispatchEx", dispatch), \
                patch.object(tempfile, "mkdtemp", owned_mkdtemp):
            report = _real_case(workdir, application, stage)
        connection.send(("result", True, report))
    except BaseException:
        connection.send(("result", False, traceback.format_exc()))
    finally:
        for child in handles:
            child._handle.Close()
        connection.close()


def _supervise(worker, connection, ownership, timeout, clock=time.monotonic):
    deadline = clock() + timeout
    result = None
    errors = []
    progress = []
    try:
        while clock() < deadline:
            if connection.poll(min(0.05, max(0, deadline - clock()))):
                try:
                    message = connection.recv()
                except EOFError:
                    worker.join(max(0, deadline - time.monotonic()))
                    break  # Socket close can precede the child's actual exit by milliseconds.
                if message[0] == "owned":
                    ownership.accept(message[1], message[2])
                    connection.send("accepted")
                elif message[0] == "released":
                    ownership.retire(message[1], message[2])
                    connection.send("retired")
                elif message[0] == "stage":
                    progress.append(message[1])
                    print("COM_STAGE " + message[1]["application"] + " " + message[1]["stage"], flush=True)
                elif message[0] == "lifecycle_error":
                    errors.append(message[1])
                elif message[0] == "result":
                    result = message
            if not worker.is_alive():
                if result is None and connection.poll():
                    continue  # Drain already-sent stage/result messages after worker exit.
                break
        if worker.is_alive():
            errors.append("COM worker hard timeout; unknown/unattributed Office processes were NOT killed")
        elif result is None:
            errors.append("COM worker exited without a result")
        elif not result[1]:
            errors.append(result[2])
        if worker.exitcode not in (None, 0):
            errors.append(f"COM worker exited with code {worker.exitcode}")
    except Exception as exc:
        errors.append(f"Ownership supervision failed closed: {exc}")
    finally:
        try:
            if worker.is_alive():
                worker.terminate()
            worker.join(2)
            if worker.is_alive():
                worker.kill()
                worker.join(2)
            if worker.is_alive():
                errors.append("COM worker could not be stopped")
        except Exception as exc:
            errors.append(f"Worker cleanup failed: {type(exc).__name__}")
        finally:
            try:
                errors.extend(ownership.cleanup())
            finally:
                connection.close()
    if errors:
        failure = ControlledCOMError("\n".join(errors))
        failure.report = {"ok": False, "progress": progress,
                          "process_lifecycle": getattr(ownership, "lifecycle", [])}
        raise failure
    return result[2]


def run_real_case(timeout=20):
    if os.environ.get("OFFICE_AGENT_REAL_COM") != "1":
        raise ControlledCOMError("Real COM requires OFFICE_AGENT_REAL_COM=1")
    if os.name != "nt":
        raise ControlledCOMError("Real COM requires Windows")
    if not 0 < timeout <= 20:
        raise ValueError("timeout must be in (0, 20]")
    report = {"ok": False, "applications": [], "timeout_seconds_per_application": timeout,
              "confirmation_scope": "User option A; fixed self-created diagnostic fixtures only"}
    for application in ("word", "excel"):
        ownership = _OwnedProcesses()  # Exclude ALL PIDs existing before this worker.
        import secrets
        import socket
        import subprocess
        token = secrets.token_hex(32)
        worker = connection = None
        # The parent owns this exact fresh directory, including on worker timeout.
        try:
            with _owned_directory(ownership) as directory:
                from examples.office_owned_launcher import NativeJob
                ownership.job = NativeJob()
                process = None
                workdir = str(Path(directory).resolve(strict=True))
                try:
                    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
                        listener.bind(("127.0.0.1", 0))
                        listener.listen(1)
                        listener.settimeout(timeout)
                        deadline = time.monotonic() + timeout
                        executable = getattr(sys, "_base_executable", sys.executable)
                        # The exact suspended child is assigned BEFORE any worker code runs.
                        process = ownership.job.spawn_worker([
                            executable, str(Path(__file__).resolve()), "--worker",
                            str(listener.getsockname()[1]), workdir, application, token])
                        worker = _SubprocessWorker(process)
                        sock, _ = listener.accept()
                        connection = _LoopbackConnection(sock, deadline)
                        hello = connection.recv()
                        if hello != ["hello", token, process.pid]:
                            raise ControlledCOMError("Worker authentication failed; Office not authorized")
                        connection.send("authenticated")
                        remaining = deadline - time.monotonic()
                        if remaining <= 0:
                            raise ControlledCOMError("Worker bootstrap exhausted the application deadline")
                        result = _supervise(worker, connection, ownership, remaining)
                        result["process_lifecycle"] = ownership.lifecycle
                        result["activation"] = "native_child_in_private_job; OBJID_NATIVEOM"
                        result["startup_switch"] = "/a" if application == "word" else "/safe"
                finally:
                    try:
                        # Job cleanup also owns unacknowledged Office children.
                        if worker is not None:
                            if worker.is_alive():
                                worker.terminate()
                            worker.join(2)
                        if connection:
                            connection.close()
                        cleanup_errors = ownership.cleanup()
                        if cleanup_errors:
                            raise ControlledCOMError("\n".join(cleanup_errors))
                    finally:
                        if process is not None:
                            process.close()
            report["applications"].append(result)  # Only after the owned directory was removed.
        except Exception as exc:
            failed = getattr(exc, "report", {})
            failed.update(application=application, ok=False, error=str(exc))
            report["applications"].append(failed)
            report["not_attempted"] = [name for name in ("word", "excel")
                                       if name not in {entry["application"] for entry in report["applications"]}]
            failure = ControlledCOMError(str(exc))
            failure.report = report
            raise failure from exc  # No automatic retries or next-application startup.
    report["ok"] = True
    return report


def readonly_report():
    """No COM activation, cache imports, filesystem writes, or secrets."""
    dependencies = {}
    for name in ("pythoncom", "win32api", "win32process", "win32com"):
        try:
            dependencies[name] = importlib.util.find_spec(name) is not None
        except (ImportError, ValueError):
            dependencies[name] = False
    registrations = {}
    if os.name == "nt":
        import winreg
        for progid in ("Word.Application", "Excel.Application"):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID", 0,
                                    winreg.KEY_READ) as key:
                    clsid = winreg.QueryValueEx(key, None)[0]
                entry = {"registered": bool(clsid),
                         "installation": "unknown (registration is not installation proof)"}
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID\\" + clsid + "\\LocalServer32", 0,
                                        winreg.KEY_READ) as server:
                        command = winreg.QueryValueEx(server, None)[0]
                    match = re.match(r'^"?(.+?\.exe)"?(?:\s|$)', os.path.expandvars(command), re.IGNORECASE)
                    if match:
                        executable = Path(match.group(1))
                        entry.update(server_executable=str(executable), server_exists=executable.is_file())
                        if executable.is_file():
                            entry["installation"] = "registered server binary present; COM operability unverified"
                except (OSError, TypeError):
                    entry["server_lookup"] = "unavailable (not proof of Office absence)"
                registrations[progid] = entry
            except OSError:
                registrations[progid] = {"registered": False, "installation": "unknown"}
    # Read OS attributes only. os.access is advisory, not an ACL/write proof.
    # tempfile.gettempdir() probes candidates by writing a file: do not call it here.
    candidates = [Path(value) / "gen_py" for key in ("TEMP", "TMP")
                  if (value := os.environ.get(key))]
    candidates += [path / f"{sys.version_info.major}.{sys.version_info.minor}" for path in candidates]
    candidates = list(dict.fromkeys(candidates))
    spec = importlib.util.find_spec("win32com") if dependencies.get("win32com") else None
    if spec and spec.submodule_search_locations:
        candidates.extend(Path(p) / "gen_py" for p in spec.submodule_search_locations)
    cache = []
    for path in candidates:
        entry = {"path": str(path), "writable": "unknown (no write attempted)"}
        try:
            info = path.stat()
            entry.update(exists=True, is_directory=path.is_dir(),
                         readonly_attribute=bool(getattr(info, "st_file_attributes", 0) & 1),
                         access_write_advisory=os.access(path, os.W_OK))
        except OSError:
            entry["exists"] = False
        cache.append(entry)
    return {"dependencies": dependencies, "office_registration": registrations,
            "cache_attributes": cache, "com_validation": "not performed; no Office instance started",
            "limitation": "cache locations are candidates, not imported gencache's actual active location"}


if __name__ == "__main__":
    import argparse
    import json
    from datetime import datetime, timezone

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--real-roundtrip", action="store_true",
                        help="User-authorized self-created Word/Excel fixtures; also requires OFFICE_AGENT_REAL_COM=1")
    parser.add_argument("--report", type=Path, help="Create a NEW UTF-8 evidence JSON (never overwrite)")
    parser.add_argument("--worker", nargs=4, help=argparse.SUPPRESS)
    options = parser.parse_args()
    if options.worker:
        import site
        import socket
        if os.environ.get("OFFICE_AGENT_REAL_COM") != "1":
            parser.error("Internal worker also requires explicit real-COM opt-in")
        port, workdir, application, token = options.worker
        if application not in {"word", "excel"}:
            parser.error("Unsupported worker application")
        # Activate the project and uv-installed dependencies without a launcher grandchild.
        project = Path(__file__).resolve().parents[1]
        sys.path.insert(0, str(project))
        site.addsitedir(str(project / ".venv" / "Lib" / "site-packages"))
        with socket.create_connection(("127.0.0.1", int(port)), timeout=20) as sock:
            channel = _LoopbackConnection(sock, time.monotonic() + 20)
            channel.send(["hello", token, os.getpid()])
            if channel.recv() != "authenticated":
                raise ControlledCOMError("COM activation denied before worker authentication")
            _worker(channel, workdir, application)
        raise SystemExit(0)
    if options.report and not options.real_roundtrip:
        parser.error("--report requires --real-roundtrip; default diagnosis remains read-only")
    if options.report and (options.report.exists() or options.report.is_symlink()):
        parser.error("Report already exists; refusing to overwrite or start Office")
    diagnostics = readonly_report()
    exit_code = 0
    if options.real_roundtrip:
        try:
            report = run_real_case(timeout=20)
        except ControlledCOMError as exc:
            exit_code = 1
            report = getattr(exc, "report", {"ok": False, "error": str(exc)})
        report["diagnostics_before"] = diagnostics
        report["recorded_at"] = datetime.now(timezone.utc).isoformat()
    else:
        report = diagnostics
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if options.report:
        with options.report.open("x", encoding="utf-8") as evidence:
            evidence.write(rendered + "\n")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    raise SystemExit(exit_code)
