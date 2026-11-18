"""Host-only Windows process supervision; not a model tool or renderer API.

This is deliberately not imported by the production UI yet. The host chooses a
fixed command; neither diagram text nor model parameters may choose a command.
No examples/diagnostic imports, shell, process-name killing or inherited handles.
"""
from dataclasses import dataclass
import os
from pathlib import Path
import subprocess
import time
import uuid
from typing import Callable

from .diagrams import DiagramError


@dataclass(frozen=True)
class OwnedRun:
    code: str
    exit_code: int | None
    pid: int
    elapsed_seconds: float
    empty_confirmed: bool


def _worker_environment() -> dict[str, str]:
    # Do not inherit provider secrets, Python search-path overrides or .env keys.
    keys = ('SystemRoot', 'WINDIR', 'TEMP', 'TMP', 'PATH', 'LOCALAPPDATA',
            'APPDATA', 'USERPROFILE', 'PROGRAMDATA')
    allowed = {k.casefold() for k in keys}
    env = {k: v for k, v in os.environ.items() if k.casefold() in allowed}
    env.update(PYTHONNET_RUNTIME='netfx', OFFICE_WEBVIEW_GUI='edgechromium',
               PYTHONUTF8='1', PYTHONDONTWRITEBYTECODE='1')
    return env


class _Process:
    def __init__(self, handle, pid, api):
        self.handle, self.pid, self.api = handle, pid, api

    def poll(self):
        code = self.api.GetExitCodeProcess(self.handle)
        return None if code == 259 else code

    def close(self):
        self.handle.Close()


class _WindowsJob:
    def __init__(self):
        if os.name != 'nt':
            raise DiagramError('renderer_unavailable')
        import win32api
        import win32con
        import win32event
        import win32job
        import win32process
        self.api, self.con, self.event = win32api, win32con, win32event
        self.job, self.process = win32job, win32process
        win32api.SetLastError(0)
        self.handle = win32job.CreateJobObject(None, 'OfficeDiagram-' + uuid.uuid4().hex)
        if win32api.GetLastError() == 183:  # Never own/reconfigure a pre-existing Job.
            self.handle.Close()
            raise DiagramError('renderer_failed')
        try:
            info = win32job.QueryInformationJobObject(
                self.handle, win32job.JobObjectExtendedLimitInformation)
            info['BasicLimitInformation']['LimitFlags'] |= (
                win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                | win32job.JOB_OBJECT_LIMIT_ACTIVE_PROCESS
                | win32job.JOB_OBJECT_LIMIT_JOB_MEMORY)
            info['BasicLimitInformation']['ActiveProcessLimit'] = 16
            info['JobMemoryLimit'] = 1024 * 1024 * 1024
            win32job.SetInformationJobObject(
                self.handle, win32job.JobObjectExtendedLimitInformation, info)
        except BaseException:
            self.handle.Close()
            raise

    def spawn(self, command):
        process, thread, pid, _ = self.process.CreateProcess(
            command[0], subprocess.list2cmdline(command), None, None, False,
            self.con.CREATE_SUSPENDED | self.con.CREATE_NO_WINDOW,
            _worker_environment(), None, self.process.STARTUPINFO())
        try:
            self.job.AssignProcessToJobObject(self.handle, process)
            if not self.job.IsProcessInJob(process, self.handle):
                raise DiagramError('renderer_failed')
            if self.process.ResumeThread(thread) != 1:
                raise DiagramError('renderer_failed')
            return _Process(process, pid, self.process)
        except BaseException:
            try:
                self.api.TerminateProcess(process, 1)
                if self.event.WaitForSingleObject(process, 2000) != self.event.WAIT_OBJECT_0:
                    raise DiagramError('renderer_cleanup_failed')
            finally:
                process.Close()
            raise
        finally:
            thread.Close()

    def close(self):
        try:
            def active():
                return self.job.QueryInformationJobObject(
                    self.handle, self.job.JobObjectBasicAccountingInformation)['ActiveProcesses']
            if active():
                self.job.TerminateJobObject(self.handle, 1)
            deadline = time.monotonic() + 2
            while active():
                if time.monotonic() >= deadline:
                    raise DiagramError('renderer_cleanup_failed')
                time.sleep(0.01)
            return True
        finally:
            self.handle.Close()  # Also kills owned descendants if polling failed.


def _run_owned(command: tuple[str, ...], *, timeout_ms: int = 8000,
               cancelled: Callable[[], bool] | None = None,
               _job_factory=_WindowsJob, _clock=time.monotonic,
               _sleep=time.sleep, _timings: dict | None = None) -> OwnedRun:
    """Internal host primitive, with injectable clock/job only for unit tests.

    Timeout includes launch time but cannot preempt a blocked Windows API call.
    During child execution the parent polls independently of the child JS/CPU.
    Cleanup failure raises, never returns an optimistic success result.
    """
    if (not isinstance(command, tuple) or not 1 <= len(command) <= 32
            or any(not isinstance(arg, str) or not arg or len(arg) > 4096
                   or any(c in arg for c in '\x00\r\n') for arg in command)
            or sum(map(len, command)) > 16384
            or not Path(command[0]).is_absolute()
            or type(timeout_ms) is not int or not 100 <= timeout_ms <= 30000
            or (cancelled is not None and not callable(cancelled))):
        raise DiagramError('unsupported_source')
    started = _clock()
    if type(_timings) is not dict:
        _timings = None
    if cancelled is not None and cancelled():
        return OwnedRun('render_cancelled', None, 0, 0, True)
    job = process = None
    stage_started = started
    try:
        job = _job_factory()
        if _timings is not None:
            _timings['job_setup_seconds'] = max(0.0, _clock() - stage_started)
        stage_started = _clock()
        process = job.spawn(command)
        if _timings is not None:
            _timings['spawn_seconds'] = max(0.0, _clock() - stage_started)
        stage_started = _clock()
        deadline = started + timeout_ms / 1000
        while True:
            if cancelled is not None and cancelled():
                code, exit_code = 'render_cancelled', None
                break
            if _clock() >= deadline:
                code, exit_code = 'render_timeout', None
                break
            exit_code = process.poll()
            if exit_code is not None:
                code = 'process_completed' if exit_code == 0 else 'process_failed'
                break
            _sleep(min(0.02, max(0, deadline - _clock())))
    finally:
        if _timings is not None:
            _timings['wait_seconds'] = max(0.0, _clock() - stage_started)
        stage_started = _clock()
        try:
            if job is not None:
                if job.close() is not True:
                    raise DiagramError('renderer_cleanup_failed')
        finally:
            if _timings is not None:
                _timings['job_cleanup_seconds'] = max(0.0, _clock() - stage_started)
            stage_started = _clock()
            if process is not None:
                process.close()
            if _timings is not None:
                _timings['handle_close_seconds'] = max(0.0, _clock() - stage_started)
    return OwnedRun(code, exit_code, process.pid, round(_clock() - started, 3), True)
