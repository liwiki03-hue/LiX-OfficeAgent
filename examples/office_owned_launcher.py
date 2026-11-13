"""Diagnostic-only native Office ownership, established before COM activation.

Never use this as a model tool. A suspended worker enters a private kill-on-close
job before executing; Office starts directly as its child, never via COM launch.
Native OM binds only a document window in the directly-created owned process.
"""
from pathlib import Path
import os
import subprocess
import time
import uuid


class NativeProcess:
    def __init__(self, handle, pid, command):
        import win32api
        import win32event
        import win32process
        self.handle, self.pid, self.command = handle, pid, command
        self.api, self.event, self.process = win32api, win32event, win32process

    def poll(self):
        code = self.process.GetExitCodeProcess(self.handle)
        return None if code == 259 else code

    def wait(self, timeout=None):
        milliseconds = 0xFFFFFFFF if timeout is None else max(0, int(timeout * 1000))
        if self.event.WaitForSingleObject(self.handle, milliseconds) == 258:
            raise subprocess.TimeoutExpired(self.command, timeout)
        return self.process.GetExitCodeProcess(self.handle)

    def terminate(self):
        self.api.TerminateProcess(self.handle, 1)

    kill = terminate

    def close(self):
        self.handle.Close()


class NativeJob:
    """Contains only a suspended child we created and its inherited descendants."""
    def __init__(self):
        import win32job
        self.api = win32job
        self.handle = win32job.CreateJobObject(None, 'OfficeAgent-' + uuid.uuid4().hex)
        try:
            info = win32job.QueryInformationJobObject(
                self.handle, win32job.JobObjectExtendedLimitInformation)
            info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            win32job.SetInformationJobObject(
                self.handle, win32job.JobObjectExtendedLimitInformation, info)
        except BaseException:
            self.handle.Close()
            raise
        self.closed = False
        self.empty_confirmed = False

    def contains(self, process_handle):
        return bool(self.api.IsProcessInJob(process_handle, self.handle))

    def spawn_worker(self, command):
        import win32api
        import win32con
        import win32event
        import win32process
        startup = win32process.STARTUPINFO()
        # GetStdHandle constants are not exported by this pywin32's win32con.
        standard = [win32api.GetStdHandle(value) for value in (-10, -11, -12)]
        if all(standard):
            startup.dwFlags |= win32con.STARTF_USESTDHANDLES
            startup.hStdInput, startup.hStdOutput, startup.hStdError = standard
        text = subprocess.list2cmdline(command)
        process, thread, pid, _ = win32process.CreateProcess(
            command[0], text, None, None, True,
            win32con.CREATE_SUSPENDED | win32con.CREATE_NO_WINDOW,
            None, None, startup)
        try:
            # Failure MUST stop this still-suspended exact child. No fallback launch.
            self.api.AssignProcessToJobObject(self.handle, process)
            if not self.contains(process):
                raise RuntimeError('Worker did not enter the private job')
            if win32process.ResumeThread(thread) != 1:
                raise RuntimeError('Unexpected worker suspension count')
            return NativeProcess(process, pid, text)
        except BaseException:
            win32api.TerminateProcess(process, 1)
            if win32event.WaitForSingleObject(process, 2000) == 258:
                raise RuntimeError('Suspended worker cleanup was not confirmed')
            process.Close()
            raise
        finally:
            thread.Close()

    def close(self):
        if self.closed:
            return []
        try:
            active = self.api.QueryInformationJobObject(
                self.handle, self.api.JobObjectBasicAccountingInformation)['ActiveProcesses']
            forced = bool(active)
            if active:
                self.api.TerminateJobObject(self.handle, 1)
            deadline = time.monotonic() + 2
            while self.api.QueryInformationJobObject(
                    self.handle, self.api.JobObjectBasicAccountingInformation)['ActiveProcesses']:
                if time.monotonic() >= deadline:
                    raise RuntimeError('Native job still has active processes after cleanup')
                time.sleep(0.02)
            self.empty_confirmed = True
            return [{'event': 'native_job_empty', 'forced': forced,
                     'active_processes': 0}]
        finally:
            self.closed = True
            self.handle.Close()  # KILL_ON_JOB_CLOSE also covers interrupted cleanup.


def owned_document_windows(pid, class_name):
    """Select document windows by the directly-created PID, never by title."""
    import win32gui
    import win32process
    result = []

    def child(hwnd, _):
        if (win32process.GetWindowThreadProcessId(hwnd)[1] == pid
                and win32gui.GetClassName(hwnd) == class_name):
            result.append(hwnd)
        return True

    def parent(hwnd, _):
        if win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
            win32gui.EnumChildWindows(hwnd, child, None)
        return True

    win32gui.EnumWindows(parent, None)
    return result


def native_window(hwnd):
    """MS documented OBJID_NATIVEOM; never calls a COM activation API."""
    import ctypes
    from ctypes import wintypes
    import pythoncom
    import win32com.client
    # IID_IDispatch: bytes_le is the native GUID layout, including Data1..3.
    iid = (ctypes.c_ubyte * 16).from_buffer_copy(
        uuid.UUID('00020400-0000-0000-c000-000000000046').bytes_le)
    pointer = ctypes.c_void_p()
    function = ctypes.WinDLL('oleacc').AccessibleObjectFromWindow
    function.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.c_void_p,
                         ctypes.POINTER(ctypes.c_void_p)]
    function.restype = ctypes.c_long
    status = function(hwnd, 0xFFFFFFF0, ctypes.byref(iid), ctypes.byref(pointer))
    if status != 0 or not pointer.value:
        raise RuntimeError(f'Owned window native object unavailable: HRESULT {status}')
    try:
        wrapped = pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch)
        return win32com.client.Dispatch(wrapped)
    finally:
        # ObjectFromAddress AddRefs its wrapper; release the API's original ref.
        table = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        release = ctypes.WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)(table[2])
        release(pointer)


def launch_office(application, executable, workdir, acknowledge):
    """Bind native OM only through a document window in our direct child."""
    import pythoncom
    import win32process

    root = Path(workdir).resolve(strict=True)
    # One supported switch per application. Word /a creates its own blank with
    # personal templates/add-ins disabled; Excel /safe bypasses startup workbooks.
    arguments = ['/a']
    if application == 'word':
        bootstrap = root / ('native-identity-' + uuid.uuid4().hex + '.rtf')
        with bootstrap.open('x', encoding='ascii') as stream:
            stream.write(r'{\rtf1\ansi Office Agent identity}')
        arguments.append(str(bootstrap))
    if application == 'excel':
        bootstrap = root / ('native-identity-' + uuid.uuid4().hex + '.csv')
        with bootstrap.open('x', encoding='ascii') as stream:
            stream.write('Office Agent identity\n')
        arguments = ['/safe', str(bootstrap)]
    elif application != 'word':
        raise ValueError('Unsupported native Office application')
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = 0
    child = subprocess.Popen([str(executable), *arguments],
                             stdin=subprocess.DEVNULL, startupinfo=startup,
                             close_fds=True)
    # CPython's retained process handle protects identity from PID reuse. The
    # parent job already contains this child even if there is no ACK or window.
    created = str(win32process.GetProcessTimes(int(child._handle))['CreationTime'])
    acknowledge(child.pid, created)
    deadline = time.monotonic() + 18
    class_name = '_WwG' if application == 'word' else 'EXCEL7'
    while True:
        windows = owned_document_windows(child.pid, class_name)
        if windows:
            window = native_window(windows[0])
            app = window.Application
            hwnd = int(window.Hwnd if application == 'word' else app.Hwnd)
            if win32process.GetWindowThreadProcessId(hwnd)[1] != child.pid:
                # Do not Quit or close anything belonging to a foreign PID.
                raise RuntimeError('Native object does not belong to the directly-created child')
            app.AutomationSecurity = 3
            app.DisplayAlerts = 0
            app.Visible = False
            document = window.Document if application == 'word' else app.ActiveWorkbook
            if Path(document.FullName).resolve() != bootstrap.resolve():
                raise RuntimeError('Native startup document is not the self-created bootstrap')
            count = app.Documents.Count if application == 'word' else app.Workbooks.Count
            if count != 1:
                raise RuntimeError('Unexpected additional startup documents in the native child')
            document.Close(SaveChanges=False)
            return app, child, child.pid, created
        if child.poll() is not None:
            raise RuntimeError(f'Owned {application} exited before its native document window appeared')
        if time.monotonic() >= deadline:
            raise RuntimeError(f'Owned {application} native document window did not appear')
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)
