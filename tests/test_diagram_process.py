"""Deterministic supervision tests; fakes do not prove Windows containment."""
import sys
from types import SimpleNamespace

import pytest

from office_agent import diagram_process as dp
from office_agent.diagrams import DiagramError

COMMAND = (sys.executable, '-c', 'pass')


class Clock:
    def __init__(self):
        self.now = 0

    def __call__(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


class Process:
    pid = 123

    def __init__(self, codes):
        self.codes = iter(codes)
        self.closed = False

    def poll(self):
        return next(self.codes, None)

    def close(self):
        self.closed = True


class Job:
    def __init__(self, process, cleanup=True, spawn_error=False):
        self.process, self.cleanup, self.spawn_error = process, cleanup, spawn_error
        self.closed = False
        self.spawned = False

    def spawn(self, command):
        self.spawned = True
        if self.spawn_error:
            raise RuntimeError('synthetic launch failure')
        return self.process

    def close(self):
        self.closed = True
        return self.cleanup


def run(job, clock=None, **kwargs):
    clock = clock or Clock()
    return dp._run_owned(COMMAND, _job_factory=lambda: job,
                         _clock=clock, _sleep=clock.sleep, **kwargs)


@pytest.mark.parametrize('exit_code,expected', [(0, 'process_completed'), (7, 'process_failed')])
def test_exit_and_always_clean(exit_code, expected):
    process = Process([None, exit_code])
    job = Job(process)
    result = run(job)
    assert result.code == expected and result.exit_code == exit_code
    assert result.pid == 123 and result.empty_confirmed
    assert job.closed and process.closed


def test_timeout_independent_of_worker_progress():
    process = Process([])
    job = Job(process)
    result = run(job, timeout_ms=100)
    assert result.code == 'render_timeout' and result.exit_code is None
    assert .1 <= result.elapsed_seconds <= .12
    assert job.closed and process.closed


def test_launch_time_counts_against_deadline():
    clock = Clock()
    job = Job(Process([0]))
    original = job.spawn
    def spawn(command):
        clock.now = 1
        return original(command)
    job.spawn = spawn
    result = run(job, clock, timeout_ms=100)
    assert result.code == 'render_timeout' and job.closed


def test_cancel_before_launch_owns_nothing():
    job = Job(Process([]))
    result = run(job, cancelled=lambda: True)
    assert result.code == 'render_cancelled' and result.pid == 0
    assert result.empty_confirmed and not job.spawned and not job.closed


def test_cancel_during_execution():
    clock = Clock()
    job = Job(Process([]))
    result = run(job, clock, cancelled=lambda: clock.now >= .04)
    assert result.code == 'render_cancelled'
    assert job.closed and job.process.closed


def test_cancel_callback_exception_still_cleans():
    count = 0
    job = Job(Process([]))
    def cancelled():
        nonlocal count
        count += 1
        if count == 3:
            raise RuntimeError('synthetic cancellation failure')
        return False
    with pytest.raises(RuntimeError, match='synthetic cancellation failure'):
        run(job, cancelled=cancelled)
    assert job.closed and job.process.closed


def test_spawn_exception_still_closes_job():
    job = Job(Process([]), spawn_error=True)
    with pytest.raises(RuntimeError, match='synthetic launch failure'):
        run(job)
    assert job.closed


@pytest.mark.parametrize('cleanup', [False, None])
def test_cleanup_not_confirmed_never_returns_success(cleanup):
    job = Job(Process([0]), cleanup=cleanup)
    with pytest.raises(DiagramError) as exc:
        run(job)
    assert exc.value.code == 'renderer_cleanup_failed'
    assert job.closed and job.process.closed


def test_cleanup_exception_still_closes_process_handle():
    job = Job(Process([0]))
    def close():
        raise RuntimeError('synthetic cleanup failure')
    job.close = close
    with pytest.raises(RuntimeError, match='synthetic cleanup failure'):
        run(job)
    assert job.process.closed


@pytest.mark.parametrize('command,timeout,cancel', [
    ('shell string', 100, None), ((), 100, None),
    (('relative.exe',), 100, None), ((sys.executable, '\x00'), 100, None),
    ((sys.executable, '\n'), 100, None), ((sys.executable, 1), 100, None),
    ((sys.executable, 'x' * 4097), 100, None),
    ((sys.executable,) * 33, 100, None),
    ((sys.executable,) + ('x' * 4000,) * 5, 100, None),
    (COMMAND, True, None), (COMMAND, 99, None), (COMMAND, 30001, None),
    (COMMAND, 100, True),
])
def test_invalid_inputs_do_not_launch(command, timeout, cancel):
    def forbidden():
        pytest.fail('must not create a job for invalid arguments')
    with pytest.raises(DiagramError) as exc:
        dp._run_owned(command, timeout_ms=timeout, cancelled=cancel, _job_factory=forbidden)
    assert exc.value.code == 'unsupported_source'


def test_environment_excludes_secrets_and_python_path(monkeypatch):
    monkeypatch.setattr(dp, 'os', SimpleNamespace(environ={
        'SystemRoot': 'C:\\Windows', 'temp': 'fixture-temp',
        'OPENAI_API_KEY': 'synthetic-secret', 'OTHER_PROVIDER_TOKEN': 'synthetic-token',
        'PYTHONPATH': 'untrusted-path', 'UV_CACHE_DIR': 'untrusted-cache'}))
    env = dp._worker_environment()
    assert env == {'SystemRoot': 'C:\\Windows', 'temp': 'fixture-temp',
                   'PYTHONNET_RUNTIME': 'netfx', 'OFFICE_WEBVIEW_GUI': 'edgechromium',
                   'PYTHONUTF8': '1', 'PYTHONDONTWRITEBYTECODE': '1'}


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows owned worker containment')
def test_windows_worker_keeps_import_source_without_bytecode_writes(tmp_path):
    module = tmp_path / 'offline_bytecode_probe.py'
    source = 'value = 125\n'
    module.write_text(source, 'utf-8')
    command = (sys.executable, '-c',
               'import sys; sys.path.insert(0, ' + repr(str(tmp_path)) + '); '
               'import offline_bytecode_probe; '
               'assert offline_bytecode_probe.value == 125; '
               'assert sys.dont_write_bytecode')
    result = dp._run_owned(command, timeout_ms=20000)
    assert result.exit_code == 0 and result.empty_confirmed
    assert module.read_text('utf-8') == source
    assert not (tmp_path / '__pycache__').exists()


def windows_fakes(monkeypatch, failure=None):
    events = []
    class Handle:
        def __init__(self, name):
            self.name = name
        def Close(self):
            events.append('close-' + self.name)
    process, thread, handle = Handle('process'), Handle('thread'), Handle('job')
    def create(*args):
        events.append(('create', args))
        return process, thread, 321, 1
    def assign(*args):
        events.append('assign')
        if failure == 'assign':
            raise RuntimeError('synthetic assignment failure')
    def contains(*args):
        events.append('contains')
        return failure != 'contains'
    def resume(*args):
        events.append('resume')
        return 0 if failure == 'resume' else 1
    modules = {
        'win32api': SimpleNamespace(SetLastError=lambda code: None,
                                     GetLastError=lambda: 183 if failure == 'existing' else 0,
                                     TerminateProcess=lambda *args: events.append('terminate-exact')),
        'win32con': SimpleNamespace(CREATE_SUSPENDED=4, CREATE_NO_WINDOW=8),
        'win32event': SimpleNamespace(WAIT_OBJECT_0=0, WaitForSingleObject=lambda *args: 0),
        'win32job': SimpleNamespace(CreateJobObject=lambda *args: handle,
            JobObjectExtendedLimitInformation=1, JobObjectBasicAccountingInformation=2,
            JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE=8192,
            JOB_OBJECT_LIMIT_ACTIVE_PROCESS=8, JOB_OBJECT_LIMIT_JOB_MEMORY=512,
            QueryInformationJobObject=lambda h, kind: {'BasicLimitInformation': {'LimitFlags': 0}} if kind == 1 else {'ActiveProcesses': 0},
            SetInformationJobObject=lambda *args: events.append('set-kill-on-close'),
            AssignProcessToJobObject=assign, IsProcessInJob=contains),
        'win32process': SimpleNamespace(STARTUPINFO=lambda: object(), CreateProcess=create,
                                       ResumeThread=resume),
    }
    for name, module in modules.items():
        monkeypatch.setitem(sys.modules, name, module)
    monkeypatch.setattr(dp, 'os', SimpleNamespace(name='nt', environ={}))
    return events


def test_windows_assign_before_resume_and_no_inherited_handles(monkeypatch):
    events = windows_fakes(monkeypatch)
    job = dp._WindowsJob()
    process = job.spawn(COMMAND)
    create = next(e for e in events if isinstance(e, tuple))
    assert create[1][4] is False and create[1][5] == 12
    assert events.index('assign') < events.index('contains') < events.index('resume')
    assert events.count('close-thread') == 1
    assert job.close() is True
    process.close()
    assert 'close-job' in events and 'close-process' in events


def test_job_name_binding_contract_uses_unique_strings(monkeypatch):
    windows_fakes(monkeypatch)
    module = sys.modules['win32job']
    original = module.CreateJobObject
    names = []
    ids = iter(['0' * 32, '1' * 32])
    monkeypatch.setattr(dp, 'uuid', SimpleNamespace(uuid4=lambda: SimpleNamespace(hex=next(ids))))
    def create(attributes, name):
        # This installed pywin32 rejects None; preserve the native regression.
        if not isinstance(name, str) or not name:
            raise TypeError('None is not a valid string in this context')
        names.append(name)
        return original(attributes, name)
    monkeypatch.setattr(module, 'CreateJobObject', create)
    first, second = dp._WindowsJob(), dp._WindowsJob()
    assert names == ['OfficeDiagram-' + '0' * 32, 'OfficeDiagram-' + '1' * 32]
    assert first.close() is True and second.close() is True


def test_existing_named_job_is_never_adopted(monkeypatch):
    events = windows_fakes(monkeypatch, 'existing')
    with pytest.raises(DiagramError) as exc:
        dp._WindowsJob()
    assert exc.value.code == 'renderer_failed'
    assert events == ['close-job']


@pytest.mark.parametrize('failure', ['assign', 'contains', 'resume'])
def test_windows_launch_failure_terminates_only_created_child(monkeypatch, failure):
    events = windows_fakes(monkeypatch, failure)
    job = dp._WindowsJob()
    with pytest.raises((RuntimeError, DiagramError)):
        job.spawn(COMMAND)
    assert 'terminate-exact' in events
    assert events.count('close-thread') == events.count('close-process') == 1
    if failure != 'resume':
        assert 'resume' not in events
    assert job.close() is True


def test_windows_job_caps_are_set_before_any_child(monkeypatch):
    events = windows_fakes(monkeypatch)
    module = sys.modules['win32job']
    values = []
    original = module.SetInformationJobObject
    def limits(handle, kind, info):
        values.append(info)
        original(handle, kind, info)
    monkeypatch.setattr(module, 'SetInformationJobObject', limits)
    job = dp._WindowsJob()
    assert not any(isinstance(event, tuple) for event in events)
    assert values[0]['BasicLimitInformation']['ActiveProcessLimit'] == 16
    assert values[0]['JobMemoryLimit'] == 1024 * 1024 * 1024
    flags = values[0]['BasicLimitInformation']['LimitFlags']
    assert flags & module.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    assert flags & module.JOB_OBJECT_LIMIT_ACTIVE_PROCESS
    assert flags & module.JOB_OBJECT_LIMIT_JOB_MEMORY
    assert job.close() is True


def test_job_limit_configuration_failure_closes_without_spawning(monkeypatch):
    events = windows_fakes(monkeypatch)
    def refused(*args):
        raise RuntimeError('synthetic limits refusal')
    monkeypatch.setattr(sys.modules['win32job'], 'SetInformationJobObject', refused)
    with pytest.raises(RuntimeError, match='limits refusal'):
        dp._WindowsJob()
    assert events == ['close-job']
