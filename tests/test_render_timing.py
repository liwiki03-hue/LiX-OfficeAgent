import json
from pathlib import Path

from office_agent import diagram_process as dp
from office_agent import fidelity_worker
from tests.test_diagram_process import Clock, Job, Process, COMMAND


def test_optional_owned_timing_records_fixed_numeric_stages_without_changing_result():
    clock = Clock()
    process = Process([None, 0])
    job = Job(process)
    timings = {}
    result = dp._run_owned(COMMAND, _job_factory=lambda: job, _clock=clock, _sleep=clock.sleep, _timings=timings)
    assert result.code == 'process_completed' and result.empty_confirmed
    assert set(timings) == {'job_setup_seconds', 'spawn_seconds', 'wait_seconds', 'job_cleanup_seconds', 'handle_close_seconds'}
    assert all(type(value) is float and value >= 0 for value in timings.values())
    assert timings['wait_seconds'] == .02 and job.closed and process.closed


def test_optional_worker_timing_never_returns_at_first_pdf_appearance(tmp_path, monkeypatch):
    clock = Clock()
    states = iter([None, None, 0, 0])
    class Child:
        waited = False
        def poll(self):
            (tmp_path / 'input.pdf').write_bytes(b'%PDF-observed-not-validated')
            return next(states, 0)
        def wait(self, timeout):
            self.waited = True
        def kill(self):
            raise AssertionError('successful owned child must not be killed')
    child = Child()
    monkeypatch.setattr(fidelity_worker.subprocess, 'Popen', lambda *args, **kwargs: child)
    monkeypatch.setattr(fidelity_worker.time, 'monotonic', clock)
    monkeypatch.setattr(fidelity_worker.time, 'sleep', clock.sleep)
    target = tmp_path / 'timing.json'
    code = fidelity_worker.measured_conversion(['fixed-command'], {}, tmp_path, Path('input.docx'), target, .001)
    timing = json.loads(target.read_text())
    assert code == 0 and child.waited
    assert timing['pdf_first_seen_seconds'] < timing['worker_conversion_seconds']
    assert timing['libreoffice_wait_seconds'] == .02
    assert all(value is None or type(value) in (int, float) and value >= 0 for value in timing.values())


def test_optional_timing_report_write_failure_does_not_replace_conversion_result(tmp_path, monkeypatch):
    class Child:
        def poll(self):
            return 0
        def wait(self, timeout):
            return 0
    monkeypatch.setattr(fidelity_worker.subprocess, 'Popen', lambda *args, **kwargs: Child())
    target = tmp_path / 'existing.json'
    target.write_text('preserve')
    assert fidelity_worker.measured_conversion(['fixed-command'], {}, tmp_path, Path('input.docx'), target, 0) == 0
    assert target.read_text() == 'preserve'


def test_timing_read_fidelity_error_does_not_replace_valid_pdf(tmp_path, monkeypatch):
    from office_agent import fidelity
    from office_agent.diagram_process import OwnedRun
    source = tmp_path / 'source.docx'
    source.write_bytes(b'owned-input')
    actual_read = fidelity.bounded_bytes
    def run(command, timeout, runs, **kwargs):
        destination = Path(command[4])
        (destination / 'source.pdf').write_bytes(b'%PDF-verified-by-later-gate')
        (destination / 'conversion-timing.json').write_bytes(b'{}')
        runs.append({'empty_confirmed': True})
    def read(path):
        if path.name == 'conversion-timing.json':
            raise fidelity.FidelityError('ordinary_file_required')
        return actual_read(path)
    monkeypatch.setattr(fidelity, '_run', run)
    monkeypatch.setattr(fidelity, 'bounded_bytes', read)
    runs = []
    result = fidelity._convert(source, tmp_path / 'output', tmp_path / 'runtime.exe', 1000, runs, timing_details=True)
    assert result.is_file() and runs == [{'empty_confirmed': True}]


def test_timing_stat_error_is_downgraded_without_changing_process_receipt():
    from office_agent import fidelity
    class UnavailableTiming:
        def is_file(self):
            raise OSError('owned telemetry stat unavailable')
    runs = [{'empty_confirmed': True, 'code': 'process_completed'}]
    fidelity._conversion_timing(UnavailableTiming(), runs)
    assert runs == [{'empty_confirmed': True, 'code': 'process_completed'}]
