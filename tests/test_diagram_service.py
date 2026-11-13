"""Channel/host regression: fake worker, no GUI/provider/Office. Native proof separate."""
import json
from pathlib import Path
import socket
import struct
import threading
import time
from types import SimpleNamespace

import pytest

from office_agent import diagram_channel as channel, diagram_service as service
from office_agent.diagram_process import OwnedRun
from office_agent.diagrams import DiagramError
from test_diagrams import SOURCES, SVG


class Wire:
    def __init__(self, data=b''):
        self.data, self.sent = data, b''
    def recv(self, size):
        chunk, self.data = self.data[:min(size, 7)], self.data[min(size, 7):]
        return chunk
    def sendall(self, data):
        self.sent += data


def packed(text):
    data = text.encode('utf-8')
    return struct.pack('!I', len(data)) + data


def test_channel_partial_chunks_roundtrip():
    output = Wire()
    payload = {'svg': SVG, 'ok': True, 'bridge_absent': True}
    channel.send_message(output, payload)
    assert channel.receive_message(Wire(output.sent), time.monotonic() + 1) == payload


@pytest.mark.parametrize('wire', [b'', b'\x00', struct.pack('!I', 0),
    struct.pack('!I', channel.MAX_MESSAGE + 1), packed('[]'), packed('{'),
    packed('{"ok":true,"ok":false}'), packed('{"a":NaN}'),
    struct.pack('!I', 2) + b'\xff\xff'])
def test_channel_rejects_bad_length_json_and_duplicate_keys(wire):
    with pytest.raises((ValueError, UnicodeError)):
        channel.receive_message(Wire(wire), time.monotonic() + 1)


def test_channel_budget_and_stop_before_read():
    stopped = threading.Event(); stopped.set()
    with pytest.raises(TimeoutError):
        channel.read_exact(Wire(b'abcd'), 4, time.monotonic() + 1, stopped)
    with pytest.raises(TimeoutError):
        channel.read_exact(Wire(b'abcd'), 4, time.monotonic() - 1)
    output = Wire()
    with pytest.raises(ValueError):
        channel.send_message(output, {'a': 'x' * channel.MAX_MESSAGE})
    assert output.sent == b''


@pytest.mark.parametrize('payload', [None, [], {}, {'ok':1}, {'ok':True,'svg':SVG},
    {'ok':True,'svg':SVG,'bridge_absent':False},
    {'ok':True,'svg':SVG,'bridge_absent':True,'source':'leak'},
    {'ok':False,'code':'unknown','detail':'private'}, {'ok':False,'code':'unknown'}])
def test_response_shape_does_not_confer_success(payload):
    with pytest.raises(DiagramError) as exc:
        service._validated_response(payload)
    assert exc.value.code == 'renderer_failed'


@pytest.mark.parametrize('svg', [None, 1, '<svg><script>bad</script></svg>',
    SVG.replace('fill="white"', 'onload="bad"')])
def test_host_revalidates_even_authenticated_svg(svg):
    with pytest.raises(DiagramError) as exc:
        service._validated_response({'ok': True, 'svg': svg, 'bridge_absent': True})
    assert exc.value.code == 'unsafe_svg'


def install_worker(monkeypatch, payload, code='process_completed', exit_code=0,
                   invalid_auth=False):
    seen = []
    def run(command, **kwargs):
        assert command[0] == service.sys.executable
        assert Path(command[1]).name == 'diagram_worker.py'
        port, key = int(command[2]), bytes.fromhex(command[3])
        if invalid_auth:
            with socket.create_connection(('127.0.0.1', port), timeout=1) as bad:
                bad.settimeout(1)
                bad.sendall(channel.MAGIC + b'0' * 32)
                assert bad.recv(1) == b''  # No source sent before authentication.
        with socket.create_connection(('127.0.0.1', port), timeout=1) as connection:
            connection.settimeout(0.1)
            connection.sendall(channel.MAGIC + key)
            seen.append(channel.receive_message(connection, time.monotonic() + 1))
            channel.send_message(connection, payload)
        return OwnedRun(code, exit_code, 123, 0.04, True)
    monkeypatch.setattr(service, '_run_owned', run)
    return seen


def assert_clean():
    assert not any(t.name == 'owned-diagram-channel' for t in threading.enumerate())


def test_host_pipeline_authenticates_and_returns_metadata_only(monkeypatch):
    seen = install_worker(monkeypatch, {'ok': True, 'svg': SVG, 'bridge_absent': True}, invalid_auth=True)
    result = service.render_diagram(SOURCES['flowchart'])
    assert seen == [{'source': SOURCES['flowchart'], 'kind': 'flowchart'}]
    assert result.svg == service.validate_svg(SVG) and result.execution.empty_confirmed
    evidence = result.evidence()
    assert evidence['rendered'] is True and evidence['svg_bytes'] == len(result.svg.encode())
    assert 'svg' not in evidence and 'source' not in evidence
    assert len(evidence['svg_sha256']) == 64
    assert_clean()


@pytest.mark.parametrize('payload,code,expected', [
    ({'ok':True,'svg':'<svg><script/></svg>','bridge_absent':True},'process_completed','unsafe_svg'),
    ({'ok':False,'code':'unsupported_source'},'process_failed','unsupported_source'),
    ({'ok':False,'code':'renderer_failed'},'process_failed','renderer_failed'),
    ({'ok':True,'svg':SVG,'bridge_absent':True},'process_failed','renderer_failed'),
    ({'ok':True,'svg':SVG,'bridge_absent':True},'render_timeout','render_timeout'),
])
def test_host_errors_cleanup_and_never_return_preview(monkeypatch, payload, code, expected):
    install_worker(monkeypatch, payload, code, 7)
    with pytest.raises(DiagramError) as exc:
        service.render_diagram(SOURCES['flowchart'])
    assert exc.value.code == expected
    assert_clean()


@pytest.mark.parametrize('kwargs', [{'timeout_ms':True}, {'timeout_ms':99},
                                   {'timeout_ms':8001}, {'cancelled':False}])
def test_invalid_options_do_not_open_a_channel(monkeypatch, kwargs):
    monkeypatch.setattr(service.socket, 'socket', lambda *a: pytest.fail('channel created'))
    with pytest.raises(DiagramError):
        service.render_diagram(SOURCES['flowchart'], **kwargs)


def test_unsafe_source_and_pre_cancel_do_not_launch(monkeypatch):
    monkeypatch.setattr(service, '_run_owned', lambda *a, **k: pytest.fail('worker launched'))
    with pytest.raises(DiagramError):
        service.render_diagram('flowchart TD\nA-->B\nclick A "https://example.com"')
    with pytest.raises(DiagramError) as exc:
        service.render_diagram(SOURCES['flowchart'], cancelled=lambda: True)
    assert exc.value.code == 'render_cancelled'


def test_running_cancel_closes_channel_thread(monkeypatch):
    calls = iter([False, True])
    def run(command, **kwargs):
        assert kwargs['cancelled']() is True
        return OwnedRun('render_cancelled', None, 123, 0.03, True)
    monkeypatch.setattr(service, '_run_owned', run)
    with pytest.raises(DiagramError) as exc:
        service.render_diagram(SOURCES['flowchart'], cancelled=lambda: next(calls))
    assert exc.value.code == 'render_cancelled'
    assert_clean()


def test_launch_exception_does_not_leak_receiver_or_slots(monkeypatch):
    def run(*a, **k): raise RuntimeError('synthetic launch failure')
    monkeypatch.setattr(service, '_run_owned', run)
    with pytest.raises(RuntimeError):
        service.render_diagram(SOURCES['flowchart'])
    assert_clean()
    assert service._slots.acquire(blocking=False)
    assert service._slots.acquire(blocking=False)
    assert not service._slots.acquire(blocking=False)
    service._slots.release(); service._slots.release()


def test_max_two_service_requests_fail_closed():
    assert service._slots.acquire(blocking=False)
    assert service._slots.acquire(blocking=False)
    try:
        with pytest.raises(DiagramError) as exc:
            service.render_diagram(SOURCES['flowchart'])
        assert exc.value.code == 'renderer_failed'
    finally:
        service._slots.release(); service._slots.release()


def test_expected_kind_mismatch_never_starts_channel(monkeypatch):
    monkeypatch.setattr(service.socket, 'socket', lambda *a: pytest.fail('channel created'))
    with pytest.raises(DiagramError):
        service.render_diagram(SOURCES['flowchart'], expected_kind='classDiagram')


def test_receiver_cleanup_failure_overrides_timeout_and_releases_slots(monkeypatch):
    monkeypatch.setattr(service, '_disabled', threading.Event())
    thread = SimpleNamespace(start=lambda: None, join=lambda _: None, is_alive=lambda: True)
    monkeypatch.setattr(service.threading, 'Thread', lambda **k: thread)
    monkeypatch.setattr(service, '_run_owned', lambda *a, **k: OwnedRun('render_timeout', None, 123, 0.1, True))
    with pytest.raises(DiagramError) as exc:
        service.render_diagram(SOURCES['flowchart'])
    assert exc.value.code == 'renderer_cleanup_failed'
    assert service._disabled.is_set()
    with pytest.raises(DiagramError) as blocked:
        service.render_diagram(SOURCES['flowchart'])
    assert blocked.value.code == 'renderer_cleanup_failed'
    assert service._slots.acquire(blocking=False)
    assert service._slots.acquire(blocking=False)
    assert not service._slots.acquire(blocking=False)
    service._slots.release(); service._slots.release()


@pytest.mark.parametrize('failure', ['false', 'unknown', 'exception'])
def test_unconfirmed_job_cleanup_disables_future_rendering(monkeypatch, failure):
    monkeypatch.setattr(service, '_disabled', threading.Event())
    def run(*args, **kwargs):
        if failure == 'exception':
            raise DiagramError('renderer_cleanup_failed')
        return OwnedRun('render_timeout', None, 123, 0.1, False if failure == 'false' else None)
    monkeypatch.setattr(service, '_run_owned', run)
    with pytest.raises(DiagramError) as exc:
        service.render_diagram(SOURCES['flowchart'])
    assert exc.value.code == 'renderer_cleanup_failed' and service._disabled.is_set()
    monkeypatch.setattr(service, '_run_owned', lambda *a, **k: pytest.fail('worker restarted after uncertain cleanup'))
    with pytest.raises(DiagramError) as blocked:
        service.render_diagram(SOURCES['flowchart'])
    assert blocked.value.code == 'renderer_cleanup_failed'
    assert_clean()


def test_receiver_deadline_before_supervisor_deadline_is_timeout(monkeypatch):
    original_thread = threading.Thread
    created = []
    def create(**kwargs):
        thread = original_thread(**kwargs)
        created.append(thread)
        return thread
    monkeypatch.setattr(service.threading, 'Thread', create)
    def expired(*args, **kwargs):
        raise TimeoutError('synthetic receiver deadline first')
    monkeypatch.setattr(service, 'receive_message', expired)
    def run(command, **kwargs):
        with socket.create_connection(('127.0.0.1', int(command[2])), timeout=1) as connection:
            connection.settimeout(0.1)
            connection.sendall(channel.MAGIC + bytes.fromhex(command[3]))
            request = channel.receive_message(connection, time.monotonic() + 1)
            assert request['source'] == SOURCES['flowchart']
            created[0].join(1.0)
            assert not created[0].is_alive()
            assert kwargs['cancelled']() is True
        return OwnedRun('render_cancelled', None, 123, 0.02, True)
    monkeypatch.setattr(service, '_run_owned', run)
    with pytest.raises(DiagramError) as exc:
        service.render_diagram(SOURCES['flowchart'])
    assert exc.value.code == 'render_timeout'
    assert_clean()


def test_last_recv_cannot_return_late_data(monkeypatch):
    ticks = iter([0.0, 2.0])
    monkeypatch.setattr(channel, 'time', SimpleNamespace(monotonic=lambda: next(ticks)))
    with pytest.raises(TimeoutError):
        channel.read_exact(Wire(b'abcd'), 4, 1.0)


def test_channel_does_not_send_nonfinite_numbers():
    output = Wire()
    with pytest.raises(ValueError):
        channel.send_message(output, {'value': float('nan')})
    assert output.sent == b''


def test_hidden_worker_shell_is_local_and_has_no_agent_api():
    html = Path(service.__file__).parent / 'webui/static/diagram_host.html'
    text = html.read_text(encoding='utf-8')
    assert 'diagram_renderer.js' in text and 'app.js' not in text and 'pywebview.api' not in text
    assert "connect-src 'none'" in text and "default-src 'none'" in text
    worker = Path(service.__file__).with_name('diagram_worker.py').read_text(encoding='utf-8')
    assert 'js_api=None' in worker and 'hidden=True' in worker
    assert 'dotenv' not in worker.replace('or dotenv loading', '')
    assert 'configure_runtime' not in worker
