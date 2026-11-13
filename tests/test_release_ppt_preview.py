import json
import threading
from types import SimpleNamespace

import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui import office_canvas
from office_agent.writes import GuardError, digest
from tests.test_office_canvas import fixtures


@pytest.fixture
def preview_service(tmp_path, monkeypatch):
    source = fixtures(tmp_path)[1]
    api = FloatingApi(WebviewApp(), settings_root=tmp_path)
    assert api._floating.admit(str(source))['ok']
    assert api._floating.read(api._floating.read_token, True)['ok']
    service = api._floating.canvas
    service.open(api._floating.document['id'], digest(source))
    monkeypatch.setattr(office_canvas, 'runtime_path', lambda: tmp_path)
    monkeypatch.setattr(office_canvas, 'environment_identity', lambda runtime: {'complete': True, 'test': 'identity'})
    monkeypatch.setattr(office_canvas, 'implementation_identity', lambda: 'test-implementation')
    monkeypatch.setattr(office_canvas, '_temporary_root', lambda: tmp_path)
    monkeypatch.setattr(office_canvas, '_convert', lambda source, target, *args: target / 'page.pdf')
    def render(command, timeout, runs):
        folder = command[-1]
        from pathlib import Path
        (Path(folder) / 'page.json').write_text(json.dumps({'pages': 2, 'page': 1}), encoding='utf8')
        (Path(folder) / 'page.png').write_bytes(b'controlled preview bytes')
    monkeypatch.setattr(office_canvas, '_run', render)
    yield service, source
    api._floating.reset()


def test_edit_during_snapshot_preview_invalidates_old_response(preview_service, monkeypatch):
    service, source = preview_service
    original = source.read_bytes()
    started, release = threading.Event(), threading.Event()
    def convert(source, target, *args):
        started.set()
        if not release.wait(5):
            raise RuntimeError('bounded preview test timed out')
        return target / 'page.pdf'
    monkeypatch.setattr(office_canvas, '_convert', convert)
    errors = []
    def pending():
        try:
            service.preview(service.document_id, 0, 'slide-0')
        except Exception as error:
            errors.append(error)
    thread = threading.Thread(target=pending)
    thread.start()
    try:
        assert started.wait(5) and service.previewing
        service.apply(service.document_id, 0, {'item_id': 'slide-0', 'target': 'run-0', 'action': 'text', 'value': '新草稿', 'expected': '实验汇报'})
        assert service.revision == 1
        with pytest.raises(GuardError):
            service.plan(service.document_id, 1)
    finally:
        release.set();thread.join(6)
    assert not thread.is_alive() and len(errors) == 1 and isinstance(errors[0], GuardError)
    assert service.preview_cache is None and not service.previewing and source.read_bytes() == original
    result = service.preview(service.document_id, 1, 'slide-0')
    assert result['draft_revision'] == 1 and service.preview_cache is not None


def test_cached_preview_rebinds_equal_bytes_to_current_revision(preview_service, monkeypatch):
    service, source = preview_service
    first = service.preview(service.document_id, 0, 'slide-0')
    cached = service.preview_cache
    service.apply(service.document_id, 0, {'item_id': 'slide-0', 'target': 'run-0', 'action': 'text', 'value': '新草稿', 'expected': '实验汇报'})
    service.undo(service.document_id, 1)
    assert service.data == service.original
    service.preview_cache = cached
    def forbidden(*args):
        raise AssertionError('equal identity must reuse the cached render')
    monkeypatch.setattr(office_canvas, '_convert', forbidden)
    result = service.preview(service.document_id, 2, 'slide-0')
    assert result['draft_revision'] == 2 and result['image'] == first['image']
    assert not service.previewing
