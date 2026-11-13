"""Temporary real file IO and fake typed renderer; no native/provider/Office proof."""
from pathlib import Path
import threading
from io import BytesIO
import hashlib
from PIL import Image
from office_agent.diagram_raster import RasterReceipt, normalize_png

import pytest

from office_agent.webui.bridge import Api
from office_agent.diagram_service import RenderedDiagram
from office_agent.diagram_process import OwnedRun
from office_agent.diagrams import validate_source

SOURCE = 'flowchart TD\nA --> B'
SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text>A</text></svg>'


class Auth:
    authenticated = True
    def status(self):
        return {'authenticated': self.authenticated}
    def logout(self):
        self.authenticated = False
        return {'ok': True}
    def login(self, *args):
        self.authenticated = True
        return {'ok': True, 'authenticated': True}
    def reset_accounts(self, confirmed=False):
        self.authenticated = not confirmed
        return {'ok': confirmed}


class App:
    busy = False
    def __init__(self, path):
        self.path = path
    def _build_model(self):
        pytest.fail('no provider during export')
    def _clear_private_state(self):
        self.path = None
    def select_file(self):
        return {'ok': True}
    def ask(self, *args):
        return {'ok': True}
    load_session = answer_clarification = execute = undo = ask


@pytest.fixture
def api(tmp_path):
    path = tmp_path / 'selected.docx'
    path.write_bytes(b'fingerprint-only NOT an Office document')
    value = Api(App(path), auth=Auth())
    value._diagram._render = lambda source, **kwargs: RenderedDiagram(
        validate_source(source, kwargs['expected_kind']), SVG, OwnedRun('process_completed', 0, 1, .1, True))
    return value


def preview(api):
    result = api.render_diagram_source('flowchart', SOURCE)
    assert result['ok']
    return result['metadata']['source_sha256'], result['metadata']['svg_sha256']


def plan(api):
    result = api.plan_diagram_save(*preview(api))
    assert result['ok']
    return result['plan']['plan_id']


def save(api):
    result = api.confirm_diagram_save(plan(api), True)
    assert result['ok'] and result['result']['saved']
    return result['result']


def test_current_native_png_plan_confirmation_and_rollback(api, monkeypatch):
    from office_agent.webui import diagram_export
    def host_raster(native):
        buffer = BytesIO()
        Image.new('RGB', (100, 100), 'red').save(buffer, format='PNG')
        data, w, h = normalize_png(buffer.getvalue())
        return RasterReceipt(native, data, hashlib.sha256(data).hexdigest(), w, h)
    monkeypatch.setattr(diagram_export, 'rasterize', host_raster)
    hashes = preview(api)
    response = api.plan_diagram_png_save(*hashes)
    assert response['ok'] and response['plan']['destination_basename'].endswith('.png')
    assert api.confirm_diagram_save(response['plan']['plan_id'], False)['reason'] == 'export_confirmation_required'
    plan = api.plan_diagram_png_save(*hashes)['plan']
    outcome = api.confirm_diagram_save(plan['plan_id'], True)['result']
    assert Path(outcome['path']).suffix == '.png' and Path(outcome['path']).exists()
    undo = api.plan_diagram_rollback(outcome['save_id'])['plan']['rollback_id']
    assert api.confirm_diagram_rollback(undo, True)['result']['removed'] is True
    assert not Path(outcome['path']).exists()


def test_authenticated_current_native_save_and_independent_rollback(api):
    source = api._app.path
    before = source.read_bytes()
    hashes = preview(api)
    result = api.plan_diagram_save(*hashes)
    assert result['ok'] and len(list(source.parent.iterdir())) == 1
    result['plan']['destination_basename'] = '../foreign.svg'
    saved = api.confirm_diagram_save(result['plan']['plan_id'], True)['result']
    output, audit = Path(saved['path']), Path(saved['audit_path'])
    assert output.is_file() and audit.is_file() and source.read_bytes() == before
    rollback = api.plan_diagram_rollback(saved['save_id'])['plan']['rollback_id']
    assert output.exists()
    rolled = api.confirm_diagram_rollback(rollback, True)['result']
    assert rolled['removed'] and rolled['audit_updated'] and not output.exists() and audit.exists()
    assert api._diagram_export._lock is api._lock


@pytest.mark.parametrize('value', [None, {}, '../file', '0'*64, 'A'*64])
def test_client_hashes_not_authority(api, value):
    hashes = preview(api)
    assert not api.plan_diagram_save(value, hashes[1])['ok']
    assert not api.plan_diagram_save(hashes[0], value)['ok']
    assert len(list(api._app.path.parent.iterdir())) == 1


@pytest.mark.parametrize('value', [False, None, 1, 'true'])
def test_literal_save_approval_consumed(api, value):
    token = plan(api)
    result = api.confirm_diagram_save(token, value)
    assert result == {'ok': False, 'reason': 'export_confirmation_required'}
    assert api.confirm_diagram_save(token, True)['reason'] == 'export_plan_invalid'
    assert len(list(api._app.path.parent.iterdir())) == 1


@pytest.mark.parametrize('action', ['cancel', 'new_preview', 'select', 'ask', 'session', 'clarify', 'execute', 'undo', 'logout', 'reset', 'path', 'busy', 'source'])
def test_pending_approval_fenced(api, action):
    token = plan(api)
    source = api._app.path
    if action == 'cancel':
        api.cancel_diagram()
    elif action == 'new_preview':
        preview(api)
    elif action == 'logout':
        api.logout()
    elif action == 'reset':
        api.reset_accounts(True)
    elif action == 'path':
        api._app.path = source.with_name('other.docx')
    elif action == 'busy':
        api._app.busy = True
    elif action == 'source':
        source.write_bytes(b'edited')
    else:
        name = {'select':'select_file', 'ask':'ask', 'session':'load_session', 'clarify':'answer_clarification', 'execute':'execute', 'undo':'undo'}[action]
        args = {'ask':('x',), 'load_session':('x',), 'answer_clarification':('id', {})}.get(name, ())
        getattr(api, name)(*args)
    assert not api.confirm_diagram_save(token, True)['ok']
    assert len(list(source.parent.iterdir())) == 1


def test_logout_relogin_same_path_never_regrants(api):
    token = plan(api)
    path = api._app.path
    api.logout()
    api.login('x', 'y')
    api._app.path = path
    preview(api)
    assert not api.confirm_diagram_save(token, True)['ok']


def test_saved_history_does_not_regrant_after_invalidation(api):
    saved = save(api)
    output = Path(saved['path'])
    token = api.plan_diagram_rollback(saved['save_id'])['plan']['rollback_id']
    api.cancel_diagram()
    preview(api)
    assert not api.confirm_diagram_rollback(token, True)['ok']
    assert not api.plan_diagram_rollback(saved['save_id'])['ok']
    assert output.exists()


def test_cleanup_uncertainty_latches_across_new_preview_and_login(api, monkeypatch):
    token = plan(api)
    service = api._diagram_export._service
    original = service._create
    def failing(path, data, created):
        if path.suffix == '.json':
            created[0].path.write_bytes(b'external edit')
            raise OSError('secret failure')
        return original(path, data, created)
    monkeypatch.setattr(service, '_create', failing)
    assert api.confirm_diagram_save(token, True)['reason'] == 'export_cleanup_uncertain'
    hashes = preview(api)
    assert api.plan_diagram_save(*hashes)['reason'] == 'export_service_disabled'
    path = api._app.path
    api.logout(); api.login('x', 'y'); api._app.path = path
    assert api.plan_diagram_save(*preview(api))['reason'] == 'export_service_disabled'


def test_partial_rollback_result_preserved_and_facade_latched(api, monkeypatch):
    saved = save(api)
    token = api.plan_diagram_rollback(saved['save_id'])['plan']['rollback_id']
    def fail(record):
        raise OSError('secret path')
    monkeypatch.setattr(api._diagram_export._service, '_update_audit', fail)
    response = api.confirm_diagram_rollback(token, True)
    assert response['ok'] and response['result']['removed'] is True
    assert response['result']['audit_updated'] is False and response['result']['service_disabled']
    assert not Path(saved['path']).exists()
    assert api.plan_diagram_save(*preview(api))['reason'] == 'export_service_disabled'


def test_unexpected_error_is_sanitized_and_uncertain(api, monkeypatch):
    token = plan(api)
    def fail(*args):
        raise RuntimeError('PRIVATE RAW TEXT')
    monkeypatch.setattr(api._diagram_export._service, 'confirm_save', fail)
    assert api.confirm_diagram_save(token, True) == {'ok': False, 'reason': 'export_failed', 'outcome_uncertain': True}
    assert api._diagram_export._disabled


def test_cancellation_serializes_after_admitted_sync_write(api, monkeypatch):
    token = plan(api)
    service = api._diagram_export._service
    entered, release, cancelled = threading.Event(), threading.Event(), threading.Event()
    original = service._create
    def blocked(*args):
        entered.set()
        assert release.wait(3)
        return original(*args)
    monkeypatch.setattr(service, '_create', blocked)
    results = []
    writer = threading.Thread(target=lambda: results.append(api.confirm_diagram_save(token, True)))
    canceller = threading.Thread(target=lambda: (api.cancel_diagram(), cancelled.set()))
    writer.start()
    try:
        assert entered.wait(3)
        canceller.start()
        assert not cancelled.wait(.05)
    finally:
        release.set()
        writer.join(4)
        if canceller.ident is not None:
            canceller.join(4)
    assert not writer.is_alive() and not canceller.is_alive()
    assert results[0]['ok'] and Path(results[0]['result']['path']).exists()
    assert cancelled.is_set() and api._diagram_export._service is None


@pytest.mark.parametrize('confirmed', [False, None, 1, 'true'])
def test_independent_literal_rollback_approval(api, confirmed):
    saved = save(api)
    token = api.plan_diagram_rollback(saved['save_id'])['plan']['rollback_id']
    assert api.confirm_diagram_rollback(token, confirmed)['reason'] == 'export_confirmation_required'
    assert api.confirm_diagram_rollback(token, True)['reason'] == 'export_plan_invalid'
    assert Path(saved['path']).exists()


def test_unauthenticated_never_plans_or_writes(api):
    hashes = preview(api)
    api._auth.authenticated = False
    assert api.plan_diagram_save(*hashes) == {'ok': False, 'reason': 'not_authenticated'}
    assert len(list(api._app.path.parent.iterdir())) == 1


def test_ambiguous_unlink_latches_even_without_core_flag(api, monkeypatch):
    saved = save(api)
    token = api.plan_diagram_rollback(saved['save_id'])['plan']['rollback_id']
    target = Path(saved['path'])
    unlink = Path.unlink
    def fail(path, *args, **kwargs):
        result = unlink(path, *args, **kwargs)
        if path == target:
            raise OSError('lost acknowledgment')
        return result
    monkeypatch.setattr(Path, 'unlink', fail)
    response = api.confirm_diagram_rollback(token, True)
    assert response['ok'] and response['result']['removed'] is None
    assert response['result']['code'] == 'export_rollback_uncertain'
    assert api._diagram_export._disabled
    assert api.plan_diagram_save(*preview(api))['reason'] == 'export_service_disabled'
