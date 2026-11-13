import hashlib
import json
from pathlib import Path
import threading

from docx import Document
from openpyxl import Workbook
import pytest

from office_agent.local_session import LocalSession
from office_agent.webui.bridge import Api, WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui.document_open import parse_payload


def make(tmp_path, suffix='.docx'):
    path = tmp_path / ('拖放样例' + suffix)
    if suffix == '.docx':
        book = Document()
        book.add_paragraph('ORB-DROP-唯一正文。温度120 ℃，压力3.5 MPa。')
        book.add_paragraph('第二段保持原样。')
        book.add_table(rows=1, cols=2).cell(0, 0).text = '原表保持'
        book.save(path)
    else:
        book = Workbook()
        sheet = book.active
        sheet.title = '实验'
        sheet['A1'], sheet['B2'], sheet['C2'] = 'ORB-DROP-唯一表格', 120, '=B2*2'
        book.create_sheet('对照')['A1'] = '第二张表'
        book.save(path)
        book.close()
    app = WebviewApp()
    api = FloatingApi(app, auth=LocalSession(tmp_path / 'appdata'), settings_root=tmp_path)
    return api, app, path


def wait(service):
    worker = service._open_worker
    if worker:
        worker.join(35)
        if worker.is_alive():
            service._open_cancel.set()
            worker.join(5)
        assert not worker.is_alive()


def send(api, text='精简第一段，保留数字和单位', client='one'):
    state = api.floating_snapshot()
    document = state['document']
    return api.floating_submit(state['session_id'], client, text,
                               document['open_request_id'], document['id'], document['version'])


@pytest.mark.parametrize('suffix', ['.docx', '.xlsx'])
def test_native_callback_opens_real_content_without_picker_or_model(tmp_path, monkeypatch, suffix):
    api, app, path = make(tmp_path, suffix)
    original = hashlib.sha256(path.read_bytes()).hexdigest()
    calls = []
    monkeypatch.setattr(app, '_start_task', lambda *a, **kw: calls.append(a))
    api._floating.dropped((str(path),))
    wait(api._floating)
    state = api.floating_snapshot()
    assert state['ok'] and state['document']['ready']
    assert state['open_request']['phase'] in {'ready', 'partial_ready'}
    assert state['document']['source_origin'] == 'native_drop'
    assert 'ORB-DROP' in json.dumps(state['view'], ensure_ascii=False)
    assert state['document']['version'] == original == hashlib.sha256(path.read_bytes()).hexdigest()
    assert app._read_authorized_path == path and not calls
    stages = [record['stage'] for record in state['open_trace']]
    assert stages.index('admission_succeeded') < stages.index('local_read_authorized') < stages.index('parse_started') < stages.index('document_content_ready')
    trace = (tmp_path / 'appdata/diagnostics/open-pipeline.json').read_text(encoding='utf-8')
    assert str(path) not in trace and 'ORB-DROP' not in trace and '温度120 ℃' not in trace
    if suffix == '.xlsx':
        assert len(state['view']['sheets']) == 2
        assert '=B2*2' in json.dumps(state['view'])


def test_parsing_messages_persist_deduplicate_and_start_once(tmp_path, monkeypatch):
    api, app, path = make(tmp_path)
    entered, release = threading.Event(), threading.Event()
    started = []
    def gated(source, cancelled, **options):
        entered.set()
        assert release.wait(20)
        return parse_payload(source)
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', gated)
    monkeypatch.setattr(app, '_start_task', lambda *a, **kw: started.append([row['text'] for row in api._floating.runtime.messages]))
    api._floating.dropped((str(path),))
    try:
        assert entered.wait(10)
        first = send(api)
        assert first['ok'] and send(api)['duplicate']
        assert send(api, '第一段语气自然一些', 'two')['ok']
        assert not started and not api._floating.document['ready']
        assert len(list((path.parent/'.office_agent/inbox').rglob('*.json'))) == 2
    finally:
        release.set()
        wait(api._floating)
    assert len(started) == 1 and len(started[0]) == 2
    assert api._floating.document['ready']


def test_failed_parse_keeps_requirements_and_retry_does_not_reselect(tmp_path, monkeypatch):
    api, app, path = make(tmp_path)
    entered, release = threading.Event(), threading.Event()
    starts = []
    def failed(source, cancelled, **options):
        entered.set()
        assert release.wait(20)
        raise ValueError('解析失败，可重试')
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', failed)
    monkeypatch.setattr(app, '_start_task', lambda *a, **kw: starts.append(a))
    api._floating.dropped((str(path),))
    try:
        assert entered.wait(10)
        assert send(api)['ok']
    finally:
        release.set()
        wait(api._floating)
    assert not starts and api.floating_snapshot()['open_request']['phase'] == 'failed'
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', lambda source, cancelled, **kw: parse_payload(source))
    assert api.floating_admit_drop()['ok']
    wait(api._floating)
    assert api._floating.document['ready'] and len(starts) == 1
    assert len(api._floating.runtime.messages) == 1


def test_busy_and_forged_drop_never_replace_document(tmp_path):
    api, app, path = make(tmp_path)
    assert not api.floating_admit_drop()['ok']
    app.busy = True
    api._floating.dropped((str(path),))
    assert api._floating.document is None and app.path is None
    state = api.floating_snapshot()
    assert any(row['event']['type'] == 'admission_failed' for row in state['events'])
    assert api._floating.local_read_grant is None


def test_default_session_does_not_read_or_modify_accounts(tmp_path, monkeypatch):
    from office_agent.auth import LocalAuth
    monkeypatch.setattr(LocalAuth, 'status', lambda self: (_ for _ in ()).throw(AssertionError('old account store used')))
    api = Api(WebviewApp(), settings_root=tmp_path)
    assert api.auth_status()['mode'] == 'single_user'
    assert api.auth_status()['authenticated']
    assert not api.register('unused', 'unused')['ok']
    assert api.logout()['ok'] and api.auth_status()['authenticated']
    assert LocalSession(tmp_path).account_key() == LocalSession(tmp_path).account_key()


@pytest.mark.parametrize('action', ['clear', 'stop', 'change'])
def test_late_parse_cannot_restore_cleared_cancelled_or_changed_source(tmp_path, monkeypatch, action):
    api, app, path = make(tmp_path)
    entered, release = threading.Event(), threading.Event()
    starts = []
    def gated(source, cancelled, **options):
        payload = parse_payload(source)
        entered.set()
        assert release.wait(20)
        return payload
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', gated)
    monkeypatch.setattr(app, '_start_task', lambda *a, **kw: starts.append(a))
    api._floating.dropped((str(path),))
    try:
        assert entered.wait(10)
        assert send(api)['ok']
        if action == 'clear':
            assert api.logout()['ok']
        elif action == 'stop':
            assert api.floating_stop()['ok']
        else:
            doc = Document(path)
            doc.paragraphs[0].text = '外部修改后的新内容'
            doc.save(path)
    finally:
        release.set()
        wait(api._floating)
    assert not starts
    if action == 'clear':
        assert api._floating.document is None and app.path is None
    else:
        assert not api._floating.document['ready']
        assert api.floating_snapshot()['open_request']['phase'] == ('cancelled' if action == 'stop' else 'failed')


@pytest.mark.parametrize('clear', [False, True])
def test_late_admission_preserves_next_drop_but_not_cleared_queue(tmp_path, monkeypatch, clear):
    api, app, first = make(tmp_path)
    second = tmp_path / '第二份.docx'
    doc = Document()
    doc.add_paragraph('SECOND-DROP-第二份正文')
    doc.save(second)
    entered, release = threading.Event(), threading.Event()
    admit = app.admit_file
    def gated(candidate):
        if Path(candidate) == first:
            entered.set()
            assert release.wait(20)
        return admit(candidate)
    monkeypatch.setattr(app, 'admit_file', gated)
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', lambda source, cancelled, **kw: parse_payload(source))
    api._floating.dropped((str(first),))
    try:
        assert entered.wait(10)
        api._floating.dropped((str(second),))
        if clear:
            assert api.logout()['ok']
    finally:
        release.set()
        wait(api._floating)
    if clear:
        assert api._floating.document is None and app.path is None
    else:
        assert app.path == second and api._floating.document['ready']
        assert 'SECOND-DROP' in json.dumps(api._floating.view, ensure_ascii=False)
        assert 'ORB-DROP' not in json.dumps(api._floating.view, ensure_ascii=False)


def test_forged_context_cannot_submit_to_native_document(tmp_path, monkeypatch):
    api, app, path = make(tmp_path)
    monkeypatch.setattr('office_agent.webui.floating_session.read_local_view', lambda source, cancelled, **kw: parse_payload(source))
    api._floating.dropped((str(path),))
    wait(api._floating)
    document = api._floating.document
    result = api.floating_submit(None, 'forged', '修改第一段', 'invented', document['id'], document['version'])
    assert not result['ok'] and api._floating.runtime is None and not app.busy
