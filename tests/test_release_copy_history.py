import json
from pathlib import Path
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui import manual_history
from office_agent.paged_documents import CandidatePublisher
from office_agent.writes import GuardError, digest
from tests.test_office_canvas import fixtures


@pytest.fixture
def copies(tmp_path, monkeypatch):
    from office_agent import fidelity
    from office_agent.webui import office_canvas
    render = lambda *a, **kw: {'inputs_unchanged': True, 'render': {'status': 'different'}}
    monkeypatch.setattr(fidelity, 'compare_documents', render)
    monkeypatch.setattr(office_canvas, 'compare_documents', render)
    xlsx, pptx = fixtures(tmp_path)
    docx = tmp_path / '手工实验报告.docx'
    doc = Document();doc.add_paragraph('实验于2025年完成，温度120℃');doc.save(docx)
    app = WebviewApp()
    api = FloatingApi(app, settings_root=tmp_path)
    receipts = []
    for source in (docx, xlsx, pptx):
        assert api._floating.admit(str(source))['ok']
        assert api._floating.read(api._floating.read_token, True)['ok']
        version = digest(source)
        if source == docx:
            import io
            candidate = Document(source);candidate.paragraphs[0].runs[0].text = '实验于2026年完成，温度120℃'
            stream = io.BytesIO();candidate.save(stream)
            publisher = CandidatePublisher(app.editor, source, version)
            plan = publisher.plan(stream.getvalue(), document_id=api._floating.document['id'], draft_revision=1)
            receipt = publisher.execute(plan['id'], document_id=plan['document_id'], draft_revision=1, confirmed=True)
        else:
            service = api._floating.canvas
            opened = service.open(api._floating.document['id'], version)
            item = opened['view']['item_id']
            operation = {'item_id': item, 'target': 'B2' if source == xlsx else 'run-0',
                         'action': 'number' if source == xlsx else 'text', 'value': '125' if source == xlsx else '复核汇报',
                         'expected': '120' if source == xlsx else '实验汇报'}
            service.apply(service.document_id, 0, operation)
            public = service.plan(service.document_id, 1)['plan']
            plan = service.plans.pop(public['id'])
            service._publish(plan, service.source_guard, app.editor, api._floating.epoch)
            receipt = api._floating.outputs[plan['id']]
            service.discard(service.document_id, 1, True)
        assert digest(source) == version
        receipts.append(receipt)
    yield api, receipts
    api._floating.reset()


def test_three_formats_reconstruct_idempotently_and_open_with_fresh_consent(copies, tmp_path):
    api, receipts = copies
    before = {path: path.read_bytes() for path in tmp_path.rglob('audit.json')}
    records = api.list_sessions()['sessions']
    manual = [row for row in records if row.get('record_kind') == 'manual_copy']
    assert len(manual) == 3 and {row['format'] for row in manual} == {'DOCX', 'XLSX', 'PPTX'}
    assert {row['transaction_id'] for row in manual} == {row['id'] for row in receipts}
    assert api.list_sessions()['sessions'] == records
    assert {path: path.read_bytes() for path in tmp_path.rglob('audit.json')} == before
    row = next(row for row in manual if row['format'] == 'DOCX')
    detail = api.load_session(row['session_dir'])
    assert detail['readonly'] and '当前摘要一致' in detail['summary']
    assert not any('plan_id' in event or 'confirmed' in event for event in detail['events'])
    assert api.floating_select_history_output(row['transaction_id'])['ok']
    assert api._app.path == Path(row['output']) and not api._floating.document['ready']
    assert api._floating.runtime is None and api._app.pending_id is None
    assert api._app.last_verified_id is None and api._floating.outputs == {}


def test_rollback_missing_changed_and_corrupt_audit_do_not_claim_current_success(copies):
    api, receipts = copies
    first, second, third = receipts
    api._app.editor.rollback(first['id'], confirmed=True)
    assert manual_history.resolve(api._app.editor, first['id'])['status'] == 'rolled_back'
    assert not api.floating_select_history_output(first['id'])['ok']
    Path(second['output']).unlink()
    assert '不存在' in api.load_session('manual:' + second['id'])['summary']
    assert not api.floating_select_history_output(second['id'])['ok']
    Path(third['output']).write_bytes(b'changed')
    assert '外部修改' in api.load_session('manual:' + third['id'])['summary']
    assert not api.floating_select_history_output(third['id'])['ok']
    artifact = api._app.editor.audit / third['id'] / 'audit.json'
    artifact.write_text('{broken', encoding='utf8')
    listing = api.list_sessions()
    assert '无法安全读取' in listing['warning']
    assert Path(third['output']).read_bytes() == b'changed'
    assert not api.load_session('manual:' + third['id'])['ok']


def test_audit_failure_states_and_path_boundaries(copies, tmp_path):
    api, receipts = copies
    receipt = receipts[0]
    artifact = api._app.editor.audit / receipt['id'] / 'audit.json'
    original = json.loads(artifact.read_text('utf8'))
    for state in ('failed', 'uncertain', 'rolling_back'):
        artifact.write_text(json.dumps({**original, 'status': state}), encoding='utf8')
        assert manual_history.resolve(api._app.editor, receipt['id'])['status'] == state
        assert not api.floating_select_history_output(receipt['id'])['ok']
    artifact.write_text(json.dumps({**original, 'output': str(tmp_path.parent / 'outside.docx')}), encoding='utf8')
    assert not api.floating_select_history_output(receipt['id'])['ok']
    assert not api.floating_select_history_output('../outside')['ok']
    artifact.write_text(json.dumps({**original, 'id': '0' * 32}), encoding='utf8')
    assert not api.load_session('manual:' + receipt['id'])['ok']
    assert len(manual_history.listing(api._app.editor)['records']) == 2


def test_no_index_write_and_bounded_audit_read(copies, monkeypatch):
    api, receipts = copies
    def fail(*args, **kwargs):
        raise AssertionError('history must not persist a second success index')
    monkeypatch.setattr(Path, 'write_text', fail)
    assert len(api.list_sessions()['sessions']) == 3
    assert api.load_session('manual:' + receipts[0]['id'])['ok']
    monkeypatch.setattr(manual_history, 'LIMIT', 8)
    result = api.list_sessions()
    assert result['sessions'] == [] and '无法安全读取' in result['warning']


def test_cancelled_publication_does_not_add_success_record(copies):
    api, receipts = copies
    service = api._floating.canvas
    service.apply(service.document_id, service.revision, {'item_id': 'slide-0', 'target': 'run-0',
                  'action': 'text', 'value': '尚未发布', 'expected': '复核汇报'})
    plan = service.plan(service.document_id, service.revision)['plan']
    assert service.publish(service.document_id, service.revision, plan['id'], False)['cancelled']
    assert not (api._app.editor.audit / plan['id']).exists()
    assert len(api.list_sessions()['sessions']) == 3
    service.discard(service.document_id, service.revision, True)


def test_open_rejects_output_changed_during_admission(copies, monkeypatch):
    api, receipts = copies
    receipt = receipts[0]
    admit = api._floating._admit
    def changed(path):
        book = Document(path);book.add_paragraph('外部并发输入');book.save(path)
        return admit(path)
    monkeypatch.setattr(api._floating, '_admit', changed)
    result = api.floating_select_history_output(receipt['id'])
    assert not result['ok'] and '准入期间' in result['error']
    assert api._floating.document is None and api._floating.read_token is None


def test_late_history_open_does_not_restore_closed_document(copies, monkeypatch):
    api, receipts = copies
    check = manual_history.checked_output
    def close(editor, identifier):
        path = check(editor, identifier)
        api._floating.reset()
        return path
    monkeypatch.setattr(manual_history, 'checked_output', close)
    result = api.floating_select_history_output(receipts[0]['id'])
    assert not result['ok'] and '已切换' in result['error']
    assert api._floating.document is None and api._floating.read_token is None


def test_history_ui_lists_all_formats_and_reopens_without_restoring_grants(copies, tmp_path, monkeypatch):
    from tests.test_paged_browser import serve_static
    from tests.test_professional_browser import EDGE
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge for isolated history UI regression')
    from playwright.sync_api import sync_playwright
    api, receipts = copies
    monkeypatch.setenv('TEMP', str(tmp_path));monkeypatch.setenv('TMP', str(tmp_path))
    monkeypatch.setattr(api, 'canvas_preview', lambda *args: {'ok': False, 'error': '本项只验证历史入口，不进行实际渲染'})
    with serve_static() as url, sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1100, 'height': 740});page.set_default_timeout(15000)
            page.expose_function('testHost', lambda name, args: getattr(api, name)(*args))
            page.add_init_script("""if(window===window.top){
              window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};
              window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};
              document.addEventListener('DOMContentLoaded',()=>{const node=document.getElementById('workspace');node.hidden=false;node.inert=false;node.setAttribute('aria-hidden','false');});
            }""")
            page.route('**/*', lambda route: route.continue_() if route.request.url.startswith(url) else route.abort())
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            errors = [];page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto(url);page.locator('#canvas-items button').first.wait_for()
            page.get_by_label('更多', exact=True).click()
            page.get_by_role('button', name='历史记录', exact=True).click()
            page.wait_for_function('document.querySelectorAll("#history-list [data-session^=\\"manual:\\"]").length===3')
            assert not page.locator('#simple-history-chats').is_visible()
            assert page.locator('#simple-history-documents').get_attribute('aria-pressed') == 'true'
            target = page.locator('#history-list [data-session="manual:' + receipts[0]['id'] + '"]')
            assert target.locator('.history-item-title').text_content() == Path(receipts[0]['output']).name
            for theme in ('light', 'dark', 'highcontrast'):
                page.evaluate('(theme)=>document.documentElement.dataset.theme=theme', theme)
                for width in (1120, 820, 520):
                    page.set_viewport_size({'width': width, 'height': 760})
                    page.get_by_role('button', name='对话', exact=True).click()
                    assert not page.locator('#simple-document-records').is_visible()
                    assert page.locator('#simple-history-chats').is_visible()
                    page.get_by_role('button', name='文档任务', exact=True).click()
                    assert not page.locator('#simple-history-chats').is_visible()
                    assert target.is_visible()
                    assert target.locator('.history-item-state').evaluate('(node)=>node.scrollWidth<=node.clientWidth+1')
                    assert page.locator('#history-list').evaluate('(node)=>node.scrollWidth<=node.clientWidth+1')
                    assert page.locator('#simple-history-content').evaluate('(node)=>node.scrollWidth<=node.clientWidth+1')
                    assert target.evaluate('(node)=>node.querySelector(".history-item-state").getBoundingClientRect().top>=node.querySelector(".history-item-title").getBoundingClientRect().bottom')
            page.set_viewport_size({'width': 1100, 'height': 740})
            target.click();page.get_by_role('button', name='重新打开副本', exact=True).click()
            page.get_by_role('button', name='同意读取并打开', exact=True).wait_for()
            assert api._app.path == Path(receipts[0]['output']) and not api._floating.document['ready']
            assert api._floating.outputs == {} and api._app.pending_id is None and errors == []
        finally:
            browser.close()
