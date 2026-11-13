from pathlib import Path

import pytest
from docx import Document

from tests.test_professional_browser import STATIC, _browser
from tests.test_range_browser import setup_range
from office_agent.writes import digest


def setup(page, api, app, calls, receipts):
    setup_range(page, api, app, calls, receipts)
    page.evaluate('window.OfficeAuth.ensure = async () => false')
    page.add_script_tag(path=str(STATIC / 'app.js'))
    page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
    page.add_style_tag(content=(STATIC / 'workspace_layout.css').read_text(encoding='utf8'))
    page.add_script_tag(path=str(STATIC / 'workspace_layout.js'))


@pytest.mark.parametrize('width', [1440, 1280, 1100, 520])
@pytest.mark.parametrize('theme', ['light', 'dark', 'highcontrast'])
def test_workspace_navigation_drafts_scroll_and_unique_bindings(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup(page, api, app, calls, receipts)
        page.set_viewport_size({'width': width, 'height': 800 if width != 1440 else 900})
        page.evaluate('theme => document.documentElement.dataset.theme = theme', theme)
        page.locator('#workspace-view-task').click()
        page.locator('#doc-task-request').fill('统一术语，保留20 ms及否定条件')
        page.locator('#nav-resources').click()
        assert page.locator('#workspace-resources').is_visible()
        page.locator('#nav-review').click()
        assert page.locator('#workspace-review').is_visible()
        page.locator('#workspace-continuation > summary').click()
        page.locator('#doc-task-return-panel > summary').click()
        page.locator('#doc-task-return-compare').scroll_into_view_if_needed()
        box = page.locator('#workspace-review').bounding_box()
        page.mouse.move(box['x'] + box['width'] - 20, box['y'] + 80)
        page.mouse.wheel(0, -10000)
        page.wait_for_timeout(100)
        assert page.locator('#workspace-review').evaluate('node => node.scrollTop < 100')
        page.locator('#nav-work').click()
        page.locator('#workspace-view-task').click()
        assert page.locator('#doc-task-request').input_value() == '统一术语，保留20 ms及否定条件'
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        assert page.locator('#nav-resources').evaluate('node => node.getBoundingClientRect().right <= node.closest(".sidebar").getBoundingClientRect().right')
        assert page.locator('#nav-settings').evaluate('node => node.getBoundingClientRect().top >= innerHeight * .8')
        assert page.evaluate('new Set(Array.from(document.querySelectorAll("[id]"), x=>x.id)).size === document.querySelectorAll("[id]").length')
        assert page.locator('nav [aria-current="page"]').count() == 1
        assert not calls and not app.editor.audit.exists()
        page.locator('#workspace-view-document').focus()
        page.keyboard.press('Tab')
        assert page.locator('#workspace-view-task').evaluate('node => document.activeElement === node')
    finally:
        fixture.close()


def prepare_range(page):
    page.locator('#range-load').click()
    page.wait_for_function("document.getElementById('range-target').options.length > 1")
    page.locator('#range-target').select_option(index=1)
    page.locator('#range-original').evaluate('(node) => {node.focus(); node.setSelectionRange(7, 13);}')
    page.locator('#workspace-view-task').click()
    page.locator('#workspace-mode-range').click()
    page.locator('#range-text').fill('second')
    page.locator('#range-reason').fill('仅修改第二次重复文字')
    page.locator('#range-add').click()
    page.locator('#workspace-see-range-review').click()


def test_migrated_controls_review_filter_locate_cancel_and_real_copy(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup(page, api, app, calls, receipts)
        source = digest(app.path)
        prepare_range(page)
        assert not app.editor.audit.exists()
        page.locator('#workspace-review-filter').select_option('pending')
        assert not page.locator('#range-queue > div').is_visible()
        page.locator('#workspace-review-filter').select_option('adopted')
        assert page.locator('#range-queue > div').is_visible()
        page.get_by_role('button', name='定位原文', exact=True).click()
        assert page.locator('#range-original').input_value().startswith('repeat repeat')
        page.locator('#nav-review').click()
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        page.evaluate('window.approved = false')
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-status').textContent.includes('已取消')")
        assert not app.editor.audit.exists()
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        page.evaluate('window.approved = true')
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已保存副本')")
        output = Path(next(reply['result']['output'] for method, reply in receipts if method == 'confirm_professional_edit' and reply.get('result', {}).get('status') == 'saved'))
        assert 'repeat second' in '\n'.join(p.text for p in Document(output).paragraphs)
        assert digest(app.path) == source
        page.locator('#range-undo').click()
        page.wait_for_function("document.getElementById('range-output').textContent.includes('已撤销')")
        assert not output.exists() and digest(app.path) == source
    finally:
        fixture.close()


def test_workspace_logout_rejects_late_read_and_clears_identity(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup(page, api, app, calls, receipts)
        page.evaluate('window.pywebview.api = {read_edit_ranges: () => new Promise(resolve => window.finishRead = resolve)}')
        page.locator('#range-load').click()
        assert page.locator('#range-load').is_disabled()
        response = api.read_edit_ranges(0)
        page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate('(reply) => window.finishRead(reply)', response)
        page.wait_for_timeout(100)
        assert page.locator('#range-original').input_value() == ''
        assert page.locator('#range-target option').count() == 0
        assert page.locator('#range-confirm').is_disabled()
        assert '重新读取' in page.locator('#workspace-document-identity').inner_text()
    finally:
        fixture.close()


def test_workspace_partial_task_is_not_complete_and_return_chain_is_filled(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup(page, api, app, calls, receipts)
        document = Document()
        for index in range(31):
            document.add_paragraph(f'第{index + 1}项旧术语，20 ms。')
        document.save(app.path)
        allowed = {'create_document_task', 'plan_document_task_group', 'confirm_document_task_group'}
        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('workspaceTaskHost', invoke)
        page.evaluate('() => { window.pywebview.api = new Proxy({}, {get: (_, name) => (...args) => window.workspaceTaskHost(name, args)}); }')
        page.locator('#doc-task-request').fill('将全文所有“旧术语”替换为“规范术语”，保留数字、单位、公式和引文')
        page.locator('#doc-task-create').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('0/3')")
        page.locator('#workspace-see-task-review').click()
        page.locator('#doc-task-plan').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        assert page.locator('#workspace-readable-plan article').count() == 12
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('1/3')")
        result = next(reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group')
        assert not result['deterministic_complete']
        assert '受限项' in page.locator('#doc-task-workflow').inner_text()
        assert page.locator('#doc-task-confirm').is_disabled()
        page.locator('#doc-task-prepare-return').click()
        assert page.locator('#workspace-continuation').get_attribute('open') is not None
        assert page.locator('#doc-task-return-baseline').input_value() == app.path.name
        assert page.locator('#doc-task-return-delivered').input_value() == result['current_name']
        assert page.locator('#doc-task-return-receipts').input_value() == result['receipts'][0]
        assert calls.count('confirm_document_task_group') == 1
    finally:
        fixture.close()


def test_workspace_source_conflict_and_declined_read_do_not_write(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup(page, api, app, calls, receipts)
        page.evaluate('''() => {
          window.savedApi = window.pywebview.api;
          window.pywebview.api = {read_edit_ranges: async () => ({ok:false, reason:'professional_source_changed'})};
        }''')
        page.locator('#range-load').click()
        page.wait_for_function("document.getElementById('range-status').textContent.includes('重新读取')")
        assert page.locator('#range-target option').count() == 0
        page.evaluate('() => { window.pywebview.api = window.savedApi; }')
        prepare_range(page)
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        document = Document(app.path)
        document.add_paragraph('External change')
        document.save(app.path)
        page.locator('#range-confirm').click()
        page.wait_for_function("document.getElementById('range-status').textContent.includes('源文件已改变')")
        assert not app.editor.audit.exists()
        assert not list(tmp_path.glob('*.agent-*.docx'))
    finally:
        fixture.close()


def test_all_existing_frontend_modules_initialize_after_dom_migration(tmp_path, monkeypatch):
    import re
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        setup(page, api, app, calls, receipts)
        loaded = {'auth.js', 'app.js', 'professional.js', 'range_edits.js', 'document_tasks.js', 'workspace_layout.js'}
        for name in re.findall(r'<script src="([^"]+)"', (STATIC / 'index.html').read_text(encoding='utf8')):
            if name not in loaded:
                page.add_script_tag(path=str(STATIC / name))
        page.locator('#nav-account').click()
        assert page.locator('#settings-pane-account').is_visible()
        page.keyboard.press('Escape')
        assert not page.locator('#settings-dialog').is_visible()
        assert page.locator('#nav-account').evaluate('node => node === document.activeElement')
        assert not errors
    finally:
        fixture.close()


def test_production_script_order_has_no_missing_bindings_or_automatic_calls(tmp_path, monkeypatch):
    import re
    fixture = _browser(tmp_path, monkeypatch)
    try:
        first, api, app, calls, receipts = next(fixture)
        page = first.context.browser.new_page(viewport={'width': 520, 'height': 800})
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.route('**/*', lambda route: route.abort())
        html = (STATIC / 'index.html').read_text(encoding='utf8')
        page.set_content(re.sub(r'<script\b[^>]*>.*?</script>', '', html, flags=re.S))
        for name in ['theme.css', 'diagram_preview.css', 'workspace_layout.css', 'workbench_ui.css']:
            page.add_style_tag(content=(STATIC / name).read_text(encoding='utf8'))
        page.evaluate('''() => {
          document.getElementById('workspace').hidden = false;
          document.getElementById('workspace').inert = false;
          document.getElementById('workspace').setAttribute('aria-hidden', 'false');
          document.getElementById('auth-screen').hidden = true;
          window.OfficeAuth = {isAuthenticated:()=>true,ensure:async()=>false,wrapApi:x=>x};
          window.unexpectedCalls = [];
          window.pywebview = {api:new Proxy({}, {get:(_, method)=>()=>{window.unexpectedCalls.push(method); throw new Error('automatic_call');}})};
        }''')
        for name in re.findall(r'<script src="([^"]+)"', html):
            if name != 'auth.js':
                page.add_script_tag(path=str(STATIC / name))
        assert not errors, errors
        page.locator('#simple-document-menu > summary').click()
        page.locator('#workspace-view-task').click()
        page.keyboard.press('Escape')
        page.locator('#workspace-quick-actions').get_by_role('button', name='统一术语').click()
        assert '规范术语' in page.locator('#doc-task-request').input_value()
        page.evaluate('() => { bind(); window.accidentalSends = 0; ask = async () => { window.accidentalSends++; }; }')
        page.locator('#doc-task-request').press('Control+Enter')
        page.locator('#doc-task-request').dispatch_event('keydown', {'key': 'Enter', 'isComposing': True, 'ctrlKey': True, 'bubbles': True})
        assert page.evaluate('window.accidentalSends') == 0
        if page.locator('#nav-resources').is_hidden():
            page.locator('#sidebar-toggle').click()
        page.locator('#nav-resources').click()
        if page.locator('#nav-work').is_hidden():
            page.locator('#sidebar-toggle').click()
        page.locator('#nav-work').click()
        page.locator('#simple-document-menu > summary').click()
        page.locator('#nav-review').click()
        page.keyboard.press('Escape')
        if page.locator('#nav-work').is_hidden():
            page.locator('#sidebar-toggle').click()
        page.locator('#nav-work').click()
        assert not errors and page.evaluate('window.unexpectedCalls') == ['get_release_info']
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        page.close()
    finally:
        fixture.close()


@pytest.mark.parametrize('width', [1440, 1100, 520])
def test_long_document_and_output_names_do_not_clip_document_bar(tmp_path, monkeypatch, width):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        setup(page, api, app, calls, receipts)
        page.set_viewport_size({'width': width, 'height': 800})
        page.evaluate("document.getElementById('file-name').textContent = 'Report_'.repeat(28) + '.docx'")
        page.wait_for_timeout(50)
        page.evaluate('''() => {
          const name = 'Report_'.repeat(28) + '.docx';
          window.dispatchEvent(new CustomEvent('officeworkspace:task', {detail:{source_name:name,current_name:name,receipts:[]}}));
        }''')
        assert page.locator('#workspace-document-bar').evaluate('node => node.scrollWidth <= node.clientWidth')
        assert page.locator('#workspace-document-identity').evaluate('node => node.scrollWidth <= node.clientWidth')
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
        assert not calls and not app.editor.audit.exists()
    finally:
        fixture.close()
