from copy import deepcopy
import json
from pathlib import Path
import re
import shutil
import time

from docx import Document
from lxml import etree as ET

from office_agent.document_return import _project
from office_agent.professional_edits import W, _canonical
from office_agent.writes import Editor, digest
from tests.test_professional_browser import STATIC, _browser
from tests.test_return_continuation import parts, rewrite


def exercise(tmp_path, monkeypatch, capture=None):
    fixture = _browser(tmp_path, monkeypatch, real_render=True)
    baseline_page = None
    try:
        page, api, app, calls, receipts = next(fixture)
        doc = Document()
        for text in ['前置上下文', '旧术语仅限室内20 ms [1]。', '普通用户说明', '后续核对区域。', '尾部上下文']:
            doc.add_paragraph(text)
        doc.save(app.path)
        def history(paragraphs):
            wrapper = ET.SubElement(paragraphs[3], W + 'ins', {W + 'id': '42', W + 'author': '另一位作者', W + 'date': '2026-01-01T00:00:00Z'})
            ET.SubElement(ET.SubElement(wrapper, W + 'r'), W + 't').text = '其他作者保留'
        rewrite(app.path, history)
        original = digest(app.path)
        source = app.path
        chosen = [source]
        def select_file():
            app.editor = Editor(chosen[0].parent, require_confirmation=True, render_validation=True)
            app.path = app.editor.space.file(str(chosen[0]))
            return {'ok': True, 'name': app.path.name, 'dir': str(app.editor.space.root)}
        app.select_file = select_file
        allowed = {'create_document_task', 'plan_document_task_group', 'confirm_document_task_group',
                   'compare_document_return', 'create_document_return_continuation', 'select_file',
                   'undo_document_task_group', 'start_range_review', 'open_range_review', 'range_review_page'}
        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('returnHost', invoke)
        page.evaluate('''(directory) => {
            document.getElementById('file-dir').textContent=directory;
            window.pywebview.api=new Proxy({}, {get: (_, method) => (...args) => window.returnHost(method,args)});
        }''', str(tmp_path))
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        page.add_script_tag(path=str(STATIC / 'range_edits.js'))
        page.locator('#doc-task-request').fill('将全文“旧术语”替换为“规范术语”')
        page.locator('#doc-task-create').click()
        page.wait_for_function("!document.getElementById('doc-task-route').disabled")
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('已执行并验证 1')")
        task = [reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group'][-1]
        delivered = tmp_path / task['current_name']
        returned = tmp_path / 'returned.docx'
        returned.write_bytes(delivered.read_bytes())
        rewrite(returned, lambda paragraphs: setattr(paragraphs[2].find('.//' + W + 't'), 'text', '用户自行补充说明'))
        stage_start = time.perf_counter()
        page.locator('#doc-task-prepare-return').click()
        fill_seconds = time.perf_counter() - stage_start
        assert page.locator('#doc-task-return-baseline').input_value() == source.name
        assert page.locator('#doc-task-return-delivered').input_value() == delivered.name
        assert calls.count('compare_document_return') == 0
        def choose(path):
            chosen[0] = path
            reply = page.evaluate('window.pywebview.api.select_file()')
            assert reply['ok'], reply
            page.evaluate('(info) => {document.getElementById("file-name").textContent=info.name;document.getElementById("file-dir").textContent=info.dir;}', reply)
            page.wait_for_function("document.getElementById('doc-task-summary').textContent === ''")
            assert not api._document_tasks.tasks and api._document_tasks.return_snapshot is None
        choose(returned)
        assert page.locator('#doc-task-return-receipts').input_value() == ', '.join(task['receipts'])
        baseline_root = Path(__file__).resolve().parents[1] / 'deliverables/medium-workflow-20261025/baseline-v0.6.6/office_agent/webui/static'
        baseline_page = page.context.browser.new_page()
        baseline_page.set_default_timeout(30000)
        baseline_page.route('**/*', lambda route: route.abort())
        baseline_page.expose_function('oldReturnHost', invoke)
        html = re.sub(r'<script\b[^>]*>.*?</script>', '', (baseline_root / 'index.html').read_text(encoding='utf8'), flags=re.S)
        baseline_page.set_content(html)
        baseline_page.evaluate('''() => {
          document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;
          document.getElementById('professional-panel').open=true;
          window.OfficeAuth={isAuthenticated:()=>true,wrapApi:value=>value};window.confirm=()=>true;
          window.pywebview={api:new Proxy({}, {get:(_, method)=>(...args)=>window.oldReturnHost(method,args)})};
        }''')
        baseline_page.add_script_tag(path=str(baseline_root / 'document_tasks.js'))
        baseline_page.locator('#doc-task-return-baseline').evaluate('node => node.closest("details").open=true')
        old_start = time.perf_counter()
        for key, value in [('baseline', source.name), ('delivered', delivered.name), ('receipts', ', '.join(task['receipts']))]:
            baseline_page.locator('#doc-task-return-' + key).fill(value)
        old_seconds = time.perf_counter() - old_start
        baseline_page.locator('#doc-task-return-compare').click()
        baseline_page.wait_for_function("document.getElementById('doc-task-return-output').textContent.includes('用户新增文字修改')")
        baseline_page.close()
        baseline_page = None
        page.locator('#doc-task-return-compare').click()
        page.wait_for_function("document.getElementById('doc-task-return-target').options.length > 1")
        comparison = [reply['result'] for method, reply in receipts if method == 'compare_document_return'][-1]
        safe = next(row for row in comparison['items'] if '后续核对' in row['current'])
        assert safe['continuation_id']
        user_row = next(row for row in comparison['items'] if row['classification'] == 'user_content_change')
        assert user_row['continuation_id'] and user_row['continuation_requires_current_review']
        page.locator('#doc-task-return-target').select_option(safe['continuation_id'])
        for key, value in [('find', '后续核对'), ('replacement', '继续核对'), ('reason', '仅调整当前措辞；保留其他作者及所有条件')]:
            page.locator('#doc-task-return-' + key).fill(value)
        page.locator('#doc-task-return-mode').select_option('tracked')
        page.locator('#doc-task-return-ack').check()
        if capture:
            page.evaluate("document.documentElement.dataset.theme='light'")
            page.set_viewport_size({'width': 520, 'height': 900})
            page.locator('#doc-task-return-target').scroll_into_view_if_needed()
            assert page.evaluate('document.documentElement.scrollWidth<=innerWidth+1')
            page.screenshot(path=str(capture / 'continuation-520.png'))
            page.set_viewport_size({'width': 1100, 'height': 900})
        before_candidate = len(list(app.editor.audit.iterdir()))
        page.locator('#doc-task-return-create').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('新候选已建立')")
        assert len(list(app.editor.audit.iterdir())) == before_candidate
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.evaluate('window.approved=false')
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('已取消')")
        assert len(list(app.editor.audit.iterdir())) == before_candidate
        page.evaluate('window.approved=true')
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('已执行并验证 1')")
        second = [reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group'][-1]
        review = tmp_path / second['current_name']
        root = ET.fromstring(parts(review)['word/document.xml'])
        history_node = next(node for node in root.iter(W + 'ins') if node.get(W + 'id') == '42')
        assert history_node.get(W + 'author') == '另一位作者'
        assert '用户自行补充说明' in ''.join(root.itertext())
        page.locator('#doc-task-deliveries button').click()
        page.locator('#range-open-review').click()
        page.wait_for_function("document.querySelectorAll('#range-before-pages img,#range-after-pages img').length === 2")
        screenshots = []
        for theme in ('dark', 'light', 'highcontrast'):
            for width in (1100, 520):
                page.evaluate('(value)=>document.documentElement.dataset.theme=value', theme)
                page.set_viewport_size({'width': width, 'height': 900})
                page.locator('#doc-task-workflow').scroll_into_view_if_needed()
                assert page.evaluate('document.documentElement.scrollWidth<=innerWidth+1')
                if capture:
                    name = f'{theme}-{width}.png'
                    page.screenshot(path=str(capture / name))
                    screenshots.append(name)
        current2 = tmp_path / 'returned2.docx'
        current2.write_bytes(review.read_bytes())
        def accept_this_round(paragraphs):
            paragraph = paragraphs[3]
            ids = [node.get(W + 'id') for node in paragraph if node.tag in {W + 'ins', W + 'del'} and node.get(W + 'id') != '42']
            paragraph.getparent().replace(paragraph, _project(paragraph, ids, ['accept'] * len(ids)))
        rewrite(current2, accept_this_round)
        page.locator('#doc-task-prepare-return').click()
        from tests.ui_confirmation import respond_to_confirmation
        respond_to_confirmation(page)
        choose(current2)
        page.locator('#doc-task-return-compare').click()
        page.wait_for_function("document.getElementById('doc-task-return-output').textContent.includes('本系统文字修订已接受')")
        compared = [reply['result'] for method, reply in receipts if method == 'compare_document_return'][-1]
        target = next(row for row in compared['items'] if '继续核对' in row['current'])
        page.locator('#doc-task-return-target').select_option(target['continuation_id'])
        for key, value in [('find', '继续核对'), ('replacement', '再次核对'), ('reason', '已接受上一轮后独立修改')]:
            page.locator('#doc-task-return-' + key).fill(value)
        page.locator('#doc-task-return-ack').check()
        page.locator('#doc-task-return-create').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('新候选已建立')")
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('已执行并验证 1')")
        third = [reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group'][-1]
        final = tmp_path / third['current_name']
        final_parts = parts(final)
        current_parts = parts(current2)
        assert all(final_parts[name] == data for name, data in current_parts.items() if name != 'word/document.xml')
        final_root = ET.fromstring(final_parts['word/document.xml'])
        assert '再次核对区域。' in ''.join(final_root.itertext())
        assert _canonical(history_node) == _canonical(next(node for node in final_root.iter(W + 'ins') if node.get(W + 'id') == '42'))
        hashes = {path.name: digest(path) for path in (source, delivered, returned, review, current2, final)}
        assert hashes[source.name] == original
        if capture:
            shutil.copyfile(final, capture / 'clean-candidate.docx')
            shutil.copyfile(review, capture / 'review-candidate.docx')
        page.locator('#doc-task-undo').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('已撤销')")
        assert not final.exists()
        assert page.locator('#doc-task-prepare-return').is_disabled()
        page.evaluate("document.getElementById('file-dir').textContent='different workspace fixture'")
        page.wait_for_function("document.getElementById('doc-task-return-receipts').value === ''")
        api._clear_private_state()
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#doc-task-return-receipts').input_value() == ''
        assert page.locator('#doc-task-return-target option').count() == 0
        report = {'status': 'completed', 'workflow': 'fresh admission, safe candidate, cancel, tracked copy, read-only pages, simulated acceptance, new admission, clean copy, independent undo',
                  'comparison': {'scope': 'same host and material; frozen v0.6.6 UI versus current UI data entry only',
                                 'baseline_fields_filled': 3, 'current_prepare_clicks': 1,
                                 'baseline_characters_entered': len(source.name + delivered.name + ', '.join(task['receipts'])),
                                 'current_characters_entered': 0, 'baseline_automation_seconds': old_seconds,
                                 'current_automation_seconds': fill_seconds, 'human_time_and_total_workflow_benefit': 'not_measured'},
                  'source_unchanged': digest(source) == original, 'other_author_preserved': True,
                  'different_directory_clears_handoff': True,
                  'original_versions_sha256': hashes, 'final_candidate_undone': True, 'screenshots': screenshots,
                  'transactions': [{'save_id': reply['result']['save_id'], 'output_name': reply['result']['task']['current_name'],
                                    'mode': reply['result']['task']['mode'], 'output_sha256': reply['result']['task']['current_sha256']}
                                   for method, reply in receipts if method == 'confirm_document_task_group' and reply.get('ok') and reply['result']['status'] == 'saved'],
                  'real_word': 'not_verified; revision acceptance simulated by explicit OOXML projection',
                  'authentication': 'fixture', 'native_file_dialog': 'fixture', 'independent_human': 'deferred_by_user',
                  'models': 'not_called', 'calls': calls}
        if capture:
            (capture / 'REPORT.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
        return report
    finally:
        if baseline_page:
            baseline_page.close()
        if 'api' in locals():
            api._clear_private_state()
        fixture.close()


def test_real_return_workflow_and_frozen_ui_comparison(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)


def test_late_return_response_duplicate_clicks_and_logout_never_restore_draft(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, _, _ = next(fixture)
        page.evaluate('''() => {
            window.returnCalls=0;
            window.pywebview.api={compare_document_return:()=>{window.returnCalls++;return new Promise(resolve=>window.resolveReturn=resolve);}};
        }''')
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        page.locator('#doc-task-return-panel').evaluate('node=>node.open=true')
        page.locator('#doc-task-return-baseline').fill('baseline.docx')
        page.locator('#doc-task-return-delivered').fill('delivered.docx')
        page.evaluate('''() => {document.getElementById('doc-task-return-compare').click();document.getElementById('doc-task-return-compare').click();}''')
        assert page.evaluate('window.returnCalls') == 0
        from tests.ui_confirmation import respond_to_confirmation
        respond_to_confirmation(page)
        page.wait_for_function('window.returnCalls===1')
        assert page.evaluate('window.returnCalls') == 1
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate('''() => window.resolveReturn({ok:true,result:{comparison_id:'old',next_offset:null,
          versions:[],classifications:{},package_changes:[],items:[{current:'late private draft',continuation_id:'old'}]}})''')
        assert page.locator('#doc-task-return-output').inner_text() == ''
        assert page.locator('#doc-task-return-target option').count() == 0
        assert page.locator('#doc-task-return-baseline').input_value() == ''
        assert page.locator('#doc-task-return-create').is_disabled()
        assert not app.editor.audit.exists()
    finally:
        fixture.close()
