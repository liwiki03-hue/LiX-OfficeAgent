import json
import shutil

from docx import Document
from lxml import etree as ET

from office_agent.professional_edits import W, _canonical
from office_agent.writes import Editor, digest
from tests.test_professional_browser import STATIC, _browser
from tests.test_return_continuation import parts, rewrite


def exercise(tmp_path, monkeypatch, capture=None):
    fixture = _browser(tmp_path, monkeypatch, real_render=True)
    try:
        page, api, app, calls, receipts = next(fixture)
        document = Document()
        document.add_paragraph('结果仅限室内，流量20 mL/min，样本8个。')
        document.add_paragraph('本报告不得用于临床决策。')
        document.add_paragraph('其他作者区域')
        table = document.add_table(rows=2, cols=2)
        for cell, text in zip(table._cells, ['指标', '条件', '旧术语', '结果仅限室内20 mL/min']):
            cell.text = text
        document.save(app.path)
        def history(paragraphs):
            wrapper = ET.SubElement(paragraphs[2], W + 'ins', {W + 'id': '42', W + 'author': '另一位作者'})
            ET.SubElement(ET.SubElement(wrapper, W + 'r'), W + 't').text = '保留历史修订'
        rewrite(app.path, history)
        source = app.path
        original = digest(source)
        chosen = [source]
        def select_file():
            app.editor = Editor(chosen[0].parent, require_confirmation=True, render_validation=True)
            app.path = app.editor.space.file(str(chosen[0]))
            return {'ok': True, 'name': app.path.name, 'dir': str(app.editor.space.root)}
        app.select_file = select_file
        allowed = {'create_document_task', 'plan_document_task_group', 'confirm_document_task_group',
                   'compare_document_return', 'create_document_return_continuation', 'select_file', 'undo_document_task_group'}
        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            reply = getattr(api, method)(*args)
            receipts.append((method, reply))
            return reply
        page.expose_function('technicalHost', invoke)
        page.evaluate('''directory => {
          document.getElementById('file-dir').textContent=directory;
          window.pywebview.api=new Proxy({}, {get:(_, method)=>(...args)=>window.technicalHost(method,args)});
        }''', str(tmp_path))
        page.add_script_tag(path=str(STATIC / 'document_tasks.js'))
        page.locator('#doc-task-profile').select_option('technical_report_v1')
        page.locator('#doc-task-request').fill('将全文“mL/min”替换为“mL/s”')
        page.locator('#doc-task-create').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('受限')")
        rejected = receipts[-1][1]['result']
        assert rejected['states']['restricted'] == 2
        assert not app.editor.audit.exists()
        page.locator('#doc-task-request').fill('将全文“旧术语”替换为“规范术语”')
        page.locator('#doc-task-create').click()
        page.wait_for_function("!document.getElementById('doc-task-route').disabled")
        assert '技术报告字面保护v1' in page.locator('#doc-task-protection').inner_text()
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('已执行并验证 1')")
        task = [reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group'][-1]
        delivered = tmp_path / task['current_name']
        returned = tmp_path / 'returned.docx'
        shutil.copyfile(delivered, returned)
        page.locator('#doc-task-prepare-return').click()
        chosen[0] = returned
        reply = page.evaluate('window.pywebview.api.select_file()')
        page.evaluate('info=>{document.getElementById("file-name").textContent=info.name;document.getElementById("file-dir").textContent=info.dir;}', reply)
        page.wait_for_function("document.getElementById('doc-task-summary').textContent === ''")
        page.locator('#doc-task-return-panel').evaluate('node=>node.open=true')
        page.locator('#doc-task-return-compare').click()
        page.wait_for_function("document.getElementById('doc-task-return-target').options.length > 1")
        ordinary = [reply['result'] for method, reply in receipts if method == 'compare_document_return'][-1]
        assert all(not row['continuation_id'] for row in ordinary['items'] if '/table/' in (row['current_location'] or ''))
        page.locator('#doc-task-return-tables').check()
        assert page.locator('#doc-task-return-target option').count() == 0
        page.locator('#doc-task-return-compare').click()
        page.wait_for_function("document.getElementById('doc-task-return-target').options.length > 1")
        comparison = [reply['result'] for method, reply in receipts if method == 'compare_document_return'][-1]
        row = next(row for row in comparison['items'] if row['current'] == '结果仅限室内20 mL/min')
        assert row['continuation_family'] == 'ordinary_table_cell' and row['continuation_id']
        page.locator('#doc-task-profile').select_option('technical_report_v1')
        page.locator('#doc-task-return-target').select_option(row['continuation_id'])
        for key, value in [('find', 'mL/min'), ('replacement', 'mL/s'), ('reason', '核对当前单元格保护')]:
            page.locator('#doc-task-return-' + key).fill(value)
        page.locator('#doc-task-return-ack').check()
        count = len(receipts)
        page.locator('#doc-task-return-create').click()
        page.wait_for_function("!document.getElementById('doc-task-return-find').disabled")
        reply = next(reply for method, reply in receipts[count:] if method == 'create_document_return_continuation')
        assert not reply['ok'] and reply['reason'] == 'professional_profile_changed_quantity', reply
        for key, value in [('find', '结果'), ('replacement', '测得结果')]:
            page.locator('#doc-task-return-' + key).fill(value)
        page.locator('#doc-task-return-mode').select_option('tracked')
        page.locator('#doc-task-return-create').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('新候选已建立')")
        assert '20 mL/min' in page.locator('#doc-task-protection').inner_text()
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.evaluate('window.approved=false')
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('取消')")
        page.evaluate('window.approved=true')
        page.locator('#doc-task-route').click()
        page.wait_for_function("!document.getElementById('doc-task-confirm').disabled")
        page.locator('#doc-task-confirm').click()
        page.wait_for_function("document.getElementById('doc-task-summary').textContent.includes('已执行并验证 1')")
        saved = [reply['result']['task'] for method, reply in receipts if method == 'confirm_document_task_group'][-1]
        output = tmp_path / saved['current_name']
        before, after = parts(returned), parts(output)
        assert all(after[name] == value for name, value in before.items() if name != 'word/document.xml')
        roots = [ET.fromstring(value['word/document.xml']) for value in (before, after)]
        histories = [next(root.iter(W + 'ins')) for root in roots]
        assert _canonical(histories[0]) == _canonical(histories[1])
        for root in roots:
            for paragraph in root.find('.//' + W + 'tbl').iter(W + 'p'):
                paragraph.clear()
        assert _canonical(roots[0].find('.//' + W + 'tbl')) == _canonical(roots[1].find('.//' + W + 'tbl'))
        assert digest(source) == original
        screenshots = []
        if capture:
            shutil.copyfile(output, capture / 'review-candidate.docx')
            shutil.copyfile(delivered, capture / 'clean-candidate.docx')
            for theme in ['light', 'dark', 'highcontrast']:
                for width in [520, 1100]:
                    page.evaluate('theme=>document.documentElement.dataset.theme=theme', theme)
                    page.set_viewport_size({'width': width, 'height': 900})
                    page.locator('#doc-task-protection').scroll_into_view_if_needed()
                    assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
                    name = f'{theme}-{width}.png'
                    page.screenshot(path=str(capture / name))
                    screenshots.append(name)
        page.locator('#doc-task-undo').click()
        page.wait_for_function("document.getElementById('doc-task-status').textContent.includes('已撤销')")
        assert not output.exists()
        report = {'status': 'completed', 'source_unchanged': True, 'ordinary_table_explicit_opt_in': True,
                  'unsafe_quantity_candidates_refused': 2, 'safe_copies': 2, 'cancel_and_undo': True,
                  'other_author_and_non_document_parts_preserved': True, 'table_geometry_preserved': True,
                  'screenshots': screenshots, 'calls': calls, 'models': 'not_called',
                  'word': 'not_verified', 'authentication_and_native_picker': 'fixture', 'independent_human': 'deferred'}
        if capture:
            (capture / 'REPORT.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
        return report
    finally:
        if 'api' in locals():
            api._clear_private_state()
        fixture.close()


def test_technical_profile_and_explicit_table_return_real_render(tmp_path, monkeypatch):
    exercise(tmp_path, monkeypatch)
