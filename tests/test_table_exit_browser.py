import base64
from contextlib import contextmanager
import hashlib
import io
import json
import os
from pathlib import Path
import shutil
import time
from types import SimpleNamespace

from docx import Document
from docx.oxml.ns import qn
import pytest

from office_agent.paged_documents import eligibility, normalize, projection
from office_agent.instruction_runtime import InstructionRuntime
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from tests.test_paged_browser import EDGE, STATIC, engine_frame, serve_static
from tests.ui_confirmation import record_browser_failure, track_browser


def table_fixture(path, case):
    doc = Document()
    if case not in ('T4',):
        doc.add_paragraph('表前正文：保留实验条件。')
    table = doc.add_table(rows=36 if case == 'T6' else 2, cols=2)
    if case == 'styled':
        table.style = 'Table Grid'
    for i, row in enumerate(table.rows):
        row.cells[0].text = f'观察{i + 1}'
        row.cells[1].text = f'{120 + i} ms'
    table.cell(0, 0).paragraphs[0].runs[0].bold = True
    if case == 'T1' or case == 'T4':
        doc.add_paragraph('表后正文：单位和数字保持。')
    elif case == 'T3':
        doc.add_paragraph()
    elif case == 'T5':
        other = doc.add_table(rows=2, cols=2)
        other.cell(0, 0).text = '另一张表'
    elif case == 'T6':
        table.cell(0, 1).add_paragraph('单元格第二段。')
    elif case == 'merged':
        table.cell(0, 0).merge(table.cell(0, 1))
    elif case == 'nested':
        table.cell(0, 0).add_table(rows=1, cols=1)
    doc.sections[0].header.paragraphs[0].text = '自建表格验收页眉'
    doc.sections[0].footer.paragraphs[0].text = '自建表格验收页脚'
    doc.save(path)
    return path


def inspect_frame(frame):
    return frame.evaluate("""() => {
      const e=window.__tableTestEngine.activeEditor,s=e.state.selection,blocks=[];
      e.state.doc.forEach((node,pos)=>blocks.push({type:node.type.name,pos,size:node.nodeSize,id:node.attrs.sdBlockId,text:node.textContent}));
      return {blocks,selection:{type:s.constructor.name,from:s.from,to:s.to,ancestors:Array.from({length:s.$from.depth+1},(_,i)=>s.$from.node(i).type.name)},
        focus:document.activeElement?.className,keys:window.__tableKeys||[],frame_state:window.__tableFrameState,navigation:{last_type:e.state.doc.lastChild.type.name,last_textblock:e.state.doc.lastChild.isTextblock,last_inline:e.state.doc.lastChild.inlineContent,end_parent:e.state.doc.resolve(e.state.doc.content.size-1).parent.type.name,end_inline:e.state.doc.resolve(e.state.doc.content.size-1).parent.inlineContent,near:typeof s.constructor.near},revision:window.__tableEvents.filter(x=>x.type==='changed').at(-1)?.value?.revision||0,
        paragraphDefaults:e.state.schema.nodes.paragraph.defaultAttrs,commands:Object.keys(e.commands).filter(x=>/paragraph|select|focus|history/i.test(x))};
    }""")


@contextmanager
def open_table(tmp_path, monkeypatch, case):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    source = table_fixture(tmp_path / f'{case}-表格验收.docx', case)
    app = WebviewApp()
    app.attach(SimpleNamespace(create_file_dialog=lambda *a, **kw: [str(source)]))
    api = FloatingApi(app, settings_root=tmp_path)
    monkeypatch.setenv('TEMP', str(tmp_path))
    monkeypatch.setenv('TMP', str(tmp_path))
    with serve_static() as url, sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1280, 'height': 820})
            page.set_default_timeout(15000)
            script = (STATIC / 'paged_editor_frame.js').read_text(encoding='utf8')
            script = script.replace('engine = new Engine(', 'engine = window.__tableTestEngine = new Engine(')
            script = script.replace('  const actions=', "  Object.defineProperty(window,'__tableFrameState',{get:()=>({ready,editable,composing})}); const actions=")
            script = script.replace("const send = (type, value) => port?.postMessage({type, value});", "window.__tableEvents=[]; const send=(type,value)=>{window.__tableEvents.push({type,value});port?.postMessage({type,value});};")
            page.route('**/paged_editor_frame.js', lambda route: route.fulfill(status=200, content_type='text/javascript', body=script))
            def host(name, args):
                if name == 'paged_sync' and args[-1] is True and os.environ.get('TABLE_EXIT_EVIDENCE'):
                    path = Path(os.environ['TABLE_EXIT_EVIDENCE'])
                    path.mkdir(parents=True, exist_ok=True)
                    data = normalize(base64.b64decode(args[2]), source.read_bytes())
                    (path / f'{case}-bootstrap.json').write_text(json.dumps({'before': projection(normalize(source.read_bytes())), 'after': projection(data)}, ensure_ascii=False, indent=2), encoding='utf8')
                return getattr(api, name)(*args)
            page.expose_function('testHost', host)
            page.add_init_script("""window.__tableKeys=[];window.addEventListener('keydown',event=>{if(!['Home','End'].includes(event.key))return;const take=phase=>{const e=window.__tableTestEngine?.activeEditor,s=e?.state.selection.$from;window.__tableKeys.push({phase,key:event.key,ctrl:event.ctrlKey,meta:event.metaKey,composing:event.isComposing,target:event.target?.id||event.target?.className,focus:document.activeElement?.className,prevented:event.defaultPrevented,ancestors:s?Array.from({length:s.depth+1},(_,depth)=>s.node(depth).type.name):[]});window.__tableKeys=window.__tableKeys.slice(-30);};take('capture');queueMicrotask(()=>take('after'));},true);""")
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.testHost(name,args)})};window.confirm=()=>true;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js', lambda route: route.fulfill(status=200, body=''))
            errors = []
            page.on('pageerror', lambda error: errors.append(str(error)))
            track_browser(page,tmp_path.name,lambda:{'document':api._floating.document,'paged':api._floating.paged.state(),'busy':app.busy})
            page.goto(url)
            page.locator('#floating-choose').click()
            page.get_by_role('button', name='同意读取并打开', exact=True).click()
            deadline = time.monotonic() + 70
            while not api._floating.paged.manifest:
                if api._floating.paged.document_id and not api._floating.paged.editable:
                    raise AssertionError({'reasons': api._floating.paged.reasons, 'feedback': page.locator('#paged-feedback').text_content(), 'errors': errors})
                if time.monotonic() > deadline:
                    raise AssertionError(page.locator('#paged-feedback').text_content())
                page.wait_for_timeout(100)
            try:
                yield page, api, source, errors
            except BaseException:
                evidence(page, engine_frame(page), tmp_path.name+'-'+case+'-failure', {'feedback': page.locator('#paged-feedback').text_content(), 'errors': errors})
                record_browser_failure(page,'table navigation and unchanged document',15)
                raise
        finally:
            browser.close()


def evidence(page, frame, name, extra=None):
    folder = os.environ.get('TABLE_EXIT_EVIDENCE') or os.environ.get('LIX_UI_FAILURE_DIR')
    if folder:
        path = Path(folder)
        path.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(path / f'{name}.png'))
        (path / f'{name}.json').write_text(json.dumps({'engine': inspect_frame(frame), 'extra': extra}, ensure_ascii=False, indent=2), encoding='utf8')


def test_existing_body_navigation_does_not_change_document(tmp_path, monkeypatch):
    with open_table(tmp_path, monkeypatch, 'T1') as (page, api, source, errors):
        frame = engine_frame(page)
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        before = inspect_frame(frame)
        frame.locator('#frame-error').evaluate("node=>node.textContent='另一项提示保留'")
        assert 'tableCell' in before['selection']['ancestors']
        frame.locator('.superdoc-page').get_by_text('表后正文：单位和数字保持。', exact=True).first.click()
        after = inspect_frame(frame)
        assert 'tableCell' not in after['selection']['ancestors']
        assert after['blocks'] == before['blocks'] and after['revision'] == before['revision'] == 0
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        frame.wait_for_function("""() => {const e=window.__tableTestEngine.activeEditor,s=e.state.selection.$from;return e.view.hasFocus()&&Array.from({length:s.depth+1},(_,depth)=>s.node(depth).type.name).includes('tableCell');}""")
        page.keyboard.press('Control+End')
        frame.wait_for_function("""() => {const s=window.__tableTestEngine.activeEditor.state.selection.$from;return !Array.from({length:s.depth+1},(_,depth)=>s.node(depth).type.name).includes('tableCell');}""")
        assert 'tableCell' not in inspect_frame(frame)['selection']['ancestors']
        page.keyboard.press('Home')
        page.keyboard.press('Shift+End')
        page.get_by_role('button', name='引用选区', exact=True).click()
        page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('已引用')")
        service = api._floating.paged
        reference_id = next(iter(service.references))
        service.proposals['selection-only'] = {'id': 'selection-only', 'status': 'pending', 'draft_revision': 0, 'instruction_revision': 0, 'old': '正文', 'new': '正文', 'reason': '自建测试'}
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        frame.get_by_role('button', name='在表格后继续正文', exact=True).click()
        assert service.proposals['selection-only']['status'] == 'pending'
        assert reference_id in service.references and inspect_frame(frame)['revision'] == 0
        assert frame.locator('#frame-error').text_content() == '另一项提示保留'
        page.keyboard.press('End')
        page.keyboard.insert_text('继续正文。')
        assert '继续正文。' in inspect_frame(frame)['blocks'][-1]['text']
        assert not errors
        evidence(page, frame, 'T1-existing-body')


def test_last_table_has_real_body_exit_and_reuses_it(tmp_path, monkeypatch):
    with open_table(tmp_path, monkeypatch, 'T2') as (page, api, source, errors):
        frame = engine_frame(page)
        assert Document(source).element.body[-2].tag == qn('w:tbl')
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        page.keyboard.press('Control+End')
        page.keyboard.press('ArrowDown')
        before = inspect_frame(frame)
        evidence(page, frame, 'T2-before-exit', {'gesture': 'last cell -> Ctrl+End -> ArrowDown'})
        frame.get_by_role('button', name='在表格后继续正文', exact=True).click()
        after = inspect_frame(frame)
        assert after['blocks'][-1]['type'] == 'paragraph'
        assert 'tableCell' not in after['selection']['ancestors']
        assert len(after['blocks']) == len(before['blocks']) + 1
        assert after['revision'] == before['revision'] + 1
        assert after['selection']['from'] == after['selection']['to']
        page.keyboard.press('Control+z')
        assert inspect_frame(frame)['blocks'] == before['blocks']
        page.keyboard.press('Control+y')
        assert len(inspect_frame(frame)['blocks']) == len(before['blocks']) + 1
        assert inspect_frame(frame)['selection']['from'] == inspect_frame(frame)['selection']['to']
        page.keyboard.insert_text('新增表外正文，保留120 ms。')
        assert inspect_frame(frame)['blocks'][-1]['text'] == '新增表外正文，保留120 ms。'
        current = inspect_frame(frame)
        assert [(b['id'], b['text']) for b in current['blocks'][:-1]] == [(b['id'], b['text']) for b in before['blocks']]
        evidence(page, frame, 'T2-before-reuse')
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        frame.get_by_role('button', name='在表格后继续正文', exact=True).click()
        assert inspect_frame(frame)['blocks'] == current['blocks']
        assert inspect_frame(frame)['revision'] == current['revision']
        page.keyboard.press('Home')
        page.keyboard.press('Shift+End')
        page.get_by_role('button', name='引用选区', exact=True).click()
        page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('已引用')")
        reference = next(iter(api._floating.paged.references.values()))
        assert reference['old'] == '新增表外正文，保留120 ms。'
        assert api._floating.paged.recovery_state == 'saved'
        original = source.read_bytes()
        page.get_by_role('button', name='保存副本', exact=True).click()
        dialog = page.get_by_role('dialog', name='生成验证副本', exact=True)
        dialog.get_by_role('button', name='取消', exact=True).click()
        assert not list(source.parent.glob('*.agent-*.docx')) and source.read_bytes() == original
        page.get_by_role('button', name='保存副本', exact=True).click()
        dialog.get_by_role('button', name='确认生成副本', exact=True).click()
        page.wait_for_function("document.getElementById('paged-save-state').textContent.startsWith('副本已验证')", timeout=180000)
        receipt = next(iter(api._floating.outputs.values()))
        output = Path(receipt['output'])
        reopened = Document(output)
        assert reopened.element.body[-2].tag == qn('w:p')
        assert reopened.paragraphs[-1].text == '新增表外正文，保留120 ms。'
        assert reopened.paragraphs[-1].style.style_id == 'Normal'
        assert not any(run.bold for run in reopened.paragraphs[-1].runs)
        assert reopened.paragraphs[-1]._p.find('./'+qn('w:pPr')+'/'+qn('w:numPr')) is None
        assert projection(original)['tables'] == projection(output.read_bytes())['tables']
        assert projection(original)['headers'] == projection(output.read_bytes())['headers']
        assert projection(original)['footers'] == projection(output.read_bytes())['footers']
        assert source.read_bytes() == original
        folder = os.environ.get('TABLE_EXIT_EVIDENCE')
        if folder:
            path = Path(folder)
            shutil.copy2(source, path/'原稿.docx')
            shutil.copy2(output, path/'表外正文副本.docx')
            (path/'PUBLISH.json').write_text(json.dumps({'source_sha256': hashlib.sha256(original).hexdigest(), 'output_sha256': hashlib.sha256(output.read_bytes()).hexdigest(), 'receipt': receipt}, ensure_ascii=False, indent=2), encoding='utf8')
        assert not errors
        evidence(page, frame, 'T2-after-exit')
        api._floating.app.editor.rollback(receipt['id'], confirmed=True)
    service = api._floating.paged
    service.close()
    reopened = service.open(api._floating.document['id'], api._floating.document['version'])
    assert reopened['recovery_available']
    restored = service.restore(service.document_id)
    assert Document(io.BytesIO(base64.b64decode(restored['data']))).paragraphs[-1].text == '新增表外正文，保留120 ms。'
    assert service.references == {} and service.proposals == {} and service.publisher.plans == {}
    assert source.read_bytes() == original


@pytest.mark.parametrize('case,side', [('T3', 'after'), ('T4', 'before'), ('T5', 'after'), ('T6', 'after')])
def test_table_boundary_reuse_order_and_long_page(tmp_path, monkeypatch, case, side):
    with open_table(tmp_path, monkeypatch, case) as (page, api, source, errors):
        frame = engine_frame(page)
        text = '155 ms' if case == 'T6' else '121 ms'
        frame.locator('.superdoc-page').get_by_text(text, exact=True).first.click()
        before = inspect_frame(frame)
        label = '在表格后继续正文' if side == 'after' else '在表格前插入正文'
        frame.get_by_role('button', name=label, exact=True).click()
        after = inspect_frame(frame)
        assert 'tableCell' not in after['selection']['ancestors']
        assert len(after['blocks']) == len(before['blocks']) + (0 if case == 'T3' else 1)
        assert after['revision'] == before['revision'] + (0 if case == 'T3' else 1)
        page.keyboard.insert_text('这是表外的正文。')
        selected = inspect_frame(frame)
        assert selected['blocks'][0 if side == 'before' else -2 if case == 'T5' else -1]['text'] == '这是表外的正文。'
        assert [(b['id'], b['text'], b['size']) for b in selected['blocks'] if b['type'] == 'table'] == [(b['id'], b['text'], b['size']) for b in before['blocks'] if b['type'] == 'table']
        for width in (1280, 520):
            page.set_viewport_size({'width': width, 'height': 820})
            collapse = page.get_by_role('button', name='收起对话侧栏', exact=True)
            if width <= 850 and collapse.is_visible():
                collapse.click()
            frame.locator('.superdoc-page').get_by_text(text, exact=True).first.click()
            button = frame.get_by_role('button', name=label, exact=True)
            button.wait_for(state='visible')
            assert button.bounding_box()['y'] >= 0
            for theme in ['light', 'dark', 'highcontrast']:
                frame.evaluate('(theme)=>document.body.dataset.theme=theme', theme)
                assert button.is_visible()
        assert not errors
        evidence(page, frame, case+'-boundary')


def test_cell_enter_tab_composition_and_stale_plans(tmp_path, monkeypatch):
    with open_table(tmp_path, monkeypatch, 'T2') as (page, api, source, errors):
        frame = engine_frame(page)
        frame.locator('.superdoc-page').get_by_text('观察1', exact=True).first.click()
        page.keyboard.press('Tab')
        assert 'tableCell' in inspect_frame(frame)['selection']['ancestors']
        page.keyboard.press('Enter')
        assert 'tableCell' in inspect_frame(frame)['selection']['ancestors']
        page.keyboard.press('Control+z')
        page.keyboard.insert_text('补充')
        page.get_by_role('button', name='保存副本', exact=True).click()
        page.get_by_role('dialog', name='生成验证副本', exact=True).wait_for()
        page.get_by_role('dialog').get_by_role('button', name='取消', exact=True).click()
        service = api._floating.paged
        plan = service.plan(service.document_id, service.revision)['plan']
        assert plan['id'] in service.publisher.plans
        api._floating.runtime = InstructionRuntime(api._floating.document['version'])
        api._floating.runtime.submit('own', '只改当前内容')
        service.proposals['old'] = {'id': 'old', 'status': 'pending', 'draft_revision': service.revision, 'instruction_revision': 1}
        frame.locator('#frame-error').evaluate("node=>node.textContent='另一项提示保留'")
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        page.get_by_role('button', name='引用整表，最多8个段落', exact=True).click()
        page.locator('#paged-reference-chip').wait_for()
        frame.locator('.superdoc-page').get_by_text('121 ms', exact=True).first.click()
        before = inspect_frame(frame)
        frame.evaluate("document.querySelector('.ProseMirror').dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,data:'正文'}))")
        frame.get_by_role('button', name='在表格后继续正文', exact=True).click()
        assert inspect_frame(frame)['blocks'] == before['blocks']
        assert '请先完成当前输入' in frame.locator('#frame-error').text_content()
        frame.evaluate("document.querySelector('.ProseMirror').dispatchEvent(new CompositionEvent('compositionend',{bubbles:true,data:''}))")
        frame.get_by_role('button', name='在表格后继续正文', exact=True).click()
        assert '请先完成当前输入' not in frame.locator('#frame-error').text_content()
        assert frame.locator('#frame-error').text_content() == '另一项提示保留'
        page.locator('#paged-reference-chip').wait_for(state='hidden')
        deadline = time.monotonic()+5
        while service.proposals['old']['status'] != 'stale':
            if time.monotonic()>deadline:
                raise AssertionError(service.state())
            page.wait_for_timeout(50)
        assert plan['id'] not in service.publisher.plans
        assert not list(source.parent.glob('*.agent-*.docx')) and not errors


@pytest.mark.parametrize('case', ['merged', 'nested'])
def test_unsupported_tables_keep_existing_readonly_policy(tmp_path, case):
    source = table_fixture(tmp_path/(case+'.docx'), case)
    assert not eligibility(source.read_bytes())['editable']
