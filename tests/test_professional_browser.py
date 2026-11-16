from pathlib import Path
from types import SimpleNamespace
import re

from docx import Document
import pytest

from office_agent.webui.bridge import Api
from office_agent.writes import Editor, digest
from test_professional_edits import fake_render
from tests.ui_confirmation import host_diagnostics,record_browser_failure

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'
from office_agent.browser_runtime import find_edge

EDGE = find_edge() or Path('__LiX_Edge_not_installed__')


def _browser(tmp_path, monkeypatch, *, real_render=False):
    if not EDGE.is_file():
        pytest.skip('Requires the existing Windows Edge; no browser runtime is installed')
    from playwright.sync_api import sync_playwright
    path = tmp_path / 'experiment.docx'
    doc = Document()
    doc.add_heading('合成技术实验报告', 1)
    p = doc.add_paragraph()
    p.add_run('样本30个，误差为').bold = True
    p.add_run('12.5% [1]，计算采用 $x=2$。').italic = True
    doc.add_paragraph('Unchanged paragraph and unchanged styling.')
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).merge(table.cell(0, 1)).text = 'Merged header'
    table.cell(1, 0).text = '20 ms'
    table.cell(1, 1).text = 'Non-target result'
    doc.save(path)
    app = SimpleNamespace(path=path, editor=Editor(tmp_path, require_confirmation=True, render_validation=True),
                          busy=False, window=None, _build_model=lambda: None, _clear_private_state=lambda: None)
    authenticated = [True]
    auth = SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': authenticated[0]})
    api = Api(app, auth=auth, settings_root=tmp_path)
    calls, receipts, corpus_trace = [], [], []
    from office_agent.research_corpus import ResearchCorpus
    inventory=ResearchCorpus.inventory
    def inspect_corpus(corpus,name):
        try:
            value=inventory(corpus,name)
            corpus_trace.append({'hash':value['hash'],'inventory':value['inventory'],
                'snapshot':corpus.store.snapshot(),'identity':corpus.rag._identity()})
            return value
        except Exception as error:
            corpus_trace.append({'error':type(error).__name__,'code':getattr(error,'code',None)})
            raise
    monkeypatch.setattr(ResearchCorpus,'inventory',inspect_corpus)
    allowed = {'enable_professional_tools', 'read_professional_structure', 'plan_professional_edit',
               'confirm_professional_edit', 'rollback_professional_edit', 'get_rag_snapshot',
               'plan_research_index', 'plan_research_query', 'confirm_rag_index', 'confirm_rag_query',
               'read_rag_evidence', 'get_rag_evidence', 'plan_rag_rollback', 'confirm_rag_rollback',
               'get_research_status', 'get_research_context'}
    allowed |= {'get_research_corpus_status', 'plan_research_corpus', 'confirm_research_corpus', 'plan_research_corpus_query'}
    def invoke(method, args):
        assert method in allowed
        calls.append(method)
        result = getattr(api, method)(*args)
        receipts.append((method, result))
        return result
    if not real_render:
        monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    monkeypatch.setenv('TEMP', str(tmp_path))
    monkeypatch.setenv('TMP', str(tmp_path))
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1100, 'height': 900})
            page._lix_case=tmp_path.name
            page._lix_api_receipts=receipts
            page._lix_page_errors=[]
            page._lix_host_diagnostics=lambda:{**host_diagnostics(api,app),'corpus_trace':corpus_trace[-12:]}
            page.on('pageerror',lambda error:page._lix_page_errors.append(str(error)))
            page.set_default_timeout(65000 if real_render else 15000)
            page.route('**/*', lambda route: route.abort())
            page.expose_function('professionalHost', invoke)
            markup = (STATIC / 'index.html').read_text(encoding='utf-8')
            html = re.sub(r'<script\b[^>]*>.*?</script>', '', markup, flags=re.S)
            page.set_content(html)
            if 'href="design_tokens.css"' in html:
                page.add_style_tag(content=(STATIC / 'design_tokens.css').read_text(encoding='utf-8'))
            page.add_style_tag(content=(STATIC / 'theme.css').read_text(encoding='utf-8'))
            page.evaluate('''() => {
              document.getElementById('workspace').hidden = false; document.getElementById('workspace').inert = false;
              document.getElementById('workspace').setAttribute('aria-hidden', 'false'); document.getElementById('auth-screen').hidden = true;
              document.getElementById('professional-panel').open = true;
              document.getElementById('file-name').textContent = 'experiment.docx';
              window.authenticated = true; window.approved = true;
              window.confirm = () => window.approved;
              window.OfficeAuth = {isAuthenticated: () => window.authenticated, wrapApi: value => value};
              window.pywebview = {api: new Proxy({}, {get: (_, method) => (...args) => window.professionalHost(method, args)})};
            }''')
            if 'src="confirm_dialog.js"' in markup:
                page.add_script_tag(path=str(STATIC / 'confirm_dialog.js'))
                page.evaluate('''() => {
                  const trace=window.__lixConfirmationTrace=[],add=value=>{trace.push({...value,at:performance.now()});if(trace.length>80)trace.shift();};
                  const ask=OfficeConfirm.ask;
                  OfficeConfirm.ask=async(...args)=>{add({kind:'request',message:String(args[0]).slice(0,1000)});const accepted=await ask(...args);add({kind:'resolved',accepted});return accepted;};
                  for(const event of ['close','cancel','officeauth:authenticated','officeauth:unauthenticated','officeworkspace:clear'])document.addEventListener(event,()=>add({kind:event}),true);
                  for(const id of ['file-name','file-dir']){const node=document.getElementById(id);if(node)new MutationObserver(()=>add({kind:'identity',id,text:node.textContent})).observe(node,{childList:true,characterData:true,subtree:true});}
                }''')
            page.add_script_tag(path=str(STATIC / 'professional.js'))
            from tests.ui_confirmation import drive_legacy_confirmations
            page._lix_real_render = real_render
            drive_legacy_confirmations(page, monkeypatch)
            yield page, api, app, calls, receipts
        finally:
            if 'page' in locals() and page._lix_case.startswith(('test_read_continue_export','test_codex_layout_controls','test_two_group_review','test_real_render_page_review','test_real_return_workflow','test_real_table_review','test_actual_directory_incremental')):
                try:record_browser_failure(page,'diagnostic fixture exit',0)
                except Exception:pass
            browser.close()


@pytest.fixture
def professional_browser(tmp_path, monkeypatch):
    yield from _browser(tmp_path, monkeypatch)


def prepare_plan(page):
    page.locator('#professional-load').click()
    page.wait_for_function("document.getElementById('professional-targets').children.length > 1")
    page.locator('#professional-targets input').nth(1).check()
    text = page.locator('#professional-text').input_value()
    page.locator('#professional-text').fill(text.replace('误差为', '测得误差为'))
    page.locator('#professional-reason').fill('Improve wording while preserving all original measurements and citations.')
    page.locator('#professional-add').click()
    page.locator('#professional-plan').click()
    page.wait_for_function("!document.getElementById('professional-confirm').disabled")


def test_real_ui_plan_separate_confirmation_save_and_rollback(professional_browser):
    page, api, app, calls, receipts = professional_browser
    version = digest(app.path)
    assert calls == []
    page.locator('#professional-enable').check()
    page.wait_for_function("document.getElementById('professional-status').textContent.includes('专业工具已绑定')")
    assert app._professional_tools_source == (app.path, id(app.editor), version)
    assert calls == ['enable_professional_tools']
    prepare_plan(page)
    assert not app.editor.audit.exists()
    assert '当前编辑目标：body:1/paragraph' in page.locator('#professional-selection').inner_text()
    assert '非原生修订' in page.locator('#professional-preview').inner_text()
    page.locator('#professional-confirm').click()
    page.wait_for_function("document.getElementById('professional-output').textContent.includes('已保存副本')")
    saved = next(reply['result'] for method, reply in receipts if method == 'confirm_professional_edit')
    output = Path(saved['output'])
    assert output.exists() and Document(output).paragraphs[1].text.startswith('样本30个，测得误差为')
    assert digest(app.path) == version and saved['source_unchanged']
    assert '渲染状态：different' in page.locator('#professional-output').inner_text()
    page.locator('#professional-rollback').click()
    page.wait_for_function("document.getElementById('professional-output').textContent.includes('已撤销')")
    assert not output.exists() and digest(app.path) == version


def test_edit_after_plan_cancel_and_logout_clear(professional_browser):
    page, api, app, calls, receipts = professional_browser
    prepare_plan(page)
    page.locator('#professional-text').fill('样本31个，误差为12.5% [1]，计算采用 $x=2$。')
    assert page.locator('#professional-confirm').is_disabled()
    page.locator('#professional-add').click()
    page.locator('#professional-plan').click()
    page.wait_for_function("document.getElementById('professional-status').textContent.includes('整组修改已拒绝')")
    assert not app.editor.audit.exists()
    assert 'confirm_professional_edit' not in calls
    page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert page.locator('#professional-text').input_value() == ''
    assert page.locator('#professional-targets').inner_text() == ''
    assert page.locator('#professional-confirm').is_disabled()


def test_research_ui_uses_persistent_profile_and_separate_evidence_confirmation(professional_browser):
    from office_agent.rag_workspace import WorkspaceRag
    from test_rag_workspace import SyntheticEmbedder

    page, api, app, calls, receipts = professional_browser
    embedder = SyntheticEmbedder()
    api._rag._factory = lambda scope_check: WorkspaceRag(app.editor.space, embedder, scope_check=scope_check)
    page.evaluate("document.getElementById('settings-dialog').hidden = false; document.getElementById('settings-dialog').inert = false; document.getElementById('settings-pane-rag').hidden = false")
    page.add_script_tag(path=str(STATIC / 'rag.js'))
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('尚未建立')")
    page.locator('#rag-research').check()
    assert embedder.calls == []
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('已保存并校验')")
    assert calls[-2:] == ['plan_research_index', 'confirm_rag_index']
    assert len(embedder.calls) == 1
    page.locator('#rag-question').fill('样本30个')
    page.locator('#rag-query').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('三路候选各50')")
    assert calls[-2:] == ['plan_research_query', 'confirm_rag_query']
    assert len(embedder.calls) == 2 and len(embedder.calls[1]) == 1
    assert page.locator('#rag-hits').inner_text()
    evidence = next(reply['result']['evidence'] for method, reply in receipts if method == 'confirm_rag_query')
    page.locator('#rag-rollback-evidence').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('已撤销')")
    assert not (app.editor.space.root / evidence).exists()
