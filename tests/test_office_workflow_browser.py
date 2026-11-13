from copy import deepcopy
import hashlib
import json
from pathlib import Path
import shutil
import time
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui.office_workflows import candidate
from tests.test_model_connections import connection, payload, save
from tests.test_paged_browser import serve_static, engine_frame
from tests.test_professional_browser import EDGE, STATIC


BEFORE = '25 ℃条件下，30个样本延迟为120 ms，尚未验证高温条件[1]。'
AFTER = '在25 ℃条件下，30个样本的延迟为120 ms，高温条件尚未验证[1]。'
OUT = STATIC.parents[2]/'deliverables/skills-office-workflow-20261113'


def make_document(path):
    book = Document()
    book.add_heading('自建技术报告',0)
    book.add_paragraph(BEFORE)
    paragraph = book.add_paragraph('非目标说明与格式保留。')
    paragraph.runs[0].bold = True
    table = book.add_table(rows=2,cols=2)
    table.cell(0,0).text='条件'
    table.cell(1,1).text='120 ms'
    book.sections[0].header.paragraphs[0].text='自建样例页眉'
    book.sections[0].footer.paragraphs[0].text='自建样例页脚'
    book.save(path)


def wait_for(page,condition,label,seconds=20):
    deadline = time.monotonic()+seconds
    while not condition():
        if time.monotonic()>deadline:
            raise AssertionError(label+': '+str(page.locator('#floating-status').text_content()))
        page.wait_for_timeout(80)


def test_lix_workflow_install_real_editor_suggestion_and_verified_copy(tmp_path,monkeypatch):
    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    from playwright.sync_api import sync_playwright
    OUT.mkdir(exist_ok=True)
    source = tmp_path/'技术报告润色验收.docx'
    make_document(source)
    original = source.read_bytes()
    app = WebviewApp()
    app.attach(SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(source)]))
    api = FloatingApi(app,settings_root=tmp_path)
    save(api._connections,payload([connection()]))
    monkeypatch.setenv('TEMP',str(tmp_path));monkeypatch.setenv('TMP',str(tmp_path))
    model_messages,host_calls,read_receipts = [],[],[]
    class Model:
        def __init__(self):self.calls=0
        def bind_tools(self,tools):return self
        def invoke(self,messages):
            self.calls+=1;model_messages.append(deepcopy(messages))
            if self.calls==1:
                reference = next(iter(api._floating.paged.references.values()))
                return {'tool_calls':[{'name':'propose_draft_edit','id':'report-polish','args':{
                    'reference_id':reference['id'],'old':reference['old'],'new':AFTER,
                    'reason':'精简语序，保留温度、样本数、延迟、引文和未验证条件。'}}]}
            return {'content':'建议待审阅；尚未写入文档副本。'}
    model = Model()
    monkeypatch.setattr(api._connections,'build_model',lambda route:model)
    original_notice = api._floating.notice
    def notice(kind,value):
        if kind=='skill_loaded':read_receipts.append(deepcopy(value))
        original_notice(kind,value)
    monkeypatch.setattr(api._floating,'notice',notice)
    def invoke(name,args):
        if name in {'list_office_workflows','resolve_public_skill','plan_public_skill_install','confirm_skill_change','paged_submit','paged_prepare_apply','paged_applied','paged_plan','paged_publish'}:
            host_calls.append(name)
        return getattr(api,name)(*args)
    with serve_static() as url,sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width':1280,'height':850})
            page.set_default_timeout(20000)
            page.expose_function('workflowHost',invoke)
            page.add_init_script("""if(window===window.top){window.OfficeAuth={isAuthenticated:()=>true,ensure:async()=>true,wrapApi:x=>x};window.pywebview={api:new Proxy({}, {get:(_,name)=>(...args)=>window.workflowHost(name,args)})};window.approved=true;window.confirm=()=>window.approved;document.addEventListener('DOMContentLoaded',()=>{document.getElementById('workspace').hidden=false;document.getElementById('workspace').inert=false;document.getElementById('workspace').setAttribute('aria-hidden','false');});}""")
            page.route('**/auth.js',lambda route:route.fulfill(status=200,body=''))
            errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
            page.goto(url)
            page.locator('#floating-choose').click()
            wait_for(page,lambda:bool(api._floating.document),'file admission')
            page.get_by_role('button',name='同意读取并打开',exact=True).click()
            wait_for(page,lambda:bool(api._floating.paged.manifest),'real editor import and bootstrap',70)
            frame = engine_frame(page)
            page.locator('#paged-top > header > details > summary').click()
            page.locator('.floating-menu').get_by_role('button',name='设置',exact=True).click()
            page.locator('#tab-skills').click();page.locator('#skills-discover').click()
            page.locator('#skills-public-provider').select_option('lix')
            page.locator('#skills-search').fill('技术报告')
            assert 'list_office_workflows' not in host_calls and model.calls==0
            page.locator('#skills-public-search').click()
            page.wait_for_selector('#skills-public-results .skill-entry')
            page.locator('#skills-public-results .skill-entry').click()
            assert 'LiX 提供' in page.locator('#skills-public-detail').inner_text()
            page.get_by_role('button',name='检查来源与兼容性',exact=True).click()
            page.get_by_role('button',name='预览安装到当前工作区',exact=True).wait_for()
            page.evaluate('window.approved=false')
            page.get_by_role('button',name='预览安装到当前工作区',exact=True).click()
            wait_for(page,lambda:not page.locator('#skills-panel').get_attribute('aria-busy')=='true','install cancellation')
            assert not (source.parent/'skills/lix-report-polish/SKILL.md').exists() and model.calls==0
            page.evaluate('window.approved=true')
            page.get_by_role('button',name='预览安装到当前工作区',exact=True).click()
            page.wait_for_selector('#skills-list .skill-entry')
            page.locator('#skills-list .skill-entry').click()
            page.wait_for_function('document.getElementById("skills-detail-meta").textContent.includes("LiX 提供 @ 1.0.0")')
            page.screenshot(path=str(OUT/'workflow-installed.png'))
            page.locator('#settings-close').click()
            page.locator('#simple-model-menu > summary').click()
            page.locator('#model-selector').select_option(json.dumps(['official','same-model','agent'],separators=(',',':')))
            page.keyboard.press('Escape')
            frame.locator('.superdoc-page').get_by_text(BEFORE,exact=True).first.click()
            page.keyboard.press('Home');page.keyboard.press('Shift+End')
            page.get_by_role('button',name='引用选区',exact=True).click()
            wait_for(page,lambda:bool(api._floating.paged.references),'selection reference')
            assert next(iter(api._floating.paged.references.values()))['old']==BEFORE
            page.locator('#floating-input').fill('技术报告润色：精简当前引用段落，保留数字、单位、引文和未验证条件，不改标题、表格与其他文字。')
            page.locator('#floating-send').click()
            page.get_by_role('button',name='采用到草稿',exact=True).wait_for(timeout=25000)
            page.wait_for_selector('[data-skill-read="user:lix-report-polish"]')
            assert page.locator('[data-skill-read="user:lix-report-polish"]').count()==1
            assert '仅代表宿主准备参考' in page.locator('[data-skill-read="user:lix-report-polish"]').text_content()
            page.wait_for_function('!document.getElementById("paged-feedback").textContent.includes("将在下一安全步骤处理")')
            assert model.calls==2 and candidate()['body'] in model_messages[0][0]['content']
            assert source.read_bytes()==original and not list(source.parent.glob('*.agent-*.docx'))
            page.screenshot(path=str(OUT/'workflow-suggestion.png'))
            page.get_by_role('button',name='采用到草稿',exact=True).click()
            wait_for(page,lambda:next(iter(api._floating.paged.proposals.values()))['status']=='adopted','adopt suggestion')
            assert source.read_bytes()==original
            page.get_by_role('button',name='保存副本',exact=True).click()
            page.get_by_role('dialog',name='生成验证副本').get_by_role('button',name='取消',exact=True).click()
            assert not list(source.parent.glob('*.agent-*.docx'))
            page.get_by_role('button',name='保存副本',exact=True).click()
            page.get_by_role('button',name='确认生成副本',exact=True).wait_for()
            plan = deepcopy(next(iter(api._floating.paged.publisher.plans.values())))
            page.get_by_role('button',name='确认生成副本',exact=True).click()
            wait_for(page,lambda:bool(api._floating.outputs),'real render and copy publish',180)
            result = next(iter(api._floating.outputs.values()))
            output = Path(result['output']);reopened=Document(output)
            assert result['status']=='verified' and result['render_diff']['inputs_unchanged']
            assert reopened.paragraphs[1].text==AFTER
            assert reopened.paragraphs[0].text=='自建技术报告' and reopened.paragraphs[2].text=='非目标说明与格式保留。'
            assert reopened.paragraphs[2].runs[0].bold and reopened.tables[0].cell(1,1).text=='120 ms'
            assert reopened.sections[0].header.paragraphs[0].text=='自建样例页眉' and reopened.sections[0].footer.paragraphs[0].text=='自建样例页脚'
            assert source.read_bytes()==original and len(plan['changes'])==1
            page.screenshot(path=str(OUT/'workflow-copy.png'))
            shutil.copyfile(source,OUT/'原稿.docx');shutil.copyfile(output,OUT/'润色副本.docx')
            actual_audit = api._floating.paged.publisher.editor.audit/plan['id']/'audit.json'
            audit_data=json.loads(actual_audit.read_text('utf-8'))
            assert Path(audit_data['backup']).read_bytes()==original
            shutil.copyfile(actual_audit,OUT/'COPY_AUDIT.json')
            (OUT/'BROWSER_FLOW.json').write_text(json.dumps({'native':False,'real_service':False,'model':'controlled tool-call Model',
                'real_editor':'bundled SuperDoc in headless Edge','render':result['render_diff'],
                'skill_source':candidate()['origin'],'model_calls':model.calls,'actual_system_sha256':hashlib.sha256(model_messages[0][0]['content'].encode()).hexdigest(),
                'skill_in_actual_request':candidate()['body'] in model_messages[0][0]['content'],'read_receipts':read_receipts,
                'host_calls':host_calls,'install_cancelled_then_confirmed':True,'adopt_does_not_write_source':True,
                'publish_cancelled_then_confirmed':True,'approved_changes':[{key:value for key,value in row.items()}for row in plan['changes']],
                'source_sha256':hashlib.sha256(original).hexdigest(),'output_sha256':hashlib.sha256(output.read_bytes()).hexdigest(),
                'backup_equal_source':True,'non_target_checks':['title','paragraph text and bold','table','header','footer'],
                'page_errors':errors},ensure_ascii=False,indent=2),'utf-8')
            assert errors==[]
        finally:
            browser.close()
