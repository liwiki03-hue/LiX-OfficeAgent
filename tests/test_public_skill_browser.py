import copy
import json

from tests.test_daily_browser import setup_daily
from tests.test_professional_browser import STATIC, _browser
from tests.test_public_skill_install import candidate
from tests.test_ui_simplify_browser import simplify


def test_browser_public_inspect_cancel_install_and_reopen(tmp_path,monkeypatch):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        app=setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page)
        source=candidate()
        monkeypatch.setattr('office_agent.webui.public_skill_resolver.resolve',lambda *args,**kwargs:copy.deepcopy(source))
        allowed={'list_skills','get_skill','plan_public_skill_install','confirm_skill_change','resolve_github_skill',
                 'plan_skill_rollback','confirm_skill_rollback','get_settings'}
        invoked=[]
        def invoke(name,args):
            assert name in allowed
            invoked.append(name)
            return getattr(api,name)(*args)
        page.expose_function('publicHost',invoke)
        page.evaluate('''names=>{const previous=window.pywebview.api;window.pywebview.api=new Proxy({},
            {get:(_,name)=>names.includes(name)?(...args)=>window.publicHost(name,args):previous[name]});}''',sorted(allowed))
        page.add_script_tag(path=str(STATIC/'skills.js'))
        page.add_style_tag(content=(STATIC/'settings_workspace.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC/'settings_workspace.js'))
        page.add_style_tag(content=(STATIC/'compact_surfaces.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC/'skill_discovery.js'))
        page.locator('#nav-settings').click()
        page.locator('#tab-skills').click()
        page.locator('#skills-discover').click()
        page.locator('#skills-discovery > details > summary').click()
        page.locator('#skills-github-url').fill(source['origin']['url'])
        page.locator('#skills-github-check').click()
        page.wait_for_function('document.getElementById("skills-public-detail").textContent.includes("单独确认") || !![...document.querySelectorAll("#skills-public-detail button")].find(n=>n.textContent.includes("预览安装"))')
        install=page.get_by_role('button',name='预览安装到当前工作区',exact=True)
        install.click()
        page.wait_for_function('document.getElementById("skills-status").textContent.includes("尚未选择文档")')
        page.wait_for_function('document.getElementById("skills-public-status").textContent.includes("尚未选择文档")')
        app.path,app.editor=original.path,original.editor
        page.evaluate("document.getElementById('file-name').textContent='自建报告.docx'")
        page.locator('#skills-installed').click()
        page.wait_for_function('document.getElementById("skills-status").textContent.includes("还没有技能")')
        page.locator('#skills-discover').click()
        page.evaluate('window.approved=false')
        install.click()
        page.wait_for_function('!document.getElementById("skills-panel").matches("[aria-busy=true]")')
        assert not (tmp_path/'skills'/source['slug']/'SKILL.md').exists()
        page.evaluate('window.approved=true')
        install.click()
        page.wait_for_function('document.querySelector("#skills-list .skill-entry") && !document.getElementById("skills-list").hidden')
        assert len(api.list_skills()['skills'])==1
        page.locator('#skills-list .skill-entry').click()
        page.wait_for_function('document.getElementById("skills-detail-meta").textContent.includes("a6bdfd9")')
        assert 'MIT' in page.locator('#skills-detail-meta').inner_text()
        out=STATIC.parents[2]/'deliverables/skills-s3-20261113'
        page.screenshot(path=str(out/'installed-public.png'))
        (out/'BROWSER.json').write_text(json.dumps({'host_calls':invoked,'native':False,'real_network':False,
            'fixture':'pinned public package through real SkillsStore; controlled confirmation',
            'cancel_writes_nothing':True,'no_document_install_blocked':True,'installed':1},ensure_ascii=False,indent=2),'utf-8')
    finally:
        fixture.close()
