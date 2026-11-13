import json

from tests.test_daily_browser import setup_daily
from tests.test_professional_browser import STATIC, _browser
from tests.test_skill_references_updates import resolved
from tests.test_ui_simplify_browser import simplify


def test_public_reference_update_cancel_confirm_and_details(tmp_path, monkeypatch):
    fixture = _browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts = next(fixture)
        app = setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page)
        app.path,app.editor = original.path,original.editor
        previous = resolved(commit='a'*40)
        following = resolved(commit='b'*40)
        install = api._skills.plan_public(previous)
        assert api.confirm_skill_change(install['plan']['plan_id'],True)['result']['verified']
        monkeypatch.setattr('office_agent.webui.public_skill_resolver.resolve',lambda *args,**kwargs:dict(following))
        names = {'list_skills','get_skill','check_public_skill_update','plan_public_skill_update','confirm_skill_change','get_settings'}
        invoked = []
        def invoke(name,args):
            assert name in names
            invoked.append(name)
            return getattr(api,name)(*args)
        page.expose_function('updateHost',invoke)
        page.evaluate('''names=>{const prior=window.pywebview.api;window.pywebview.api=new Proxy({},
            {get:(_,name)=>names.includes(name)?(...args)=>window.updateHost(name,args):prior[name]});}''',sorted(names))
        page.add_script_tag(path=str(STATIC/'skills.js'))
        page.add_style_tag(content=(STATIC/'settings_workspace.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC/'settings_workspace.js'))
        page.locator('#nav-settings').click()
        page.locator('#tab-skills').click()
        page.wait_for_selector('#skills-list .skill-entry')
        page.locator('#skills-list .skill-entry').click()
        page.wait_for_selector('#skills-public-management')
        assert '包内参考 · 1 项' in page.locator('#skills-public-management').inner_text()
        page.get_by_role('button',name='检查更新',exact=True).click()
        page.wait_for_selector('[data-update]')
        assert 'a'*40 in page.locator('[data-update]').inner_text()
        assert 'b'*40 in page.locator('[data-update]').inner_text()
        page.evaluate('window.approved=false')
        page.get_by_role('button',name='预览更新',exact=True).click()
        page.wait_for_function('!document.getElementById("skills-panel").matches("[aria-busy=true]")')
        assert api.get_skill(previous['slug'],'user')['skill']['origin']['commit']=='a'*40
        page.evaluate('window.approved=true')
        page.get_by_role('button',name='预览更新',exact=True).click()
        page.wait_for_function('!document.getElementById("skills-list").hidden && !!document.querySelector("#skills-list .skill-entry")')
        page.locator('#skills-list .skill-entry').click()
        page.wait_for_function('document.getElementById("skills-detail-meta").textContent.includes("'+('b'*40)+'")')
        assert invoked.count('plan_public_skill_update')==2 and invoked.count('check_public_skill_update')==1
        out = STATIC.parents[2]/'deliverables/skills-s3-references-20261113'
        out.mkdir(exist_ok=True)
        page.screenshot(path=str(out/'reference-update.png'))
        (out/'BROWSER.json').write_text(json.dumps({'host_calls':invoked,'native':False,'real_network':False,
            'confirmation':'controlled browser confirm', 'cancel_preserved_old_version':True,
            'confirmed_version':following['origin']['commit'], 'reference_preview':True},ensure_ascii=False,indent=2),'utf-8')
    finally:
        fixture.close()
