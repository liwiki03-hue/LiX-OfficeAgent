import copy
import json
import re

from tests.test_daily_browser import setup_daily
from tests.test_professional_browser import STATIC,_browser
from tests.test_public_skill_install import candidate
from tests.test_ui_simplify_browser import simplify


def test_curated_browser_has_real_host_catalog_and_fixed_inspection(tmp_path,monkeypatch):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page)
        assets=[]
        def asset(route):
            name=route.request.url.removeprefix('http://lix-preview.local/')
            if re.fullmatch(r'vendor/lucide/[a-z0-9-]+\.svg',name) and (STATIC/name).is_file():
                assets.append(name)
                route.fulfill(body=(STATIC/name).read_bytes(),content_type='image/svg+xml')
            else:route.abort()
        page.route('http://lix-preview.local/**',asset)
        page.evaluate('''()=>{const base=document.createElement('base');base.href='http://lix-preview.local/';document.head.prepend(base);}''')
        names={'list_skills','list_featured_skills','resolve_public_skill','get_settings'}
        invoked=[]
        def invoke(name,args):
            assert name in names
            invoked.append(name)
            return getattr(api,name)(*args)
        page.expose_function('libraryHost',invoke)
        page.evaluate('''names=>{const prior=window.pywebview.api;window.pywebview.api=new Proxy({},
            {get:(_,name)=>names.includes(name)?(...args)=>window.libraryHost(name,args):prior[name]});}''',sorted(names))
        page.add_script_tag(path=str(STATIC/'skills.js'))
        page.add_style_tag(content=(STATIC/'settings_workspace.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC/'settings_workspace.js'))
        page.add_style_tag(content=(STATIC/'compact_surfaces.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC/'skill_discovery.js'))
        page.locator('#nav-settings').click();page.locator('#tab-skills').click();page.locator('#skills-discover').click()
        page.wait_for_function('document.querySelectorAll("#skills-public-results .skill-entry").length===6')
        assert '离线来源快照' in page.locator('#skills-public-status').inner_text()
        assert invoked.count('list_featured_skills')==1 and 'resolve_public_skill' not in invoked
        page.wait_for_function('getComputedStyle(document.querySelector("#skills-public-results .skill-entry-icon"),"::before").maskImage.includes("file-text.svg")')
        page.wait_for_timeout(150)
        assert assets
        out=STATIC.parents[2]/'deliverables/skills-discovery-library-20261113'
        page.screenshot(path=str(out/'curated-library.png'))
        page.locator('#skills-search').fill('方案')
        assert page.locator('#skills-public-results .skill-entry').count()==1
        page.locator('#skills-public-results .skill-entry').click()
        assert candidate()['origin']['commit'] in page.locator('#skills-public-detail').inner_text()
        monkeypatch.setattr('office_agent.webui.public_skill_resolver.resolve',lambda *args,**kwargs:copy.deepcopy(candidate()))
        page.get_by_role('button',name='检查来源与兼容性',exact=True).click()
        page.wait_for_selector('#skills-public-detail button:text-is("预览安装到当前工作区")')
        assert invoked.count('resolve_public_skill')==1
        assert not (tmp_path/'skills').exists()
        (out/'BROWSER.json').write_text(json.dumps({'native':False,'cloud_model':False,'host_calls':invoked,
            'curated_entries':6,'offline_catalog':True,'fixed_source_inspection':True,
            'installation_not_implicitly_started':True,'source_resolution':'controlled fixed MIT fixture',
            'local_icon_assets':assets},ensure_ascii=False,indent=2),'utf-8')
    finally:fixture.close()
