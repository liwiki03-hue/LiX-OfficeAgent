import json

from tests.test_daily_browser import setup_daily
from tests.test_professional_browser import STATIC, _browser
from tests.test_skill_plugin_sources import plugin_reader, presentation_package
from tests.test_ui_simplify_browser import simplify
from office_agent.webui.public_skill_resolver import resolve


def test_plugin_source_switch_search_and_installed_display_are_host_bound(tmp_path,monkeypatch):
    fixture = _browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts = next(fixture)
        app = setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page)
        app.path,app.editor = original.path,original.editor
        package = presentation_package()
        plan = api._skills.plan_public(package)
        assert api.confirm_skill_change(plan['plan']['plan_id'],True)['result']['verified']
        read,requests = plugin_reader(dependencies=True)
        monkeypatch.setattr('office_agent.webui.public_skill_resolver.read_api',read)
        monkeypatch.setattr('office_agent.webui.public_skill_resolver.resolve',lambda *args,**kwargs:resolve(*args,**{**kwargs,'reader':read}))
        names = {'list_skills','get_skill','get_settings','list_featured_skills','search_plugin_skills','resolve_public_skill'}
        invoked = []
        def invoke(name,args):
            assert name in names
            invoked.append(name)
            return getattr(api,name)(*args)
        page.expose_function('pluginHost',invoke)
        page.evaluate('''names=>{const prior=window.pywebview.api;window.pywebview.api=new Proxy({},
          {get:(_,name)=>names.includes(name)?(...args)=>window.pluginHost(name,args):prior[name]});}''',sorted(names))
        page.add_script_tag(path=str(STATIC/'skills.js'))
        page.add_style_tag(content=(STATIC/'settings_workspace.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC/'settings_workspace.js'))
        page.add_script_tag(path=str(STATIC/'skill_discovery.js'))
        page.locator('#nav-settings').click();page.locator('#tab-skills').click()
        page.wait_for_selector('#skills-list .skill-entry')
        page.locator('#skills-list .skill-entry').click()
        page.wait_for_selector('#skills-detail-icon img')
        assert page.locator('#skills-detail-title').inner_text()=='Report helper'
        assert page.locator('#skills-detail-icon img').get_attribute('src')==package['icon']['data_url']
        assert '展示信息与图标' in page.locator('#skills-public-management').inner_text()
        out = STATIC.parents[2]/'deliverables/skills-plugin-sources-20261113'
        out.mkdir(exist_ok=True)
        page.screenshot(path=str(out/'installed-display.png'))
        page.locator('#skills-discover').click()
        page.wait_for_selector('#skills-public-results .skill-entry')
        page.locator('#skills-search').fill('reports')
        page.locator('#skills-public-provider').select_option('openai/plugins')
        assert page.locator('#skills-search').input_value()=='reports' and not requests
        page.locator('#skills-search').press('Enter')
        page.wait_for_selector('#skills-public-results .skill-entry')
        assert invoked.count('search_plugin_skills')==1
        page.locator('#skills-public-results .skill-entry').click()
        page.get_by_role('button',name='检查来源与兼容性',exact=True).click()
        page.wait_for_function('document.getElementById("skills-public-detail").textContent.includes("当前无法完整安装")')
        assert 'mcpServers' in page.locator('#skills-public-detail').inner_text()
        assert page.get_by_role('button',name='预览安装到当前工作区',exact=True).count()==0
        page.screenshot(path=str(out/'plugin-requirements.png'))
        page.locator('#skills-search').fill('保留的查询草稿')
        page.locator('#skills-public-provider').select_option('skills.sh')
        assert page.locator('#skills-search').input_value()=='保留的查询草稿'
        assert invoked.count('search_plugin_skills')==1 and invoked.count('resolve_public_skill')==1
        (out/'BROWSER.json').write_text(json.dumps({'native':False,'live_network':False,'cloud_model':False,
            'data':'synthetic plugin and raster; real host SkillsStore and headless Edge',
            'host_calls':invoked,'source_requests':requests,'provider_switch_does_not_search':True,
            'query_preserved':True,'display_reopened_from_v4':True,'mcp_install_blocked':True,
            'implicit_install_or_model_requests':0},ensure_ascii=False,indent=2),'utf-8')
    finally:
        fixture.close()
