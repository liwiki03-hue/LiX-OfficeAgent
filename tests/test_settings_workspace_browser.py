import json

from tests.test_daily_browser import setup_daily
from tests.test_professional_browser import STATIC, _browser
from tests.test_ui_simplify_browser import simplify


OUT = STATIC.parents[2] / 'deliverables/settings-skills-s1-20261113'


def test_settings_skills_basic_real_store_flow(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        app = setup_daily(page, api, original, calls, receipts, monkeypatch)
        simplify(page)
        names = {'list_skills', 'get_skill', 'plan_skill_change', 'confirm_skill_change',
                 'plan_skill_import', 'plan_skill_rollback', 'confirm_skill_rollback', 'get_settings'}
        skill_calls = []

        def invoke(method, args):
            assert method in names
            skill_calls.append(method)
            return getattr(api, method)(*args)

        page.expose_function('skillsHost', invoke)
        page.evaluate('''names => {
          const previous = window.pywebview.api;
          window.pywebview.api = new Proxy({}, {get: (_, name) => names.includes(name) ? (...args) => window.skillsHost(name, args) : previous[name]});
        }''', sorted(names))
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.evaluate('''() => { const base = document.createElement('base'); base.href = 'http://lix.invalid/'; document.head.prepend(base); }''')
        page.route('http://lix.invalid/vendor/lucide/file-text.svg', lambda route: route.fulfill(path=str(STATIC / 'vendor/lucide/file-text.svg'), content_type='image/svg+xml'))
        page.add_script_tag(path=str(STATIC / 'skills.js'))
        page.add_script_tag(path=str(STATIC / 'settings_editor.js'))
        page.add_style_tag(content=(STATIC / 'settings_workspace.css').read_text('utf-8'))
        page.add_script_tag(path=str(STATIC / 'settings_workspace.js'))
        page.set_viewport_size({'width':1120, 'height':760})
        page.locator('#prompt').fill('保留未发送的文档修改要求')
        page.locator('#nav-settings').click()
        page.wait_for_function('!document.getElementById("settings-provider").disabled')
        page.locator('#tab-skills').click()
        page.wait_for_function('document.getElementById("skills-status").textContent.includes("尚未选择文档")')
        assert page.locator('#skills-import').is_disabled()
        page.locator('#skills-discover').click()
        assert page.locator('#skills-discovery').is_visible()
        assert not page.locator('#skills-discovery .skill-entry').count()
        app.path, app.editor = original.path, original.editor
        page.evaluate('''() => {
          document.getElementById('file-name').textContent = '自建技能验收.docx';
          window.dispatchEvent(new CustomEvent('officeapp:state', {detail:{hasFile:true}}));
        }''')
        page.locator('#skills-installed').click()
        page.wait_for_function('document.getElementById("skills-status").textContent.includes("还没有技能")')
        page.locator('#skills-add-menu > summary').click()
        page.locator('#skills-new').click()
        page.locator('#skills-name').fill('技术报告措辞核对')
        page.locator('#skills-description').fill('修改技术报告时保留数值和限定条件')
        page.locator('#skills-body').fill('先定位明确范围，再核对数字、单位与否定；所有写入仍需用户确认。')
        page.evaluate('window.approved = false')
        page.locator('#tab-machine').click()
        assert page.locator('#skills-editor').is_visible()
        page.locator('#settings-close').click()
        assert page.locator('#settings-dialog').is_visible()
        page.locator('#skills-create').click()
        page.wait_for_function('document.getElementById("skills-panel").getAttribute("aria-busy")==="false"')
        assert api.list_skills()['skills'] == []
        page.evaluate('window.approved = true')
        page.locator('#skills-create').click()
        page.wait_for_function('document.querySelector("#skills-list .skill-entry")')
        assert len(api.list_skills()['skills']) == 1
        assert page.locator('#skills-editor').is_hidden()
        page.locator('#skills-search').fill('不存在')
        assert '没有匹配' in page.locator('#skills-list').inner_text()
        page.locator('#skills-search').fill('数值')
        assert page.locator('#skills-list .skill-entry').count() == 1
        page.locator('#skills-list .skill-entry').click()
        page.wait_for_function('document.getElementById("skills-detail-body").textContent.includes("限定") || document.getElementById("skills-detail-body").textContent.includes("先定位")')
        page.locator('#skills-toggle').click()
        page.wait_for_function('document.querySelector("#skills-list .skill-entry")?.textContent.includes("已停用")')
        assert api.list_skills()['skills'][0]['enabled'] is False
        page.locator('#skills-rollback').click()
        page.wait_for_function('document.querySelector("#skills-list .skill-entry")?.textContent.includes("已启用")')
        page.locator('#skills-list .skill-entry').click()
        page.wait_for_function('!document.getElementById("skills-open-editor").disabled')
        page.locator('#skills-open-editor').click()
        page.locator('#skills-name').fill('保留尚未保存的技能草稿')
        page.evaluate('document.getElementById("file-dir").textContent = "changed-test-workspace"')
        page.wait_for_function('!document.getElementById("skills-recover-draft").hidden')
        assert page.locator('#skills-editor').is_hidden()
        page.locator('#skills-recover-draft').click()
        assert page.locator('#skills-name').input_value() == '保留尚未保存的技能草稿'
        assert page.locator('#skills-create').is_visible() and page.locator('#skills-edit').is_hidden()
        page.locator('#skills-editor-cancel').click()
        assert api.list_skills()['skills'][0]['name'] == '技术报告措辞核对'
        page.locator('#settings-search').fill('主题')
        page.locator('.settings-search-result').first.click()
        assert page.locator('#settings-theme-choice').is_visible()
        assert page.locator('#settings-pane-model').is_hidden()
        output = OUT / 'screenshots'
        output.mkdir(parents=True, exist_ok=True)
        for key in ['machine', 'model', 'rag', 'account', 'updates', 'skills']:
            page.locator('#tab-' + key).click()
            page.screenshot(path=str(output / ('desktop-' + key + '.png')))
        for width, theme in [(1440,'dark'), (1120,'light'), (900,'highcontrast'), (520,'dark')]:
            page.set_viewport_size({'width':width, 'height':800})
            page.evaluate('theme => document.documentElement.dataset.theme = theme', theme)
            assert page.locator('#settings-close').is_visible()
            assert page.evaluate('document.querySelector(".settings-panes").scrollWidth <= document.querySelector(".settings-panes").clientWidth + 1')
            assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
            page.screenshot(path=str(output / f'skills-{width}-{theme}.png'))
        assert page.evaluate('new Set([...document.querySelectorAll("[id]")].map(n=>n.id)).size===document.querySelectorAll("[id]").length')
        page.keyboard.press('Escape')
        assert page.locator('#settings-dialog').is_hidden()
        assert page.locator('#prompt').input_value() == '保留未发送的文档修改要求'
        assert not errors
        assert not [name for name in calls if name in {'ask', 'test_model_connection', 'confirm_image_generation'}]
        (OUT / 'BROWSER_FLOW.json').write_text(json.dumps({'skills_calls':skill_calls, 'page_errors':errors,
            'real_workspace_store':True, 'native_windows':False, 'real_model':False,
            'explicit_confirm_dialog_response':'controlled browser fixture', 'installed':len(api.list_skills()['skills'])}, ensure_ascii=False, indent=2), 'utf-8')
    finally:
        if 'errors' in locals():
            OUT.mkdir(parents=True, exist_ok=True)
            (OUT / 'BROWSER_ERRORS.json').write_text(json.dumps(errors, ensure_ascii=False), 'utf-8')
        fixture.close()
