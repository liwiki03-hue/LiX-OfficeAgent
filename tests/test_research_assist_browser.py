from test_professional_browser import _browser, STATIC
from test_research_assist import Model


def exercise_assist_browser(session, capture=None, *, tracked=False):
    page, api, app, calls, receipts = session
    model = Model()
    api._research_assist.model_factory = lambda: model
    methods = {'read_research_assist', 'start_research_assist', 'research_assist_status', 'cancel_research_assist',
               'regenerate_research_suggestion', 'plan_assisted_edit', 'confirm_assisted_edit', 'rollback_assisted_edit',
               'plan_research_task_save', 'confirm_research_task_save', 'restore_research_task', 'rollback_research_task'}

    def invoke(method, args):
        assert method in methods
        calls.append(method)
        result = getattr(api, method)(*args)
        receipts.append((method, result))
        return result

    page.expose_function('researchHost', invoke)
    page.evaluate('''() => {
      window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.researchHost(method, args)});
      document.getElementById('assist-panel').open = true;
    }''')
    page.add_script_tag(path=str(STATIC / 'research_assist.js'))
    assert not calls and not model.inputs
    page.locator('#assist-load').click()
    page.locator('#assist-targets input').first.check()
    page.locator('#assist-question').fill('Clarify the original observation while preserving measurements.')
    page.locator('#assist-start').click()
    page.locator('#assist-suggestions fieldset textarea').wait_for()
    assert len(model.inputs) == 1
    text = page.locator('#assist-suggestions fieldset textarea').first
    text.fill(text.input_value() + ' Human edit.')
    page.locator('#assist-suggestions input[value=accept]').first.check()
    page.locator('#assist-plan').click()
    assert '逐项完成审阅' in page.locator('#assist-status').inner_text()
    page.locator('#assist-suggestions input[value=review]').first.check()
    if tracked:
        page.locator('#assist-tracked').check()
    page.locator('#assist-plan').click()
    page.wait_for_function('!document.getElementById("assist-confirm").disabled')
    assert '修改前：' in page.locator('#assist-preview').inner_text()
    assert '修改后：' in page.locator('#assist-preview').inner_text()
    assert ('原生 Word 插入/删除修订' if tracked else '保护检查不证明语义等价') in page.locator('#assist-preview').inner_text()
    assert not page.locator('#assist-preview details').evaluate('(node) => node.open')
    page.evaluate('window.approved = false')
    page.locator('#assist-confirm').click()
    page.wait_for_function('document.getElementById("assist-status").textContent.includes("已取消")')
    assert not list(app.path.parent.glob('*modified*.docx'))
    page.evaluate('window.approved = true')
    page.locator('#assist-plan').click()
    page.wait_for_function('!document.getElementById("assist-confirm").disabled')
    page.locator('#assist-confirm').click()
    page.wait_for_function('document.getElementById("assist-output").textContent.includes("已核验副本")')
    if capture:
        capture(page, calls, model)
    for theme in ('dark', 'light', 'highcontrast'):
        page.set_viewport_size({'width': 520, 'height': 900})
        page.evaluate('(value) => document.documentElement.dataset.theme = value', theme)
        assert page.locator('#assist-panel').evaluate('(node) => node.scrollWidth <= node.clientWidth')
    page.locator('#assist-undo').click()
    page.wait_for_function('document.getElementById("assist-output").textContent.includes("已撤销")')
    page.evaluate('''() => { window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated')); }''')
    assert page.locator('#assist-suggestions').inner_text() == '' and page.locator('#assist-question').input_value() == ''
    assert page.locator('#assist-start').is_disabled()


def test_assist_browser_review_cancel_copy_undo_and_logout(tmp_path, monkeypatch):
    session = _browser(tmp_path, monkeypatch)
    try:
        exercise_assist_browser(next(session))
    finally:
        session.close()


def test_changed_context_clears_reviews_and_disables_old_plan(tmp_path, monkeypatch):
    def changed_context(page, calls, model):
        assert page.locator('#assist-suggestions input[value=review]').first.is_checked()
        page.locator('#assist-terms').fill('New protected phrase')
        assert not page.locator('#assist-suggestions input[value=review]').first.is_checked()
        assert page.locator('#assist-plan').is_disabled()
        assert '要求已改变' in page.locator('#assist-status').inner_text()

    session = _browser(tmp_path, monkeypatch)
    try:
        exercise_assist_browser(next(session), changed_context)
    finally:
        session.close()
