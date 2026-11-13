from office_agent.rag_workspace import WorkspaceRag
from test_rag_workspace import SyntheticEmbedder
from test_professional_browser import STATIC, professional_browser


def exercise_research_context_browser(browser, capture=None):
    page, api, app, calls, receipts = browser
    embedder = SyntheticEmbedder()
    api._rag._factory = lambda scope_check: WorkspaceRag(app.editor.space, embedder, scope_check=scope_check)
    page.evaluate("document.getElementById('settings-dialog').hidden = false; document.getElementById('settings-dialog').inert = false; document.getElementById('settings-pane-rag').hidden = false")
    page.evaluate("document.querySelectorAll('.settings-tab').forEach(node => { node.classList.toggle('active', node.id === 'tab-rag'); if (node.id === 'tab-rag') node.setAttribute('aria-current','true'); else node.removeAttribute('aria-current'); }); document.querySelectorAll('.settings-pane').forEach(node => node.hidden = node.id !== 'settings-pane-rag')")
    page.add_script_tag(path=str(STATIC / 'rag.js'))
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('尚未建立')")
    page.locator('#rag-research').check()
    assert calls == ['get_rag_snapshot'] and embedder.calls == []
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-profile').textContent.includes('未建立当前文件')")
    assert calls[-1] == 'get_research_status' and embedder.calls == []
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('已保存并校验')")
    assert len(embedder.calls) == 1
    snapshot = api._rag._store.snapshot()
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-profile').textContent.includes('可复用')")
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('未重复嵌入')")
    assert len(embedder.calls) == 1 and api._rag._store.snapshot() == snapshot
    page.locator('#rag-question').fill('样本30个')
    page.locator('#rag-query').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('三路候选各50')")
    assert len(embedder.calls) == 2
    page.locator('[data-rag-context]').first.click()
    page.wait_for_function("document.getElementById('rag-context').textContent.includes('只读上下文')")
    assert '命中所在段落/行' in page.locator('#rag-context').inner_text()
    assert len(embedder.calls) == 2 and calls[-1] == 'get_research_context'
    if capture is not None:
        capture(page, embedder, calls)
    page.locator('#rag-rollback-evidence').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('已撤销')")
    assert page.locator('#rag-context').inner_text() == ''
    page.evaluate("window.authenticated = false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert not page.locator('[data-rag-context]').count()
    assert page.locator('#rag-profile').inner_text() == ''


def test_research_context_ui_reuse_read_and_rollback(professional_browser):
    exercise_research_context_browser(professional_browser)
