from docx import Document

from office_agent.rag_workspace import WorkspaceRag
from test_rag_workspace import SyntheticEmbedder
from test_professional_browser import STATIC, professional_browser


def exercise_corpus_browser(browser, capture=None):
    page, api, app, calls, receipts = browser
    second = app.path.parent / 'second.docx'
    document = Document()
    document.add_paragraph('EXPT-02 30 ms synthetic experiment with a limitation.')
    document.save(second)
    source = app.path.read_bytes()
    embedder = SyntheticEmbedder()
    api._rag._factory = lambda scope_check: WorkspaceRag(app.editor.space, embedder, scope_check=scope_check)
    page.evaluate("document.getElementById('settings-dialog').hidden = false; document.getElementById('settings-dialog').inert = false; document.querySelectorAll('.settings-pane').forEach(node => node.hidden = node.id !== 'settings-pane-rag'); document.querySelectorAll('.settings-tab').forEach(node => node.classList.toggle('active', node.id === 'tab-rag'))")
    page.evaluate("document.querySelectorAll('.settings-tab').forEach(node => { if (node.id === 'tab-rag') node.setAttribute('aria-current','true'); else node.removeAttribute('aria-current'); })")
    page.add_script_tag(path=str(STATIC / 'rag.js'))
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('尚未建立')")
    page.locator('#rag-research').check()
    page.locator('#rag-directory').check()
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-profile').textContent.includes('目录纳入 2')")
    assert embedder.calls == []
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('已保存并校验')")
    assert len(embedder.calls) == 1
    page.get_by_text('目录专业检索选项', exact=True).click()
    page.locator('#rag-rerank').check()
    page.locator('#rag-expand').check()
    page.locator('#rag-question').fill('比较 样本30个 以及 EXPT-02')
    page.locator('#rag-query').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('已发布目录索引')")
    assert '未配置，保留融合排序' in page.locator('#rag-profile').inner_text()
    assert len(embedder.calls) == 2 and 1 < len(embedder.calls[1]) <= 3
    page.locator('[data-rag-context]').first.click()
    page.wait_for_function("document.getElementById('rag-context').textContent.includes('只读上下文')")
    if capture is not None:
        capture(page, calls, embedder)
    second.unlink()
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-profile').textContent.includes('需移除 1')")
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-snapshot').textContent.includes('已保存并校验')")
    assert len(embedder.calls) == 2
    page.locator('#rag-rollback-index').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('已撤销')")
    assert app.path.read_bytes() == source
    assert api._rag._store.corpus_status(app.path.name, api._rag._session)['removed'] == 1


def test_actual_directory_incremental_rerank_fallback_context_and_rollback(professional_browser):
    exercise_corpus_browser(professional_browser)
