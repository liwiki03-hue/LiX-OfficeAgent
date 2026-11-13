from office_agent.rag_workspace import WorkspaceRag
from test_rag_workspace import SyntheticEmbedder
from test_professional_browser import STATIC, professional_browser
from test_research_reading import long_pdf


def exercise_pdf_range(browser, capture=None):
    page, api, app, calls, receipts = browser
    path = app.editor.space.root / 'paper.pdf'
    original = long_pdf()
    path.write_bytes(original)
    app.path = path
    embedder = SyntheticEmbedder()
    api._rag._factory = lambda check: WorkspaceRag(app.editor.space, embedder, scope_check=check)
    page.evaluate("""() => {
      document.getElementById('file-name').textContent = 'paper.pdf';
      document.getElementById('settings-dialog').hidden = false;
      document.getElementById('settings-dialog').inert = false;
      document.querySelectorAll('.settings-pane').forEach(node => node.hidden = node.id !== 'settings-pane-rag');
      document.querySelectorAll('.settings-tab').forEach(node => { node.classList.toggle('active', node.id === 'tab-rag'); if (node.id === 'tab-rag') node.setAttribute('aria-current','true'); else node.removeAttribute('aria-current'); });
    }""")
    page.add_script_tag(path=str(STATIC / 'rag.js'))
    page.locator('#rag-research').check()
    page.locator('#rag-pdf-range').check()
    page.locator('#rag-pdf-start').fill('6')
    page.locator('#rag-pdf-end').fill('7')
    page.locator('#rag-pdf-end').blur()
    assert not calls and not embedder.calls
    page.locator('#rag-load').click()
    page.wait_for_function("document.getElementById('rag-profile').textContent.includes('仅部分覆盖')")
    assert '未建立' in page.locator('#rag-profile').inner_text()
    page.evaluate('window.approved = false')
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('没有写入')")
    assert not embedder.calls and not api._rag._store.snapshot()['exists']
    page.evaluate('window.approved = true')
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('索引已保存')")
    assert '仅部分覆盖' in page.locator('#rag-status').inner_text()
    page.locator('#rag-question').fill('measured latency')
    page.locator('#rag-topk').fill('10')
    page.locator('#rag-query').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('完整结果已保存')")
    assert 'pdf-page:7' in page.locator('#rag-hits').inner_text()
    assert '仅部分覆盖' in page.locator('#rag-profile').inner_text()
    assert '仅部分覆盖' in page.locator('#rag-status').inner_text()
    if capture:
        capture(page, calls, receipts)
    page.locator('#rag-rollback-evidence').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('已撤销')")
    page.locator('#rag-rollback-index').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('已撤销') && document.getElementById('rag-rollback-index').disabled")
    count = len(calls)
    page.locator('#rag-pdf-end').fill('28')
    page.locator('#rag-pdf-end').blur()
    page.locator('#rag-index').click()
    page.wait_for_function("document.getElementById('rag-status').textContent.includes('每次最多20页')")
    assert len(calls) == count
    assert path.read_bytes() == original
    return {'calls': calls, 'synthetic_embedding_calls': len(embedder.calls), 'source_unchanged': True,
            'late_page_retrieved': True, 'partial_coverage_visible': True, 'cancel_and_independent_undo_verified': True}


def test_actual_browser_late_range_cancel_query_and_undo(professional_browser):
    exercise_pdf_range(professional_browser)
