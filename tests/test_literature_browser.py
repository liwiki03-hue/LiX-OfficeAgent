import json
import threading

import pytest

import httpx2

from test_literature import client_fixture, crossref, openalex, response
from test_fidelity import pdf_bytes
from test_professional_browser import _browser, STATIC


def literature_browser(tmp_path, monkeypatch, fixture=None):
    session = _browser(tmp_path, monkeypatch)
    page, api, app, old_calls, receipts = next(session)
    client, requests, now = fixture or client_fixture()
    api._literature.factory = lambda: client
    calls = []
    allowed = {'search_literature','verify_literature','review_literature','expand_literature_citations',
               'plan_literature_export','plan_literature_fulltext','confirm_literature_save',
               'read_literature_fulltext','preview_literature_local_pdf','check_literature_citations','rollback_literature_save'}
    def invoke(method, args):
        assert method in allowed
        calls.append(method)
        value = getattr(api, method)(*args)
        receipts.append((method, value))
        return value
    page.expose_function('literatureHost', invoke)
    page.evaluate("window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.literatureHost(method, args)}); document.getElementById('professional-panel').open=false; document.getElementById('literature-panel').open=true")
    page.add_script_tag(path=str(STATIC / 'literature.js'))
    try:
        yield page, api, app, calls, receipts, requests, now
    finally:
        session.close()


def exercise_literature_browser(browser, capture=None):
    page, api, app, calls, receipts, requests, now = browser
    source = app.path.read_bytes()
    assert not requests and not calls
    page.locator('#literature-query').fill('thermal experiment')
    page.locator('#literature-search').click()
    page.wait_for_function("document.querySelectorAll('#literature-results article').length === 2")
    assert len(requests) == 2
    page.evaluate('''() => {
      const raw = window.pywebview.api;
      window.pywebview.api = new Proxy({}, {get: (_, method) => method === 'plan_literature_export' ?
        async (...args) => { if (window.failLiteraturePlan) throw Error('fixture'); return raw[method](...args); } : raw[method]});
    }''')
    assert '同 DOI 来源字段存在冲突' in page.locator('#literature-results').inner_text()
    assert page.locator('#literature-results').get_by_text('未知，未确认未撤稿', exact=False).count() == 1
    first = page.locator('#literature-results article').first
    first.locator('input').fill('Methods relevant to technical experiment report')
    first.get_by_role('button', name='收藏 / 纳入', exact=True).click()
    page.wait_for_function("document.getElementById('literature-bibliography').textContent.includes('[1]')")
    page.get_by_text('收藏、引用预览与导出', exact=True).click()
    page.locator('#literature-citation-text').fill('Known [1], missing [9].')
    page.locator('#literature-check').click()
    page.wait_for_function("document.getElementById('literature-citation-report').textContent.includes('无对应条目的编号：9')")
    page.evaluate('window.failLiteraturePlan=true')
    page.locator('#literature-csl').click()
    page.wait_for_function("document.getElementById('literature-status').textContent.includes('尚未批准保存')")
    assert page.locator('#literature-csl').is_enabled()
    page.evaluate('window.failLiteraturePlan=false')
    page.evaluate('window.approved=false')
    page.locator('#literature-csl').click()
    page.wait_for_function("document.getElementById('literature-status').textContent.includes('已取消')")
    assert not list(app.path.parent.glob('.office-rag-evidence-*'))
    page.evaluate('window.approved=true')
    page.locator('#literature-csl').click()
    page.wait_for_function("document.getElementById('literature-outputs').textContent.includes('references.csl.json')")
    saved = next(reply['result'] for method, reply in reversed(receipts) if method == 'confirm_literature_save' and reply['result']['status'] == 'saved')
    target = app.path.parent / saved['evidence']
    assert json.loads(target.read_bytes())[0]['title'] == 'Synthetic thermal experiment'
    if capture:
        capture(page, calls, requests, saved)
    page.locator('#literature-outputs').get_by_role('button', name='撤销本次保存…', exact=True).click()
    page.wait_for_function("document.getElementById('literature-status').textContent.includes('已撤销')")
    assert not target.exists() and app.path.read_bytes() == source
    page.evaluate("window.authenticated=false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert page.locator('#literature-query').input_value() == ''
    assert page.locator('#literature-results').inner_text() == ''
    assert page.locator('#literature-search').is_disabled()


def test_actual_page_sources_conflicts_collection_export_cancel_rollback_and_logout(tmp_path, monkeypatch):
    session = literature_browser(tmp_path, monkeypatch)
    try:
        exercise_literature_browser(next(session))
    finally:
        session.close()


def test_late_frontend_response_is_discarded_after_logout(tmp_path, monkeypatch):
    session = literature_browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts, requests, now = next(session)
        page.evaluate('''() => {
          const raw = window.pywebview.api;
          window.pywebview.api = {search_literature: async (...args) => {
            const reply = await raw.search_literature(...args);
            await new Promise(resolve => { window.releaseLiterature = resolve; });
            window.literatureDelayedDone = true;
            return reply;
          }};
        }''')
        page.locator('#literature-openalex').uncheck()
        page.locator('#literature-query').fill('thermal')
        page.locator('#literature-search').click()
        page.wait_for_function("document.getElementById('literature-panel').getAttribute('aria-busy') === 'true'")
        page.wait_for_function("typeof window.releaseLiterature === 'function'")
        assert page.locator('#literature-search').is_disabled()
        page.evaluate("window.authenticated=false; window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        page.evaluate('window.releaseLiterature()')
        page.wait_for_function("document.getElementById('literature-panel').getAttribute('aria-busy') === 'false'")
        page.wait_for_function('window.literatureDelayedDone === true')
        assert page.locator('#literature-results').inner_text() == '' and page.locator('#literature-query').input_value() == ''
        assert len(requests) == 1
    finally:
        session.close()


def test_pdf_ui_save_parse_metadata_refresh_and_rollback_keep_state_consistent(tmp_path, monkeypatch):
    data = pdf_bytes(text='Synthetic public body evidence.')
    def handler(request):
        if request.url.host == 'arxiv.org':
            return httpx2.Response(200, headers={'content-type':'application/pdf'}, stream=httpx2.ByteStream(data))
        return response({'results':[openalex()], 'meta':{'count':1}})
    session = literature_browser(tmp_path, monkeypatch, client_fixture(handler))
    try:
        page, api, app, calls, receipts, requests, now = next(session)
        source = app.path.read_bytes()
        page.locator('#literature-crossref').uncheck()
        page.locator('#literature-query').fill('thermal')
        page.locator('#literature-search').click()
        page.locator('#literature-results').get_by_role('button', name='取得许可范围内全文…', exact=True).click()
        page.wait_for_function("document.getElementById('literature-outputs').textContent.includes('fulltext.pdf')")
        assert '已保存，尚未解析' in page.locator('#literature-results').inner_text()
        page.locator('#literature-outputs').get_by_role('button', name='解析已保存全文', exact=True).click()
        page.wait_for_function("document.getElementById('literature-fulltext').textContent.includes('Synthetic public body evidence')")
        assert '文本层已解析' in page.locator('#literature-results').inner_text()
        page.locator('#literature-search').click()
        page.wait_for_function("document.getElementById('literature-reports').textContent.includes('会话缓存')")
        assert '文本层已解析' in page.locator('#literature-results').inner_text() and len(requests) == 2
        page.locator('#literature-outputs').get_by_role('button', name='撤销本次保存…', exact=True).click()
        page.wait_for_function("document.getElementById('literature-status').textContent.includes('已撤销')")
        assert '尚未取得' in page.locator('#literature-results').inner_text() and app.path.read_bytes() == source
    finally:
        session.close()
