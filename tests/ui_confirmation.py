import hashlib
import json
import os
from pathlib import Path
import time


def record_browser_failure(page, expression, seconds):
    directory=os.environ.get('LIX_UI_FAILURE_DIR')
    if not directory or getattr(page,'_lix_failure_recorded',False):return
    page._lix_failure_recorded=True
    root=Path(directory).resolve();root.mkdir(parents=True,exist_ok=True)
    name=getattr(page,'_lix_case','browser')+'-'+hashlib.sha256(str(expression).encode()).hexdigest()[:12]
    def bounded(value):
        if isinstance(value,dict):return {str(key):('[redacted]' if any(word in str(key).lower() for word in ('secret','password','api_key','authorization')) else bounded(item)) for key,item in list(value.items())[:60]}
        if isinstance(value,(list,tuple)):return [bounded(item) for item in value[-40:]]
        if isinstance(value,str):return '[image bytes omitted]' if value.startswith('data:image/') else value[:2000]
        if value is None or isinstance(value,(int,float,bool)):return value
        return str(value)[:2000]
    evidence={'expression':str(expression),'total_budget_seconds':seconds,'native_desktop':False,
        'confirmations':getattr(page,'_lix_confirmation_receipts',[]),
        'host_receipts':bounded(getattr(page,'_lix_api_receipts',[])),
        'page_errors':getattr(page,'_lix_page_errors',[])}
    try:
        evidence['page']=page.evaluate('''() => ({url:location.href,ready:document.readyState,viewport:[innerWidth,innerHeight,devicePixelRatio],
          labels:Object.fromEntries(['file-name','file-dir','range-status','range-before-caption','range-after-caption','range-review-note','rag-status','rag-snapshot','rag-profile'].map(id=>[id,document.getElementById(id)?.textContent?.slice(0,2000)])),
          images:document.querySelectorAll('#range-before-pages img,#range-after-pages img').length,
          dialogs:[...document.querySelectorAll('dialog[open]')].map(node=>({label:node.getAttribute('aria-label'),text:node.textContent.slice(0,1200)})),
          trace:window.__lixConfirmationTrace||[]})''')
        page.screenshot(path=str(root/(name+'.png')))
    except Exception as error:evidence['capture_error']=type(error).__name__
    (root/(name+'.json')).write_text(json.dumps(evidence,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'browser_failure_evidence':name,'total_budget_seconds':seconds},ensure_ascii=True))


def respond_to_confirmation(page, *, approved=None):
    previous = getattr(page, '_lix_confirmation_responding', False)
    page._lix_confirmation_responding = True
    try:
        dialog = page.locator('dialog.lix-confirm-dialog[open]')
        dialog.wait_for(state='visible')
        if approved is None:
            approved = page.evaluate('window.approved !== false')
        assert dialog.get_attribute('aria-label') == 'LiX'
        dialog.get_by_role('button', name='确认' if approved else '取消', exact=True).click()
    finally:
        page._lix_confirmation_responding = previous


def click_with_confirmation(page, selector, *, approved=None):
    page.locator(selector).click()
    respond_to_confirmation(page, approved=approved)


def drive_legacy_confirmations(page, monkeypatch):
    from playwright.sync_api import TimeoutError as BrowserTimeout
    receipts = []
    dialog = page.locator('dialog.lix-confirm-dialog[open]')

    def respond():
        if getattr(page, '_lix_confirmation_responding', False):
            return
        if not dialog.count() or not dialog.is_visible():
            return
        approved = page.evaluate('window.approved !== false')
        message = dialog.locator('p').inner_text()
        receipts.append(dict(message=message, approved=approved))
        page.evaluate('(message) => { if (Array.isArray(window.modalTexts)) window.modalTexts.push(message); }', message)
        respond_to_confirmation(page, approved=approved)

    page.add_locator_handler(dialog, lambda locator: respond(), no_wait_after=True)
    original = page.wait_for_function

    def wait(expression, *, arg=None, timeout=None, polling=None):
        seconds = (timeout if timeout is not None else 65000 if getattr(page, '_lix_real_render', False) else 15000) / 1000
        deadline = time.monotonic() + seconds
        while True:
            respond()
            remaining = max(1, (deadline - time.monotonic()) * 1000)
            try:
                return original(expression, arg=arg, timeout=min(250, remaining), polling=polling)
            except BrowserTimeout:
                if time.monotonic() >= deadline:
                    record_browser_failure(page,expression,seconds)
                    raise

    monkeypatch.setattr(page, 'wait_for_function', wait)
    page._lix_confirmation_receipts = receipts
