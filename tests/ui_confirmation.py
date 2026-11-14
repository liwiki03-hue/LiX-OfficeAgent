import time


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
                    raise

    monkeypatch.setattr(page, 'wait_for_function', wait)
    page._lix_confirmation_receipts = receipts
