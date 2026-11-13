from pathlib import Path

from office_agent.update_core import write_json
from test_app_updates import fixture, package, service
from test_professional_browser import _browser, STATIC


def browser(tmp_path, monkeypatch, *, configured=True, automatic=False):
    session = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(session)
        key = fixture(tmp_path)
        updates, state = service(tmp_path, package(tmp_path, key))
        api._updates = updates
        if not configured:
            write_json(tmp_path / 'office_agent/update_channel.json', {'repository': '', 'public_key': ''})
        allowed = {'get_release_info', 'check_app_updates', 'download_app_update', 'confirm_app_update'}
        def invoke(method, args):
            assert method in allowed
            calls.append(method)
            result = getattr(api, method)(*args)
            receipts.append(result)
            return result
        page.expose_function('updatesHost', invoke)
        page.evaluate('''automatic => {
          const saved = {'office-agent-updates-auto-check': String(automatic), 'office-agent-updates-auto-download': String(automatic)};
          Object.defineProperty(window, 'localStorage', {value: {getItem: key => saved[key], setItem: (key, value) => {saved[key] = value;}}});
          window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.updatesHost(method, args)});
        }''', automatic)
        page.add_script_tag(path=str(STATIC / 'settings_dialog.js'))
        page.add_script_tag(path=str(STATIC / 'updates.js'))
        page.locator('#nav-settings').click()
        page.locator('#tab-updates').click()
        page.wait_for_function('document.getElementById("updates-version").textContent.includes("0.2.0")')
        yield page, api, app, calls, receipts, state
    finally:
        session.close()


def exercise(page, calls, state, capture=None):
    assert calls == ['get_release_info'] and not state['calls']
    page.locator('#updates-check').click()
    page.wait_for_function('!document.getElementById("updates-download").disabled')
    assert page.locator('#updates-release-notes').text_content() == '<script>untrusted notes</script>'
    assert page.locator('#updates-release-notes script').count() == 0
    page.locator('#updates-download').click()
    page.wait_for_function('!document.getElementById("updates-install").disabled')
    if capture:
        capture(page)
    page.evaluate('window.approved = false')
    page.locator('#updates-install').click()
    page.wait_for_function('document.getElementById("updates-status").textContent.includes("已取消")')
    assert not state['handoffs']
    page.locator('#updates-download').click()
    page.wait_for_function('!document.getElementById("updates-install").disabled')
    page.evaluate('window.approved = true')
    page.locator('#updates-install').click()
    page.wait_for_function('document.getElementById("updates-status").textContent.includes("安装已安排")')
    assert len(state['handoffs']) == 1 and page.locator('#updates-install').is_disabled()


def test_browser_download_cancel_confirm_and_themes(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, state in browser(tmp_path, monkeypatch):
        def capture(page):
            for theme in ('dark', 'light', 'highcontrast'):
                page.set_viewport_size({'width': 520, 'height': 900})
                page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
                assert page.locator('#settings-dialog').evaluate('(node) => node.scrollWidth <= node.clientWidth')
        exercise(page, calls, state, capture)


def test_unconfigured_release_is_readable_and_never_requests_network(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, state in browser(tmp_path, monkeypatch, configured=False):
        assert page.locator('#updates-check').is_disabled()
        assert page.locator('#updates-auto-check').is_disabled()
        assert '尚未配置' in page.locator('#updates-source').inner_text()
        assert '版本管理与受控更新' in page.locator('#updates-changelog').inner_text()
        assert not state['calls'] and calls == ['get_release_info']


def test_opted_in_automatic_check_and_download_never_installs(tmp_path, monkeypatch):
    for page, api, app, calls, receipts, state in browser(tmp_path, monkeypatch, automatic=True):
        page.wait_for_function('!document.getElementById("updates-install").disabled')
        assert calls.count('check_app_updates') == 1 and calls.count('download_app_update') == 1
        assert not state['handoffs']
        page.evaluate('window.authenticated = false; window.dispatchEvent(new Event("officeauth:unauthenticated"))')
        assert page.locator('#updates-check').is_disabled() and not page.locator('#updates-version').inner_text()
