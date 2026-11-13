import hashlib
import json
from pathlib import Path
import sys
from types import SimpleNamespace

from PIL import Image
import pytest

from office_agent.branding import APP_ICON, APP_LOGO, APP_NAME, ASSET_ROOT


def test_web_entry_passes_actual_icon_to_existing_backend(monkeypatch):
    from office_agent import app
    from office_agent.webui import bridge, runtime

    calls = {}
    monkeypatch.setattr(runtime, 'configure_runtime', lambda: 'qt')
    monkeypatch.setattr(bridge, 'WebviewApp', lambda **kwargs: SimpleNamespace(attach=lambda window: None))
    monkeypatch.setattr(bridge, 'Api', lambda target: object())
    def create(title, *args, **kwargs):
        calls['title'] = title
        return object()
    monkeypatch.setitem(sys.modules, 'webview', SimpleNamespace(settings={}, create_window=create,
                                                              start=lambda **kwargs: calls.update(kwargs)))
    app._launch_web()
    assert calls['title'] == 'LiX — 审核后写入'
    assert Path(calls['icon']) == APP_ICON
    assert calls['gui'] == 'qt'


def test_brand_assets_are_decodable_packaged_and_match_supplied_art():
    root = Path(__file__).resolve().parents[1]
    manifest = json.loads((ASSET_ROOT / 'manifest.json').read_text('utf-8'))
    public = json.loads((root / 'release/PUBLIC_FILES.json').read_text('utf-8'))['files']
    assert APP_NAME == 'LiX'
    for name, digest in manifest['files'].items():
        path = ASSET_ROOT / name
        assert hashlib.sha256(path.read_bytes()).hexdigest() == digest
        assert path.relative_to(root).as_posix() in public
        with Image.open(path) as image:
            image.load()
    with Image.open(APP_ICON) as icon:
        assert icon.ico.sizes() == {(size, size) for size in manifest['icon_sizes']}
        for size in manifest['icon_sizes']:
            frame = icon.ico.getimage((size, size)).convert('RGBA')
            assert all(frame.getpixel(point)[3] == 0 for point in (
                (0, 0), (size - 1, 0), (0, size - 1), (size - 1, size - 1)))
            red, green, blue, alpha = frame.getpixel((size * 3 // 4, size // 6))
            assert red > green > blue and alpha >= 250
    with Image.open(ASSET_ROOT / 'lix-transparent.png') as source, Image.open(APP_LOGO) as logo:
        expected = source.convert('RGBA').resize((256, 256), Image.Resampling.LANCZOS)
        assert expected.tobytes() == logo.convert('RGBA').tobytes()
        assert all(logo.getpixel(point)[3] == 0 for point in ((0, 0), (255, 0), (0, 255), (255, 255)))
        assert logo.getpixel((128, 128))[3] >= 250
        left, top, right, bottom = logo.getchannel('A').point(lambda value: 255 if value >= 128 else 0).getbbox()
        assert right - left >= 240 and bottom - top >= 240


def test_browser_brand_and_existing_chrome_with_real_local_assets(tmp_path):
    from playwright.sync_api import sync_playwright
    from tests.test_professional_browser import EDGE, STATIC

    if not EDGE.is_file():
        pytest.skip('Requires existing Windows Edge')
    with sync_playwright() as pw:
        browser = pw.chromium.launch(executable_path=str(EDGE), headless=True, args=['--disable-background-networking'])
        try:
            page = browser.new_page(viewport={'width': 1120, 'height': 760})
            page.set_default_timeout(15000)
            page.route('http://**/*', lambda route: route.abort())
            page.route('https://**/*', lambda route: route.abort())
            page.goto((STATIC / 'index.html').as_uri())
            assert page.title() == 'LiX — 审核后写入'
            page.evaluate('''() => {
              window.OfficeAuth = {wrapApi: x => x};
              window.pywebview = {api: {floating_window_state: async () => ({ok:true, available:true})}};
              const header = document.createElement('header');
              const title = document.createElement('strong'); title.id = 'floating-title'; title.textContent = 'LiX';
              const close = document.createElement('button'); header.append(title, close); document.body.append(header);
              window.OfficeWindowChrome.attach(header, close);
              window.OfficeWindowChrome.secondary(true);
            }''')
            page.wait_for_selector('#desktop-window-bar:not([hidden])')
            assert page.locator('#desktop-window-bar > span').inner_text() == 'LiX'
            for selector in ('.logo .lix-brand-icon', '#desktop-window-bar .lix-brand-icon'):
                assert page.locator(selector).evaluate('(img) => img.complete && img.naturalWidth === 256')
            assert page.locator('#desktop-window-bar button[aria-label="最大化窗口"]').count() == 1
            assert page.locator('link[rel="icon"]').get_attribute('href') == 'brand/lix.png'
            for width, theme in ((1120, 'dark'), (820, 'light'), (520, 'highcontrast')):
                page.set_viewport_size({'width': width, 'height': 760})
                page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
                assert page.locator('#desktop-window-bar img').is_visible()
        finally:
            browser.close()
