import os
from pathlib import Path
import threading

from .bridge import WebviewApp
from .floating_session import FloatingApi
from .orb_instance import SingleInstance
from .orb_native import NativeOrb
from .window_chrome import WindowChrome
from ..branding import APP_ICON


def launch_floating(webview, gui, ocr_provider=None):
    instance = SingleInstance()
    if not instance.primary:
        instance.close()
        return
    app = WebviewApp(ocr_provider=ocr_provider)
    api = FloatingApi(app)
    static = Path(__file__).resolve().parent / 'static' / 'index.html'
    window = webview.create_window('LiX · 文档助手', str(static), js_api=api,
                                  width=420, height=640, min_size=(360, 440), hidden=True,
                                  background_color='#15171a', frameless=True, easy_drag=False)
    api._window_chrome = WindowChrome(window)
    app.attach(window)
    native = NativeOrb(window, dropped=api._floating.dropped, can_exit=lambda: not app.busy and not (api._floating.paged.active() and api._floating.paged.revision!=api._floating.paged.published_revision) and api._floating.canvas.switch_allowed(),
                       choose=lambda: window.evaluate_js("document.getElementById('floating-choose')?.click()"),
                       settings=lambda: window.evaluate_js('window.OfficeSettings?.show()'),
                       recent=lambda: window.evaluate_js('window.OfficeFloating?.history()'),
                       position=lambda value: api._floating.notice('orb_position', value),
                       hotkey=os.getenv('OFFICE_ORB_HOTKEY', ''), wake=instance.poll_wake,
                       warning=lambda text: api._floating.notice('native_warning', {'text': text}))
    api._floating.native = native
    def ready():
        if window.events.shown.wait(20):
            try:
                api._window_chrome.install()
            except Exception:
                try:
                    api._window_chrome.fallback()
                    text = '自定义边界未加载，已使用系统窗框；拖动与缩放可使用系统边框'
                except Exception:
                    text = '窗口边界与系统窗框回退未就绪；请通过托盘恢复并查看诊断'
                api._floating.notice('native_warning', {'text': text})
            try:
                native.install()
            except Exception:
                window.show()
                api._floating.notice('native_warning', {'text': '悬浮球初始化失败；面板可用，请查看启动诊断'})
    try:
        webview.start(ready, gui=gui, icon=str(APP_ICON))
    finally:
        api._window_chrome.dispose()
        instance.close()
