import json
import os
from pathlib import Path
import re
import threading

from .bridge import WebviewApp
from .floating_session import FloatingApi
from .orb_instance import SingleInstance
from .orb_native import NativeOrb, drop_candidates
from .window_chrome import WindowChrome
from ..branding import APP_ICON


DROP_MESSAGE = 'lix-native-file-drop'


def _document_drop(api, paths, request=None, owner=None):
    floating = api._floating
    if request is None:
        if not drop_candidates(paths):
            floating.notice('admission_failed', {'text': '请拖入一份本机 DOCX、XLSX 或 PPTX 文档。'})
            return
        floating.dropped(paths)
    else:
        if len(paths) != 1:
            result = {'ok': False, 'error': '一次只能拖入一份参考文件。'}
        else:
            result = api._floating_call(floating.conversation.scopes.drop_file, *owner, paths[0])
        floating.notice('composer_file_dropped', {'request': request, **result})


def _receive_drop(api, source, args, file_type):
    if not source or str(args.Source) != source:
        return
    raw = str(args.WebMessageAsJson)
    if len(raw) > 2048 or DROP_MESSAGE not in raw:
        return
    try:
        handler, value, marker = json.loads(raw)
        if handler != 'pywebviewEventHandler' or marker != DROP_MESSAGE:
            return
        value = json.loads(value)
        if not isinstance(value, dict) or value.get('nodeId') != DROP_MESSAGE:
            return
        request, owner = value.get('request'), value.get('owner')
        if value.get('route') == 'reference':
            if (not isinstance(request, str) or not re.fullmatch(r'[a-zA-Z0-9-]{1,64}', request)
                    or not isinstance(owner, list) or len(owner) != 3 or type(owner[0]) is not int
                    or not all(isinstance(item, str) and 1 <= len(item) <= 160 for item in owner[1:])):
                return
        elif value.get('route') != 'open' or request is not None or owner is not None:
            return
        objects = list(args.get_AdditionalObjects() or ())
        paths = tuple(str(item.Path) for item in objects if isinstance(item, file_type))
        if not paths or len(paths) != len(objects):
            return
    except (ValueError, TypeError, AttributeError):
        return
    threading.Thread(target=_document_drop, args=(api, paths, request, owner), daemon=True).start()


def _install_file_drop(window, api):
    from System import Action
    from Microsoft.Web.WebView2.Core import CoreWebView2File

    form = window.native

    def receive(_, args):
        _receive_drop(api, window.real_url, args, CoreWebView2File)

    def install():
        form.webview.WebMessageReceived += receive

    form.Invoke(Action(install))


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
                       warning=lambda text: api._floating.notice('native_warning', {'text': text}),
                       preferences=api._ui_preferences)
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
            if gui == 'edgechromium':
                try:
                    _install_file_drop(window, api)
                except Exception:
                    api._floating.notice('native_warning', {'text': '面板文件拖放未就绪，可继续选择参考文件或从悬浮球打开。'})
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
