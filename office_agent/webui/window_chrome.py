import ctypes
from ctypes import wintypes
import sys


SIZE_ACTIONS = {'resize-left': 10, 'resize-right': 11, 'resize-top': 12, 'resize-top-left': 13,
                'resize-top-right': 14, 'resize-bottom': 15, 'resize-bottom-left': 16, 'resize-bottom-right': 17}
WINDOW_ACTIONS = frozenset({'minimize', 'maximize', 'drag'} | SIZE_ACTIONS.keys())


def resize_edge(x, y, bounds, border=6):
    left, top, right, bottom = bounds
    if not left <= x < right or not top <= y < bottom:
        return 0
    horizontal = -1 if x < left + border else 1 if x >= right - border else 0
    vertical = -1 if y < top + border else 1 if y >= bottom - border else 0
    return {(-1, -1): 13, (1, -1): 14, (-1, 1): 16, (1, 1): 17,
            (-1, 0): 10, (1, 0): 11, (0, -1): 12, (0, 1): 15}.get((horizontal, vertical), 0)


class WindowChrome:
    def __init__(self, window):
        self.window = window
        self._old_proc = self._proc = None
        self._hwnd = None

    def state(self):
        form = self.window.native
        if form is None:
            maximized = bool(self.window.maximized)
        elif form.InvokeRequired:
            from System import Boolean, Func
            maximized = bool(form.Invoke(Func[Boolean](lambda: str(form.WindowState) == 'Maximized')))
        else:
            maximized = str(form.WindowState) == 'Maximized'
        return {'ok': True, 'available': True, 'maximized': maximized}

    def action(self, action):
        if type(action) is not str or action not in WINDOW_ACTIONS:
            raise ValueError('不支持的窗口操作')
        if action == 'drag' or action in SIZE_ACTIONS:
            if self.state()['maximized']:
                raise ValueError('请先还原窗口，再拖动或调整大小')
            self._move_size(2 if action == 'drag' else SIZE_ACTIONS[action])
        elif action == 'minimize':
            self.window.minimize()
        elif self.state()['maximized']:
            self.window.restore()
        else:
            if sys.platform == 'win32' and self.window.native is not None:
                from System import Action
                from System.Windows.Forms import Screen
                form = self.window.native
                form.Invoke(Action(lambda: setattr(form, 'MaximizedBounds', Screen.FromControl(form).WorkingArea)))
            self.window.maximize()
        return self.state()

    def _move_size(self, hit):
        if sys.platform != 'win32' or self.window.native is None:
            raise ValueError('当前窗口不支持原生拖动')
        from System import Action
        form = self.window.native
        started = []
        def move():
            user = ctypes.windll.user32
            user.GetAsyncKeyState.argtypes = [ctypes.c_int]
            user.GetAsyncKeyState.restype = ctypes.c_short
            if not user.GetAsyncKeyState(1) & 0x8000:
                return
            point = wintypes.POINT()
            user.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
            if not user.GetCursorPos(ctypes.byref(point)):
                return
            user.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
            user.SendMessageW.restype = ctypes.c_ssize_t
            user.ReleaseCapture()
            started.append(True)
            user.SendMessageW(ctypes.c_void_p(int(form.Handle.ToInt64())), 0x00a1, hit, ((point.y & 0xffff) << 16) | (point.x & 0xffff))
        form.Invoke(Action(move))
        if not started:
            raise ValueError('已松开鼠标；请按住顶部文件名或窗口边缘拖动')

    def install(self):
        if sys.platform != 'win32':
            return
        from System import Action
        form = self.window.native
        if form is None:
            raise RuntimeError('原生窗口尚未就绪')
        form.Invoke(Action(lambda: self._install(form)))

    def _install(self, form):
        from System.Windows.Forms import Screen
        form.MaximizedBounds = Screen.FromControl(form).WorkingArea
        user = ctypes.windll.user32
        self._hwnd = ctypes.c_void_p(int(form.Handle.ToInt64()))
        get = user.GetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.GetWindowLongW
        set_value = user.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.SetWindowLongW
        get.argtypes = [ctypes.c_void_p, ctypes.c_int]
        get.restype = ctypes.c_ssize_t
        set_value.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]
        set_value.restype = ctypes.c_ssize_t
        user.CallWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
        user.CallWindowProcW.restype = ctypes.c_ssize_t
        user.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.RECT)]
        user.SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
        callback = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t)

        def dispatch(hwnd, message, wparam, lparam):
            if message == 0x0083:
                return 0
            if message == 0x0084 and str(form.WindowState) != 'Maximized':
                rect = wintypes.RECT()
                if user.GetWindowRect(hwnd, ctypes.byref(rect)):
                    x = ctypes.c_short(lparam & 0xffff).value
                    y = ctypes.c_short((lparam >> 16) & 0xffff).value
                    edge = resize_edge(x, y, (rect.left, rect.top, rect.right, rect.bottom), max(6, round(float(form.DeviceDpi) / 16)))
                    if edge:
                        return edge
            return user.CallWindowProcW(self._old_proc, hwnd, message, wparam, lparam)

        self._proc = callback(dispatch)
        self._old_proc = set_value(self._hwnd, -4, ctypes.cast(self._proc, ctypes.c_void_p).value)
        if not self._old_proc:
            raise OSError('无法安装原生窗口边界')
        style = get(self._hwnd, -16)
        set_value(self._hwnd, -16, (style | 0x00070000) & ~0x00c00000)
        user.SetWindowPos(self._hwnd, None, 0, 0, 0, 0, 0x0037)
        form.HandleDestroyed += lambda sender, event: self.dispose()

    def dispose(self):
        if self._old_proc and self._hwnd:
            user = ctypes.windll.user32
            user.IsWindow.argtypes = [ctypes.c_void_p]
            if user.IsWindow(self._hwnd):
                setter = user.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.SetWindowLongW
                setter(self._hwnd, -4, self._old_proc)
            self._old_proc = None
