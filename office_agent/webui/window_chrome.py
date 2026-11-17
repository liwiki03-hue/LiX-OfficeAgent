import ctypes
from ctypes import wintypes
import math
import re
import sys
import time


SIZE_ACTIONS = {'resize-left': 10, 'resize-right': 11, 'resize-top': 12, 'resize-top-left': 13,
                'resize-top-right': 14, 'resize-bottom': 15, 'resize-bottom-left': 16, 'resize-bottom-right': 17}
WINDOW_ACTIONS = frozenset({'minimize', 'maximize', 'drag', 'standard-frame'} | SIZE_ACTIONS.keys())


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
        self._fallback = False
        self._filter = None
        self._message_proc = None
        self._drag_regions = None
        self._regions_version = -1
        self._original_style = None
        self._caption_failed = False
        self._mouse_messages = self._caption_hits = self._move_requests = 0
        self._move_result = 'idle'
        self._drag_events = []
        self._pointer_drag = None

    def trace(self, value):
        if not isinstance(value, dict) or set(value) != {'phase', 'driver', 'region', 'prevented', 'buttons'}:
            raise ValueError('标题输入记录无效')
        if (value['phase'] not in {'pointerdown', 'mousedown', 'pointermove', 'pointerup', 'mouseup'}
                or value['driver'] not in {'unset', 'pywebview', 'pointer'} or type(value['region']) is not bool
                or type(value['prevented']) is not bool or type(value['buttons']) is not int
                or not 0 <= value['buttons'] <= 31):
            raise ValueError('标题输入记录超出范围')
        if len(self._drag_events) < 20:
            self._drag_events.append(dict(value))
            from ..diagnostics import record
            record('window', 'caption_input', 'observed', code='_'.join(map(str, (
                value['phase'], value['driver'], int(value['region']), int(value['prevented']), value['buttons']))))
        return {'ok': True}

    def state(self):
        form = self.window.native
        if form is None:
            maximized = bool(self.window.maximized)
        elif form.InvokeRequired:
            from System import Boolean, Func
            maximized = bool(form.Invoke(Func[Boolean](lambda: str(form.WindowState) == 'Maximized')))
        else:
            maximized = str(form.WindowState) == 'Maximized'
        return {'ok': True, 'available': True, 'maximized': maximized, 'standard_frame': self._fallback,
                'native_caption': False, 'drag_driver': 'pointer',
                'caption_failed': self._caption_failed, 'drag_regions_version': self._regions_version,
                'title_events': list(self._drag_events),
                'drag': {'mouse_messages': self._mouse_messages, 'caption_hits': self._caption_hits,
                         'move_requests': self._move_requests, 'last_result': self._move_result}}

    def pointer(self, value):
        if not isinstance(value, dict) or set(value) != {'id', 'phase', 'seq', 'screen', 'client', 'viewport'}:
            raise ValueError('拖动事件格式无效')
        if (not isinstance(value['id'], str) or not re.fullmatch(r'[A-Za-z0-9-]{1,64}', value['id'])
                or value['phase'] not in {'start', 'move', 'end', 'cancel'}
                or type(value['seq']) is not int or not 0 <= value['seq'] <= 2000000):
            raise ValueError('拖动事件身份无效')
        for key, maximum in (('screen', 100000), ('client', 32768), ('viewport', 32768)):
            row = value[key]
            if (not isinstance(row, list) or len(row) != 2 or any(type(n) not in (int, float)
                    or not math.isfinite(n) or abs(n) > maximum for n in row)):
                raise ValueError('拖动坐标无效')
        width, height = value['viewport']
        client_x, client_y = value['client']
        if not 1 <= width <= 32768 or not 1 <= height <= 32768 or not 0 <= client_x < width or not 0 <= client_y < height:
            raise ValueError('拖动起点超出窗口')
        form = self.window.native
        if sys.platform != 'win32' or form is None:
            raise ValueError('当前窗口不能移动')
        from System import Action
        from System.Drawing import Point
        from System.Windows.Forms import FormWindowState, Screen
        def apply():
            now = time.monotonic()
            current = self._pointer_drag
            if value['phase'] == 'start':
                scale = (form.ClientSize.Width / width, form.ClientSize.Height / height)
                anchor = form.PointToScreen(Point(round(client_x * scale[0]), round(client_y * scale[1])))
                self._pointer_drag = {'id': value['id'], 'seq': value['seq'], 'screen': tuple(value['screen']),
                    'origin': (form.Left, form.Top), 'scale': scale, 'deadline': now + 30,
                    'anchor': (anchor.X, anchor.Y), 'maximized': str(form.WindowState) == 'Maximized', 'started': False}
                self._caption_hits = min(2147483647, self._caption_hits + 1)
                self._move_result = 'pointer_started'
                return
            if not current or current['id'] != value['id'] or value['seq'] <= current['seq']:
                return
            if current['deadline'] < now or value['phase'] == 'cancel':
                self._pointer_drag = None
                self._move_result = 'pointer_cancelled'
                return
            dx = (value['screen'][0] - current['screen'][0]) * current['scale'][0]
            dy = (value['screen'][1] - current['screen'][1]) * current['scale'][1]
            if not current['started']:
                if max(abs(dx), abs(dy)) < 3:
                    current['seq'] = value['seq']
                    if value['phase'] == 'end':
                        self._pointer_drag = None
                    return
                if current['maximized'] and str(form.WindowState) == 'Maximized':
                    form.WindowState = FormWindowState.Normal
                    form.Location = Point(round(current['anchor'][0] - form.Width * client_x / width),
                                          round(current['anchor'][1] - 20 * form.DeviceDpi / 96))
                    current['origin'] = (form.Left, form.Top)
                current['started'] = True
            x = round(current['origin'][0] + dx)
            y = round(current['origin'][1] + dy)
            area = Screen.FromPoint(Point(x + min(80, form.Width // 2), y + 20)).WorkingArea
            form.Location = Point(max(area.Left, min(x, area.Right - form.Width)), max(area.Top, min(y, area.Bottom - form.Height)))
            self._move_requests = min(2147483647, self._move_requests + 1)
            self._move_result = 'moved' if current['origin'] != (form.Left, form.Top) else 'unchanged'
            current.update(seq=value['seq'], deadline=now + 30)
            if value['phase'] == 'end':
                self._pointer_drag = None
        form.Invoke(Action(apply))
        return self.state()

    def set_drag_regions(self, value):
        if not isinstance(value, dict) or set(value) != {'version', 'viewport', 'regions', 'exclude'}:
            raise ValueError('窗口标题区域格式无效')
        def number(item, minimum, maximum):
            if type(item) not in (int, float) or not math.isfinite(item) or not minimum <= item <= maximum:
                raise ValueError('窗口标题区域超出范围')
            return float(item)
        version = number(value['version'], 0, 2e15)
        viewport = value['viewport']
        if not isinstance(viewport, list) or len(viewport) != 2:
            raise ValueError('窗口可视区域无效')
        width, height = [number(item, 1, 32768) for item in viewport]
        groups = []
        for key, limit in (('regions', 16), ('exclude', 64)):
            items = value[key]
            if not isinstance(items, list) or len(items) > limit:
                raise ValueError('窗口标题区域过多')
            rectangles = []
            for item in items:
                if not isinstance(item, list) or len(item) != 4:
                    raise ValueError('窗口标题矩形无效')
                left, top, right, bottom = [number(coordinate, 0, 32768) for coordinate in item]
                if not 0 <= left < right <= width or not 0 <= top < bottom <= height:
                    raise ValueError('窗口标题矩形越界')
                rectangles.append((left, top, right, bottom))
            groups.append(tuple(rectangles))
        if version <= self._regions_version:
            return self.state()
        form = self.window.native
        if sys.platform == 'win32' and form is not None:
            from System import Action
            def apply():
                if version > self._regions_version:
                    size = (form.ClientSize.Width, form.ClientSize.Height)
                    self._drag_regions = (width, height, groups[0], groups[1], size)
                    self._regions_version = version
            form.Invoke(Action(apply))
        return self.state()

    def action(self, action):
        if type(action) is not str or action not in WINDOW_ACTIONS:
            raise ValueError('不支持的窗口操作')
        self._pointer_drag = None
        if action == 'standard-frame':
            return self.fallback()
        if action == 'drag' or action in SIZE_ACTIONS:
            if self.state()['maximized']:
                if action != 'drag':
                    return self.state()
                self._restore_for_drag()
            started = self._move_size(2 if action == 'drag' else SIZE_ACTIONS[action])
            if started is False:
                return dict(self.state(), drag_started=False)
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

    def _filter_message(self, handle, message):
        if message != 0x0201 or not self._drag_regions or self._fallback:
            return False
        from System.Windows.Forms import Cursor
        form = self.window.native
        if form is None or form.IsDisposed or not form.Enabled:
            return False
        user = ctypes.windll.user32
        handle = ctypes.c_void_p(handle)
        if handle.value != self._hwnd.value and not user.IsChild(self._hwnd, handle):
            return False
        self._mouse_messages = min(2147483647, self._mouse_messages + 1)
        width, height, regions, exclude, size = self._drag_regions
        if size != (form.ClientSize.Width, form.ClientSize.Height):
            return False
        point = form.PointToClient(Cursor.Position)
        x, y = point.X * width / max(1, size[0]), point.Y * height / max(1, size[1])
        inside = lambda rect: rect[0] <= x < rect[2] and rect[1] <= y < rect[3]
        if not any(map(inside, regions)) or any(map(inside, exclude)):
            return False
        self._caption_hits = min(2147483647, self._caption_hits + 1)
        self._move_size(2)
        return True

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
                self._move_result = 'pointer_released'
                return
            point = wintypes.POINT()
            user.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
            if not user.GetCursorPos(ctypes.byref(point)):
                return
            user.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
            user.SendMessageW.restype = ctypes.c_ssize_t
            user.ReleaseCapture()
            started.append(True)
            self._move_requests = min(2147483647, self._move_requests + 1)
            before = (form.Left, form.Top, form.Width, form.Height)
            user.SendMessageW(ctypes.c_void_p(int(form.Handle.ToInt64())), 0x00a1, hit, ((point.y & 0xffff) << 16) | (point.x & 0xffff))
            self._move_result = 'moved_or_resized' if before != (form.Left, form.Top, form.Width, form.Height) else 'unchanged'
        form.Invoke(Action(move))
        return bool(started)

    def _restore_for_drag(self):
        if sys.platform != 'win32' or self.window.native is None:
            self.window.restore()
            return
        from System import Action
        from System.Drawing import Point
        from System.Windows.Forms import Cursor, FormWindowState, Screen
        form = self.window.native
        def restore():
            point = Cursor.Position
            ratio = max(0, min(1, (point.X - form.Left) / max(1, form.Width)))
            form.WindowState = FormWindowState.Normal
            area = Screen.FromPoint(point).WorkingArea
            x = max(area.Left, min(int(point.X - form.Width * ratio), area.Right - form.Width))
            y = max(area.Top, min(point.Y - int(20 * form.DeviceDpi / 96), area.Bottom - form.Height))
            form.Location = Point(x, y)
        form.Invoke(Action(restore))

    def fallback(self):
        self.dispose()
        from System import Action
        from System.Windows.Forms import FormBorderStyle
        form = self.window.native
        def apply():
            form.FormBorderStyle = FormBorderStyle.Sizable
            form.MaximizeBox = True
            form.MinimizeBox = True
            user = ctypes.windll.user32
            if self._original_style is not None:
                setter = user.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.SetWindowLongW
                setter.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]
                setter.restype = ctypes.c_ssize_t
                setter(ctypes.c_void_p(int(form.Handle.ToInt64())), -16, self._original_style | 0x00c70000)
            user.SetWindowPos(ctypes.c_void_p(int(form.Handle.ToInt64())), None, 0, 0, 0, 0, 0x0037)
        form.Invoke(Action(apply))
        self._fallback = True
        return self.state()

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
        user.IsChild.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        user.IsChild.restype = wintypes.BOOL
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
        self._original_style = style
        set_value(self._hwnd, -16, (style | 0x00070000) & ~0x00c00000)
        user.SetWindowPos(self._hwnd, None, 0, 0, 0, 0, 0x0037)
        form.HandleDestroyed += lambda sender, event: self.dispose()

    def dispose(self):
        self._pointer_drag = None
        if self._filter is not None:
            saved, self._filter = self._filter, None
            ctypes.windll.user32.UnhookWindowsHookEx(saved)
        self._drag_regions = None
        if self._old_proc and self._hwnd:
            user = ctypes.windll.user32
            user.IsWindow.argtypes = [ctypes.c_void_p]
            if user.IsWindow(self._hwnd):
                setter = user.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.SetWindowLongW
                setter(self._hwnd, -4, self._old_proc)
            self._old_proc = None
