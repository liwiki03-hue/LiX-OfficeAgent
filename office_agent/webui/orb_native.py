import ctypes
import math
from pathlib import Path
import re
import threading

from .orb_visuals import STATUS_LABELS


def clamp_position(x, y, width, height, area):
    left, top, right, bottom = area
    return (max(left, min(int(x), max(left, right - width))),
            max(top, min(int(y), max(top, bottom - height))))


def drop_candidates(paths):
    if not isinstance(paths, (tuple, list)) or not 1 <= len(paths) <= 8:
        return False
    from .document_open import OPEN_FORMATS
    return all(isinstance(path, str) and 0 < len(path) <= 32768
               and Path(path).suffix.lower() in OPEN_FORMATS
               and not path.startswith(('\\\\', 'http:', 'https:')) for path in paths)


def parse_hotkey(value):
    if value == '':
        return None
    if not isinstance(value, str) or len(value) > 64:
        raise ValueError('Invalid hotkey')
    parts = value.upper().split('+')
    key, modifiers = parts[-1], parts[:-1]
    choices = {'CTRL': 2, 'ALT': 1, 'SHIFT': 4}
    if len(modifiers) != len(set(modifiers)) or not modifiers or any(p not in choices for p in modifiers):
        raise ValueError('Use CTRL/ALT/SHIFT and a letter or F1-F12')
    if re.fullmatch('[A-Z]', key):
        code = ord(key)
    elif re.fullmatch(r'F(?:[1-9]|1[0-2])', key):
        code = 111 + int(key[1:])
    else:
        raise ValueError('Unsupported hotkey key')
    return sum(choices[p] for p in modifiers) | 0x4000, code


def orb_points(count=140):
    result = []
    for index in range(count):
        y = 1-2*(index+.5)/count
        radius = math.sqrt(max(0,1-y*y))
        angle = index*math.pi*(3-math.sqrt(5))
        x,z = math.cos(angle)*radius, math.sin(angle)*radius
        result.append((x,y,z))
    return sorted(result,key=lambda point:point[2])


class NativeOrb:
    def __init__(self, window, *, dropped, presentation=None, can_exit=None, choose=None,
                 settings=None, recent=None, position=None, hotkey='', wake=None, warning=None,
                 observation=None, tray_label='LiX 文档助手', preferences=None):
        self.window = window
        self.dropped = dropped
        self.presentation = presentation or (lambda value: None)
        self.can_exit = can_exit or (lambda: True)
        self.choose = choose or (lambda: None)
        self.settings = settings or (lambda: None)
        self.recent = recent or (lambda: None)
        self.position = position or (lambda value: None)
        self.hotkey = hotkey
        self.wake = wake or (lambda: False)
        self.warning = warning or (lambda value: None)
        self.observation = observation or (lambda kind, payload: None)
        self.tray_label = tray_label[:63]
        self._old_proc = self._wndproc = None
        self._hotkey_registered = False
        self.orb = self.tray = self.timer = None
        self.mode = 'orb'
        self.exiting = False
        self.hidden = False
        from .window_preferences import WindowPreferences
        self._preferences = preferences if preferences is not None else WindowPreferences()
        self._bounds = self._preferences.read()
        self.topmost = self._bounds.get('topmost', False)
        self._last_panel_mode = 'compact'
        self._maximized_modes = {}
        self._positioning = False
        self._drag = None
        self._moved = False
        self._accept = False
        self._status = 'ready'
        self._handlers = []
        self.actions = None
        self.visual = None
        self.menu_style = None

    def install(self):
        import clr
        clr.AddReference('System.Windows.Forms')
        clr.AddReference('System.Drawing')
        from System import Action
        self._Action = Action
        self.form = self.window.native
        if self.form is None:
            raise RuntimeError('WinForms native window is unavailable')
        self.form.Invoke(Action(self._create))

    def _ui(self, action):
        if self.form.IsDisposed:
            return
        if self.form.InvokeRequired:
            self.form.BeginInvoke(self._Action(action))
        else:
            action()

    def _create(self):
        from System.Drawing import Color, Font, Point, Rectangle, Region, Size
        from System.Drawing.Drawing2D import GraphicsPath, SmoothingMode
        from System.Windows.Forms import (ContextMenuStrip, Cursors, DataFormats, DragDropEffects,
            Form, FormBorderStyle, FormStartPosition, MouseButtons, NotifyIcon, Screen,
            TextFormatFlags, TextRenderer, Timer, ToolTip, SystemInformation)
        self._Point, self._Size = Point, Size
        self._Screen, self._Color = Screen, Color
        self._Region, self._Path = Region, GraphicsPath
        self._DataFormats, self._Effects = DataFormats, DragDropEffects
        self._MouseButtons, self._SystemInformation = MouseButtons, SystemInformation
        self.orb = Form()
        self.orb.Text = 'LiX · 拖入文档'
        self.orb.AccessibleName = 'LiX 悬浮球；拖入文档或单击打开助手'
        self.orb.FormBorderStyle = getattr(FormBorderStyle, 'None')
        self.orb.ShowInTaskbar = True
        self.orb.StartPosition = FormStartPosition.Manual
        self.orb.TopMost = self.topmost
        self.orb.AllowDrop = True
        self.orb.BackColor = Color.FromArgb(28, 31, 36)
        self.orb.Cursor = Cursors.Hand
        self.font = Font('Microsoft YaHei UI', 9)
        self.tooltip = ToolTip()
        self.tooltip.ShowAlways = True
        self.tooltip.InitialDelay = 500
        self.tooltip.ReshowDelay = 500
        self.tooltip.AutoPopDelay = 10000
        self.tooltip.SetToolTip(self.orb, '打开助手；右键展开快捷菜单，拖动可移动，拖入文档可打开')
        self._shape()
        from .orb_morph import MorphVisual
        self.visual = MorphVisual(self)
        work = Screen.PrimaryScreen.WorkingArea
        self.orb.Location = Point(work.Right - self.orb.Width - 16, work.Top + work.Height // 2)

        def paint(sender, event):
            self.visual.redraw()

        def down(sender, event):
            self.tooltip.Hide(self.orb)
            if event.Button == MouseButtons.Left:
                self._drag = (self.orb.PointToScreen(event.Location), self.orb.Location)
                self._moved = False
                self.orb.Capture = True

        def move(sender, event):
            if self._drag is None:
                return
            now = self.orb.PointToScreen(event.Location)
            origin, location = self._drag
            dx, dy = now.X - origin.X, now.Y - origin.Y
            threshold = SystemInformation.DragSize
            if self._moved or abs(dx) >= threshold.Width or abs(dy) >= threshold.Height:
                first_move = not self._moved
                self._moved = True
                if first_move:
                    self.visual.hide_actions(immediate=True)
                    self.visual.refresh()
                area = Screen.FromPoint(now).WorkingArea
                x, y = clamp_position(location.X + dx, location.Y + dy, self.orb.Width, self.orb.Height,
                                      (area.Left, area.Top, area.Right, area.Bottom))
                self.orb.Location = Point(x, y)

        def up(sender, event):
            if self._drag is None or event.Button != MouseButtons.Left:
                return
            moved = self._moved
            self._drag = None
            self.orb.Capture = False
            self.visual.refresh()
            if moved:
                self.position(self.location())
            elif self.visual.expanded:
                self.visual.hide_actions(focus=True)
            else:
                self.restore()

        def enter(sender, event):
            try:
                paths = event.Data.GetData(DataFormats.FileDrop) if event.Data.GetDataPresent(DataFormats.FileDrop) else None
            except Exception:
                paths = None
            self._accept = drop_candidates(tuple(str(p) for p in paths)) if paths is not None else False
            event.Effect = DragDropEffects.Copy if self._accept else getattr(DragDropEffects, 'None')
            self.tooltip.SetToolTip(self.orb, '松开以打开' if self._accept else '请拖入一个本地 DOCX、XLSX 或 PPTX 文件')
            self.visual.refresh()

        def leave(sender, event):
            self._accept = False
            self.tooltip.SetToolTip(self.orb, '收起快捷菜单' if self.visual.expanded else '打开助手；右键展开快捷菜单，拖动可移动，拖入文档可打开')
            self.visual.refresh()

        def drop(sender, event):
            try:
                paths = event.Data.GetData(DataFormats.FileDrop) if event.Data.GetDataPresent(DataFormats.FileDrop) else None
                accepted = tuple(str(p) for p in paths) if paths is not None else ()
            except Exception:
                accepted = ()
            leave(sender, event)
            if drop_candidates(accepted):
                self.present('review')
                try:
                    result = self.dropped(accepted)
                    self.observation('native_drop_received', result if isinstance(result, dict) else {'count': len(accepted)})
                except Exception:
                    self.warning('文件接收未完成，请再次拖入；尚未启动模型或修改文件')
            else:
                self.warning('未打开：悬浮球仅支持本地DOCX/XLSX/PPTX；拒绝目录、链接与旧格式')

        self._handlers.extend((paint, down, move, up, enter, leave, drop))
        self.orb.Paint += paint
        self.orb.MouseDown += down
        self.orb.MouseMove += move
        self.orb.MouseUp += up
        self.orb.DragEnter += enter
        self.orb.DragOver += enter
        self.orb.DragLeave += leave
        self.orb.DragDrop += drop
        self.menu = ContextMenuStrip()

        def add(label, callback, parent=None):
            item = (parent.DropDownItems if parent is not None else self.menu.Items).Add(label)
            handler = lambda sender, event: callback()
            item.Click += handler
            self._handlers.append(handler)
            return item

        add('打开助手', lambda: self.restore())
        add('选择文档…', lambda: self._action(self.choose))
        add('最近文档', lambda: self._action(self.recent))
        self.menu.Items.Add('-')
        ball_menu = self.menu.Items.Add('悬浮球')
        self.ball_item = add('显示悬浮球', self.toggle_ball, ball_menu)
        self.ball_item.Checked = not self.hidden
        self.pin = add('置顶', self.toggle_topmost, ball_menu)
        self.pin.Checked = self.topmost
        self.motion_item = add('减少动态效果', self.toggle_motion, ball_menu)
        self.motion_item.ToolTipText = '系统减少动态、高对比度或节电模式始终优先；不写入配置文件'
        add('设置', lambda: self._action(self.settings))
        self.menu.Items.Add('-')
        add('退出 LiX', self.exit)
        from .native_menu import NativeMenuStyle
        self.menu_style = NativeMenuStyle(self.menu)
        def opening(sender, event):
            self.ball_item.Checked = not self.hidden
            self.pin.Checked = self.topmost
            self._motion_menu()
        self.menu.Opening += opening
        self._handlers.append(opening)
        self.orb.MouseUp += lambda sender,event: self.show_actions() if event.Button==MouseButtons.Right else None
        self.tray = NotifyIcon()
        self.tray.Icon = self.form.Icon
        self.orb.Icon = self.tray.Icon
        self.tray.Text = self.tray_label
        self.tray.ContextMenuStrip = self.menu
        self.tray.Visible = True
        def restore(sender, event):
            if event.Button == MouseButtons.Left:
                self.observation('tray_restore_requested', {'mode': self.mode})
                self.restore()
        self._handlers.append(restore)
        self.tray.MouseClick += restore
        self.form.FormClosing += self._closing
        self.form.ResizeEnd += lambda sender, event: self._remember_panel()
        self.timer = Timer()
        self.timer.Interval = 1500
        self.timer.Tick += self._screens_changed
        self.timer.Start()
        self.form.Hide()
        self._show_ball()
        self.visual.check_environment()
        self._motion_menu()
        self._register_hotkey()

    def _register_hotkey(self):
        try:
            parsed = parse_hotkey(self.hotkey)
        except ValueError:
            self.warning('快捷键格式无效；请使用托盘或悬浮球打开助手。')
            return
        if parsed is None:
            return
        from ctypes import wintypes
        user = ctypes.windll.user32
        handle = ctypes.c_void_p(int(self.orb.Handle.ToInt64()))
        if not user.RegisterHotKey(handle, 0x4F41, *parsed):
            self.warning('唤出快捷键已被占用；请使用托盘或悬浮球。')
            return
        callback_type = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        setter = user.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.SetWindowLongW
        setter.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        setter.restype = ctypes.c_void_p
        caller = user.CallWindowProcW
        caller.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        caller.restype = ctypes.c_ssize_t
        def procedure(hwnd, message, wparam, lparam):
            if message == 0x0312 and wparam == 0x4F41:
                try:
                    self.restore()
                except Exception:
                    self.warning('快捷键唤出失败；请使用托盘。')
                return 0
            return caller(self._old_proc, hwnd, message, wparam, lparam)
        self._wndproc = callback_type(procedure)
        self._old_proc = setter(handle, -4, ctypes.cast(self._wndproc, ctypes.c_void_p))
        if not self._old_proc:
            user.UnregisterHotKey(handle, 0x4F41)
            self._wndproc = None
            self.warning('快捷键消息接收不可用；请使用托盘。')
            return
        self._hotkey_registered = True

    def _shape(self):
        scale = max(1.0, float(self.orb.DeviceDpi) / 96)
        size = max(56, int(math.ceil(64 * scale)))
        height = size
        self.orb.MinimumSize = self._Size(size, height)
        self.orb.MaximumSize = self._Size(size, height)
        self.orb.Size = self._Size(size, height)
        path = self._Path()
        old = self.orb.Region
        try:
            path.AddEllipse(0, 0, size, size)
            self.orb.Region = self._Region(path)
        finally:
            path.Dispose()
            if old is not None:
                old.Dispose()

    def _show_ball(self):
        if self.hidden or self.exiting:
            return
        handle = int(self.orb.Handle.ToInt64())
        user = ctypes.windll.user32
        style = user.GetWindowLongW(ctypes.c_void_p(handle), -20)
        user.SetWindowLongW(ctypes.c_void_p(handle), -20, style | 0x08000000)
        self.orb.Show()
        user.ShowWindow(ctypes.c_void_p(handle), 4)
        user.SetWindowPos(ctypes.c_void_p(handle), ctypes.c_void_p(-1 if self.topmost else -2),
                          0, 0, 0, 0, 0x0010 | 0x0001 | 0x0002 | 0x0040)

    def _screens_changed(self, sender, event):
        if self.exiting or self.orb is None:
            return
        if self.wake():
            self.restore()
        self.visual.check_environment()
        self._motion_menu()
        area = self._Screen.FromControl(self.orb).WorkingArea
        x, y = clamp_position(self.orb.Left, self.orb.Top, self.orb.Width, self.orb.Height,
                              (area.Left, area.Top, area.Right, area.Bottom))
        self.orb.Location = self._Point(x, y)
        if self.orb.Width != int(math.ceil(64 * max(1.0, float(self.orb.DeviceDpi) / 96))):
            self._shape()
            self.visual.refresh()

    def location(self):
        area = self._Screen.FromControl(self.orb).WorkingArea
        return {'screen': str(self._Screen.FromControl(self.orb).DeviceName),
                'x': (self.orb.Left - area.Left) / max(1, area.Width),
                'y': (self.orb.Top - area.Top) / max(1, area.Height)}

    def diagnostics(self):
        result = {}
        def inspect():
            result.update(visible=bool(self.orb.Visible), size=[self.orb.Width, self.orb.Height],
                location=[self.orb.Left, self.orb.Top], dpi=int(self.orb.DeviceDpi),
                corner_in_region=bool(self.orb.Region.IsVisible(0, 0)),
                center_in_region=bool(self.orb.Region.IsVisible(self.orb.Width // 2, self.orb.Height // 2)))
            if self.visual is not None:
                times = sorted(self.visual.frame_times)
                intervals = sorted(self.visual.frame_intervals)
                result.update(visual='morph-per-pixel', status=self._status,
                              redraws=self.visual.surface.draws, animation=bool(self.visual.timer.Enabled),
                              actions_visible=bool(self.visual.expanded), action_windows=len(self.visual.buttons),
                              reduced_motion=self.visual.reduced_motion, system_reduced=self.visual.system_reduced,
                              motion_reason=self.visual.system_reason,
                              animation_interval_ms=int(self.visual.timer.Interval),
                              frame_samples=len(times), frame_ms_p95=times[max(0, math.ceil(len(times) * .95) - 1)] if times else None,
                              interval_ms_p95=intervals[max(0, math.ceil(len(intervals) * .95) - 1)] if intervals else None,
                              decoration_redraws=self.visual.glow.draws,
                              action_redraws=sum(surface.draws for form, surface, icon, callback in self.visual.buttons))
                result.update(hovered=self.visual.hovered, focused=self.visual.focused, drop_candidate=self._accept,
                              menu_position=self.visual.motion.menu.value(self.visual.motion.clock()))
        self.form.Invoke(self._Action(inspect))
        return result

    def restore_position(self, position):
        if not isinstance(position, dict) or set(position) != {'screen', 'x', 'y'}:
            return False
        if not isinstance(position['screen'], str) or len(position['screen']) > 256:
            return False
        if any(type(position[k]) not in (int, float) or not math.isfinite(position[k]) or not 0 <= position[k] <= 1 for k in ('x', 'y')):
            return False
        def apply():
            screen = next((s for s in self._Screen.AllScreens if s.DeviceName == position['screen']), self._Screen.PrimaryScreen)
            area = screen.WorkingArea
            x, y = clamp_position(area.Left + position['x'] * area.Width, area.Top + position['y'] * area.Height,
                self.orb.Width, self.orb.Height, (area.Left, area.Top, area.Right, area.Bottom))
            self.orb.Location = self._Point(x, y)
        self._ui(apply)
        return True

    def set_status(self, status):
        if status not in STATUS_LABELS:
            return
        def apply():
            if self.exiting or status == self._status:
                return
            self._status = status
            if self.visual is not None:
                self.visual.refresh()
                self.observation('orb_status', {'status': status})
        self._ui(apply)

    def _motion_menu(self):
        self.motion_item.Checked = self.visual.reduced_motion
        self.motion_item.Enabled = not self.visual.system_reduced
        reasons = {'high-contrast': '系统高对比度已启用', 'system-menu': '系统菜单动画已关闭',
                   'system-client': '系统动态效果已关闭', 'battery-saver': '系统正在节电', 'unavailable': '系统动态设置暂不可读取'}
        self.motion_item.ToolTipText = reasons.get(self.visual.system_reason, '仅本次运行；不写入配置文件')

    def toggle_motion(self):
        def apply():
            self.visual.set_reduced_motion(not self.visual.user_reduced)
            self._motion_menu()
        self._ui(apply)

    def show_actions(self):
        if self.visual is not None:
            self._ui(self.visual.toggle_actions)

    def _action(self, callback):
        if callback is not self.choose:
            self.restore()
        threading.Thread(target=callback, daemon=True).start()

    def restore(self):
        self.present(self._last_panel_mode)

    def _remember_panel(self):
        if self._positioning or self.mode not in {'compact', 'review'}:
            return
        maximum = str(self.form.WindowState) == 'Maximized'
        if str(self.form.WindowState) == 'Minimized':
            return
        bounds = self.form.RestoreBounds if maximum else self.form.Bounds
        self._bounds[self.mode] = {'x': int(bounds.X), 'y': int(bounds.Y), 'width': int(bounds.Width),
                                  'height': int(bounds.Height), 'dpi': int(self.form.DeviceDpi),
                                  'screen': str(self._Screen.FromControl(self.form).DeviceName)}
        self._maximized_modes[self.mode] = maximum
        try:
            self._preferences.write(self._bounds)
        except (OSError, ValueError):
            from ..diagnostics import record
            record('window', 'save_bounds', 'failed', code='preference_io')

    def present(self, mode):
        if mode not in {'orb', 'compact', 'review', 'hidden'}:
            raise ValueError('Unknown shell presentation')
        def apply():
            if self.exiting:
                return
            if mode in {'compact', 'review'} and mode == self.mode and self.form.Visible:
                self.form.Activate()
                return
            self._remember_panel()
            self.mode = mode
            if mode in {'orb', 'hidden'}:
                self.form.Hide()
                if mode == 'orb':
                    self._show_ball()
                else:
                    self.orb.Hide()
            else:
                from System.Windows.Forms import FormWindowState
                self._positioning = True
                self._last_panel_mode = mode
                self.orb.Hide()
                saved = self._bounds.get(mode)
                screen = next((screen for screen in self._Screen.AllScreens if saved and screen.DeviceName == saved['screen']), self._Screen.FromControl(self.orb))
                area = screen.WorkingArea
                scale = max(1.0, float(self.form.DeviceDpi) / 96)
                width, height = ((1280, 820) if getattr(self, 'paged_document', False) else (1120, 760)) if mode == 'review' else (420, 640)
                ratio = float(self.form.DeviceDpi) / saved['dpi'] if saved else scale
                width, height = min(int((saved['width'] if saved else width) * ratio), area.Width), min(int((saved['height'] if saved else height) * ratio), area.Height)
                self.form.MinimumSize = self._Size(min(int(360 * scale), area.Width), min(int(440 * scale), area.Height))
                self.form.WindowState = FormWindowState.Normal
                self.form.Size = self._Size(width, height)
                x, y = clamp_position(saved['x'] if saved else self.orb.Left - width - 12, saved['y'] if saved else self.orb.Top - 40, width, height,
                                      (area.Left, area.Top, area.Right, area.Bottom))
                self.form.Location = self._Point(x, y)
                self.form.Show()
                self.form.TopMost = self.topmost
                if self._maximized_modes.get(mode):
                    self.form.WindowState = FormWindowState.Maximized
                self._positioning = False
                self.form.Activate()
                self.observation('panel_visible', {'visible': bool(self.form.Visible),
                    'state': str(self.form.WindowState), 'mode': mode})
            self.presentation(mode)
        self._ui(apply)

    def toggle_ball(self):
        self.hidden = not self.hidden
        self.observation('ball_visibility_requested', {'hidden': self.hidden})
        self._ui(lambda: self.orb.Hide() if self.hidden else self._show_ball())

    def toggle_topmost(self):
        self.topmost = not self.topmost
        def apply():
            self.orb.TopMost = self.topmost
            self.form.TopMost = self.topmost
            self.pin.Checked = self.topmost
            self._bounds['topmost'] = self.topmost
            try:
                self._preferences.write(self._bounds)
            except (OSError, ValueError):
                pass
        self._ui(apply)

    def _closing(self, sender, event):
        if not self.exiting:
            event.Cancel = True
            self.present('orb')

    def exit(self):
        def apply():
            if not self.can_exit():
                self.tooltip.Show('请先完成当前步骤或处理未保存草稿；在助手“更多→关闭当前文档”中保存或保留恢复草稿后再退出。', self.orb, 0, 0, 5000)
                self.present('review')
                return
            self.exiting = True
            self.observation('exit_started', {})
            self.dispose()
            self.form.Close()
        self._ui(apply)

    def dispose(self):
        if self.menu_style is not None:
            self.menu_style.dispose()
            self.menu_style = None
        if self.visual is not None:
            self.visual.dispose()
            self.visual = None
        if self.actions is not None:
            for control in self.actions.Controls:
                if control.Region is not None:
                    control.Region.Dispose()
                    control.Region = None
            self.actions.Dispose()
            self.actions = None
        if self._hotkey_registered and self.orb is not None:
            handle = ctypes.c_void_p(int(self.orb.Handle.ToInt64()))
            user = ctypes.windll.user32
            user.UnregisterHotKey(handle, 0x4F41)
            setter = user.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8 else user.SetWindowLongW
            setter.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
            setter.restype = ctypes.c_void_p
            setter(handle, -4, self._old_proc)
            self._hotkey_registered = False
            self._old_proc = None
        if self.timer is not None:
            self.timer.Stop()
            self.timer.Dispose()
            self.timer = None
        if self.tray is not None:
            self.tray.Visible = False
            self.tray.Dispose()
            self.tray = None
        if self.orb is not None:
            region = self.orb.Region
            self.orb.Region = None
            if region is not None:
                region.Dispose()
            self.orb.Dispose()
            self.orb = None
        for value in ('font', 'tooltip', 'menu'):
            resource = getattr(self, value, None)
            if resource is not None:
                resource.Dispose()
                setattr(self, value, None)
