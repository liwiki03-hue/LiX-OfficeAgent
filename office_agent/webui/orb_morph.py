from collections import deque
import ctypes
import time

from .native_surface import PixelSurface
from .orb_motion import OrbMotion
from .orb_visuals import STATUS_LABELS, action_image, action_positions, decoration_image, orb_image


def motion_preference(info, user=None, kernel=None):
    if info.HighContrast:
        return True, 'high-contrast'
    if not info.IsMenuAnimationEnabled:
        return True, 'system-menu'
    user = user or ctypes.windll.user32
    animation = ctypes.c_int()
    if not user.SystemParametersInfoW(0x1042, 0, ctypes.byref(animation), 0):
        return True, 'unavailable'
    if not animation.value:
        return True, 'system-client'
    class PowerStatus(ctypes.Structure):
        _fields_ = [('ac', ctypes.c_ubyte), ('battery', ctypes.c_ubyte), ('percent', ctypes.c_ubyte),
                    ('saving', ctypes.c_ubyte), ('life', ctypes.c_ulong), ('full', ctypes.c_ulong)]
    power = PowerStatus()
    kernel = kernel or ctypes.windll.kernel32
    if not kernel.GetSystemPowerStatus(ctypes.byref(power)):
        return True, 'unavailable'
    return (True, 'battery-saver') if power.saving == 1 else (False, 'enabled')


class MorphVisual:
    def __init__(self, owner):
        from System.Drawing import Point, Size
        from System.Windows.Forms import (AccessibleRole, Form, FormBorderStyle, FormStartPosition,
                                         Keys, MouseButtons, SystemInformation, Timer)
        self.owner, self.Point, self.Size = owner, Point, Size
        self.Form, self.Keys, self.MouseButtons = Form, Keys, MouseButtons
        self.handlers, self.buttons = [], []
        self.hovered = self.focused = self.expanded = False
        self.disposed = False
        self.SystemInformation = SystemInformation
        self.user_reduced = False
        self.system_reduced, self.system_reason = motion_preference(SystemInformation)
        self.reduced_motion = bool(self.system_reduced)
        self.highcontrast = bool(SystemInformation.HighContrast)
        self.motion = OrbMotion()
        self.frame_times, self.frame_intervals = deque(maxlen=240), deque(maxlen=240)
        self.last_frame = None
        self.decoration_key = None
        self.surface = PixelSurface(owner.orb)
        self.decoration = Form()
        self.decoration.Text = 'LiX · 悬浮球状态'
        self.decoration.FormBorderStyle = getattr(FormBorderStyle, 'None')
        self.decoration.StartPosition = FormStartPosition.Manual
        self.decoration.ShowInTaskbar = False
        self.decoration.TopMost = True
        self.glow = PixelSurface(self.decoration, clickthrough=True)
        self.timer = Timer()
        self.timer.Interval = 83
        self.timer.Tick += self._tick
        owner.orb.KeyPreview = True
        owner.orb.AccessibleRole = AccessibleRole.PushButton
        self._bind(owner.orb.MouseEnter, lambda sender, event: self._hover(True))
        self._bind(owner.orb.MouseLeave, lambda sender, event: self._hover(False))
        self._bind(owner.orb.GotFocus, lambda sender, event: self._focus(True))
        self._bind(owner.orb.LostFocus, lambda sender, event: self._focus(False))
        self._bind(owner.orb.KeyDown, self._ball_key)
        self._bind(owner.orb.VisibleChanged, lambda sender, event: self.refresh())
        self._bind(owner.orb.LocationChanged, lambda sender, event: self._relocate())

    def _bind(self, event, handler):
        event += handler
        self.handlers.append(handler)

    def _hover(self, value):
        self.hovered = value
        self.refresh()

    def _focus(self, value):
        self.focused = value
        self.refresh()

    def _ball_key(self, sender, event):
        keys = self.Keys
        if self.expanded and event.KeyCode in (keys.Tab, keys.Left, keys.Right, keys.Up, keys.Down):
            event.Handled = event.SuppressKeyPress = True
            self.buttons[3 if event.Shift or event.KeyCode in (keys.Left, keys.Up) else 0][0].Activate()
        elif event.KeyCode in (keys.Apps, keys.Space) or (event.KeyCode == keys.F10 and event.Shift):
            event.Handled = event.SuppressKeyPress = True
            self.toggle_actions()
        elif event.KeyCode == keys.Enter:
            event.Handled = event.SuppressKeyPress = True
            if self.expanded:
                self.hide_actions(focus=True)
            else:
                self.owner.present('compact')
        elif event.KeyCode == keys.Escape and self.expanded:
            event.Handled = event.SuppressKeyPress = True
            self.hide_actions(focus=True)

    def refresh(self):
        if self.disposed:
            return
        visible = self.owner.orb.Visible and not self.owner.hidden and not self.owner.exiting
        self.motion.update(visible=visible, reduced=self.reduced_motion, dragging=self.owner._moved and self.owner._drag is not None,
                           hovered=self.hovered, focused=self.focused, drop=self.owner._accept, expanded=self.expanded, status=self.owner._status)
        if not visible:
            self.timer.Stop()
            self.last_frame = None
            self.decoration.Hide()
            self.hide_actions(immediate=True)
            return
        self.redraw()
        self._schedule()

    def _schedule(self):
        interval = self.motion.interval()
        if not interval:
            self.timer.Stop()
            self.last_frame = None
        else:
            if self.timer.Interval != interval:
                self.timer.Interval = interval
            self.timer.Start()

    def _tick(self, sender, event):
        if self.disposed or not self.owner.orb.Visible or self.owner.hidden or self.owner.exiting:
            self.timer.Stop()
            return
        self.redraw()
        self._schedule()

    def redraw(self):
        if self.disposed or not self.owner.orb.Visible:
            return
        started = time.perf_counter()
        now = time.monotonic()
        orb = self.owner.orb
        size, status = orb.Width, self.owner._status
        frame = self.motion.frame()
        image = action_image(size, 'x', self.focused, self.highcontrast) if self.expanded or frame['menu'] > 0 else orb_image(size, status, frame['phase'], self.hovered or self.owner._accept, self.focused, self.highcontrast, frame['pulse'], frame['emphasis'])
        self.surface.draw(image)
        state = '松开即可打开' if self.owner._accept else STATUS_LABELS[status]
        orb.AccessibleDescription = state + '；左键打开，右键或Shift+F10展开操作，拖动可移动'
        self._relocate()
        if self.buttons:
            if self.expanded or frame['menu'] > 0:
                self._animate(frame['menu'])
            else:
                for form, surface, icon, callback in self.buttons:
                    if form.Visible:
                        form.Hide()
        if self.last_frame is not None:
            self.frame_intervals.append((now - self.last_frame) * 1000)
        self.last_frame = now
        self.frame_times.append((time.perf_counter() - started) * 1000)

    def _relocate(self):
        if self.disposed or not self.owner.orb.Visible:
            return
        orb = self.owner.orb
        width, height = orb.Width * 2, round(orb.Width * 1.95)
        if self.decoration.Width != width or self.decoration.Height != height:
            self.decoration.Size = self.Size(width, height)
        area = self.owner._Screen.FromControl(orb).WorkingArea
        x = max(area.Left, min(orb.Left - orb.Width // 2, area.Right - width))
        y = max(area.Top, min(orb.Top - round(orb.Width * .22), area.Bottom - height))
        if self.decoration.Left != x or self.decoration.Top != y:
            self.decoration.Location = self.Point(x, y)
        key = (orb.Width, self.owner._status, self.hovered or self.focused, self.owner._accept,
               (orb.Left - x + orb.Width / 2, orb.Top - y + orb.Width / 2))
        if key != self.decoration_key:
            self.glow.draw(decoration_image(*key))
            self.decoration_key = key
        if not self.decoration.Visible:
            self.decoration.Show(orb)
        if self.expanded:
            self._place_actions()
            self._animate(self.motion.menu.value(time.monotonic()))

    def _create_actions(self):
        from System.Windows.Forms import AccessibleRole, FormBorderStyle, FormStartPosition
        actions = [('选择文档', 'folder-open', self.owner.choose),
                   ('返回工作区', 'message-square', lambda: self.owner.present('compact')),
                   ('历史记录', 'history', self.owner.recent), ('设置', 'settings', self.owner.settings)]
        descriptions = ('选择文档：打开本地文件选择窗口', '返回工作区：打开助手对话与文档工作区',
                        '历史记录：查看最近文档与历史任务', '设置：管理模型接口与应用偏好')
        for index, (label, icon, callback) in enumerate(actions):
            form = self.Form()
            form.Text = 'LiX · ' + label
            form.AccessibleName = label
            form.AccessibleRole = AccessibleRole.PushButton
            form.AccessibleDescription = descriptions[index] + '；方向键或Tab切换，Enter打开，Esc收起'
            form.FormBorderStyle = getattr(FormBorderStyle, 'None')
            form.StartPosition = FormStartPosition.Manual
            form.ShowInTaskbar = False
            form.TopMost = True
            form.KeyPreview = True
            surface = PixelSurface(form, noactivate=False)
            self.buttons.append((form, surface, icon, callback))
            self.owner.tooltip.SetToolTip(form, descriptions[index])
            self._bind(form.MouseDown, lambda sender, event: self.owner.tooltip.Hide(sender))
            self._bind(form.MouseUp, lambda sender, event, i=index: self._activate(i) if event.Button == self.MouseButtons.Left else None)
            self._bind(form.KeyDown, lambda sender, event, i=index: self._action_key(i, event))
            self._bind(form.Activated, lambda sender, event, i=index: self._paint_action(i, True))
            self._bind(form.Deactivate, lambda sender, event, i=index: self._deactivated(i))
            self._bind(form.MouseEnter, lambda sender, event, i=index: self._paint_action(i, True))
            self._bind(form.MouseLeave, lambda sender, event, i=index: self._paint_action(i, self.Form.ActiveForm == self.buttons[i][0]))

    def _paint_action(self, index, focused=False):
        if self.disposed or not self.buttons:
            return
        form, surface, icon, callback = self.buttons[index]
        surface.draw(action_image(form.Width, icon, focused, self.highcontrast))

    def _deactivated(self, index):
        if self.disposed:
            return
        self._paint_action(index)
        def check():
            if not self.disposed and self.expanded and self.Form.ActiveForm not in [self.owner.orb] + [entry[0] for entry in self.buttons]:
                self.hide_actions()
        self.owner.form.BeginInvoke(self.owner._Action(check))

    def _action_key(self, index, event):
        keys = self.Keys
        if event.KeyCode in (keys.Tab, keys.Left, keys.Right, keys.Up, keys.Down):
            event.Handled = event.SuppressKeyPress = True
            previous = event.Shift or event.KeyCode in (keys.Left, keys.Up)
            self.buttons[(index + (-1 if previous else 1)) % 4][0].Activate()
        elif event.KeyCode in (keys.Enter, keys.Space):
            event.Handled = event.SuppressKeyPress = True
            self._activate(index)
        elif event.KeyCode == keys.Escape:
            event.Handled = event.SuppressKeyPress = True
            self.hide_actions(focus=True)

    def _activate(self, index):
        if not self.expanded or self.motion.menu.running(time.monotonic()):
            return
        callback = self.buttons[index][3]
        self.hide_actions()
        self.owner.observation('orb_action', {'action': ('choose', 'workspace', 'history', 'settings')[index]})
        self.owner._action(callback)

    def _place_actions(self):
        orb = self.owner.orb
        area = self.owner._Screen.FromControl(orb).WorkingArea
        scale = max(1, float(orb.DeviceDpi) / 96)
        self.targets = action_positions((orb.Left, orb.Top), scale, (area.Left, area.Top, area.Right, area.Bottom))
        for (form, surface, icon, callback), (x, y) in zip(self.buttons, self.targets):
            size = self.Size(round(40 * scale), round(40 * scale))
            if form.Size != size:
                form.Size = size

    def toggle_actions(self):
        if self.expanded:
            self.hide_actions(focus=True)
            return
        if not self.buttons:
            self._create_actions()
        self.expanded = True
        try:
            self._place_actions()
        except ValueError as error:
            self.expanded = False
            self.owner.warning(str(error))
            return
        self.owner.tooltip.Hide(self.owner.orb)
        self.owner.tooltip.SetToolTip(self.owner.orb, '收起快捷菜单')
        self.motion.update(expanded=True)
        for index, (form, surface, icon, callback) in enumerate(self.buttons):
            self._paint_action(index)
            if not form.Visible:
                form.Show(self.owner.orb)
        self._animate(self.motion.menu.value(time.monotonic()))
        self.buttons[0][0].Activate()
        self.refresh()
        self.owner.observation('orb_actions_visible', {'visible': True, 'count': len(self.buttons)})

    def _animate(self, fraction):
        if not self.owner.orb.Visible:
            return
        orb = self.owner.orb
        for (form, surface, icon, callback), (x, y) in zip(self.buttons, self.targets):
            start_x, start_y = orb.Left + (orb.Width - form.Width) / 2, orb.Top + (orb.Height - form.Height) / 2
            next_position = self.Point(round(start_x + (x - start_x) * fraction), round(start_y + (y - start_y) * fraction))
            if form.Location != next_position:
                form.Location = next_position

    def hide_actions(self, focus=False, immediate=False):
        if not self.expanded and not immediate:
            return
        was_expanded = self.expanded
        self.expanded = False
        self.owner.tooltip.Hide(self.owner.orb)
        self.owner.tooltip.SetToolTip(self.owner.orb, '打开助手；右键展开快捷菜单，拖动可移动，拖入文档可打开')
        for form, surface, icon, callback in self.buttons:
            self.owner.tooltip.Hide(form)
        self.motion.update(expanded=False)
        if immediate or self.reduced_motion:
            self.motion.menu.move(0, time.monotonic(), 0)
            for form, surface, icon, callback in self.buttons:
                form.Hide()
        if focus and self.owner.orb.Visible:
            user = self.surface.user
            user.SetWindowLongW(self.surface.hwnd, -20, user.GetWindowLongW(self.surface.hwnd, -20) & ~0x08000000)
            self.owner.orb.Activate()
        if was_expanded:
            self.owner.observation('orb_actions_visible', {'visible': False})
        if self.owner.orb.Visible and not self.owner.hidden and not self.owner.exiting:
            self.refresh()

    def set_reduced_motion(self, value):
        self.user_reduced = bool(value)
        self.reduced_motion = self.user_reduced or self.system_reduced
        self.owner.observation('orb_motion_preference', {'reduced': self.reduced_motion, 'user': self.user_reduced, 'system': self.system_reduced})
        self.refresh()

    def check_environment(self):
        if self.disposed:
            return
        highcontrast = bool(self.SystemInformation.HighContrast)
        system, self.system_reason = motion_preference(self.SystemInformation)
        if system != self.system_reduced or highcontrast != self.highcontrast:
            self.system_reduced, self.highcontrast = bool(system), highcontrast
            self.decoration_key = None
            for index in range(len(self.buttons)):
                self._paint_action(index)
            self.set_reduced_motion(self.user_reduced)

    def dispose(self):
        if self.disposed:
            return
        self.disposed = True
        self.expanded = False
        self.timer.Stop()
        self.timer.Dispose()
        for form, surface, icon, callback in self.buttons:
            form.Dispose()
        self.buttons.clear()
        self.decoration.Dispose()
        self.handlers.clear()
