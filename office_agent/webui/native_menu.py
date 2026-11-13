from functools import lru_cache


@lru_cache(maxsize=1)
def _renderer_type():
    from System.Drawing import Color, Pen, Rectangle, RectangleF, SolidBrush, StringAlignment, StringFormat, StringFormatFlags, StringTrimming
    from System.Drawing.Drawing2D import GraphicsPath, SmoothingMode
    from System.Windows.Forms import ToolStripProfessionalRenderer

    def rounded(rect, radius):
        path = GraphicsPath()
        diameter = max(2, radius * 2)
        path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90)
        path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90)
        path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90)
        path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90)
        path.CloseFigure()
        return path

    class MenuRenderer:
        def __init__(self):
            self.renderer = ToolStripProfessionalRenderer()
            self.renderer.RenderToolStripBackground += self.OnRenderToolStripBackground
            self.renderer.RenderImageMargin += self.OnRenderImageMargin
            self.renderer.RenderToolStripBorder += self.OnRenderToolStripBorder
            self.renderer.RenderMenuItemBackground += self.OnRenderMenuItemBackground
            self.renderer.RenderItemText += self.OnRenderItemText
            self.renderer.RenderSeparator += self.OnRenderSeparator
            self.renderer.RenderItemCheck += self.OnRenderItemCheck
            self.renderer.RenderArrow += self.OnRenderArrow

        def OnRenderToolStripBackground(self, sender, event):
            event.Graphics.Clear(Color.FromArgb(32, 32, 32))

        def OnRenderImageMargin(self, sender, event):
            brush = SolidBrush(Color.FromArgb(32, 32, 32))
            try:
                event.Graphics.FillRectangle(brush, event.AffectedBounds)
            finally:
                brush.Dispose()

        def OnRenderToolStripBorder(self, sender, event):
            rect = event.ToolStrip.ClientRectangle
            rect.Width -= 1
            rect.Height -= 1
            path = rounded(rect, round(9 * event.ToolStrip.DeviceDpi / 96))
            pen = Pen(Color.FromArgb(54, 54, 54))
            state = event.Graphics.Save()
            try:
                event.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                event.Graphics.DrawPath(pen, path)
            finally:
                event.Graphics.Restore(state)
                pen.Dispose()
                path.Dispose()

        def OnRenderMenuItemBackground(self, sender, event):
            brush = SolidBrush(Color.FromArgb(32, 32, 32))
            try:
                event.Graphics.FillRectangle(brush, Rectangle(0, 0, event.Item.Width, event.Item.Height))
            finally:
                brush.Dispose()
            if event.Item.Enabled and event.Item.Selected:
                rect = Rectangle(2, 1, event.Item.Width - 4, event.Item.Height - 2)
                path = rounded(rect, round(5 * event.ToolStrip.DeviceDpi / 96))
                brush = SolidBrush(Color.FromArgb(40, 40, 40))
                state = event.Graphics.Save()
                try:
                    event.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                    event.Graphics.FillPath(brush, path)
                finally:
                    event.Graphics.Restore(state)
                    brush.Dispose()
                    path.Dispose()
            if event.Item.Checked:
                scale = event.ToolStrip.DeviceDpi / 96
                pen = Pen(Color.FromArgb(238, 238, 238), max(1.5, 1.5 * scale))
                x, y = round(10 * scale), event.Item.Height // 2
                try:
                    event.Graphics.DrawLine(pen, x, y, x + round(3 * scale), y + round(3 * scale))
                    event.Graphics.DrawLine(pen, x + round(3 * scale), y + round(3 * scale), x + round(9 * scale), y - round(3 * scale))
                finally:
                    pen.Dispose()

        def OnRenderItemText(self, sender, event):
            if not event.Text:
                return
            self.OnRenderMenuItemBackground(sender, event)
            color = Color.FromArgb(238, 238, 238) if event.Item.Enabled else Color.FromArgb(176, 176, 176)
            scale = event.ToolStrip.DeviceDpi / 96
            rect = RectangleF(round(28 * scale), 0, event.Item.Width - round(52 * scale), event.Item.Height)
            brush = SolidBrush(color)
            formatting = StringFormat()
            formatting.Alignment = StringAlignment.Near
            formatting.LineAlignment = StringAlignment.Center
            formatting.Trimming = StringTrimming.EllipsisCharacter
            formatting.FormatFlags = StringFormatFlags.NoWrap
            try:
                event.Graphics.DrawString(event.Text, event.TextFont, brush, rect, formatting)
            finally:
                formatting.Dispose()
                brush.Dispose()

        def OnRenderSeparator(self, sender, event):
            scale = event.ToolStrip.DeviceDpi / 96
            inset = round(10 * scale)
            brush = SolidBrush(Color.FromArgb(32, 32, 32))
            try:
                event.Graphics.FillRectangle(brush, Rectangle(0, 0, event.Item.Width, event.Item.Height))
            finally:
                brush.Dispose()
            pen = Pen(Color.FromArgb(64, 65, 68))
            try:
                event.Graphics.DrawLine(pen, inset, event.Item.Height // 2, event.Item.Width - inset, event.Item.Height // 2)
            finally:
                pen.Dispose()

        def OnRenderItemCheck(self, sender, event):
            rect = event.ImageRectangle
            brush = SolidBrush(Color.FromArgb(40, 40, 40) if event.Item.Selected else Color.FromArgb(32, 32, 32))
            try:
                event.Graphics.FillRectangle(brush, Rectangle(rect.Left - 2, rect.Top - 2, rect.Width + 4, rect.Height + 4))
            finally:
                brush.Dispose()
            pen = Pen(Color.FromArgb(238, 238, 238), max(1.5, 1.5 * event.ToolStrip.DeviceDpi / 96))
            state = event.Graphics.Save()
            try:
                event.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                event.Graphics.DrawLine(pen, rect.Left + rect.Width // 4, rect.Top + rect.Height // 2,
                                        rect.Left + rect.Width * 5 // 12, rect.Top + rect.Height * 2 // 3)
                event.Graphics.DrawLine(pen, rect.Left + rect.Width * 5 // 12, rect.Top + rect.Height * 2 // 3,
                                        rect.Left + rect.Width * 4 // 5, rect.Top + rect.Height // 3)
            finally:
                event.Graphics.Restore(state)
                pen.Dispose()

        def OnRenderArrow(self, sender, event):
            rect = event.ArrowRectangle
            brush = SolidBrush(Color.FromArgb(40, 40, 40) if event.Item.Selected and event.Item.Enabled else Color.FromArgb(32, 32, 32))
            try:
                event.Graphics.FillRectangle(brush, rect)
            finally:
                brush.Dispose()
            scale = event.Item.Owner.DeviceDpi / 96
            x, y = event.Item.Width - round(14 * scale), event.Item.Height // 2
            offset = max(2, round(3 * scale))
            pen = Pen(Color.FromArgb(194, 196, 201) if event.Item.Enabled else Color.FromArgb(176, 176, 176), 1.5)
            try:
                event.Graphics.DrawLine(pen, x - offset // 2, y - offset, x + offset // 2, y)
                event.Graphics.DrawLine(pen, x + offset // 2, y, x - offset // 2, y + offset)
            finally:
                pen.Dispose()

        def dispose(self):
            self.renderer.RenderToolStripBackground -= self.OnRenderToolStripBackground
            self.renderer.RenderImageMargin -= self.OnRenderImageMargin
            self.renderer.RenderToolStripBorder -= self.OnRenderToolStripBorder
            self.renderer.RenderMenuItemBackground -= self.OnRenderMenuItemBackground
            self.renderer.RenderItemText -= self.OnRenderItemText
            self.renderer.RenderSeparator -= self.OnRenderSeparator
            self.renderer.RenderItemCheck -= self.OnRenderItemCheck
            self.renderer.RenderArrow -= self.OnRenderArrow

    return MenuRenderer, rounded


class NativeMenuStyle:
    def __init__(self, menu):
        from System.Drawing import Font
        from System.Windows.Forms import ToolStripSystemRenderer
        renderer, self._rounded = _renderer_type()
        self._painter = renderer()
        self._renderer = self._painter.renderer
        self._system_renderer = ToolStripSystemRenderer()
        self._font = Font('Microsoft YaHei UI', 9.5)
        self._menus = []
        self._handlers = []
        self._attach(menu)

    def _attach(self, menu):
        from System.Windows.Forms import ToolStripMenuItem
        self._menus.append(menu)
        menu.Font = self._font
        menu.ShowImageMargin = False
        menu.ShowCheckMargin = False
        menu.DropShadowEnabled = True
        menu.ShowItemToolTips = True
        opening = lambda sender, event: self._layout(menu)
        resized = lambda sender, event: self._shape(menu)
        menu.Opening += opening
        menu.SizeChanged += resized
        self._handlers.extend((opening, resized))
        for item in menu.Items:
            if isinstance(item, ToolStripMenuItem) and item.HasDropDownItems:
                self._attach(item.DropDown)
        self._layout(menu)

    def _layout(self, menu):
        from System.Drawing import Color, Size, SystemColors
        from System.Windows.Forms import Padding, SystemInformation, ToolStripSeparator
        if menu.IsDisposed:
            return
        scale = menu.DeviceDpi / 96
        high_contrast = bool(SystemInformation.HighContrast)
        menu.ShowCheckMargin = high_contrast
        menu.Renderer = self._system_renderer if high_contrast else self._renderer
        menu.BackColor = SystemColors.Menu if high_contrast else Color.FromArgb(32, 32, 32)
        menu.ForeColor = SystemColors.MenuText if high_contrast else Color.FromArgb(238, 238, 238)
        menu.Padding = Padding(round(6 * scale))
        menu.MinimumSize = Size(round(260 * scale), 0)
        for item in menu.Items:
            item.AutoSize = False
            item.Size = Size(round(248 * scale), round((9 if isinstance(item, ToolStripSeparator) else 34) * scale))
            item.Margin = Padding(0)
        self._shape(menu)

    def _shape(self, menu):
        from System.Drawing import Region
        from System.Windows.Forms import SystemInformation
        if menu.IsDisposed:
            return
        old = menu.Region
        if SystemInformation.HighContrast or menu.Width < 24 or menu.Height < 24:
            menu.Region = None
        else:
            path = self._rounded(menu.ClientRectangle, round(9 * menu.DeviceDpi / 96))
            try:
                menu.Region = Region(path)
            finally:
                path.Dispose()
        if old is not None:
            old.Dispose()

    def dispose(self):
        for index, menu in enumerate(self._menus):
            if not menu.IsDisposed:
                menu.Opening -= self._handlers[index * 2]
                menu.SizeChanged -= self._handlers[index * 2 + 1]
            if not menu.IsDisposed and menu.Region is not None:
                region = menu.Region
                menu.Region = None
                region.Dispose()
        self._font.Dispose()
        self._painter.dispose()
        self._menus.clear()
        self._handlers.clear()
