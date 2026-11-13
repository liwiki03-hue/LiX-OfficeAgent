import json
from pathlib import Path
import sys


def main():
    import clr
    clr.AddReference('System.Windows.Forms')
    clr.AddReference('System.Drawing')
    from System.Drawing import Bitmap, Rectangle, Size
    from System.Drawing.Imaging import ImageFormat
    from System.Threading import ApartmentState, Thread, ThreadStart
    from System.Windows.Forms import Application, ContextMenuStrip
    from office_agent.webui.native_menu import NativeMenuStyle

    output = Path(sys.argv[1]).resolve()
    output.mkdir(parents=True, exist_ok=True)
    report = {}

    def run():
        menu, style = None, None
        try:
            Application.EnableVisualStyles()
            menu = ContextMenuStrip()
            for label in ['打开助手', '选择文档…', '最近文档', '-']:
                menu.Items.Add(label)
            ball = menu.Items.Add('悬浮球')
            show = ball.DropDownItems.Add('显示悬浮球')
            show.Checked = True
            ball.DropDownItems.Add('置顶').Checked = True
            ball.DropDownItems.Add('减少动态效果').Enabled = False
            menu.Items.Add('设置')
            menu.Items.Add('-')
            quit_item = menu.Items.Add('退出 LiX')
            clicked = []
            handler = lambda sender, event: clicked.append('settings')
            menu.Items[5].Click += handler
            style = NativeMenuStyle(menu)

            def capture(target, name):
                target.CreateControl()
                target.PerformLayout()
                target.Size = target.GetPreferredSize(Size.Empty)
                bitmap = Bitmap(target.Width, target.Height)
                try:
                    target.DrawToBitmap(bitmap, Rectangle(0, 0, target.Width, target.Height))
                    bitmap.Save(str(output / name), ImageFormat.Png)
                    pixel = bitmap.GetPixel(target.Width // 2, 3)
                    return {'width': target.Width, 'height': target.Height,
                            'background': [pixel.R, pixel.G, pixel.B]}
                finally:
                    bitmap.Dispose()

            quit_item.Select()
            report['root'] = capture(menu, 'tray-menu.png')
            report['submenu'] = capture(ball.DropDown, 'tray-submenu.png')
            menu.Items[5].PerformClick()
            report['actions'] = clicked
            report['checked'] = bool(show.Checked)
            report['disabled_motion'] = not ball.DropDownItems[2].Enabled
            report['root_items'] = [str(item.Text) for item in menu.Items]
            report['renderer'] = str(menu.Renderer.GetType().FullName)
            report['native_render'] = True
            report['physical_tray'] = False
        except BaseException as exc:
            report['error'] = type(exc).__name__ + ': ' + str(exc)
        finally:
            if style is not None:
                style.dispose()
            if menu is not None:
                menu.Dispose()

    thread = Thread(ThreadStart(run))
    thread.IsBackground = True
    thread.SetApartmentState(ApartmentState.STA)
    thread.Start()
    if not thread.Join(20000):
        raise TimeoutError('Native menu probe exceeded deadline')
    (output / 'NATIVE_RENDER.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), 'utf-8')
    if 'error' in report:
        raise RuntimeError(report['error'])
    assert report['actions'] == ['settings']
    assert report['checked'] and report['disabled_motion']
    assert report['root']['background'] == [32, 32, 32]
    print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    main()
