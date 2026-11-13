from types import SimpleNamespace

import pytest

from office_agent.webui.bridge import WebviewApp
from tests.test_professional_browser import STATIC, _browser
from tests.test_web_images import PNG


def test_file_picker_filter_matches_installed_webview_parser():
    from webview.util import parse_file_type
    from webview import FileDialog

    app = WebviewApp()
    calls = []

    def dialog(kind, **kwargs):
        assert kind == FileDialog.OPEN
        calls.extend(parse_file_type(value) for value in kwargs['file_types'])
        return None

    app.window = SimpleNamespace(create_file_dialog=dialog)
    assert app.select_file() == {'ok': False, 'reason': 'cancelled'}
    assert calls[0][1] == '*.doc;*.docx;*.xls;*.xlsx;*.pptx;*.pdf'
    assert calls[1][1] == '*.*'
    assert app.path is None and app.editor is None


@pytest.mark.parametrize('width,theme', [(1440, 'light'), (1280, 'dark'), (520, 'highcontrast')])
def test_image_confirm_layout_keyboard_and_preview_scope(tmp_path, monkeypatch, width, theme):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, _, _, calls, _ = next(fixture)
        page.set_viewport_size({'width': width, 'height': 800})
        page.evaluate('(theme) => document.documentElement.dataset.theme = theme', theme)
        page.add_script_tag(path=str(STATIC / 'images.js'))
        page.add_script_tag(path=str(STATIC / 'imagegen.js'))
        page.evaluate('''png => {
          const normal = document.createElement('div'); normal.id = 'ordinary-preview'; document.body.append(normal);
          OfficeImages.appendImages(normal, [{data_url: png}]);
          const card = document.createElement('article'); card.id = 'generation-preview'; document.body.append(card);
          window.imageCalls = [];
          OfficeImageGen.handleEvent({type: 'image_result', plan_id: 'test-plan', images: [{data_url: png}]},
            async (...args) => {imageCalls.push(args); return {ok: true};}, node => card.append(node));
        }''', PNG)
        assert page.locator('#ordinary-preview .image-save').is_disabled()
        assert page.locator('#generation-preview .image-save').count() == 0
        assert '独立确认入口' in page.locator('#generation-preview').inner_text()
        page.evaluate('''() => {
          const card = document.createElement('article'); card.id = 'generation-plan'; document.body.append(card);
          OfficeImageGen.handleEvent({type: 'image_plan', plan: {plan_id: 'test-plan', effective_prompt: '独立确认验收', n: 1}},
            async (...args) => {imageCalls.push(args); return {ok: true};}, node => card.append(node));
        }''')
        trigger = page.locator('#generation-plan').get_by_role('button', name='确认生成（可能产生费用）', exact=True)
        trigger.click()
        dialog = page.get_by_role('dialog', name='确认联网生成图片？')
        cancel = dialog.get_by_role('button', name='取消', exact=True)
        assert cancel.evaluate('node => node === document.activeElement')
        bounds = dialog.bounding_box()
        assert bounds['x'] >= 15 and bounds['y'] >= 15
        assert abs(bounds['x'] + bounds['width'] / 2 - width / 2) < 2
        assert bounds['width'] <= min(560, width - 32) + 1
        assert dialog.evaluate('node => node.scrollWidth <= node.clientWidth')
        page.keyboard.press('Shift+Tab')
        assert dialog.get_by_role('button', name='确认生成（可能产生费用）', exact=True).evaluate('node => node === document.activeElement')
        page.keyboard.press('Tab')
        assert cancel.evaluate('node => node === document.activeElement')
        page.screenshot(path=str(tmp_path / 'image-confirm.png'))
        page.keyboard.press('Escape')
        assert not page.locator('.imagegen-confirm').count()
        assert not page.evaluate('imageCalls') and not calls
        assert trigger.evaluate('node => node === document.activeElement')
    finally:
        fixture.close()
