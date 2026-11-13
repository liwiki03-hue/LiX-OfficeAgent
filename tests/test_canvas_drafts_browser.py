import json
import os
from pathlib import Path
import shutil
import time

from openpyxl import load_workbook
from pptx import Presentation

from tests.test_office_canvas_browser import native_browser, idle


def test_checkpoint_reopen_restore_copy_and_independent_rollback(native_browser, monkeypatch):
    try:
        checkpoint_workflow(native_browser, monkeypatch)
    except BaseException:
        page, api, source, errors = native_browser
        folder = Path(os.environ.get('CANVAS_DRAFT_EVIDENCE', str(source.parent))) / ('failure-' + source.suffix[1:])
        folder.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(folder / 'screen.png'))
        observed = page.evaluate("({dialogs:[...document.querySelectorAll('dialog')].map(n=>({name:n.getAttribute('aria-label'),open:n.open,text:n.innerText})),status:document.getElementById('canvas-state')?.textContent,feedback:document.getElementById('canvas-feedback')?.textContent})")
        (folder / 'STATE.json').write_text(json.dumps({'ui': observed, 'revision': api._floating.canvas.revision, 'errors': errors}, ensure_ascii=False, indent=2), encoding='utf8')
        raise


def checkpoint_workflow(native_browser, monkeypatch):
    page, api, source, errors = native_browser
    original = source.read_bytes()
    service = api._floating.canvas
    accepted = []
    confirm_draft = api.canvas_draft_confirm
    def counted(*args):
        if args[-1] is True: accepted.append(args[2])
        return confirm_draft(*args)
    monkeypatch.setattr(api, 'canvas_draft_confirm', counted)
    if source.suffix == '.xlsx':
        page.locator('.canvas-grid button[data-target="B2"]').dblclick()
        value = '128'
    else:
        page.get_by_role('button', name='编辑片段 run-0 实验汇报', exact=True).click()
        value = '实验汇报（恢复复核）'
    page.locator('#canvas-value').fill(value)
    page.locator('#canvas-value').press('Enter')
    page.wait_for_function("window.pywebview && document.getElementById('canvas-state').textContent.includes('未保存')")
    idle(page)
    page.locator('#canvas-draft-save').locator('..').locator('..').locator('summary').click()
    page.get_by_role('button', name='保存本机草稿', exact=True).click()
    modal = page.get_by_role('dialog', name='保存本机草稿', exact=True)
    modal.get_by_role('button', name='取消', exact=True).click()
    modal.wait_for(state='detached')
    assert not service.drafts.target.exists()
    page.get_by_role('button', name='保存本机草稿', exact=True).click()
    modal.get_by_role('button', name='确认保存本机草稿', exact=True).evaluate('node=>{node.click();node.click();}')
    modal.wait_for(state='detached')
    page.wait_for_function("document.getElementById('canvas-draft-status').textContent.includes('已保存')")
    assert service.drafts.target.is_file() and source.read_bytes() == original
    assert len(accepted) == 1
    assert not list(source.parent.glob('*.agent-*'))
    page.locator('#floating-choose').click()
    page.get_by_role('dialog', name='处理未保存编辑').get_by_role('button', name='放弃并继续', exact=True).click()
    page.get_by_role('button', name='同意读取并打开', exact=True).click()
    page.get_by_role('button', name='恢复本机草稿', exact=True).wait_for()
    if source.suffix == '.pptx': page.locator('#canvas-slide img').wait_for()
    idle(page)
    assert service.revision == 0 and service.data == original
    assert not service.ai.proposals and service.ai.runtime is None
    page.get_by_role('button', name='恢复本机草稿', exact=True).click()
    restore = page.get_by_role('dialog', name='恢复本机草稿', exact=True)
    restore.get_by_role('button', name='取消', exact=True).click()
    restore.wait_for(state='detached')
    assert service.revision == 0
    page.get_by_role('button', name='恢复本机草稿', exact=True).click()
    restore.get_by_role('button', name='确认恢复内容', exact=True).click()
    restore.wait_for(state='detached')
    page.wait_for_function("document.getElementById('canvas-draft-status').textContent.includes('已恢复')")
    idle(page)
    assert service.revision == 1 and len(service.operations) == 1 and not service.plans
    assert len(accepted) == 2
    assert source.read_bytes() == original
    if source.suffix == '.xlsx':
        assert page.locator('.canvas-grid button[data-target="B2"]').text_content() == value
    else:
        page.get_by_role('button', name='编辑片段 run-0 '+value, exact=True).wait_for()
    for width in (1280, 520):
        for theme in ('light', 'dark', 'highcontrast'):
            page.set_viewport_size({'width': width, 'height': 820})
            page.evaluate('(theme)=>document.documentElement.dataset.theme=theme', theme)
            details = page.locator('#canvas-draft-panel').locator('..')
            if not details.evaluate('node=>node.open'): details.locator('summary').click()
            box = page.locator('#canvas-draft-panel').bounding_box()
            assert box and box['x'] >= 0 and box['x']+box['width'] <= width
            assert page.get_by_label('保存副本', exact=True).is_visible()
            details.locator('summary').click()
    page.set_viewport_size({'width': 1280, 'height': 820})
    evidence = os.environ.get('CANVAS_DRAFT_EVIDENCE')
    folder = Path(evidence)/source.suffix[1:] if evidence else None
    if folder:
        folder.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, folder/source.name)
        shutil.copy2(service.drafts.target, folder/'saved-draft.json')
        page.screenshot(path=str(folder/'restored.png'))
    page.get_by_label('保存副本', exact=True).click()
    confirm = page.get_by_role('dialog', name='确认保存副本', exact=True)
    confirm.get_by_role('button', name='取消', exact=True).click()
    confirm.wait_for(state='detached')
    assert not list(source.parent.glob('*.agent-*'))
    page.get_by_label('保存副本', exact=True).click()
    confirm.get_by_role('button', name='确认生成副本', exact=True).click()
    confirm.wait_for(state='detached')
    deadline = time.monotonic()+90
    while service.publishing:
        assert time.monotonic() < deadline
        page.wait_for_timeout(100)
    assert not service.uncertain and not service.last_error
    page.wait_for_function("document.getElementById('canvas-state').textContent==='副本已验证'")
    page.wait_for_function("document.getElementById('canvas-feedback').textContent.startsWith('副本已生成：')")
    receipt = next(iter(api._floating.outputs.values()))
    output = Path(receipt['output'])
    assert receipt['render_diff']['inputs_unchanged'] and Path(receipt['backup']).read_bytes() == original
    if source.suffix == '.xlsx':
        book = load_workbook(output)
        assert book['实验']['B2'].value == 128 and book['实验']['B2'].font.bold and book['实验']['C2'].value == '=B2*2'
    else:
        book = Presentation(output)
        assert book.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].text == value
    if folder:
        shutil.copy2(output, folder/('恢复编辑副本'+source.suffix))
        shutil.copy2(Path(receipt['backup']), folder/('BACKUP'+source.suffix))
        (folder/'PUBLISH.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2), encoding='utf8')
        page.screenshot(path=str(folder/'published.png'))
    rollback = api._app.editor.rollback(receipt['id'], confirmed=True)
    assert rollback['status'] == 'rolled_back' and not output.exists() and source.read_bytes() == original
    assert not errors
    if folder: (folder/'ROLLBACK.json').write_text(json.dumps(rollback, ensure_ascii=False, indent=2), encoding='utf8')
