import json
from copy import deepcopy
import os
from pathlib import Path
import shutil

from tests.test_office_canvas_browser import native_browser, idle
from tests.test_canvas_ai import Connections
from tests.test_model_connections import connection, payload, save, select


def test_current_unsaved_draft_ai_through_existing_composer_and_copy(native_browser, monkeypatch):
    page, api, source, errors = native_browser
    service = api._floating.canvas
    configured = connection('local-test')
    configured['name'] = '隔离模拟接口'
    configured['models'] = [{'id': 'simulated-model', 'capabilities': ['agent']}]
    save(api._connections, payload([configured]))
    route = select(api._connections, 'local-test', 'agent', 'simulated-model')
    page.evaluate('window.OfficeModelSelector.refresh()')
    original = source.read_bytes()
    numeric = source.suffix == '.xlsx'
    if numeric:
        page.locator('.canvas-grid button[data-target="B2"]').dblclick()
        page.locator('#canvas-value').fill('125'); page.locator('#canvas-value').press('Enter')
        page.wait_for_function("document.querySelector('.canvas-grid button[data-target=B2]').textContent==='125'")
        new = '130'
    else:
        page.get_by_role('button', name='编辑片段 run-0 实验汇报', exact=True).click()
        page.locator('#canvas-value').fill('实验汇报（复核）'); page.locator('#canvas-value').press('Enter')
        page.get_by_role('button', name='编辑片段 run-0 实验汇报（复核）', exact=True).wait_for()
        new = '实验汇报（复核修订）'
    idle(page)
    stale = {'document': deepcopy(api._floating.document), 'canvas': deepcopy(service.state())}
    connections = Connections(service, new)
    connections.route = route
    monkeypatch.setattr(api._connections, 'build_model', connections.build_model)
    page.get_by_label('引用给AI', exact=True).click()
    if numeric: page.get_by_role('button', name='允许本格数值建议', exact=True).click()
    page.locator('#canvas-ai-panel').get_by_text('引用：', exact=False).wait_for()
    page.evaluate('(value)=>window.OfficeNative.update(value)', stale)
    page.locator('#canvas-ai-panel').get_by_text('引用：', exact=False).wait_for()
    page.locator('#floating-input').fill('仅修改所选内容，保留其他位置和单位。')
    page.locator('#floating-input').press('Enter')
    try:
        page.get_by_role('button', name='采用到草稿', exact=True).wait_for(timeout=15000)
    except Exception:
        raise AssertionError({'calls': connections.calls, 'ai': service.ai.state(),
                              'feedback': page.locator('#floating-status').text_content(),
                              'panel': page.locator('#canvas-ai-panel').text_content()}) from None
    assert connections.calls == 2 and not list(source.parent.glob('*.agent-*'))
    assert service.revision == 1 and not errors
    evidence = os.environ.get('CANVAS_AI_BROWSER_EVIDENCE')
    folder = Path(evidence)/source.suffix[1:] if evidence else None
    if folder:
        folder.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(folder/'proposal.png'))
    page.get_by_role('button', name='采用到草稿', exact=True).click()
    page.wait_for_function("document.getElementById('canvas-feedback').textContent.includes('已采用到草稿')")
    assert service.revision == 2 and source.read_bytes() == original
    page.get_by_label('保存副本', exact=True).click()
    confirmation = page.get_by_role('dialog', name='确认保存副本', exact=True)
    confirmation.get_by_role('button', name='取消', exact=True).click()
    confirmation.wait_for(state='detached')
    assert not list(source.parent.glob('*.agent-*'))
    page.get_by_label('保存副本', exact=True).click()
    confirmation.get_by_role('button', name='确认生成副本', exact=True).click()
    page.wait_for_function("document.getElementById('canvas-state').textContent.includes('副本已验证')", timeout=90000)
    receipt = next(iter(api._floating.outputs.values()))
    output = Path(receipt['output'])
    assert output.is_file() and source.read_bytes() == original and connections.calls == 2
    assert receipt['instruction_revision'] == 1
    assert page.get_by_text('这是上一版要求的结果，有新要求待处理；尚未完成最新要求。', exact=True).count() == 0
    if not numeric:
        page.locator('#canvas-surface').evaluate('node=>node.scrollTop=node.scrollHeight')
        scroll_before = page.locator('#canvas-surface').evaluate('node=>({top:node.scrollTop,left:node.scrollLeft})')
        input_before = page.locator('#canvas-value').input_value()
        page.get_by_label('刷新页图', exact=True).click()
        page.locator('#canvas-slide img').wait_for()
        page.wait_for_function("document.querySelector('#canvas-slide img')?.complete && document.querySelector('#canvas-slide img')?.naturalWidth>0")
        idle(page)
        assert page.locator('#canvas-surface').evaluate('node=>({top:node.scrollTop,left:node.scrollLeft})') == scroll_before
        assert page.locator('#canvas-value').input_value() == input_before
        assert page.locator('#canvas-slide').get_attribute('data-revision') == str(service.revision)
    if folder:
        shutil.copy2(source, folder/source.name); shutil.copy2(output, folder/('编辑副本'+source.suffix))
        page.screenshot(path=str(folder/'published.png'))
    for width in (1280, 520):
        page.set_viewport_size({'width': width, 'height': 820})
        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth+1')
        page.locator('#floating-input').fill('下一条未发送要求')
        page.locator('#canvas-toggle-ai').click(); page.locator('#canvas-toggle-ai').click()
        assert page.locator('#floating-input').input_value() == '下一条未发送要求'
    api._floating.app.editor.rollback(receipt['id'], confirmed=True)
    assert not output.exists() and connections.calls == 2 and not errors
