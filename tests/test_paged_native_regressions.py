import time

from docx import Document

from tests.test_paged_browser import engine_frame, paged_browser
from tests.test_paged_editor import draft, sync


def test_recovery_status_does_not_claim_output_is_absent_after_publish(draft, monkeypatch):
    service, floating, source = draft
    sync(service)
    assert service.state()['recovery_state'] == 'saved'
    plan = service.plan(service.document_id, 1)['plan']
    frozen = service.publisher.claim(plan['id'], document_id=service.document_id, draft_revision=1, confirmed=True)
    monkeypatch.setattr(service.publisher, 'execute', lambda *a, **kw: {'id': plan['id'], 'status': 'verified', 'output': 'controlled-test-copy', 'output_sha256': 'controlled-test'})
    service._publish(frozen, service.publisher)
    state = service.state()
    assert state['published_revision'] == state['draft_revision'] == 1
    assert not state['dirty'] and state['recovery_state'] == 'saved'
    assert '尚未生成' not in state['recovery_status']
    assert '批准' in state['recovery_status']


def test_recovery_failure_is_explicit_and_keeps_current_memory_snapshot(draft, monkeypatch):
    service, floating, source = draft
    original = source.read_bytes()

    def fail():
        raise OSError('controlled own-fixture write failure')

    monkeypatch.setattr(service, '_persist_recovery', fail)
    state = sync(service)
    assert state['recovery_state'] == 'failed'
    assert '未保存' in state['recovery_status']
    assert state['snapshot_revision'] == state['draft_revision'] == 1 and state['dirty']
    assert source.read_bytes() == original and service.data


def test_native_findings_confirmation_layout_cancel_and_published_feedback(paged_browser):
    page, api, source, errors = paged_browser
    original = source.read_bytes()
    frame = engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。', exact=True).first.click()
    page.keyboard.press('End')
    page.keyboard.insert_text(' 原生界面验收。' + '需要保留完整审阅说明。' * 80)
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('未保存')")
    for width, height in [(1280, 820), (520, 800), (360, 440)]:
        page.set_viewport_size({'width': width, 'height': height})
        for theme in ['light', 'dark', 'highcontrast']:
            page.evaluate('(theme)=>document.documentElement.dataset.theme=theme', theme)
            sidebar_closed = page.locator('#floating-root').evaluate("node=>node.classList.contains('paged-ai-closed')")
            page.get_by_role('button', name='保存副本', exact=True).click()
            dialog = page.get_by_role('dialog', name='生成验证副本', exact=True)
            dialog.wait_for()
            bounds = dialog.bounding_box()
            assert bounds and abs(bounds['x'] + bounds['width'] / 2 - width / 2) <= 2
            assert abs(bounds['y'] + bounds['height'] / 2 - height / 2) <= 2
            assert bounds['height'] <= height - 30
            cancel = dialog.get_by_role('button', name='取消', exact=True)
            cancel.scroll_into_view_if_needed()
            control = cancel.bounding_box()
            assert control and control['y'] >= 0 and control['y'] + control['height'] <= height
            assert cancel.evaluate("node=>parseFloat(getComputedStyle(node).paddingLeft)>=10 && parseFloat(getComputedStyle(node).borderRadius)>=6")
            assert dialog.locator('.paged-change-list').evaluate('node=>node.scrollHeight>node.clientHeight')
            cancel.focus()
            cancel.press('Tab')
            assert page.evaluate("document.activeElement.textContent") == '确认生成副本'
            page.keyboard.press('Escape')
            assert page.get_by_role('dialog').count() == 0
            page.wait_for_function("document.activeElement.id==='paged-save'")
            assert page.locator('#floating-root').evaluate("node=>node.classList.contains('paged-ai-closed')") == sidebar_closed
            assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))
    page.set_viewport_size({'width': 1280, 'height': 820})
    page.get_by_role('button', name='保存副本', exact=True).click()
    page.get_by_role('dialog', name='生成验证副本').get_by_role('button', name='确认生成副本', exact=True).click()
    deadline = time.monotonic() + 180
    while not api._floating.outputs:
        if time.monotonic() > deadline:
            raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(100)
    page.wait_for_function("document.getElementById('paged-save-state').textContent.startsWith('副本已验证')")
    page.wait_for_function("document.getElementById('paged-feedback').textContent.includes('副本已通过验证')")
    assert '尚未生成' not in page.locator('#paged-feedback').text_content()
    assert '原稿' not in page.locator('#paged-save-state').text_content()
    assert page.get_by_role('button', name='保存副本', exact=True).is_disabled()
    page.evaluate("document.getElementById('paged-feedback').textContent='新的操作失败；请保留当前草稿'")
    page.wait_for_timeout(1000)
    assert page.locator('#paged-feedback').text_content() == '新的操作失败；请保留当前草稿'
    result = next(iter(api._floating.outputs.values()))
    assert '原生界面验收。' in Document(result['output']).paragraphs[1].text
    assert source.read_bytes() == original and errors == []
    api._floating.outputs[result['id']]['undone'] = True
    api._floating.paged.output_rolled_back('unrelated-output')
    assert not api._floating.paged.state()['dirty']
    api._floating.paged.output_rolled_back(result['id'])
    page.wait_for_function("document.getElementById('paged-save-state').textContent.startsWith('副本已撤销')")
    assert api._floating.paged.state()['dirty']
    assert page.get_by_role('button', name='保存副本', exact=True).is_enabled()
    assert page.locator('#paged-feedback').text_content() == '新的操作失败；请保留当前草稿'
