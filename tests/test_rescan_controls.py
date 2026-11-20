import pytest

from tests.test_floating_browser import floating_browser
from tests.test_model_connections import connection, payload, save, select


@pytest.mark.parametrize('model_state', ['unconfigured', 'invalidated'])
@pytest.mark.parametrize('at_capacity', [False, True], ids=['below-limit', 'at-limit'])
def test_floating_controls_without_available_model(floating_browser, monkeypatch, model_state, at_capacity):
    page, api, path, calls, errors = floating_browser
    original = path.read_bytes()
    if model_state == 'invalidated':
        select(api._connections, 'official', 'agent')
        save(api._connections, payload([{**connection(), 'enabled': False}]))
        assert api._connections.snapshot()['catalog']['current']['agent']['reason'] == 'model_selection_invalid'
    else:
        save(api._connections, payload([]))
    page.evaluate('OfficeModelSelector.refresh()')
    page.locator('#floating-choose').click()
    page.get_by_role('button', name='同意读取并打开', exact=True).click()
    page.wait_for_function("!document.getElementById('floating-send').disabled")
    assert not page.evaluate('OfficeModelSelector.selection().valid || OfficePaged?.active() || OfficeNative?.active()')
    model_builds = []
    monkeypatch.setattr(api._connections, 'build_model', lambda *args: model_builds.append(args))

    prompt = page.locator('#floating-input')
    send = page.locator('#floating-send')
    prompt.fill('请解释实验条件。')
    send.click()
    page.locator('#floating-status').filter(has_text='请先在输入区选择有效的文档 Agent 模型').wait_for()
    assert prompt.input_value() == '请解释实验条件。'
    assert calls.count('conversation_begin') == 1
    conversation = api._floating.conversation
    assert not conversation.turns and not page.locator('.conversation-turn').count()

    if at_capacity:
        document = api._floating.document
        seed = api.conversation_begin('old-progress', '进度如何？', document['id'], document['version'])['turn']
        with conversation.lock:
            conversation.turns.update({f'old-{i}': dict(seed, id=f'old-{i}', question=f'已处理问题 {i}') for i in range(79)})
        page.evaluate('OfficeFloating.refresh()')
        page.get_by_role('button', name='开始新对话', exact=True).wait_for()

    prompt.fill('进度如何？')
    send.click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    progress = conversation.control_turn if at_capacity else list(conversation.turns.values())[-1]
    assert progress['intent'] == 'status' and progress['status'] == 'complete'
    assert progress['answer'] == '当前没有正在执行的文档修改任务。'
    page.locator(f'[data-turn-id="{progress["id"]}"] [role="status"]').filter(has_text=progress['answer']).wait_for()
    assert '选择有效' not in page.locator('#floating-status').inner_text()

    prompt.fill('停止')
    send.click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    stop = conversation.control_turn if at_capacity else list(conversation.turns.values())[-1]
    assert stop['intent'] == 'stop' and stop['stage'] == 'stopped'
    assert api._floating.stage == 'stopped' and '已停止后续派发' in stop['text']
    page.locator(f'[data-turn-id="{stop["id"]}"] [role="status"]').filter(has_text=stop['text']).wait_for()
    assert calls.count('conversation_begin') == 3 and not model_builds
    assert len(conversation.turns) == (80 if at_capacity else 2)
    assert not {'floating_submit', 'paged_submit', 'canvas_submit'}.intersection(calls)
    assert path.read_bytes() == original and not api._floating.outputs and not errors
