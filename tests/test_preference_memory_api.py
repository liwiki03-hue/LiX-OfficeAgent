"""Mounted Web API opt-in and repeated human turn gate without providers."""
from pathlib import Path

from office_agent.webui.bridge import Api

FORMAT = '请把文档格式改为Markdown'


class Auth:
    def __init__(self, root):
        self.root, self.user = root, True
    def status(self):
        return {'authenticated': self.user}
    def account_key(self):
        return 'a' * 64 if self.user else None
    def logout(self):
        self.user = False
        return {'ok': True}


class App:
    busy = False
    def __init__(self, path):
        self.path, self.events = path, []
    def ask(self, text):
        return {'ok': True}
    def emit(self, *args):
        self.events.append(args)
    def _clear_private_state(self):
        self.path = None


def test_api_memory_optin_recurrence_and_logout(tmp_path):
    path = tmp_path / 'selected.docx'
    path.write_bytes(b'temporary test source')
    app = App(path)
    api = Api(app, auth=Auth(tmp_path))
    assert api.memory_snapshot()['result']['enabled'] is False
    api._dispatch('ask', FORMAT)
    api._dispatch('ask', FORMAT)
    assert api.memory_snapshot()['result']['suggestion'] is None
    proposal = api.plan_memory_change('enable', True)['plan']['plan_id']
    assert api.confirm_memory_change(proposal, False)['reason'] == 'memory_confirmation_required'
    assert api.memory_snapshot()['result']['enabled'] is False
    proposal = api.plan_memory_change('enable', True)['plan']['plan_id']
    assert api.confirm_memory_change(proposal, True)['result']['saved'] is True
    api._dispatch('ask', FORMAT)
    assert api.memory_snapshot()['result']['suggestion'] is None
    api._dispatch('ask', FORMAT)
    assert api.memory_snapshot()['result']['suggestion']['repeated'] is True
    proposal = api.plan_memory_change('save', FORMAT)['plan']['plan_id']
    assert api.confirm_memory_change(proposal, True)['result']['saved'] is True
    api._dispatch('ask', '请把文档格式改为表格')
    assert len(app._memory_context) == 1 and app._memory_context[0]['text'] == FORMAT
    assert api.logout()['ok'] is True
    assert api.memory_snapshot()['reason'] == 'not_authenticated'
    assert app._memory_context is None


def test_manual_memory_bridge_requires_auth_confirmation_and_does_not_observe_tool_events(tmp_path):
    app = App(tmp_path / 'selected.docx')
    api = Api(app, auth=Auth(tmp_path))
    proposed = api.plan_memory_change('enable', True)['plan']['plan_id']
    assert api.confirm_memory_change(proposed, True)['ok']
    proposed = api.plan_memory_change('set_manual', {'kind': 'persona', 'text': '用简洁中文回答'})['plan']['plan_id']
    assert api.memory_snapshot()['result']['items'] == []
    assert api.confirm_memory_change(proposed, True)['ok']
    assert api.memory_snapshot()['result']['items'][0]['source'] == 'manual'
    app.emit('text', '以后处理报告先列提纲再检查标题')
    assert api.memory_snapshot()['result']['suggestion'] is None
    for _ in range(2):
        api._dispatch('ask', '以后处理报告先列提纲再检查标题')
    assert api.memory_snapshot()['result']['suggestion']['kind'] == 'skill'
    api._app.busy = True
    assert api.plan_memory_change('set_manual', {'kind': 'work', 'text': '负责图表样式'})['reason'] == 'busy'
    api._app.busy = False
    assert api.logout()['ok']
    assert api.memory_snapshot()['reason'] == 'not_authenticated'
