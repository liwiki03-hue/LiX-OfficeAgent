import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from office_agent.webui import conversations
from office_agent.webui.bridge import Api, WebviewApp
from tests.test_daily_flows import finish, setup


def test_full_capacity_delete_current_frees_slot_and_preserves_existing(tmp_path):
    app, api, _ = setup(tmp_path, True)
    first = api.conversation_snapshot()['result']['current']
    for _ in range(31):
        assert api.new_conversation()['ok']
    before = api.conversation_snapshot()['result']
    assert before['capacity']['sessions'] == 32
    assert api.new_conversation()['reason'] == 'chat_history_capacity'
    removed = api.delete_conversation(before['current'], True)
    assert removed['ok'] and len(removed['result']['sessions']) == 31
    assert removed['result']['current'] in {row['id'] for row in before['sessions'][:-1]}
    assert first in {row['id'] for row in removed['result']['sessions']}
    assert api.new_conversation()['ok'] and not app.busy


def test_new_and_renamed_sessions_restore_without_requests_or_permissions(tmp_path, monkeypatch):
    app, api, _ = setup(tmp_path, True)
    current = api.new_conversation()['result']['current']
    assert api.rename_conversation(current, '技术报告 📄')['ok']
    calls = []
    monkeypatch.setattr(api._connections, 'build_model', lambda route: SimpleNamespace(invoke=lambda messages: calls.append(messages) or SimpleNamespace(content='answer')))
    assert api.ask('first question', {'stream':False})['ok']; finish(app)
    assert api.ask('second question', {'stream':False})['ok']; finish(app)
    state = api.conversation_snapshot()['result']
    assert next(row for row in state['sessions'] if row['id'] == current)['title'] == '技术报告 📄'
    restored_app = WebviewApp()
    restored = Api(restored_app, auth=api._auth, settings_root=tmp_path)
    result = restored.conversation_snapshot()['result']
    assert result['sessions'] == state['sessions'] and result['save_state'] == 'saved'
    assert len(calls) == 2 and restored_app.path is None and not restored_app.busy
    assert restored_app.pending_id is None and restored_app._read_authorized_path is None


@pytest.mark.parametrize('operation', ['delete', 'rename', 'new'])
def test_failed_precommit_mutations_preserve_memory_and_disk(tmp_path, monkeypatch, operation):
    app, api, _ = setup(tmp_path, True)
    state = api.new_conversation()['result']
    before = Path(state['location']).read_bytes()
    monkeypatch.setattr(conversations.os, 'fsync', lambda *args: (_ for _ in ()).throw(OSError('private failure detail')))
    result = (api.delete_conversation(state['current'], True) if operation == 'delete' else
              api.rename_conversation(state['current'], 'changed') if operation == 'rename' else api.new_conversation())
    assert not result['ok'] and result['reason'] == 'chat_history_write_failed'
    after = api.conversation_snapshot()['result']
    assert after['sessions'] == state['sessions'] and after['current'] == state['current']
    assert Path(state['location']).read_bytes() == before
    assert 'private failure detail' not in json.dumps(result)
    assert after['save_state'] == 'failed' and not app.busy


@pytest.mark.parametrize('failure', [OSError, PermissionError])
def test_uncertain_commit_fuses_writes_until_explicit_reload(tmp_path, monkeypatch, failure):
    _, api, _ = setup(tmp_path, True)
    state = api.new_conversation()['result']
    original = conversations.os.replace
    writes = []
    def replace(source, target):
        writes.append(str(target))
        original(source, target)
        raise failure('exception after replacement')
    monkeypatch.setattr(conversations.os, 'replace', replace)
    result = api.rename_conversation(state['current'], '已提交但回执不确定')
    assert not result['ok'] and result['outcome_uncertain']
    assert result['result']['save_state'] == 'uncertain'
    assert api.save_conversation_history()['reason'] == 'chat_history_uncertain'
    assert len(writes) == 1
    monkeypatch.setattr(conversations.os, 'replace', original)
    restored = api.reload_conversations()['result']
    assert restored['save_state'] == 'saved'
    assert next(row for row in restored['sessions'] if row['id'] == state['current'])['title'] == '已提交但回执不确定'
    assert len(writes) == 1


def test_external_change_rejects_delete_and_keeps_memory(tmp_path):
    _, api, _ = setup(tmp_path, True)
    state = api.new_conversation()['result']
    path = Path(state['location'])
    body = json.loads(path.read_text(encoding='utf8'))
    body['sessions'][0]['title'] = 'external'
    path.write_text(json.dumps(body), encoding='utf8')
    before = path.read_bytes()
    result = api.delete_conversation(state['current'], True)
    assert not result['ok'] and result['reason'] == 'chat_history_conflict'
    assert result['result']['sessions'] == state['sessions']
    assert path.read_bytes() == before


def test_disabled_delete_does_not_persist_volatile_conversation(tmp_path):
    _, api, _ = setup(tmp_path, True)
    saved = api.new_conversation()['result']
    plan = api.plan_settings_update({'chat_history':False})
    assert api.confirm_settings_update(plan['plan']['plan_id'], True)['ok']
    volatile = api.new_conversation()['result']['current']
    assert api.rename_conversation(volatile, '仅本窗口')['ok']
    assert api.delete_conversation(saved['current'], True)['ok']
    disk = json.loads(Path(saved['location']).read_text(encoding='utf8'))
    assert volatile not in {row['id'] for row in disk['sessions']}
    assert saved['current'] not in {row['id'] for row in disk['sessions']}
    assert api.conversation_snapshot()['result']['save_state'] == 'disabled'


@pytest.mark.parametrize('title,reason', [('', 'chat_title_invalid'), ('x'*81, 'chat_title_invalid'),
                                       ('bad\nname', 'chat_title_invalid'), ('bad'+chr(0xD800), 'chat_title_invalid'),
                                       ('fixture-private-one', 'chat_sensitive_input')])
def test_rename_validation_and_credentials_never_persist(tmp_path, title, reason):
    _, api, _ = setup(tmp_path, True)
    state = api.new_conversation()['result']
    before = Path(state['location']).read_bytes()
    reply = api.rename_conversation(state['current'], title)
    assert reply == {'ok': False, 'reason': reason}
    assert Path(state['location']).read_bytes() == before


def test_manual_save_retry_does_not_call_model(tmp_path, monkeypatch):
    _, api, active = setup(tmp_path, True)
    monkeypatch.setattr(api._connections, 'build_model', lambda *args: pytest.fail('history operation called model'))
    identifier = api.new_conversation()['result']['current']
    original = conversations.os.fsync
    monkeypatch.setattr(conversations.os, 'fsync', lambda *args: (_ for _ in ()).throw(OSError('failed')))
    assert api.rename_conversation(identifier, '未落盘标题')['reason'] == 'chat_history_write_failed'
    monkeypatch.setattr(conversations.os, 'fsync', original)
    result = api.save_conversation_history()
    assert result['ok'] and result['result']['save_state'] == 'saved'
    active[0] = False
    assert api.rename_conversation(identifier, 'unauthorized')['reason'] == 'not_authenticated'
