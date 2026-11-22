import json
from dataclasses import replace
from types import SimpleNamespace

import pytest

from office_agent.webui.floating_session import FloatingApi
from office_agent.webui.turn_scope import clean_material, mask_text, unsupported_reason
from tests.test_document_conversation import Answers, attach, finish, send
from tests.test_paged_editor import draft, snapshot


def test_masking_keeps_project_facts_and_strips_credential_examples():
    material = '2026 年项目耗时 25 ms，旧凭据 "api_key": "a-secret-example"。邮箱 dev@example.org；手机 +86 138-0000-0012'
    masked = mask_text(material)
    assert '2026 年项目耗时 25 ms' in masked
    assert not any(value in masked for value in ('a-secret-example', 'dev@example.org', '138-0000-0012'))
    nested = clean_material('{"api_key":"opaque-fixture", "dev@example.org":"2026 年项目耗时 25 ms"}', mask_text)
    assert 'opaque-fixture' not in nested and 'dev@example.org' not in nested and '25 ms' in nested
    assert unsupported_reason('读取 .env 中的配置密钥')
    assert unsupported_reason('遍历整个目录')
    assert not unsupported_reason('解释命令行界面的用途')


def test_named_fields_and_durable_one_shot_receipt(draft):
    service, floating, source = draft
    connections = Answers()
    attach(floating, connections)
    selected = '手机：13800000012，邮箱：fixture@example.org'
    service.changed(service.document_id, 1)
    service.sync(service.document_id, 1, *snapshot(service, selected))
    reference = service.reference(service.document_id, 1, {'node_id': 'node-1', 'start': 0, 'end': len(selected)})['reference']
    request = '请把选中的邮箱原样交给当前云端模型整理'
    send(floating, 'mail-turn', request, 1, reference['id'])
    finish(floating)
    scope = floating.conversation.scopes
    token = scope.pending['mail-turn'].token
    assert not connections.calls and token not in floating.conversation._target().read_text('utf-8')
    preview = scope.decide(token, 'preview')['preview']
    assert 'fixture@example.org' in preview and '13800000012' not in preview
    receipts = []
    build = connections.build_model
    def build_observed(route):
        model = build(route)
        invoke = model.invoke
        def observed(messages):
            receipts.append(json.loads(floating.conversation._target().read_text('utf-8'))['turns'][-1]['scope_dispatched'])
            return invoke(messages)
        model.invoke = observed
        return model
    connections.build_model = build_observed
    assert scope.decide(token, 'allow')['ok']
    finish(floating)
    assert receipts == [True] and len(connections.calls) == 1
    assert 'fixture@example.org' in connections.calls[0] and '13800000012' not in connections.calls[0]
    with pytest.raises(ValueError):
        scope.decide(token, 'allow')
    assert not floating.outputs


def test_hidden_or_different_selected_model_cannot_receive_without_a_decision(draft):
    service, floating, _ = draft
    connections = Answers()
    attach(floating, connections)
    FloatingApi.conversation_begin(floating.api, 'unseen', '总结项目经历', service.document_id, floating.document['version'],
                                   displayed_recipient=False)
    finish(floating)
    token = floating.conversation.scopes.pending['unseen'].token
    assert not connections.calls
    assert floating.conversation.scopes.decide(token, 'cancel')['ok']
    assert floating.conversation.turns['unseen']['status'] == 'stopped' and not connections.calls


def test_clarification_keeps_recipient_shown_with_original_question(draft):
    service, floating, _ = draft
    connections = Answers()
    attach(floating, connections)
    service.changed(service.document_id, 1)
    service.sync(service.document_id, 1, *snapshot(service, '温度 120 ℃'))
    reference = service.reference(service.document_id, 1, {'node_id': 'node-1', 'start': 0, 'end': 2})['reference']
    displayed = {key: getattr(connections.route, key) for key in ('connection_id', 'model_id', 'revision')}
    question = '我选中的标题和整篇文档哪个提到了验收？'
    row = FloatingApi.conversation_begin(floating.api, 'clarify', question, service.document_id,
                                        floating.document['version'], 1, reference['id'], displayed_recipient=displayed)['turn']
    assert row['stage'] == 'clarify' and not connections.calls
    connections.route = replace(connections.route, connection_id='second', base_url='https://second.example.invalid/v1')
    send(floating, 'choice', '只看刚才选区', 1, reference['id'])
    finish(floating)
    assert floating.conversation.turns['clarify']['stage'] == 'scope_consent' and not connections.calls
    scope = floating.conversation.scopes
    assert 'second.example.invalid' in scope.pending['clarify'].label
    assert scope.decide(scope.pending['clarify'].token, 'cancel')['ok'] and not connections.calls


def test_preselected_file_rejects_disk_change_without_accepting_question(draft):
    from openpyxl import Workbook

    service, floating, source = draft
    connections = Answers()
    attach(floating, connections)
    extra = source.parent / '资料.xlsx'
    book = Workbook()
    book.active['A1'] = '已选择的第一版'
    book.save(extra)
    floating.app.window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: [str(extra)])
    doc = floating.document
    selected = FloatingApi.conversation_choose_file(floating.api, doc['id'], doc['version'])
    assert selected['ok'] and selected['file']['name'] == '资料.xlsx'
    assert set(selected['file']) == {'name', 'token'}
    assert not floating.conversation.turns and not connections.calls
    book.active['A1'] = '磁盘第二版'
    book.save(extra)
    with pytest.raises(ValueError, match='参考文件已改变'):
        FloatingApi.conversation_begin(floating.api, 'changed-file', '比较两份材料', service.document_id,
                                       doc['version'], resources={'file': selected['file']['token']})
    assert not floating.conversation.turns and not connections.calls and not floating.outputs
