from copy import deepcopy
import time

from docx import Document
import pytest

from office_agent.instruction_runtime import InstructionRuntime
from office_agent.paged_documents import GuardError
from tests.test_paged_editor import draft, sync
from tests.test_canvas_ai import canvas, prepare, send, finish


def authorize(floating, level='edit'):
    document = floating.document
    return floating.access.set(document['id'], document['version'], floating.epoch, level, True)['access']


def test_default_copy_mode_activates_only_for_a_user_instruction(draft):
    service, floating, source = draft
    document = floating.document
    assert floating.access.state()['level'] == 'save_copy'
    assert not floating.access.state()['active'] and not floating.access.state()['id']
    for level, confirmed in [('edit', False), ('save_copy', False), ('review', True), ('full_disk', True), (True, True), ('edit', 1)]:
        with pytest.raises(ValueError):
            floating.access.set(document['id'], document['version'], floating.epoch, level, confirmed)
    assert floating.access.state()['level'] == 'save_copy' and not floating.access.state()['active']
    with pytest.raises(ValueError):
        floating.access.require('invented', 'publish')
    grant = floating.access.for_instruction()
    assert grant['level'] == 'save_copy' and grant['origin'] == 'user_instruction'
    assert floating.access.for_instruction()['id'] == grant['id']
    assert floating.access.require(grant['id'], 'publish')['active']
    assert not list(source.parent.glob('*.agent-*.docx'))
    access = authorize(floating)
    assert floating.access.require(access['id'], 'edit')['level'] == 'edit'
    with pytest.raises(ValueError):
        floating.access.require(access['id'], 'publish')
    with pytest.raises(ValueError):
        floating.access.require(access['id'], 'shell')


def test_revoke_stop_and_file_switch_invalidate_grants(draft):
    service, floating, source = draft
    grant = authorize(floating, 'save_copy')
    floating.access.reset()
    with pytest.raises(ValueError):
        floating.access.require(grant['id'], 'publish')
    grant = authorize(floating)
    assert floating.stop()['ok']
    assert floating.access.state()['level'] == 'edit' and not floating.access.state()['active']
    floating.collect()
    assert not floating.access.state()['id']
    with pytest.raises(ValueError):
        floating.access.require(grant['id'], 'edit')
    grant = floating.access.for_instruction()
    assert grant['id'] and grant['level'] == 'edit'
    other = source.with_name('另一份材料.docx')
    book = Document(); book.add_paragraph('另一份材料'); book.save(other)
    assert floating.admit(str(other))['ok']
    assert floating.access.state()['level'] == 'save_copy' and not floating.access.state()['active']
    with pytest.raises(ValueError):
        floating.access.require(grant['id'], 'edit')


def test_grant_rejects_stale_document_epoch_external_change_and_uncertain_result(draft):
    service, floating, source = draft
    document = deepcopy(floating.document)
    for identifier, version, epoch in [('other', document['version'], floating.epoch),
                                       (document['id'], 'stale', floating.epoch),
                                       (document['id'], document['version'], floating.epoch + 1)]:
        with pytest.raises(ValueError):
            floating.access.set(identifier, version, epoch, 'edit', True)
    grant = authorize(floating)
    service.uncertain = True
    with pytest.raises(ValueError, match='待核实'):
        floating.access.require(grant['id'], 'edit')
    assert not floating.access.state()['active']
    service.uncertain = False
    grant = authorize(floating)
    source.write_bytes(source.read_bytes() + b'external-change')
    with pytest.raises(ValueError, match='外部改变'):
        floating.access.require(grant['id'], 'edit')
    assert not floating.access.state()['active']


def test_authorized_copy_publish_keeps_original_and_rejects_duplicate_and_revoked_grant(draft, monkeypatch):
    service, floating, source = draft
    original = source.read_bytes()
    sync(service)
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **kw: {'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = service.plan(service.document_id, 1)['plan']
    grant = authorize(floating)
    with pytest.raises(ValueError):
        service.publish(plan['id'], service.document_id, 1, False, grant['id'])
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))
    grant = authorize(floating, 'save_copy')
    assert service.publish(plan['id'], service.document_id, 1, False, grant['id'])['publishing']
    end = time.monotonic() + 10
    while service.publishing and time.monotonic() < end:
        time.sleep(.01)
    assert not service.publishing and not service.uncertain
    assert source.read_bytes() == original
    output = next(iter(floating.outputs.values()))
    assert output['authorization']['id'] == grant['id']
    assert Document(output['output']).paragraphs[1].text == '实验于2026年完成，温度120 ℃。'
    assert Document(output['backup']).paragraphs[1].text == '实验于2025年完成，温度120 ℃。'
    second = service.plan(service.document_id, 1)['plan']
    with pytest.raises(GuardError, match='重复自动保存'):
        service.publish(second['id'], service.document_id, 1, False, grant['id'])
    assert len(list(source.parent.glob('*.agent-*.docx'))) == 1
    floating.access.reset()
    with pytest.raises(ValueError, match='授权已失效'):
        service.publish(second['id'], service.document_id, 1, False, grant['id'])


def test_new_requirements_block_automatic_publication_of_an_older_draft(draft):
    service, floating, source = draft
    sync(service)
    floating.runtime = InstructionRuntime(floating.document['version'])
    floating.runtime.submit('new', '补充说明当前实验条件')
    plan = service.plan(service.document_id, 1)['plan']
    assert plan['pending_requirement_count'] == 1
    grant = authorize(floating, 'save_copy')
    with pytest.raises(GuardError, match='要求待处理'):
        service.publish(plan['id'], service.document_id, 1, False, grant['id'])
    assert not list(source.parent.glob('*.agent-*.docx'))


def test_canvas_authorized_adoption_and_copy_keep_source_protection(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    original = source.read_bytes()
    grant = authorize(floating)
    assert send(service, reference)['ok']
    finish(service.ai)
    proposal = next(row for row in service.ai.proposals.values() if row['status'] == 'pending')
    with pytest.raises(ValueError):
        service.ai.adopt(service.document_id, service.revision, proposal['id'], 'invented')
    assert service.ai.adopt(service.document_id, service.revision, proposal['id'], grant['id'])['ok']
    assert service.summaries[-1]['ai']['authorization']['id'] == grant['id']
    assert source.read_bytes() == original
    monkeypatch.setattr('office_agent.webui.office_canvas.compare_documents', lambda *a, **kw: {'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = service.plan(service.document_id, service.revision)['plan']
    with pytest.raises(ValueError):
        service.publish(service.document_id, service.revision, plan['id'], False, grant['id'])
    publish_grant = authorize(floating, 'save_copy')
    assert service.publish(service.document_id, service.revision, plan['id'], False, publish_grant['id'])['publishing']
    end = time.monotonic() + 10
    while service.publishing and time.monotonic() < end:
        time.sleep(.01)
    assert not service.publishing and not service.uncertain
    output = next(iter(floating.outputs.values()))
    assert output['authorization']['id'] == publish_grant['id']
    assert source.read_bytes() == original


def test_canvas_user_request_activates_default_copy_mode(canvas, monkeypatch):
    service, floating, source, reference, connections = prepare(canvas, monkeypatch)
    original = source.read_bytes()
    assert not floating.access.state()['active']
    result = send(service, reference)
    assert result['ok'] and result['started']
    grant = result['ai']['access']
    assert grant['level'] == 'save_copy' and grant['origin'] == 'user_instruction'
    finish(service.ai)
    assert floating.access.require(grant['id'], 'publish')['active']
    assert source.read_bytes() == original and not floating.outputs
    floating.stop()
    floating.collect()
    assert not floating.access.state()['active']
    with pytest.raises(ValueError, match='授权已失效'):
        floating.access.require(grant['id'], 'edit')
