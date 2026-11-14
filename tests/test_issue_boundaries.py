import base64
import io
import json
from pathlib import Path
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.diagnostics import Diagnostics
from office_agent.model_content import check_scope
from office_agent.paged_documents import GuardError
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from tests.test_paged_editor import draft, sync


def test_recovery_in_new_host_has_no_old_authorization_or_task(draft):
    service, floating, source = draft
    sync(service); grant = floating.access.for_instruction()
    old_inbox = service.inbox_id
    new_app = WebviewApp()
    api = SimpleNamespace(_app=new_app, _auth=floating.api._auth, _memory_key=floating.api._memory_key, _connections=floating.api._connections)
    reopened = FloatingSession(api)
    assert reopened.admit(str(source))['ok'] and reopened.read(reopened.read_token, True)['ok']
    restored = reopened.paged; restored.open(reopened.document['id'], reopened.document['version'])
    assert restored.state()['recovery_available']
    assert not reopened.access.state()['active'] and reopened.runtime is None
    with pytest.raises(ValueError): reopened.access.require(grant['id'], 'edit')
    result = restored.restore(restored.document_id)
    assert result['dirty'] and restored.inbox_id != old_inbox
    assert not restored.proposals and not restored.references and not restored.publisher.plans
    assert '2026' in Document(io.BytesIO(base64.b64decode(result['data']))).paragraphs[1].text


@pytest.mark.parametrize('target', ['output', 'backup'])
def test_changed_copy_or_backup_refuses_rollback_and_preserves_source(draft, monkeypatch, target):
    service, floating, source = draft; original = source.read_bytes(); sync(service)
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **kw: {'inputs_unchanged': True, 'render': {'status': 'different'}})
    plan = service.plan(service.document_id, 1)['plan']
    service.publisher.execute(plan['id'], document_id=service.document_id, draft_revision=1, confirmed=True)
    audit = floating.app.editor.audit / plan['id'] / 'audit.json'
    record = json.loads(audit.read_text('utf8')); output = Path(record['output']); changed = Path(record[target])
    changed.write_bytes(changed.read_bytes() + b'external-change')
    with pytest.raises(GuardError, match='changed'):
        floating.app.editor.rollback(plan['id'], confirmed=True)
    assert output.exists() and source.read_bytes() == original
    assert json.loads(audit.read_text('utf8'))['status'] == 'verified'


def test_model_scope_and_diagnostic_export_are_separate_privacy_boundaries(draft, tmp_path):
    from office_agent.memory import classify_candidate
    service, _, _ = draft; sync(service)
    reference = service.reference(service.document_id, 1, {'node_id': 'node-1', 'start': 0, 'end': 2})['reference']
    assert reference['context'] == reference['old'] == '实验'
    for content in ['api_key=synthetic-private', '身份证号：110101199001011234', '手机：13800138000']:
        with pytest.raises(GuardError): check_scope(content)
        assert not classify_candidate(content)['allowed']
    logger = Diagnostics(tmp_path)
    logger.record(module='model', action='request', result='refused', code='private_scope', content='synthetic-private', request=reference)
    assert logger.close()
    exported = json.dumps(logger.preview(), ensure_ascii=False)
    assert 'synthetic-private' not in exported and '实验' not in exported and 'api_key' not in exported
