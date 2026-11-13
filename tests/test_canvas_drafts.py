from copy import deepcopy
import json
import os
from pathlib import Path
import sys
import time

import pytest

from office_agent import office_canvas as engine
from office_agent.writes import GuardError
from office_agent.fidelity import _run
from tests.test_office_canvas import canvas, operation


def save(service):
    plan = service.draft_plan(service.document_id, service.revision, 'save')['plan']
    return service.draft_confirm(service.document_id, service.revision, plan['id'], True)


def reopen(service, floating, source):
    service.close()
    assert floating.admit(str(source))['ok']
    assert floating.read(floating.read_token, True)['ok']
    return service.open(floating.document['id'], floating.document['version'])


def test_save_cancel_restart_restore_replay_and_no_old_approvals(canvas):
    service, floating, source, _, op = canvas
    original = source.read_bytes()
    service.apply(service.document_id, 0, op)
    plan = service.draft_plan(service.document_id, 1, 'save')['plan']
    assert not service.drafts.target.exists()
    assert service.draft_confirm(service.document_id, 1, plan['id'], False)['cancelled']
    assert not service.drafts.target.exists()
    saved = save(service)
    assert saved['checkpoint_revision'] == 1 and source.read_bytes() == original
    expected = engine.load(service.data, source.suffix)
    encoded = service.drafts.target.read_text(encoding='utf8')
    assert all(word not in encoded for word in ['api_key', 'approval', 'plan_id', 'model_id', 'base64'])
    opened = reopen(service, floating, source)
    assert opened['recovery_available'] and not opened['dirty'] and service.data == original
    with pytest.raises(GuardError, match='恢复'): service.apply(service.document_id, 0, op)
    plan = service.draft_plan(service.document_id, 0, 'restore')['plan']
    assert service.draft_confirm(service.document_id, 0, plan['id'], False)['cancelled']
    assert service.data == original
    plan = service.draft_plan(service.document_id, 0, 'restore')['plan']
    restored = service.draft_confirm(service.document_id, 0, plan['id'], True)
    assert restored['dirty'] and restored['recovery_state'] == 'restored'
    assert engine.load(service.data, source.suffix) == expected
    assert not service.plans and service.checkpoint is None and service.ai.reference is None
    assert service.ai.runtime is None and not service.ai.proposals and floating.runtime is None
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*'))
    with pytest.raises(GuardError): service.draft_confirm(service.document_id, 0, plan['id'], True)
    assert service.plan(service.document_id, service.revision)['plan']['confirmation_required']


def test_old_cancel_cannot_clear_new_plan_and_mutation_invalidates(canvas):
    service, _, source, _, op = canvas
    service.apply(service.document_id, 0, op)
    first = service.draft_plan(service.document_id, 1, 'save')['plan']
    second = service.draft_plan(service.document_id, 1, 'save')['plan']
    with pytest.raises(GuardError): service.draft_confirm(service.document_id, 1, first['id'], False)
    assert service.checkpoint['id'] == second['id']
    service.undo(service.document_id, 1)
    with pytest.raises(GuardError): service.draft_confirm(service.document_id, 2, second['id'], True)
    assert not service.drafts.target.exists()


def test_expiry_external_change_and_write_failure_keep_memory(canvas, monkeypatch):
    service, _, source, _, op = canvas
    service.apply(service.document_id, 0, op)
    plan = service.draft_plan(service.document_id, 1, 'save')['plan']
    service.checkpoint['deadline'] = time.monotonic()-1
    with pytest.raises(GuardError): service.draft_confirm(service.document_id, 1, plan['id'], True)
    before = service.data
    def denied(*args): raise PermissionError('owned fixture write failure')
    monkeypatch.setattr(service.drafts, 'write', denied)
    with pytest.raises(GuardError, match='未确认'): save(service)
    assert service.data == before and service.revision == 1 and service.state()['recovery_state'] == 'failed'
    plan = service.draft_plan(service.document_id, 1, 'save')['plan']
    source.write_bytes(source.read_bytes()+b'changed')
    with pytest.raises(ValueError): service.draft_confirm(service.document_id, 1, plan['id'], True)
    assert service.data == before and not service.drafts.target.exists()


@pytest.mark.parametrize('damage', ['schema', 'source', 'operation', 'hash', 'extra', 'oversize', 'hardlink'])
def test_damaged_or_untrusted_record_never_restores(canvas, damage):
    service, floating, source, _, op = canvas
    service.apply(service.document_id, 0, op); save(service)
    path = service.drafts.target
    record = json.loads(path.read_text(encoding='utf8'))
    if damage == 'schema': record['schema_version'] = 2
    if damage == 'source': record['source_version'] = '0'*64
    if damage == 'operation': record['operations'][0]['target'] = '../../outside'
    if damage == 'hash': record['candidate_package'] = '0'*64
    if damage == 'extra': record['approval'] = True
    if damage == 'oversize': path.write_bytes(b' '* (512*1024+1))
    elif damage == 'hardlink': os.link(path, path.with_name('linked.json'))
    else: path.write_text(json.dumps(record), encoding='utf8')
    opened = reopen(service, floating, source)
    original = source.read_bytes()
    if opened['recovery_available']:
        with pytest.raises(ValueError): service.draft_plan(service.document_id, 0, 'restore')
    else:
        assert opened['recovery_state'] == 'invalid'
    assert service.data == original and service.revision == 0 and not service.operations


def test_saved_record_change_conflicts_with_pending_plan(canvas):
    service, floating, source, _, op = canvas
    service.apply(service.document_id, 0, op); save(service)
    reopen(service, floating, source)
    plan = service.draft_plan(service.document_id, 0, 'restore')['plan']
    path = service.drafts.target
    record = json.loads(path.read_text(encoding='utf8')); record['saved_at'] = '2026-11-09T12:00:00+00:00'
    path.write_text(json.dumps(record), encoding='utf8')
    with pytest.raises(GuardError, match='改变'): service.draft_confirm(service.document_id, 0, plan['id'], True)
    assert service.revision == 0 and not service.operations
    service.use_original(service.document_id, 0)
    service.apply(service.document_id, 0, op)
    plan = service.draft_plan(service.document_id, 1, 'save')['plan']
    record['saved_at'] = '2026-11-09T12:01:00+00:00'; path.write_text(json.dumps(record), encoding='utf8')
    with pytest.raises(GuardError, match='未确认'): service.draft_confirm(service.document_id, 1, plan['id'], True)
    assert json.loads(path.read_text(encoding='utf8')) == record


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_current_protection_blocks_restore(canvas):
    service, floating, source, _, op = canvas
    service.apply(service.document_id, 0, op); save(service)
    reopen(service, floating, source)
    floating.protected_cells.add(('实验', 'B2'))
    with pytest.raises(GuardError, match='保护'): service.draft_plan(service.document_id, 0, 'restore')
    assert service.revision == 0 and not service.operations


@pytest.mark.parametrize('damage', ['duplicate', 'nesting', 'nan'])
def test_invalid_json_keeps_original_usable(canvas, damage):
    service, floating, source, _, op = canvas
    service.apply(service.document_id, 0, op); save(service)
    path = service.drafts.target
    if damage == 'duplicate':
        path.write_text('{"schema_version":1,'+path.read_text(encoding='utf8')[1:], encoding='utf8')
    elif damage == 'nesting':
        path.write_text('['*2000+']'*2000, encoding='utf8')
    else:
        path.write_text('{"value":NaN}', encoding='utf8')
    opened = reopen(service, floating, source)
    assert opened['recovery_state'] == 'invalid' and not opened['recovery_available']
    assert service.data == source.read_bytes()
    assert service.apply(service.document_id, 0, op)['draft_revision'] == 1


@pytest.mark.parametrize('canvas', ['.xlsx'], indirect=True)
def test_formula_replay_and_source_replacement(canvas):
    service, floating, source, _, _ = canvas
    formula = operation('C2', '=B2+5', '=B2*2', action='formula')
    service.apply(service.document_id, 0, formula); save(service)
    reopen(service, floating, source)
    plan = service.draft_plan(service.document_id, 0, 'restore')['plan']
    service.draft_confirm(service.document_id, 0, plan['id'], True)
    assert next(cell for cell in engine.view(service.data, '.xlsx')['cells'] if cell['id'] == 'C2')['formula'] == 'B2+5'
    saved_ops = deepcopy(service.operations)
    assert saved_ops == [formula]
    service.undo(service.document_id, 1)
    service.close()
    before = source.read_bytes()
    replacement = source.with_name('replacement.xlsx'); replacement.write_bytes(before)
    os.replace(replacement, source)
    assert floating.admit(str(source))['ok'] and floating.read(floating.read_token, True)['ok']
    assert not service.open(floating.document['id'], floating.document['version'])['recovery_available']
    assert not service.operations and service.revision == 0


def test_changed_source_rejects_restore_confirmation(canvas):
    service, floating, source, _, op = canvas
    service.apply(service.document_id, 0, op); save(service)
    reopen(service, floating, source)
    plan = service.draft_plan(service.document_id, 0, 'restore')['plan']
    source.write_bytes(source.read_bytes()+b'changed')
    with pytest.raises(ValueError): service.draft_confirm(service.document_id, 0, plan['id'], True)
    assert service.revision == 0 and not service.operations


@pytest.mark.parametrize('canvas', ['.pptx'], indirect=True)
def test_large_checkpoint_refuses_without_losing_memory(canvas):
    service, _, source, _, _ = canvas
    old = '实验汇报'
    for index in range(70):
        value = str(index)+'x'*3998
        service.apply(service.document_id, index, operation('run-0', value, old, 'slide-0'))
        old = value
    before = service.data
    with pytest.raises(GuardError, match='512 KiB'): service.draft_plan(service.document_id, 70, 'save')
    assert service.data == before and service.state()['dirty'] and not service.drafts.target.exists()


def test_fresh_process_reads_only_saved_operations(canvas):
    service, _, source, _, op = canvas
    service.apply(service.document_id, 0, op); save(service)
    data = {'audit': str(service.drafts.audit), 'actor': 'self-created', 'key': service.drafts.key,
            'version': service.drafts.version, 'identity': service.drafts.file_identity, 'suffix': source.suffix, 'source': str(source)}
    script = """import json,sys
from pathlib import Path
sys.path.insert(0,sys.argv[1])
from office_agent.canvas_drafts import CanvasDraftStore,package_identity
from office_agent import office_canvas as e
p=json.loads(Path(sys.argv[2]).read_text(encoding='utf8'))
s=CanvasDraftStore(Path(p['audit']),p['actor'],p['key'],p['version'],p['identity'],p['suffix'])
r,_=s.read(); data=Path(p['source']).read_bytes()
for op in r['operations']: data,_=e.transform(data,p['suffix'],op)
assert package_identity(data,p['suffix'])==r['candidate_package']
Path(sys.argv[3]).write_text(json.dumps({'count':len(r['operations']),'matches':True}),encoding='utf8')
"""
    worker = source.parent/'fresh-process.py'; worker.write_text(script, encoding='utf8')
    inputs = source.parent/'fresh-input.json'; inputs.write_text(json.dumps(data), encoding='utf8')
    output = source.parent/'fresh-output.json'
    runs = []
    _run((sys.executable, str(worker), str(Path(__file__).resolve().parents[1]), str(inputs), str(output)), 20000, runs)
    assert runs and all(item['empty_confirmed'] for item in runs)
    assert json.loads(output.read_text(encoding='utf8')) == {'count': 1, 'matches': True}
