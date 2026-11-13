"""Deterministic temp-file IO, typed FAKE renderer/Job; no Office/provider/UI proof."""
from dataclasses import replace
import hashlib
import json
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from office_agent import diagram_save as module
from office_agent.diagram_process import OwnedRun
from office_agent.diagram_service import RenderedDiagram
from office_agent.diagram_save import DiagramSaveError, DiagramSaveService
from office_agent.diagrams import validate_source, validate_svg

SOURCE = 'flowchart TD\nA --> B'
SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text>A to B</text></svg>'


@pytest.fixture
def preview():
    return RenderedDiagram(validate_source(SOURCE, 'flowchart'), SVG,
                           OwnedRun('process_completed', 0, 1, 0.1, True))


@pytest.fixture
def service(tmp_path):
    source = tmp_path / 'selected.docx'
    source.write_bytes(b'fingerprint-only, NOT an Office document')
    return DiagramSaveService(tmp_path, source, 'session_1')


def save(service, preview):
    plan = service.plan_save(preview)
    result = service.confirm_save(plan['plan_id'], True)
    return result, Path(result['path']), Path(result['audit_path'])


def assert_code(code, call):
    with pytest.raises(DiagramSaveError) as caught:
        call()
    assert caught.value.code == code
    assert str(caught.value) == code


def test_planning_metadata_only_zero_writes(service, preview, monkeypatch):
    before = set(service._root.iterdir())
    monkeypatch.setattr(module.os, 'open', lambda *a, **k: pytest.fail('planning must not create'))
    plan = service.plan_save(preview)
    assert plan['confirmation_required'] is True
    assert plan['backup'] == 'not_applicable_new_artifact'
    assert 'svg' not in plan and 'source' not in plan and 'execution' not in plan
    assert set(service._root.iterdir()) == before
    plan['destination_basename'] = '../foreign.svg'
    assert service._plans[plan['plan_id']].target[1]['destination_basename'].startswith('office-diagram-')


def test_save_reopen_audit_source_untouched(service, preview):
    source = service._guard._source
    before = source.read_bytes(), source.stat().st_mtime_ns
    result, output, audit = save(service, preview)
    expected = validate_svg(SVG).encode('utf-8')
    assert result['saved'] is True and result['source_unchanged'] is True
    assert output.read_bytes() == expected
    assert result['sha256'] == hashlib.sha256(expected).hexdigest()
    assert output.parent == source.parent and output.name.isascii()
    record = json.loads(audit.read_text(encoding='utf-8'))
    assert record['saved'] is True and record['output_removed'] is False
    assert record['svg_sha256'] == result['sha256'] and record['bytes'] == len(expected)
    assert 'svg' not in record and 'source' not in record and 'session_id' not in record
    assert len(audit.read_bytes()) <= module.MAX_AUDIT_BYTES
    assert (source.read_bytes(), source.stat().st_mtime_ns) == before
    assert_code('export_plan_invalid', lambda: service.confirm_save(result['save_id'], True))


@pytest.mark.parametrize('confirmed', [False, None, 0, 1, 'true', [], {}])
def test_literal_true_consumed(service, preview, confirmed):
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_confirmation_required', lambda: service.confirm_save(plan, confirmed))
    assert_code('export_plan_invalid', lambda: service.confirm_save(plan, True))
    assert list(service._root.iterdir()) == [service._guard._source]


def test_wrong_session_consumed(service, preview):
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_operation_invalidated', lambda: service.confirm_save(plan, True, session_id='other'))
    assert_code('export_plan_invalid', lambda: service.confirm_save(plan, True))


def test_expired_consumed_and_purge(service, preview, monkeypatch):
    clock = [100.0]
    monkeypatch.setattr(module.time, 'monotonic', lambda: clock[0])
    plan = service.plan_save(preview)['plan_id']
    clock[0] = 400.0
    assert_code('export_plan_expired', lambda: service.confirm_save(plan, True))
    assert_code('export_plan_invalid', lambda: service.confirm_save(plan, True))
    clock[0] = 500.0
    for _ in range(8):
        service.plan_save(preview)
    assert_code('export_plan_limit', lambda: service.plan_save(preview))
    clock[0] = 800.0
    assert service.plan_save(preview)['confirmation_required'] is True


@pytest.mark.parametrize('mode', ['exception', 'malformed', 'collision', 'consumed_collision'])
def test_uuid_sanitized_and_never_reissued(service, preview, monkeypatch, mode):
    plan = service.plan_save(preview)['plan_id']
    if mode == 'consumed_collision':
        service.cancel(plan)
    if mode == 'exception':
        def broken():
            raise RuntimeError('PRIVATE RAW FAILURE')
        monkeypatch.setattr(module.uuid, 'uuid4', broken)
    else:
        monkeypatch.setattr(module.uuid, 'uuid4', lambda: SimpleNamespace(hex='bad' if mode == 'malformed' else plan))
    assert_code('export_plan_invalid', lambda: service.plan_save(preview))


@pytest.mark.parametrize('bad', [None, {}, 'svg', 1])
def test_client_objects_rejected_before_source_read(service, bad, monkeypatch):
    monkeypatch.setattr(service._guard, '_source_guard', lambda: pytest.fail('invalid preview must fail first'))
    assert_code('export_preview_invalid', lambda: service.plan_save(bad))


@pytest.mark.parametrize('change', ['empty_job', 'exit_bool', 'active_svg'])
def test_untrusted_native_shape_rejected(service, preview, change):
    if change == 'active_svg':
        preview = replace(preview, svg='<svg><script>evil</script></svg>')
    else:
        preview = replace(preview, execution=replace(preview.execution, **(
            {'empty_confirmed': False} if change == 'empty_job' else {'exit_code': True})))
    assert_code('export_preview_invalid', lambda: service.plan_save(preview))


def test_source_change_consumes_before_any_writes(service, preview):
    plan = service.plan_save(preview)['plan_id']
    service._guard._source.write_bytes(b'changed source')
    assert_code('export_source_changed', lambda: service.confirm_save(plan, True))
    assert_code('export_plan_invalid', lambda: service.confirm_save(plan, True))
    assert len(list(service._root.iterdir())) == 1


@pytest.mark.parametrize('destination', ['svg', 'audit'])
def test_collision_never_overwrites_or_deletes(service, preview, destination):
    plan = service.plan_save(preview)
    path = service._root / (plan['destination_basename'] if destination == 'svg' else
                            f"office-diagram-{plan['plan_id']}.audit.json")
    path.write_bytes(b'PREEXISTING')
    assert_code('export_save_failed', lambda: service.confirm_save(plan['plan_id'], True))
    assert path.read_bytes() == b'PREEXISTING'
    assert len(list(service._root.iterdir())) == 2
    assert not service._disabled


def test_owned_cleanup_after_audit_failure(service, preview, monkeypatch):
    original = service._create
    def fail_audit(path, data, created):
        if path.suffix == '.json':
            raise OSError('raw file details')
        return original(path, data, created)
    monkeypatch.setattr(service, '_create', fail_audit)
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_save_failed', lambda: service.confirm_save(plan, True))
    assert list(service._root.iterdir()) == [service._guard._source]
    assert not service._disabled


def test_uncertain_cleanup_preserves_changed_output_disables(service, preview, monkeypatch):
    original = service._create
    def mutate_before_audit(path, data, created):
        if path.suffix == '.json':
            created[0].path.write_bytes(b'EXTERNAL CHANGE')
            raise OSError('fail')
        return original(path, data, created)
    monkeypatch.setattr(service, '_create', mutate_before_audit)
    plan = service.plan_save(preview)
    assert_code('export_cleanup_uncertain', lambda: service.confirm_save(plan['plan_id'], True))
    assert (service._root / plan['destination_basename']).read_bytes() == b'EXTERNAL CHANGE'
    assert service._disabled and not service._plans
    assert_code('export_service_disabled', lambda: service.plan_save(preview))


def test_short_write_then_failure_cleans_only_known_prefix(service, preview, monkeypatch):
    original = module.os.write
    calls = [0]
    def failing(fd, data):
        calls[0] += 1
        if calls[0] == 1:
            return original(fd, data[:7])
        raise OSError('write failed')
    monkeypatch.setattr(module.os, 'write', failing)
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_save_failed', lambda: service.confirm_save(plan, True))
    assert list(service._root.iterdir()) == [service._guard._source]


def test_rollback_fresh_approval_and_audit_retained(service, preview):
    result, output, audit = save(service, preview)
    plan = service.plan_rollback(result['save_id'])
    assert output.exists()
    assert_code('export_confirmation_required', lambda: service.confirm_rollback(plan['rollback_id']))
    assert output.exists()
    assert_code('export_plan_invalid', lambda: service.confirm_rollback(plan['rollback_id'], True))
    plan = service.plan_rollback(result['save_id'])
    rolled = service.confirm_rollback(plan['rollback_id'], True)
    assert rolled['removed'] is True and rolled['audit_updated'] is True
    assert not output.exists() and audit.exists()
    record = json.loads(audit.read_text(encoding='utf-8'))
    assert record['output_removed'] is True and record['saved'] is False
    assert record['independent_human_rollback_confirmed'] is True
    assert_code('export_rollback_invalid', lambda: service.plan_rollback(result['save_id']))


@pytest.mark.parametrize('target', ['output', 'audit'])
@pytest.mark.parametrize('change', ['modified', 'missing', 'replacement', 'multilink'])
def test_rollback_changed_owned_files_fail_closed(service, preview, target, change):
    result, output, audit = save(service, preview)
    plan = service.plan_rollback(result['save_id'])['rollback_id']
    path = output if target == 'output' else audit
    if change == 'modified':
        path.write_bytes(b'EDITED')
    elif change == 'missing':
        path.unlink()  # Test owns these exact temporary paths.
    elif change == 'replacement':
        foreign = service._root / 'replacement.tmp'
        foreign.write_bytes(path.read_bytes())
        os.replace(foreign, path)
    else:
        os.link(path, service._root / 'extra-link.tmp')
    assert_code('export_rollback_changed', lambda: service.confirm_rollback(plan, True))
    if target == 'audit':
        assert output.exists()
    if change != 'missing':
        assert path.exists()
    assert_code('export_plan_invalid', lambda: service.confirm_rollback(plan, True))


def test_partial_rollback_reports_actual_deletion(service, preview, monkeypatch):
    result, output, audit = save(service, preview)
    plan = service.plan_rollback(result['save_id'])['rollback_id']
    def fail(saved):
        raise OSError('private raw failure')
    monkeypatch.setattr(service, '_update_audit', fail)
    rolled = service.confirm_rollback(plan, True)
    assert rolled['removed'] is True and rolled['audit_updated'] is False
    assert rolled['code'] == 'export_rollback_partial' and rolled['service_disabled'] is True
    assert not output.exists() and audit.exists()
    assert_code('export_service_disabled', lambda: service.plan_save(preview))


def test_mismatched_opened_audit_never_truncated(service, preview, monkeypatch):
    result, output, audit = save(service, preview)
    plan = service.plan_rollback(result['save_id'])['rollback_id']
    foreign = service._root / 'foreign.json'
    foreign.write_bytes(b'NEVER TRUNCATE OR WRITE THIS')
    original = Path.open
    def substituted(self, mode='r', *args, **kwargs):
        if self == audit and mode == 'r+b':
            return original(foreign, mode, *args, **kwargs)
        return original(self, mode, *args, **kwargs)
    before = audit.read_bytes()
    monkeypatch.setattr(Path, 'open', substituted)
    rolled = service.confirm_rollback(plan, True)
    assert rolled['removed'] is True and rolled['audit_updated'] is False
    assert foreign.read_bytes() == b'NEVER TRUNCATE OR WRITE THIS'
    assert audit.read_bytes() == before and not output.exists()


def test_invalidate_drops_authority_preserves_files(service, preview):
    result, output, audit = save(service, preview)
    rollback = service.plan_rollback(result['save_id'])['rollback_id']
    pending = service.plan_save(preview)['plan_id']
    service.invalidate()
    assert not service._saved and not service._plans and not service._rollbacks
    assert output.exists() and audit.exists()
    assert_code('export_operation_invalidated', lambda: service.confirm_save(pending, True))
    assert_code('export_operation_invalidated', lambda: service.confirm_rollback(rollback, True))
    fresh = DiagramSaveService(service._root, service._guard._source, 'session_1')
    assert_code('export_rollback_invalid', lambda: fresh.plan_rollback(result['save_id']))


def test_registry_bounds(service, preview):
    for _ in range(8):
        save(service, preview)
    assert_code('export_saved_limit', lambda: service.plan_save(preview))
    save_id = next(iter(service._saved))
    for _ in range(8):
        service.plan_rollback(save_id)
    assert_code('export_plan_limit', lambda: service.plan_rollback(save_id))


def test_root_identity_change_prevents_write(service, preview, monkeypatch):
    plan = service.plan_save(preview)['plan_id']
    service._guard._root_identity = (-1, -1)
    assert_code('export_source_changed', lambda: service.confirm_save(plan, True))
    assert len(list(service._root.iterdir())) == 1


def test_identifier_ledger_bounded(service, preview, monkeypatch):
    monkeypatch.setattr(module, 'MAX_IDENTIFIERS', 1)
    token = service.plan_save(preview)['plan_id']
    service.cancel(token)
    assert_code('export_id_limit', lambda: service.plan_save(preview))


@pytest.mark.parametrize('expanded', ['x' * 1_048_577, 'é' * 524_289], ids=['ascii-byte-limit', 'unicode-byte-limit'])
def test_canonical_byte_bound_before_source_io(service, preview, monkeypatch, expanded):
    # Synthetic validator expansion, not a grammar or native renderer claim.
    monkeypatch.setattr(module, 'validate_svg', lambda text: expanded)
    monkeypatch.setattr(service._guard, '_source_guard', lambda: pytest.fail('reject oversized first'))
    assert_code('export_preview_invalid', lambda: service.plan_save(preview))


def test_audit_bound_failure_cleans_owned_output(service, preview, monkeypatch):
    monkeypatch.setattr(module, 'MAX_AUDIT_BYTES', 1)
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_save_failed', lambda: service.confirm_save(plan, True))
    assert list(service._root.iterdir()) == [service._guard._source]


def test_unknown_creation_identity_retains_disables(service, preview, monkeypatch):
    original = module.os.fstat
    def fail_owned_fd(fd):
        # Source snapshot uses a readable fd; creation opens its output write-only.
        stat = original(fd)
        if stat.st_size == 0:
            raise OSError('cannot prove ownership')
        return stat
    monkeypatch.setattr(module.os, 'fstat', fail_owned_fd)
    plan = service.plan_save(preview)
    assert_code('export_cleanup_uncertain', lambda: service.confirm_save(plan['plan_id'], True))
    assert (service._root / plan['destination_basename']).exists() and service._disabled


def test_multilink_cleanup_never_deletes_either_entry(service, preview, monkeypatch):
    original = service._create
    linked = service._root / 'linked.tmp'
    def fail_after_link(path, data, created):
        if path.suffix == '.json':
            os.link(created[0].path, linked)
            raise OSError('audit failure')
        return original(path, data, created)
    monkeypatch.setattr(service, '_create', fail_after_link)
    plan = service.plan_save(preview)
    assert_code('export_cleanup_uncertain', lambda: service.confirm_save(plan['plan_id'], True))
    assert linked.exists() and (service._root / plan['destination_basename']).exists()


def test_fsync_failure_cleanup(service, preview, monkeypatch):
    original = module.os.fsync
    calls = [0]
    def fail_second(fd):
        calls[0] += 1
        if calls[0] == 2:
            raise OSError('audit fsync failed')
        return original(fd)
    monkeypatch.setattr(module.os, 'fsync', fail_second)
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_save_failed', lambda: service.confirm_save(plan, True))
    assert list(service._root.iterdir()) == [service._guard._source]


def test_source_changed_during_save_cleans_only_new_artifacts(service, preview, monkeypatch):
    original = service._create
    def change_source(path, data, created):
        result = original(path, data, created)
        if path.suffix == '.json':
            service._guard._source.write_bytes(b'EXTERNAL SOURCE EDIT')
        return result
    monkeypatch.setattr(service, '_create', change_source)
    plan = service.plan_save(preview)['plan_id']
    assert_code('export_save_failed', lambda: service.confirm_save(plan, True))
    assert service._guard._source.read_bytes() == b'EXTERNAL SOURCE EDIT'
    assert list(service._root.iterdir()) == [service._guard._source]


@pytest.mark.parametrize('reason', ['wrong_session', 'expired'])
def test_rollback_approval_consumed_on_scope_or_expiry(service, preview, monkeypatch, reason):
    result, output, audit = save(service, preview)
    clock = [100.0]
    monkeypatch.setattr(module.time, 'monotonic', lambda: clock[0])
    token = service.plan_rollback(result['save_id'])['rollback_id']
    if reason == 'expired':
        clock[0] = 400.0
        assert_code('export_plan_expired', lambda: service.confirm_rollback(token, True))
    else:
        assert_code('export_operation_invalidated', lambda: service.confirm_rollback(token, True, session_id='other'))
    assert output.exists() and audit.exists()
    assert_code('export_plan_invalid', lambda: service.confirm_rollback(token, True))


def test_partial_audit_write_reports_unknown_not_false(service, preview, monkeypatch):
    result, output, audit = save(service, preview)
    token = service.plan_rollback(result['save_id'])['rollback_id']
    def partial(saved):
        saved.audit.path.write_bytes(b'PARTIAL AUDIT WRITE')
        raise OSError('failed after mutation')
    monkeypatch.setattr(service, '_update_audit', partial)
    result = service.confirm_rollback(token, True)
    assert result['removed'] is True and result['audit_updated'] is None
    assert result['service_disabled'] and not output.exists()
    assert audit.read_bytes() == b'PARTIAL AUDIT WRITE'


def test_post_update_scope_failure_does_not_hide_verified_audit(service, preview, monkeypatch):
    result, output, audit = save(service, preview)
    token = service.plan_rollback(result['save_id'])['rollback_id']
    original = service._update_audit
    def update_then_change_source(saved):
        original(saved)
        service._guard._source.write_bytes(b'LATE EXTERNAL EDIT')
    monkeypatch.setattr(service, '_update_audit', update_then_change_source)
    rolled = service.confirm_rollback(token, True)
    assert rolled['removed'] is True and rolled['audit_updated'] is True
    assert rolled['code'] == 'export_rollback_partial' and rolled['service_disabled']
    assert not output.exists() and json.loads(audit.read_text(encoding='utf-8'))['output_removed']


def test_failed_unlink_preserves_authority_for_new_approval(service, preview, monkeypatch):
    result, output, audit = save(service, preview)
    token = service.plan_rollback(result['save_id'])['rollback_id']
    original = Path.unlink
    def deny(self, *args, **kwargs):
        if self == output:
            raise PermissionError('locked by another program')
        return original(self, *args, **kwargs)
    monkeypatch.setattr(Path, 'unlink', deny)
    assert_code('export_rollback_failed', lambda: service.confirm_rollback(token, True))
    assert output.exists() and audit.exists() and not service._disabled
    assert service.plan_rollback(result['save_id'])['confirmation_required']


def test_saved_limit_rechecked_at_confirmation(service, preview):
    for _ in range(7):
        save(service, preview)
    first = service.plan_save(preview)['plan_id']
    second = service.plan_save(preview)['plan_id']
    service.confirm_save(first, True)
    assert_code('export_saved_limit', lambda: service.confirm_save(second, True))
    assert_code('export_plan_invalid', lambda: service.confirm_save(second, True))
