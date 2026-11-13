"""Deterministic host-only capsule tests; no renderer, Office or export writes."""
from dataclasses import replace
import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from office_agent import diagram_transactions as tx
from office_agent.diagram_process import OwnedRun
from office_agent.diagram_service import RenderedDiagram
from office_agent.diagrams import validate_source

SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100"><rect width="10" height="10"/></svg>'


@pytest.fixture
def capsule(tmp_path):
    source = tmp_path / 'source.docx'
    source.write_bytes(b'PK00')  # Fingerprint-only fixture, not an Office document.
    preview = RenderedDiagram(validate_source('flowchart TD\nA --> B'), SVG,
                              OwnedRun('process_completed', 0, 123, 0.01, True))
    return tx.DiagramExportService(tmp_path, source, 'session_1'), source, preview


def error(code, call):
    with pytest.raises(tx.DiagramExportError) as caught:
        call()
    assert caught.value.code == code
    assert str(caught.value) == code
    return caught.value


def test_metadata_copies_and_confirmation_never_save(capsule):
    service, source, preview = capsule
    before = set(source.parent.iterdir())
    plan = service.plan_export(preview)
    assert set(source.parent.iterdir()) == before == {source}
    assert preview.source.source not in str(plan) and SVG not in str(plan)
    assert plan['saving_implemented'] is False and plan['confirmation_required'] is True
    token = plan['plan_id']
    plan['kind'] = 'forged'
    assert service._plans[token].metadata['kind'] == 'flowchart'
    assert service.confirm_export(token, True) == {
        'plan_id': token, 'saved': False, 'code': 'export_saving_unimplemented'}
    error('export_plan_invalid', lambda: service.confirm_export(token, True))
    assert set(source.parent.iterdir()) == before and source.read_bytes() == b'PK00'


@pytest.mark.parametrize('confirmation', ['default', False, None, 1, 'true'])
def test_strict_confirmation_consumes(capsule, confirmation):
    service, _, preview = capsule
    token = service.plan_export(preview)['plan_id']
    call = (lambda: service.confirm_export(token)) if confirmation == 'default' else (
        lambda: service.confirm_export(token, confirmation))
    error('export_confirmation_required', call)
    error('export_plan_invalid', lambda: service.confirm_export(token, True))


def test_wrong_session_consumes(capsule):
    service, _, preview = capsule
    token = service.plan_export(preview)['plan_id']
    error('export_operation_invalidated', lambda: service.confirm_export(token, True, session_id='other'))
    error('export_plan_invalid', lambda: service.confirm_export(token, True))


def test_expiry(capsule, monkeypatch):
    service, _, preview = capsule
    monkeypatch.setattr(tx.time, 'monotonic', lambda: 10.0)
    token = service.plan_export(preview)['plan_id']
    monkeypatch.setattr(tx.time, 'monotonic', lambda: 10.0 + tx.APPROVAL_TTL)
    error('export_plan_expired', lambda: service.confirm_export(token, True))
    error('export_plan_invalid', lambda: service.confirm_export(token, True))


def test_cancel_invalidate(capsule):
    service, _, preview = capsule
    token = service.plan_export(preview)['plan_id']
    assert service.cancel(token) == {'plan_id': token, 'cancelled': True}
    error('export_plan_invalid', lambda: service.cancel(token))
    service.plan_export(preview)
    assert service.invalidate() == {'invalidated': True}
    assert not service._plans
    error('export_operation_invalidated', lambda: service.plan_export(preview))


def test_eight_slots_purge_reuse(capsule, monkeypatch):
    service, _, preview = capsule
    monkeypatch.setattr(tx.time, 'monotonic', lambda: 0.0)
    plans = [service.plan_export(preview) for _ in range(8)]
    assert len({p['plan_id'] for p in plans}) == 8
    error('export_plan_limit', lambda: service.plan_export(preview))
    monkeypatch.setattr(tx.time, 'monotonic', lambda: tx.APPROVAL_TTL)
    service.plan_export(preview)
    assert len(service._plans) == 1


@pytest.mark.parametrize('digest_only', [False, True])
def test_changed_source_consumes(capsule, monkeypatch, digest_only):
    service, source, preview = capsule
    token = service.plan_export(preview)['plan_id']
    stamp = source.stat()
    source.write_bytes(b'PK11')
    os.utime(source, ns=(stamp.st_atime_ns, stamp.st_mtime_ns + 2_000_000_000))
    assert source.stat().st_mtime_ns != stamp.st_mtime_ns
    if digest_only:
        # Pin observed metadata to isolate the same-length content digest guard.
        monkeypatch.setattr(tx, '_stamp', lambda info: service._source_stamp)
    error('export_source_changed', lambda: service.confirm_export(token, True))
    error('export_plan_invalid', lambda: service.confirm_export(token, True))
    assert source.read_bytes() == b'PK11' and set(source.parent.iterdir()) == {source}


def test_source_size_bound(capsule, monkeypatch):
    service, source, preview = capsule
    monkeypatch.setattr(tx, 'MAX_SOURCE_BYTES', 3)
    error('export_source_changed', lambda: service.plan_export(preview))
    error('export_source_invalid', lambda: tx.DiagramExportService(source.parent, source, 's'))


def test_root_identity_changed(capsule, monkeypatch):
    service, _, preview = capsule
    monkeypatch.setattr(tx, '_identity', lambda info: ('different', 0))
    error('export_source_changed', lambda: service.plan_export(preview))


@pytest.mark.parametrize('guard', ['audit', 'workspace', 'links'])
def test_path_risk_guards_without_privileged_links(capsule, monkeypatch, guard):
    service, _, preview = capsule
    if guard == 'audit':
        monkeypatch.setattr(tx, '_audit_path', Mock(side_effect=ValueError('private path')))
    elif guard == 'workspace':
        monkeypatch.setattr(service._space, 'file', Mock(side_effect=ValueError('outside')))
    else:
        original = Path.lstat
        def linked(path, *args, **kwargs):
            info = original(path, *args, **kwargs)
            return SimpleNamespace(st_dev=info.st_dev, st_ino=info.st_ino, st_size=info.st_size,
                                   st_mtime_ns=info.st_mtime_ns, st_ctime_ns=info.st_ctime_ns, st_nlink=2)
        monkeypatch.setattr(Path, 'lstat', linked)
    error('export_source_changed', lambda: service.plan_export(preview))


@pytest.mark.parametrize('invalid', ['dict', 'subclass', 'source_type', 'canonical', 'run_type',
                                   'run_code', 'bool_exit', 'nonempty', 'unsafe_svg', 'byte_bound'])
def test_invalid_preview_before_source_io(capsule, monkeypatch, invalid):
    service, _, preview = capsule
    if invalid == 'dict':
        preview = {'source': preview.source, 'svg': SVG, 'execution': preview.execution}
    elif invalid == 'subclass':
        class Derived(RenderedDiagram):
            pass
        preview = Derived(preview.source, SVG, preview.execution)
    elif invalid == 'source_type':
        preview = replace(preview, source=SimpleNamespace(**vars(preview.source)))
    elif invalid == 'canonical':
        preview = replace(preview, source=replace(preview.source, sha256='0' * 64))
    elif invalid == 'run_type':
        preview = replace(preview, execution=SimpleNamespace(**vars(preview.execution)))
    elif invalid == 'run_code':
        preview = replace(preview, execution=replace(preview.execution, code='process_failed'))
    elif invalid == 'bool_exit':
        preview = replace(preview, execution=replace(preview.execution, exit_code=False))
    elif invalid == 'nonempty':
        preview = replace(preview, execution=replace(preview.execution, empty_confirmed=False))
    elif invalid == 'unsafe_svg':
        preview = replace(preview, svg=SVG.replace('<rect', '<script'))
    else:
        monkeypatch.setattr(tx, 'MAX_SVG_BYTES', 1)
    spy = Mock(side_effect=AssertionError('source I/O forbidden'))
    monkeypatch.setattr(service, '_snapshot', spy)
    error('export_preview_invalid', lambda: service.plan_export(preview))
    spy.assert_not_called()


@pytest.mark.parametrize('mode', ['exception', 'malformed', 'collision'])
def test_uuid_failure_normalized(capsule, monkeypatch, mode):
    service, _, preview = capsule
    if mode == 'exception':
        monkeypatch.setattr(tx.uuid, 'uuid4', Mock(side_effect=RuntimeError('sensitive')))
    elif mode == 'malformed':
        monkeypatch.setattr(tx.uuid, 'uuid4', lambda: SimpleNamespace(hex='BAD'))
    else:
        token = service.plan_export(preview)['plan_id']
        monkeypatch.setattr(tx.uuid, 'uuid4', lambda: SimpleNamespace(hex=token))
    caught = error('export_plan_invalid', lambda: service.plan_export(preview))
    assert caught.__cause__ is None
    assert len(service._plans) == (1 if mode == 'collision' else 0)


def test_rollback_fixed_zero_io(capsule, monkeypatch):
    service, source, _ = capsule
    spy = Mock(side_effect=AssertionError('I/O forbidden'))
    monkeypatch.setattr(service, '_snapshot', spy)
    error('export_rollback_unimplemented', lambda: service.plan_rollback('../arbitrary'))
    error('export_rollback_unimplemented', lambda: service.confirm_rollback('../arbitrary', True))
    spy.assert_not_called()
    assert source.read_bytes() == b'PK00' and set(source.parent.iterdir()) == {source}


def test_constructor_rejects_wrong_parent_suffix_session(capsule):
    _, source, _ = capsule
    error('export_source_invalid', lambda: tx.DiagramExportService(source.parent.parent, source, 's'))
    error('export_source_invalid', lambda: tx.DiagramExportService(source.parent, source.with_suffix('.txt'), 's'))
    error('export_session_invalid', lambda: tx.DiagramExportService(source.parent, source, True))
    assert set(source.parent.iterdir()) == {source}

def test_empty_svg_bytes_before_io(capsule, monkeypatch):
    service, _, preview = capsule
    monkeypatch.setattr(tx, 'validate_svg', lambda svg: '')
    spy = Mock(side_effect=AssertionError('I/O forbidden'))
    monkeypatch.setattr(service, '_snapshot', spy)
    error('export_preview_invalid', lambda: service.plan_export(preview))
    spy.assert_not_called()
