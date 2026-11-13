from copy import deepcopy
import json

import pytest

from office_agent.output_dependencies import OutputDependencies
from office_agent.review_archive import validate
from office_agent.rag_evidence import EvidenceError
from office_agent.writes import digest
from test_output_dependencies import family
from test_output_review import reviewer


def prepared(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = output
    review = reviewer(second, manager)
    session = review.start(child, [])['result']
    return editor, source, manager, second, child, child_output, review, session


def save(review, session):
    plan = review.plan_archive(session['review_id'], session['revision'])
    assert plan['ok'], plan
    result = review.confirm_archive(plan['result']['plan_id'], True)
    assert result['ok'] and result['result']['status'] == 'saved', result
    return result['result']['save_id']


def test_review_set_save_restore_rechecks_without_restoring_permissions(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    original = digest(source)
    receipt = save(review, session)
    path = tmp_path / ('.office-rag-evidence-' + receipt) / 'output-review-set.json'
    document = json.loads(path.read_bytes())
    assert set(document) == {'format', 'scope', 'target', 'receipts', 'stamps'}
    assert len(document['receipts']) == 2 and child in document['receipts']
    assert str(tmp_path) not in path.read_text() and 'review_token' not in path.read_text()
    second._saved.clear()
    fresh_manager = OutputDependencies()
    restored_review = reviewer(second, fresh_manager)
    restored = restored_review.restore_archive(receipt)
    assert restored['ok'], restored
    restored = restored['result']
    assert restored['changes'] == {'added': [], 'changed': [], 'removed': []}
    assert restored['historical_only'] and not restored['target_has_current_undo_authority']
    assert not restored['review_acknowledged'] and not restored['restores_approval'] and fresh_manager.known == {}
    ack = restored_review.acknowledge(restored['review_id'], 1, True)['result']
    assert restored_review.rollback(restored['review_id'], ack['review_token'], True)['reason'] == 'professional_review_read_only'
    assert digest(source) == original and output.exists()


def test_archive_cancel_replay_changed_snapshot_and_logout(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    plan = review.plan_archive(session['review_id'], 1)['result']
    assert review.confirm_archive(plan['plan_id'], False)['result']['wrote_nothing']
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    assert review.confirm_archive(plan['plan_id'], True)['reason'] == 'professional_plan_invalid'
    plan = review.plan_archive(session['review_id'], 1)['result']
    output.write_bytes(output.read_bytes() + b'\n')
    assert review.confirm_archive(plan['plan_id'], True)['reason'] == 'professional_review_changed'
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    session = review.refresh(session['review_id'])['result']
    plan = review.plan_archive(session['review_id'], session['revision'])['result']
    review.invalidate()
    assert review.confirm_archive(plan['plan_id'], True)['reason'] == 'professional_plan_invalid'


@pytest.mark.parametrize('mutation', ['missing_audit', 'changed_output', 'new_known'])
def test_restoration_reports_current_changes_without_admitting_candidates(tmp_path, monkeypatch, mutation):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    receipt = save(review, session)
    manager = OutputDependencies()
    if mutation == 'missing_audit':
        (editor.audit / child / 'audit.json').unlink()
    elif mutation == 'changed_output':
        output.write_bytes(output.read_bytes() + b'\n')
    else:
        manager.admit(editor, 'e' * 32, source)
    before = deepcopy(manager.known)
    restored = reviewer(second, manager).restore_archive(receipt)['result']
    assert manager.known == before
    if mutation == 'new_known':
        assert restored['changes']['added'] == ['e' * 32] and restored['review_blocked']
    else:
        assert child in restored['changes']['changed']
    if mutation == 'missing_audit':
        assert child in restored['unresolved_receipts'] and restored['review_blocked']


def test_review_archive_removal_cancel_and_external_change_refused(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    receipt = save(review, session)
    fresh = reviewer(second, OutputDependencies())
    plan = fresh.plan_remove_archive(receipt)['result']
    cancelled = fresh.confirm_remove_archive(plan['rollback_id'], False)['result']
    assert cancelled['status'] == 'rejected' and cancelled['code'] == 'evidence_denied' and cancelled['wrote_nothing']
    plan = fresh.plan_remove_archive(receipt)['result']
    path = tmp_path / ('.office-rag-evidence-' + receipt) / 'output-review-set.json'
    path.write_bytes(path.read_bytes() + b'\n')
    assert not fresh.confirm_remove_archive(plan['rollback_id'], True)['ok']
    assert path.exists() and output.exists() and source.exists()


def test_fresh_instance_can_remove_only_archive_with_audit_retained(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    receipt = save(review, session)
    fresh = reviewer(second, OutputDependencies())
    plan = fresh.plan_remove_archive(receipt)['result']
    result = fresh.confirm_remove_archive(plan['rollback_id'], True)
    assert result['ok'] and result['result']['status'] == 'rolled_back', result
    directory = tmp_path / ('.office-rag-evidence-' + receipt)
    assert not (directory / 'output-review-set.json').exists()
    assert (directory / 'backup').exists() and (directory / 'audit-removed.json').exists()
    assert not fresh.restore_archive(receipt)['ok'] and output.exists() and source.exists()


def test_archive_plan_expiry_scope_and_extra_file_refused(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    plan = review.plan_archive(session['review_id'], 1)['result']
    review.archive_backend._plans[plan['plan_id']]['time'] -= 301
    expired = review.confirm_archive(plan['plan_id'], True)
    assert expired['result']['code'] == 'evidence_expired' and expired['result']['wrote_nothing'], expired
    receipt = save(review, session)
    directory = tmp_path / ('.office-rag-evidence-' + receipt)
    monkeypatch.setattr(review, '_hash', lambda value: '0' * 64)
    assert review.restore_archive(receipt)['reason'] == 'professional_review_archive_scope'
    (directory / 'unexpected').write_bytes(b'owned')
    assert review.restore_archive(receipt)['reason'] == 'professional_review_archive_evidence_path_changed'


def test_full_reviewed_collection_restores_within_existing_capacity(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    for index in range(30):
        manager.admit(editor, f'{index:032x}', source)
    session = review.refresh(session['review_id'])['result']
    assert len(session['records']) == 32 and session['review_blocked']
    receipt = save(review, session)
    fresh_manager = OutputDependencies()
    fresh = reviewer(second, fresh_manager)
    restored = fresh.restore_archive(receipt)
    assert restored['ok'] and len(restored['result']['records']) == 32, restored
    assert len(restored['result']['unresolved_receipts']) == 30 and fresh_manager.known == {}
    fresh_manager.admit(editor, 'f' * 32, source)
    assert fresh.restore_archive(receipt)['reason'] == 'professional_dependency_capacity'
    assert len(fresh_manager.known) == 1


def test_uncertain_archive_write_fuses_without_retry_or_document_change(tmp_path, monkeypatch):
    editor, source, manager, second, child, output, review, session = prepared(tmp_path, monkeypatch)
    original = digest(source)
    plan = review.plan_archive(session['review_id'], 1)['result']
    def failed(*args):
        raise OSError('owned injected failure')
    monkeypatch.setattr(review.archive_backend, '_exclusive', failed)
    result = review.confirm_archive(plan['plan_id'], True)['result']
    assert result['status'] == 'uncertain' and not result['wrote_nothing']
    assert review._disabled and review.confirm_archive(plan['plan_id'], True)['reason'] == 'professional_uncertain'
    assert digest(source) == original and output.exists()


@pytest.mark.parametrize('change', ['extra', 'duplicate', 'too_many', 'boolean', 'bad_stamp'])
def test_review_archive_schema_rejects_untrusted_metadata(change):
    document = {'format': 'output-review-set-v1', 'scope': 'a' * 64, 'target': 'b' * 32,
                'receipts': ['b' * 32], 'stamps': {'b' * 32: 'c' * 64}}
    if change == 'extra':
        document['review_token'] = 'd' * 32
    elif change == 'duplicate':
        document['receipts'] *= 2
    elif change == 'too_many':
        document['receipts'] = [f'{index:032x}' for index in range(33)]
    elif change == 'boolean':
        document['target'] = True
    else:
        document['stamps']['b' * 32] = '../outside'
    with pytest.raises(EvidenceError):
        validate(document)
