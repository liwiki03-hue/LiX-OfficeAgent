from copy import deepcopy
import json

import pytest

from office_agent.output_dependencies import OutputDependencies
from office_agent.webui.output_review import OutputReview
from office_agent.writes import digest
from test_output_dependencies import family


def reviewer(service, manager):
    return OutputReview(service._lock, service._scope, service._editor, service._selected, manager, service.rollback,
                        lambda receipt: service._saved.get(receipt) == service._binding()[2])


def test_child_first_review_refresh_and_independent_single_output_undo(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    original = digest(source)
    selected[0] = source
    parent_review = reviewer(first, manager)
    started = parent_review.start(parent, [])['result']
    assert started['suggested_child_first_order'] == [child, parent]
    assert started['active_descendants'] == [child] and started['review_blocked']
    assert parent_review.acknowledge(started['review_id'], 1, True)['reason'] == 'professional_review_blocked'
    selected[0] = output
    child_review = reviewer(second, manager)
    current = child_review.start(child, [])['result']
    assert current['target_has_current_undo_authority'] and not current['file_content_revalidated']
    ack = child_review.acknowledge(current['review_id'], 1, True)['result']
    assert not ack['review_is_write_approval'] and not ack['write_approval_granted']
    token = ack['review_token']
    assert child_review.rollback(current['review_id'], token, False)['result']['status'] == 'cancelled'
    assert child_output.exists()
    assert child_review.rollback(current['review_id'], token, True)['reason'] == 'professional_review_not_acknowledged'
    ack = child_review.acknowledge(current['review_id'], 1, True)['result']
    assert child_review.rollback(current['review_id'], ack['review_token'], True)['result']['status'] == 'rolled_back'
    selected[0] = source
    refreshed = parent_review.refresh(started['review_id'])['result']
    assert refreshed['revision'] == 2 and child in refreshed['changes']['changed']
    assert not refreshed['review_blocked'] and not refreshed['review_acknowledged']
    ack = parent_review.acknowledge(started['review_id'], 2, True)['result']
    assert parent_review.rollback(started['review_id'], ack['review_token'], True)['result']['status'] == 'rolled_back'
    assert digest(source) == original and not output.exists() and not child_output.exists()


def test_bad_candidate_does_not_poison_live_scope_and_replaced_session_is_retired(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = output
    review = reviewer(second, manager)
    before = deepcopy(manager.known)
    invalid = review.start(child, ['f' * 32])['result']
    assert invalid['unresolved_receipts'] == ['f' * 32] and invalid['review_blocked']
    assert manager.known == before
    valid = review.start(child, [])['result']
    assert not valid['review_blocked'] and manager.known == before
    assert review.refresh(invalid['review_id'])['reason'] == 'professional_review_invalid'
    assert review.cancel(valid['review_id'])['result']['wrote_nothing'] and child_output.exists()


@pytest.mark.parametrize('change', ['audit', 'output', 'selection', 'new_known', 'missing_audit'])
def test_changed_snapshot_rejects_old_ticket_before_undo(tmp_path, monkeypatch, change):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = output
    review = reviewer(second, manager)
    current = review.start(child, [])['result']
    ack = review.acknowledge(current['review_id'], 1, True)['result']
    audit = editor.audit / child / 'audit.json'
    if change == 'audit':
        audit.write_bytes(audit.read_bytes() + b'\n')
    elif change == 'output':
        child_output.write_bytes(child_output.read_bytes() + b'\n')
    elif change == 'selection':
        output.write_bytes(output.read_bytes() + b'\n')
    elif change == 'missing_audit':
        audit.unlink()
    else:
        manager.admit(editor, 'e' * 32, output)
    denied = review.rollback(current['review_id'], ack['review_token'], True)
    assert denied['reason'] == 'professional_review_changed', denied
    assert child_output.exists() and output.exists() and not second._disabled
    refreshed = review.refresh(current['review_id'])['result']
    assert refreshed['revision'] == 2 and not refreshed['review_acknowledged']
    assert refreshed['changes']['added'] or refreshed['changes']['changed']


def test_review_does_not_restore_lost_undo_grant(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = output
    review = reviewer(second, manager)
    current = review.start(child, [])['result']
    second._saved.clear()
    ack = review.acknowledge(current['review_id'], 1, True)['result']
    assert not ack['target_has_current_undo_authority'] and not ack['restores_rollback_authority']
    assert review.rollback(current['review_id'], ack['review_token'], True)['reason'] == 'professional_plan_invalid'
    assert child_output.exists()


def test_refresh_does_not_extend_expiry_and_logout_retires_review(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = output
    authenticated = [True]
    review = reviewer(second, manager)
    review._scope = lambda: (authenticated[0], False, second._scope()[2])
    current = review.start(child, [])['result']
    entry = review.reviews[current['review_id']]
    started = entry['created']
    ack = review.acknowledge(current['review_id'], 1, True)['result']
    assert review.refresh(current['review_id'])['ok'] and entry['created'] == started
    assert review.rollback(current['review_id'], ack['review_token'], True)['reason'] == 'professional_review_not_acknowledged'
    entry['created'] -= 301
    assert review.refresh(current['review_id'])['reason'] == 'professional_review_expired'
    current = review.start(child, [])['result']
    authenticated[0] = False
    assert review.refresh(current['review_id'])['reason'] == 'not_authenticated'
    authenticated[0] = True
    assert review.refresh(current['review_id'])['reason'] == 'professional_review_invalid'
    assert child_output.exists()


def test_cycles_are_reported_without_scanning_or_approval(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    path = editor.audit / parent / 'audit.json'
    audit = json.loads(path.read_bytes())
    after = json.loads(audit['after'])
    after['items'].append({'citation_manifest': {'parent_receipt': child}})
    audit['after'] = json.dumps(after)
    path.write_text(json.dumps(audit), encoding='utf8')
    fresh = OutputDependencies()
    selected[0] = source
    review = reviewer(first, fresh)
    from pathlib import Path
    monkeypatch.setattr(Path, 'iterdir', lambda *args: pytest.fail('Review must not scan directories'))
    result = review.start(parent, [child])['result']
    assert result['cycle_or_cycle_dependent_receipts'] == sorted([parent, child])
    assert result['suggested_child_first_order'] == [] and result['review_blocked']
    assert not result['write_approval_granted'] and fresh.known == {}


def test_active_grandchild_blocks_even_when_child_was_removed_outside_host(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = child_output
    data = second.read_structure()['result']
    plan = second.plan_structure({'source_sha256': data['document']['source_sha256'],
        'insertions': [{'anchor': data['anchors'][0]['locator_id'], 'side': 'after', 'text': 'Grandchild context.'}]})
    assert plan['ok'], plan
    grandchild = plan['result']['id']
    assert second.confirm(grandchild, True)['result']['status'] == 'saved'
    assert editor.rollback(child, confirmed=True)['status'] == 'rolled_back'
    selected[0] = source
    review = reviewer(first, manager)
    current = review.start(parent, [])['result']
    assert current['active_descendants'] == [grandchild] and current['review_blocked']
    assert review.acknowledge(current['review_id'], 1, True)['reason'] == 'professional_review_blocked'
    assert output.exists()
