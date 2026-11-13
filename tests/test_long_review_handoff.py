from copy import deepcopy

import pytest

from office_agent import render_handoff as handoff
from tests.test_render_handoff import prepared


@pytest.fixture(autouse=True)
def isolated():
    handoff.clear()
    yield
    handoff.clear()


def reports(tmp_path):
    first, pdfs = prepared(tmp_path)
    second = {**deepcopy(first), 'source_sha256': first['output_sha256'], 'output_sha256': 'second-output'}
    third = {**deepcopy(second), 'source_sha256': second['output_sha256'], 'output_sha256': 'third-output'}
    return (first, second, third), pdfs


@pytest.mark.parametrize('order', [(0, 1), (1, 0)])
def test_two_groups_consumed_independently_with_same_total_budget(tmp_path, order):
    items, pdfs = reports(tmp_path)
    expected = tuple(path.read_bytes() for path in pdfs)
    for item in items[:2]:
        assert handoff.offer(tmp_path, item, pdfs)
    assert len(handoff._entries()) == 2
    assert sum(len(data) for entry in handoff._entries() for data in entry['payloads']) <= handoff.MAX_BYTES
    for index in order:
        assert handoff.take(tmp_path, items[index]) == expected
    assert not handoff._entries()
    assert handoff.take(tmp_path, items[order[0]]) is None


def test_third_group_evicts_oldest_without_increasing_count_or_budget(tmp_path):
    items, pdfs = reports(tmp_path)
    for item in items:
        assert handoff.offer(tmp_path, item, pdfs)
    assert len(handoff._entries()) == 2
    assert all(entry['key'] != handoff._key(tmp_path, items[0]) for entry in handoff._entries())
    assert handoff.take(tmp_path, items[1])
    assert handoff.take(tmp_path, items[2])


def test_combined_capacity_and_oversize_checked_before_loading(tmp_path, monkeypatch):
    items, pdfs = reports(tmp_path)
    size = sum(path.stat().st_size for path in pdfs)
    monkeypatch.setattr(handoff, 'MAX_BYTES', size)
    assert handoff.offer(tmp_path, items[0], pdfs)
    assert handoff.offer(tmp_path, items[1], pdfs)
    assert len(handoff._entries()) == 1
    assert handoff.take(tmp_path, items[1])
    monkeypatch.setattr(handoff, 'MAX_BYTES', size - 1)
    def forbidden(path):
        raise AssertionError('oversize PDF must not be loaded')
    monkeypatch.setattr('office_agent.fidelity.bounded_bytes', forbidden)
    assert not handoff.offer(tmp_path, items[0], pdfs)


def test_expiry_does_not_refresh_or_remove_another_group(tmp_path, monkeypatch):
    items, pdfs = reports(tmp_path)
    clock = [100.0]
    monkeypatch.setattr(handoff.time, 'monotonic', lambda: clock[0])
    assert handoff.offer(tmp_path, items[0], pdfs)
    old_key, old_deadline = handoff._entry['key'], handoff._entry['deadline']
    clock[0] = 200.0
    assert handoff.offer(tmp_path, items[1], pdfs)
    assert handoff._previous['deadline'] == old_deadline
    clock[0] = 401.0
    handoff._expire(old_key, old_deadline)
    assert len(handoff._entries()) == 1
    assert handoff.take(tmp_path, items[1])


def test_corruption_and_timer_failure_keep_unrelated_valid_group(tmp_path, monkeypatch):
    items, pdfs = reports(tmp_path)
    assert handoff.offer(tmp_path, items[0], pdfs)
    assert handoff.offer(tmp_path, items[1], pdfs)
    handoff._previous['payloads'] = (b'corrupt', b'corrupt')
    assert handoff.take(tmp_path, items[0]) is None
    assert handoff.take(tmp_path, items[1])
    assert handoff.offer(tmp_path, items[0], pdfs)
    def unavailable(timer):
        raise RuntimeError('thread unavailable')
    monkeypatch.setattr(handoff.threading.Timer, 'start', unavailable)
    assert not handoff.offer(tmp_path, items[1], pdfs)
    assert handoff.take(tmp_path, items[0])


def test_identity_failure_and_logout_clear_both_groups(tmp_path):
    items, pdfs = reports(tmp_path)
    for item in items[:2]:
        assert handoff.offer(tmp_path, item, pdfs)
    assert handoff.take(tmp_path, {**items[0], 'runtime_sha256': 'changed'}) is None
    assert not handoff._entries()
    for item in items[:2]:
        assert handoff.offer(tmp_path, item, pdfs)
    handoff.clear()
    assert not handoff._entries()
