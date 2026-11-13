import json
import threading

import pytest

from office_agent.review_diagnostics import initialize_review_diagnostics, review_phase, review_request, review_stage


class Review:
    def __init__(self): initialize_review_diagnostics(self)
    @review_stage('identity')
    def check(self, failure=False):
        if failure: raise ValueError('sensitive-path secret-value')
    @review_request
    def page(self, failure=False, blocked=False):
        self.check(failure)
        with review_phase('render'):
            return {'ok': not blocked}


def test_diagnostics_bounded_redacted_and_nested_failure():
    service = Review()
    with pytest.raises(ValueError): service.page(True)
    row = service.diagnostics[-1]
    assert row['status'] == 'failed' and row['stages'][0]['stage'] == 'identity'
    assert row['stages'][0]['error_type'] == 'ValueError'
    assert 'secret-value' not in json.dumps(row) and 'sensitive-path' not in json.dumps(row)
    for _ in range(20): service.page(blocked=True)
    assert len(service.diagnostics) == 16 and all(row['status'] == 'blocked' for row in service.diagnostics)
    assert all(row['milliseconds'] >= 0 for row in service.diagnostics)


def test_diagnostics_are_isolated_between_threads():
    service = Review()
    workers = [threading.Thread(target=service.page) for _ in range(4)]
    for worker in workers: worker.start()
    for worker in workers: worker.join(2); assert not worker.is_alive()
    assert len(service.diagnostics) == 4 and len({row['id'] for row in service.diagnostics}) == 4
    assert all([stage['stage'] for stage in row['stages']] == ['identity', 'render'] for row in service.diagnostics)
