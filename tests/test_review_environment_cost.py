from copy import deepcopy
import json
from pathlib import Path
from types import SimpleNamespace

from tests.test_range_delivery import prepared


def test_page_shares_only_current_request_observation_and_checks_again_after_render(tmp_path, monkeypatch):
    review, writer, editor, source, receipt, state = prepared(tmp_path, monkeypatch)
    identity = [{'complete': True, 'sha256': 'same'}]
    calls, cleanup, change_during_page = [], [], [False]
    def environment(runtime):
        calls.append(True)
        return deepcopy(identity[0])
    monkeypatch.setattr('office_agent.fidelity.runtime_path', lambda: Path('owned-runtime'))
    monkeypatch.setattr('office_agent.render_identity.environment_identity', environment)
    audit_path = editor.audit / receipt / 'audit.json'
    audit = json.loads(audit_path.read_text(encoding='utf8'))
    audit['render_diff']['render_environment'] = deepcopy(identity[0])
    audit_path.write_text(json.dumps(audit, ensure_ascii=False), encoding='utf8')
    identifier = review.start(receipt, [])['result']['review_id']
    opened = review.open(identifier, True)
    assert opened['ok'], opened
    review.cache = {'key': [identifier, opened['result']['pair']], 'pdfs': [Path('left.pdf'), Path('right.pdf')],
                    'hashes': ['a', 'b'], 'projections': ['c', 'd'], 'processes': [],
                    'timer': SimpleNamespace(cancel=lambda: cleanup.append('timer')),
                    'directory': SimpleNamespace(cleanup=lambda: cleanup.append('directory'))}
    def page(*args):
        if change_during_page[0]:
            identity[0] = {'complete': True, 'sha256': 'changed'}
        return {'page': 1, 'image': 'owned-fixture'}
    monkeypatch.setattr('office_agent.research_page.render_page', page)
    calls.clear()
    assert review.page(identifier, 0, 'before')['ok']
    assert len(calls) == 2
    change_during_page[0] = True
    calls.clear()
    assert review.page(identifier, 0, 'after')['reason'] == 'professional_review_environment_changed'
    assert len(calls) == 2 and review.cache is None
    assert cleanup == ['timer', 'directory']
    review.invalidate()
