from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace
import threading

from docx import Document
import pytest

from office_agent.webui.research_assist import ResearchAssist
from office_agent.writes import digest
from test_grounded import setup_service


class Model:
    identity = {'provider': 'local-fixture', 'model': 'synthetic-test-only', 'tool_access': False}

    def __init__(self, mutate=lambda value: value):
        self.inputs = []
        self.mutate = mutate

    def suggest(self, payload, cancelled):
        self.inputs.append(deepcopy(payload))
        row = payload['targets'][0]
        return self.mutate({'cards': [], 'suggestions': [{'locator_id': row['locator_id'],
                          'new_text': row['text'] + ' The observation is reported.', 'reason': 'Clarify the existing statement.',
                          'evidence_ids': [], 'risks': [], 'needs_evidence': False}],
                          'missing_evidence': [], 'stop_reason': 'bounded_review_complete'})


def setup(tmp_path, monkeypatch, model=None):
    grounded, path, record_id, auth, busy = setup_service(tmp_path, monkeypatch)
    model = model or Model()
    rag = SimpleNamespace(_research_dependencies={})
    service = ResearchAssist(grounded._lock, grounded._scope, grounded._editor, lambda: path, grounded, rag, lambda: model)
    return service, grounded, path, record_id, auth, busy, model


def request(service, **changes):
    source = service.read()['result']
    return {'question': 'Clarify the original observation.', 'mode': 'polish',
            'source_sha256': source['document']['source_sha256'], 'locator_ids': [source['items'][0]['locator_id']],
            'capture_ids': [], 'evidence_save_id': '', 'protected_terms': [], 'outbound_approved': True, **changes}


def complete(service, payload):
    reply = service.start(payload)
    assert reply['ok'], reply
    task_id = reply['result']['id']
    service.jobs[task_id]['thread'].join(5)
    assert not service.jobs[task_id]['thread'].is_alive()
    reply = service.status(task_id)
    assert reply['ok'], reply
    return reply['result']


def test_unreviewed_model_to_human_plan_cancel_copy_and_undo(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    original = digest(path)
    assert service.read()['ok'] and not model.inputs
    task = complete(service, request(service))
    assert task['status'] == 'completed' and task['calls'] == 1
    proposal = task['result']['suggestions'][0]
    assert proposal['assessment'] == 'unverified' and proposal['protection_status'] == 'passed'
    edits = [{'locator_id': proposal['locator_id'], 'new_text': proposal['new_text'], 'reason': proposal['reason'], 'reviewed': False}]
    assert not service.plan({'task_id': task['id'], 'edits': edits})['ok']
    edits[0]['reviewed'] = True
    plan = service.plan({'task_id': task['id'], 'edits': edits})
    assert plan['ok'], plan
    assert service.confirm(plan['result']['id'], False)['result']['status'] == 'cancelled'
    assert digest(path) == original and len(list(tmp_path.glob('*.docx'))) == 1
    plan = service.plan({'task_id': task['id'], 'edits': edits})
    saved = service.confirm(plan['result']['id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    output = Path(saved['result']['output'])
    assert Document(output).paragraphs[0].text == proposal['new_text']
    assert digest(path) == original
    assert service.rollback(saved['result']['save_id'], True)['result']['status'] == 'rolled_back'
    assert not output.exists()


@pytest.mark.parametrize('mutation', [
    lambda value: {**value, 'assessment': 'user_reviewed'},
    lambda value: {**value, 'proof': {'approved': True}},
    lambda value: {**value, 'cards': [{'claim': 'Fabricated', 'evidence_ids': ['unknown'], 'relation': 'supports', 'applicability': '', 'limitation': '', 'conflicts': []}]},
    lambda value: {**value, 'suggestions': [{**value['suggestions'][0], 'locator_id': 'loc:' + 'f' * 64}]},
])
def test_model_cannot_supply_authority_or_unknown_evidence(tmp_path, monkeypatch, mutation):
    service, *rest = setup(tmp_path, monkeypatch, Model(mutation))
    task = complete(service, request(service))
    assert task['status'] == 'failed' and task['result'] is None
    assert len(list(tmp_path.glob('*.docx'))) == 1


def test_insufficient_evidence_stops_before_model_and_explicit_approval(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    assert not service.start(request(service, outbound_approved=False))['ok']
    task = complete(service, request(service, mode='supplement'))
    assert task['stop_reason'] == 'insufficient_materials' and task['calls'] == 0 and not model.inputs
    assert service.plan({'task_id': task['id'], 'edits': []})['reason'] == 'professional_assist_citation_review_required'


def test_duplicate_cancel_and_late_model_response_are_not_republished(tmp_path, monkeypatch):
    entered, released = threading.Event(), threading.Event()
    model = Model()
    original = model.suggest

    def blocked(payload, cancelled):
        entered.set()
        assert released.wait(5)
        return original(payload, cancelled)

    model.suggest = blocked
    service, grounded, path, record, auth, busy, _ = setup(tmp_path, monkeypatch, model)
    payload = request(service)
    task = service.start(payload)['result']
    try:
        assert entered.wait(3)
        assert not service.start(payload)['ok']
        assert service.cancel(task['id'])['result']['status'] == 'cancelled'
        assert service.start(payload)['reason'] == 'busy'
    finally:
        released.set()
        service.jobs[task['id']]['thread'].join(5)
    state = service.status(task['id'])['result']
    assert state['result'] is None and state['status'] == 'cancelled'
    service.invalidate()
    assert not service.status(task['id'])['ok']


def test_evidence_and_source_changes_invalidate_suggestions(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    capture = grounded.capture(record, 'paper.pdf')['result']
    task = complete(service, request(service, capture_ids=[capture['capture_id']]))
    assert model.inputs[0]['evidence'] and task['evidence']
    (tmp_path / 'paper.pdf').write_bytes(b'changed')
    assert service.status(task['id'])['reason'] == 'professional_evidence_changed'
    auth[0] = False
    assert service.status(task['id'])['reason'] == 'not_authenticated'


def test_regeneration_is_per_item_and_budgeted_without_review_authority(tmp_path, monkeypatch):
    model = Model()
    original = model.suggest

    def all_targets(payload, cancelled):
        result = original(payload, cancelled)
        template = result['suggestions'][0]
        result['suggestions'] = [{**template, 'locator_id': row['locator_id'], 'new_text': row['text'] + ' Clarified.'} for row in payload['targets']]
        return result

    model.suggest = all_targets
    service, *rest = setup(tmp_path, monkeypatch, model)
    source = service.read()['result']
    task = complete(service, request(service, locator_ids=[row['locator_id'] for row in source['items'][:2]]))
    other = deepcopy(task['result']['suggestions'][1])
    for call in (2, 3):
        reply = service.regenerate({'task_id': task['id'], 'locator_id': task['targets'][0]['locator_id'], 'feedback': 'Keep the original scope.', 'outbound_approved': True})
        assert reply['ok'], reply
        service.jobs[task['id']]['thread'].join(5)
        task = service.status(task['id'])['result']
        assert task['calls'] == call and task['result']['suggestions'][1] == other
        assert task['result']['suggestions'][0]['assessment'] == 'unverified'
    assert service.regenerate({'task_id': task['id'], 'locator_id': task['targets'][0]['locator_id'], 'feedback': '', 'outbound_approved': True})['reason'] == 'professional_assist_budget_exhausted'
    assert len(model.inputs) == 3 and all(len(row['targets']) == 1 for row in model.inputs[1:])


def test_bounded_planning_reads_only_known_context_and_preserves_disagreement(tmp_path, monkeypatch):
    class Planner(Model):
        def plan(self, payload, cancelled):
            return {'queries': ['latency conditions', '适用条件与限制'], 'read_ids': [payload['catalog'][0]['id']]}

        def suggest(self, payload, cancelled):
            response = super().suggest(payload, cancelled)
            evidence = payload['evidence'][0]
            response['cards'] = [{'claim': 'The evidence has limited applicability.', 'evidence_ids': [evidence['id']], 'relation': relation,
                                  'applicability': 'The original experimental conditions.', 'limitation': 'Independent corroboration required.', 'conflicts': ['Disagreement retained.']}
                                 for relation in ('supports', 'contradicts')]
            return response

    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch, Planner())
    capture = grounded.capture(record, 'paper.pdf')['result']
    task = complete(service, request(service, capture_ids=[capture['capture_id']]))
    assert task['status'] == 'completed' and task['calls'] == 2 and task['rounds'] == 2
    assert task['retrieval']['queries'] == ['latency conditions', '适用条件与限制']
    assert task['retrieval']['published_index'] is False and task['retrieval']['quality'] == 'lexical_only_no_vector_model'
    assert {row['relation'] for row in task['result']['cards']} == {'supports', 'contradicts'}
    assert all(row['assessment'] == 'unverified' for row in task['result']['cards'])
    assert all(key in {row['id'] for row in capture['fragments']} for key in task['evidence'])


def test_planner_cannot_request_unknown_sources(tmp_path, monkeypatch):
    class Planner(Model):
        def plan(self, payload, cancelled):
            return {'queries': ['fixture'], 'read_ids': ['../private.pdf']}

    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch, Planner())
    capture = grounded.capture(record, 'paper.pdf')['result']
    task = complete(service, request(service, capture_ids=[capture['capture_id']]))
    assert task['status'] == 'failed' and task['calls'] == 1 and not model.inputs


def test_excluding_a_source_consumes_existing_plan_before_write(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    capture = grounded.capture(record, 'paper.pdf')['result']
    task = complete(service, request(service, capture_ids=[capture['capture_id']]))
    proposal = task['result']['suggestions'][0]
    planned = service.plan({'task_id': task['id'], 'edits': [{'locator_id': proposal['locator_id'], 'new_text': proposal['new_text'], 'reason': proposal['reason'], 'reviewed': True}]})
    assert planned['ok'], planned
    grounded.literature.decisions[record]['decision'] = 'exclude'
    assert service.confirm(planned['result']['id'], True)['reason'] == 'professional_evidence_invalid'
    assert planned['result']['id'] not in service._editor().plans and len(list(tmp_path.glob('*.docx'))) == 1


def test_cancel_cannot_hide_uncertain_inference_cleanup(tmp_path, monkeypatch):
    class Planner(Model):
        def plan(self, payload, cancelled):
            return {'queries': ['fixture'], 'read_ids': []}

    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch, Planner())
    capture = grounded.capture(record, 'paper.pdf')['result']

    def uncertain(*args, **kwargs):
        service.cancel(next(iter(service.jobs)))
        raise ValueError('research_retrieval_cleanup_failed')

    monkeypatch.setattr('office_agent.research_capture_retrieval.retrieve_captures', uncertain)
    payload = request(service, capture_ids=[capture['capture_id']])
    task = complete(service, payload)
    assert task['status'] == 'cancelled' and service._retrieval_disabled
    service.invalidate()
    assert service.start(payload)['reason'] == 'professional_assist_retrieval_disabled'


def test_user_can_lower_hard_call_budget_but_cannot_expand_it(tmp_path, monkeypatch):
    class Planner(Model):
        def plan(self, payload, cancelled):
            raise AssertionError('one-call budget must skip the planning model call')

    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch, Planner())
    capture = grounded.capture(record, 'paper.pdf')['result']
    assert service.start(request(service, max_calls=4))['reason'] == 'professional_invalid_request'
    task = complete(service, request(service, max_calls=1, capture_ids=[capture['capture_id']]))
    assert task['status'] == 'completed' and task['calls'] == 1 and len(model.inputs) == 1
    assert service.regenerate({'task_id': task['id'], 'locator_id': task['targets'][0]['locator_id'], 'feedback': '', 'outbound_approved': True})['reason'] == 'professional_assist_budget_exhausted'
