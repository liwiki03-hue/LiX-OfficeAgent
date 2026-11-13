from copy import deepcopy
from pathlib import Path

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.professional_edits import W
from office_agent.writes import _apply_revision_decision, _word_state, digest
from test_professional_edits import parts
from test_research_assist import setup, request, complete


def planned(service):
    task = complete(service, request(service))
    proposal = task['result']['suggestions'][0]
    payload = {'task_id': task['id'], 'tracked': True, 'edits': [{'locator_id': proposal['locator_id'], 'new_text': proposal['new_text'], 'reason': proposal['reason'], 'reviewed': True}]}
    result = service.plan(payload)
    return payload, result


def test_native_revisions_accept_clean_reject_original_and_copy_undo(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    original = digest(path)
    payload, planned_reply = planned(service)
    assert planned_reply['ok'], planned_reply
    plan = planned_reply['result']
    assert plan['items'][0]['native_revisions']['accept_all_matches_approved']
    saved = service.confirm(plan['id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    output = Path(saved['result']['output'])
    before, after = parts(path), parts(output)
    assert all(value == after[name] for name, value in before.items() if name != 'word/document.xml')
    root = ET.fromstring(after['word/document.xml'])
    revisions = list(root.iter(W + 'ins')) + list(root.iter(W + 'del'))
    assert len(revisions) == 2 and len({row.get(W + 'id') for row in revisions}) == 2
    assert all(row.get(W + 'author') == 'Office Agent' and row.get(W + 'date').endswith('Z') for row in revisions)
    accepted = Document(output)
    rejected = Document(output)
    _apply_revision_decision(accepted, 'accept_all')
    _apply_revision_decision(rejected, 'reject_all')
    assert accepted.paragraphs[0].text == payload['edits'][0]['new_text']
    assert _word_state(rejected) == _word_state(Document(path))
    clean_plan = service.plan({**payload, 'tracked': False})
    assert clean_plan['ok'], clean_plan
    clean_saved = service.confirm(clean_plan['result']['id'], True)
    assert clean_saved['result']['status'] == 'saved'
    assert _word_state(accepted) == _word_state(Document(clean_saved['result']['output']))
    assert digest(path) == original
    assert service.rollback(plan['id'], True)['result']['status'] == 'rolled_back'


@pytest.mark.parametrize('part', ['body', 'header'])
def test_prior_revisions_anywhere_refuse_without_writing(tmp_path, monkeypatch, part):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    document = Document(path)
    paragraph = document.paragraphs[-1] if part == 'body' else document.sections[0].header.paragraphs[0]
    insertion = ET.SubElement(paragraph._p, W + 'ins', {W + 'id': '12', W + 'author': 'Existing', W + 'date': '2024-01-01T00:00:00Z'})
    run = ET.SubElement(insertion, W + 'r')
    ET.SubElement(run, W + 't').text = 'Existing revision'
    document.save(path)
    payload, reply = planned(service)
    assert reply['reason'] == 'professional_existing_revisions'
    assert len(list(tmp_path.glob('*.docx'))) == 1
