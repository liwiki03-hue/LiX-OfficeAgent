import json
from pathlib import Path

from docx import Document
import pytest

from office_agent.structured_edits import RECEIPT_VERSION
from office_agent.webui.citations import Citations
from office_agent.webui.structured import Structured
from office_agent.writes import digest
from test_grounded import setup_service


def saved_fixture(tmp_path, monkeypatch, real_render=False):
    grounded, source, record, auth, busy = setup_service(tmp_path, monkeypatch, real_render=real_render)
    doc = Document()
    doc.add_paragraph('Experiment overview.')
    doc.add_paragraph('References')
    doc.save(source)
    capture = grounded.capture(record, 'paper.pdf')['result']
    selected = [source]
    scope = lambda: (auth[0], busy[0], (selected[0], tmp_path))
    grounded._scope = scope
    grounded._selected = lambda: selected[0]
    def services():
        citations = Citations(grounded._lock, scope, grounded._editor, lambda: selected[0], grounded)
        structured = Structured(grounded._lock, scope, grounded._editor, lambda: selected[0])
        structured.grounded, structured.citations = grounded, citations
        return structured, citations
    service, citations = services()
    rows = service.read_structure()['result']['anchors']
    payload = {'source_sha256': digest(source), 'insertions': [{'anchor': rows[0]['locator_id'], 'side': 'after',
               'text': 'The measured latency was 20 ms.', 'attributions': [{'capture_id': capture['capture_id'],
               'evidence_ids': [capture['fragments'][0]['id']], 'relation': 'supports', 'applicability': 'Laboratory only',
               'limitation': 'Developer fixture', 'assessment': 'user_reviewed', 'identity_reviewed': True}]}],
               'bibliography_anchor': rows[1]['locator_id']}
    plan = service.plan_structure(payload)
    assert plan['ok'], plan
    result = service.confirm(plan['result']['id'], True)
    assert result['ok'] and result['result']['status'] == 'saved', result
    selected[0] = Path(result['result']['output'])
    return selected, services, service, plan['result']['id'], source


def test_restart_restore_actual_inserted_positions_without_approval(tmp_path, monkeypatch):
    selected, services, first, receipt, source = saved_fixture(tmp_path, monkeypatch)
    original = digest(source)
    service, citations = services()
    result = citations.read(receipt)
    assert result['ok'], result
    data = result['result']
    assert data['manifest']['version'] == RECEIPT_VERSION
    assert data['manifest']['slots'] == ['body:3/paragraph']
    assert data['citation_numbers'] == [1] and not data['receipt_restores_write_approval']
    assert not service._pending and not service._saved
    rows = service.read_structure()['result']['anchors']
    payload = {'source_sha256': digest(selected[0]), 'insertions': [{'anchor': rows[0]['locator_id'], 'side': 'before', 'text': 'Additional context.'}]}
    plan = service.plan_structure(payload)
    assert plan['ok'], plan
    result = service.confirm(plan['result']['id'], True)
    assert result['result']['status'] == 'saved', result
    audit = json.loads((tmp_path / '.office_agent' / plan['result']['id'] / 'audit.json').read_bytes())
    assert audit['status'] == 'verified' and digest(source) == original
    selected[0] = Path(result['result']['output'])
    again, restored = services()
    second = restored.read(plan['result']['id'])
    assert second['ok'], second
    assert second['result']['manifest']['slots'] == ['body:4/paragraph']
    assert second['result']['manifest']['parent_receipt'] == receipt
    selected[0] = Path(audit['source']) if 'source' in audit else selected[0]
    assert service.rollback(plan['result']['id'], True)['result']['status'] == 'rolled_back'
    assert source.exists()


@pytest.mark.parametrize('failure', ['external_output', 'rolled_back', 'different_file', 'bad_manifest'])
def test_receipt_rejects_changed_output_or_audit(tmp_path, monkeypatch, failure):
    selected, services, first, receipt, source = saved_fixture(tmp_path, monkeypatch)
    if failure == 'external_output':
        doc = Document(selected[0])
        doc.paragraphs[0].text = 'External edit'
        doc.save(selected[0])
    elif failure == 'rolled_back':
        assert first._editor().rollback(receipt, confirmed=True)['status'] == 'rolled_back'
        selected[0] = source
    elif failure == 'different_file':
        selected[0] = source
    else:
        path = tmp_path / '.office_agent' / receipt / 'audit.json'
        data = json.loads(path.read_bytes())
        after = json.loads(data['after'])
        manifest = next(r['citation_manifest'] for r in after['items'] if 'citation_manifest' in r)
        manifest['slots'] = ['body:1/paragraph']
        data['after'] = json.dumps(after)
        path.write_text(json.dumps(data), encoding='utf8')
    service, citations = services()
    assert citations.read(receipt)['reason'] == 'professional_citation_receipt'
    assert not citations._restored and not service._pending


def test_parent_receipt_changed_after_plan_is_refused(tmp_path, monkeypatch):
    selected, services, first, receipt, source = saved_fixture(tmp_path, monkeypatch)
    service, citations = services()
    assert citations.read(receipt)['ok']
    row = service.read_structure()['result']['anchors'][0]
    plan = service.plan_structure({'source_sha256': digest(selected[0]), 'insertions': [{'anchor': row['locator_id'], 'side': 'before', 'text': 'Additional context.'}]})
    assert plan['ok'], plan
    path = tmp_path / '.office_agent' / receipt / 'audit.json'
    path.write_bytes(path.read_bytes() + b'\n')
    reply = service.confirm(plan['result']['id'], True)
    assert not reply['ok'], reply
    assert len(list(tmp_path.glob('*.agent-*.docx'))) == 1


def test_restored_mapping_new_evidence_and_bibliography_remain_one_group(tmp_path, monkeypatch):
    selected, services, first, receipt, source = saved_fixture(tmp_path, monkeypatch)
    service, citations = services()
    restored = citations.read(receipt)['result']['manifest']
    record = restored['catalog'][restored['number_to_key']['1']]['id']
    capture = service.grounded.capture(record, 'paper.pdf')['result']
    rows = service.read_structure()['result']['anchors']
    payload = {'source_sha256': digest(selected[0]), 'insertions': [{'anchor': rows[0]['locator_id'], 'side': 'before',
               'text': 'The measured latency was 20 ms.', 'attributions': [{'capture_id': capture['capture_id'],
               'evidence_ids': [capture['fragments'][0]['id']], 'relation': 'supports', 'applicability': 'Laboratory only',
               'limitation': 'Developer fixture', 'assessment': 'user_reviewed', 'identity_reviewed': True}]}],
               'number_records': {'1': record}, 'bibliography': [rows[3]['locator_id']], 'bibliography_anchor': rows[2]['locator_id']}
    payload['insertions'][0]['attributions'][0]['identity_reviewed'] = False
    assert service.plan_structure(payload)['reason'] == 'professional_identity_required'
    payload['insertions'][0]['attributions'][0]['identity_reviewed'] = True
    planned = service.plan_structure(payload)
    assert planned['ok'], planned
    saved = service.confirm(planned['result']['id'], True)
    assert saved['result']['status'] == 'saved', saved
    text = [p.text for p in Document(saved['result']['output']).paragraphs]
    assert text[0].endswith('[1]') and text[2].endswith('[1]')
    assert sum('10.1234/expt.01' in p for p in text) == 1
    assert service.rollback(planned['result']['id'], True)['result']['status'] == 'rolled_back'
