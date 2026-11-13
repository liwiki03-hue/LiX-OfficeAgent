from pathlib import Path
import hashlib
import threading

from docx import Document
import pytest

from office_agent.webui.citations import Citations
from office_agent.webui.grounded import Grounded
from office_agent.webui.literature import Literature
from office_agent.professional_edits import snapshot
from office_agent.writes import digest
from test_citation_lifecycle import fixture


def long_fixture(tmp_path, monkeypatch):
    service, grounded, path, original = fixture(tmp_path, monkeypatch)
    doc = Document()
    doc.add_paragraph('The measured latency was 20 ms [4].')
    for index in range(79):
        doc.add_paragraph('Unchanged technical discussion ' + str(index) + '.')
    doc.add_paragraph('Independent result [2], repeated [7].')
    doc.add_paragraph('[2] First experiment. doi: 10.1234/expt.01')
    doc.add_paragraph('[4] Second experiment. doi: 10.1234/expt.02')
    doc.add_paragraph('[7] Duplicate first experiment. doi: 10.1234/expt.01')
    doc.add_paragraph('Reference placeholder')
    doc.save(path)
    reply = service.read()
    assert reply['ok'], reply
    first = reply['result']
    last = service.read('', 80, first['document']['source_sha256'])['result']
    items = [first['items'][0], *last['items']]
    payload = {**original, 'source_sha256': digest(path), 'authorized': [row['locator_id'] for row in items],
               'bibliography': [row['locator_id'] for row in last['items'][1:]]}
    return service, grounded, path, payload, first, last


def test_long_doc_cross_page_transaction_resume_and_rollback(tmp_path, monkeypatch):
    service, grounded, path, payload, first, last = long_fixture(tmp_path, monkeypatch)
    assert first['total_targets'] == 85 and len(first['items']) == 40 and first['next_offset'] == 40
    assert first['citation_numbers'] == [2, 4, 7] and first['citation_numbers_scope'] == 'body_inventory'
    assert last['offset'] == 80 and last['next_offset'] is None
    original = digest(path)
    plan = service.plan(payload)
    assert plan['ok'], plan
    assert service.confirm(plan['result']['id'], False)['result']['status'] == 'cancelled'
    assert digest(path) == original and len(list(tmp_path.glob('*.docx'))) == 1
    plan = service.plan(payload)['result']
    saved = service.confirm(plan['id'], True)['result']
    assert saved['status'] == 'saved'
    output = Path(saved['output'])
    paragraphs = Document(output).paragraphs
    assert paragraphs[0].text == 'The measured latency was 20 ms [1].'
    assert paragraphs[80].text == 'Independent result [2], repeated [2].'
    assert [row.text for row in paragraphs[1:80]] == [row.text for row in Document(path).paragraphs[1:80]]
    assert '10.1234/expt.02' in paragraphs[81].text and '10.1234/expt.01' in paragraphs[82].text
    editor = grounded._editor()
    scope = lambda: (True, False, (output, editor.space.root))
    lock = threading.RLock()
    literature = Literature(lock, scope, lambda: editor, tmp_path)
    resumed = Citations(lock, scope, lambda: editor, lambda: output, Grounded(lock, scope, lambda: editor, lambda: output, literature))
    restored = resumed.read(plan['id'])
    assert restored['ok'] and restored['result']['manifest']['number_to_key'] == plan['manifest']['number_to_key'], restored
    assert not restored['result']['receipt_restores_write_approval']
    tail = resumed.read(plan['id'], 80, digest(output))
    assert tail['ok'] and tail['result']['manifest'] == restored['result']['manifest']
    rows = tail['result']['items']
    metadata = restored['result']['manifest']['catalog']
    descendant_payload = {'source_sha256': digest(output), 'authorized': [row['locator_id'] for row in rows],
                          'bibliography': [row['locator_id'] for row in rows[1:]],
                          'number_records': {number: metadata[key]['id'] for number, key in plan['manifest']['number_to_key'].items()},
                          'removals': {rows[0]['locator_id']: [2]}, 'mapping_reviewed': True}
    descendant = resumed.plan(descendant_payload)
    assert descendant['ok'], descendant
    next_saved = resumed.confirm(descendant['result']['id'], True)['result']
    assert next_saved['status'] == 'saved' and Document(next_saved['output']).paragraphs[80].text == 'Independent result , repeated .'
    assert resumed.rollback(descendant['result']['id'], True)['result']['status'] == 'rolled_back'
    assert output.exists() and service.rollback(plan['id'], True)['result']['status'] == 'rolled_back'
    assert not output.exists() and digest(path) == original


def test_version_change_between_pages_and_before_plan_refuses_without_writes(tmp_path, monkeypatch):
    service, grounded, path, payload, first, last = long_fixture(tmp_path, monkeypatch)
    doc = Document(path)
    doc.paragraphs[40].text = 'Externally revised discussion.'
    doc.save(path)
    assert service.read('', 40, payload['source_sha256'])['reason'] == 'professional_source_changed'
    assert service.plan(payload)['reason'] == 'professional_source_changed'
    assert not grounded._editor().plans and not grounded._editor().audit.exists()


@pytest.mark.parametrize('offset,version', [(True, ''), (-1, '0' * 64), (4001, '0' * 64), (1, ''), (0, 'invalid'), (0, None)])
def test_invalid_pagination_is_rejected(tmp_path, monkeypatch, offset, version):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    assert service.read('', offset, version)['reason'] == 'professional_invalid_page'
    assert not grounded._editor().audit.exists()


def test_text_budget_advances_by_actual_returned_rows(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    doc = Document()
    for index in range(10):
        text = ' '.join(hashlib.sha256(f'{index}:{part}'.encode()).hexdigest() for part in range(124))
        doc.add_paragraph(text[:7999])
    doc.save(path)
    assert snapshot(grounded._editor(), path, limit=40)['next_offset'] == 3
    reply = service.read()
    assert reply['ok'], reply
    first = reply['result']
    assert len(first['items']) == 3 and first['next_offset'] == 3
    next_page = service.read('', first['next_offset'], first['document']['source_sha256'])['result']
    assert next_page['offset'] == 3 and next_page['next_offset'] == 6
    assert set(row['locator_id'] for row in first['items']).isdisjoint(row['locator_id'] for row in next_page['items'])


def test_cross_page_selection_does_not_bypass_twelve_target_limit(tmp_path, monkeypatch):
    service, grounded, path, payload, first, last = long_fixture(tmp_path, monkeypatch)
    payload['authorized'].extend(row['locator_id'] for row in first['items'][1:9])
    assert len(payload['authorized']) == 14
    assert service.plan(payload)['reason'] == 'professional_invalid_request'
    assert not grounded._editor().plans and not grounded._editor().audit.exists()


def test_omitted_far_reference_still_refuses_whole_group(tmp_path, monkeypatch):
    service, grounded, path, payload, first, last = long_fixture(tmp_path, monkeypatch)
    payload['authorized'].remove(last['items'][0]['locator_id'])
    assert service.plan(payload)['reason'] == 'professional_citation_scope'
    assert not grounded._editor().plans and not grounded._editor().audit.exists()
