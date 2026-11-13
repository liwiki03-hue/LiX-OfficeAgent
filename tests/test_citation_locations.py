import hashlib

from docx import Document
import pytest

from office_agent.writes import digest
from test_citation_pagination import long_fixture
from test_citation_lifecycle import fixture


def test_locations_find_far_repeated_clusters_and_skip_formula(tmp_path, monkeypatch):
    service, grounded, path, payload, first, last = long_fixture(tmp_path, monkeypatch)
    doc = Document(path)
    doc.paragraphs[80].text = 'Result [2,7], repeated [7]; formula $v=[7]$.'
    doc.save(path)
    before = digest(path)
    reply = service.locations(7, before)
    assert reply['ok'], reply
    result = reply['result']
    assert [row['index'] for row in result['items']] == [80, 83]
    assert result['total_locations'] == 2 and result['total_occurrences'] == 3
    assert result['items'][1]['kind'] == 'bibliography_candidate'
    assert result['items'][0]['eligible_not_authorized']
    assert result['read_only'] and not result['write_approval_granted']
    assert result['impact_analysis'] == 'occurrences_only_not_complete_renumbering_impact'
    assert not service._pending and not grounded._editor().audit.exists() and digest(path) == before


def test_locations_page_counts_and_empty_result(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    doc = Document()
    for index in range(85):
        doc.add_paragraph('Technical observation [4].')
    doc.save(path)
    version = digest(path)
    first = service.locations(4, version)['result']
    last = service.locations(4, version, 80)['result']
    assert first['total_locations'] == 85 and first['next_offset'] == 40 and len(first['items']) == 40
    assert [row['index'] for row in last['items']] == list(range(80, 85)) and last['next_offset'] is None
    empty = service.locations(999, version)['result']
    assert empty['total_locations'] == 0 and empty['items'] == [] and empty['next_offset'] is None
    assert service.locations(4, version, 85)['reason'] == 'professional_invalid_page'


def test_readonly_explanations_and_redaction_do_not_enable_writes(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    doc = Document()
    doc.add_heading('Styled reference [4]', 1)
    doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Table reference [4]'
    special = doc.add_paragraph('Special reference [4]')
    special.add_run().add_break()
    secret = 'api_key=' + 'sk-' + hashlib.sha256(b'synthetic').hexdigest()
    doc.add_paragraph(secret + ' [4]')
    doc.save(path)
    reply = service.locations(4, digest(path))
    assert reply['ok'], reply
    result = reply['result']
    assert [row['citation_refusal'] for row in result['items']] == ['explicit_paragraph_style', 'not_body_paragraph', 'field_image_note_break_or_special_run']
    assert all(not row['eligible_not_authorized'] for row in result['items'])
    assert result['omitted_targets'] == 1 and secret not in str(reply)
    source = service.read()['result']
    assert source['items'][0]['citation_refusal'] == 'explicit_paragraph_style'
    assert source['items'][-1]['text'] == '' and source['items'][-1]['citation_refusal'] == 'sensitive_or_oversize'
    assert not grounded._editor().audit.exists()


@pytest.mark.parametrize('number,version,offset', [(True, '0' * 64, 0), (0, '0' * 64, 0), (1000, '0' * 64, 0),
    ('4', '0' * 64, 0), (4, '', 0), (4, '0' * 64, True), (4, '0' * 64, -1)])
def test_locations_strict_inputs(tmp_path, monkeypatch, number, version, offset):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    assert service.locations(number, version, offset)['reason'] == 'professional_invalid_request'


def test_changed_source_and_auth_busy_guards(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    assert service.locations(4, '0' * 64)['reason'] == 'professional_source_changed'
    original_scope = service._scope
    service._scope = lambda: (True, True, original_scope()[2])
    assert service.locations(4, digest(path))['reason'] == 'busy'
    service._scope = lambda: (False, False, None)
    assert service.locations(4, digest(path))['reason'] == 'not_authenticated'
    assert not grounded._editor().audit.exists()
