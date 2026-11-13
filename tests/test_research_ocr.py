from copy import deepcopy

from office_agent.grounded_edits import validate_proof, GroundedDocx, seal
from office_agent.provenance import ClaimSupport
from office_agent.writes import GuardError
import pytest

from test_grounded import setup_service


def test_ocr_is_explicit_raw_and_human_correction_remain_separate(tmp_path, monkeypatch):
    service, path, record, auth, busy = setup_service(tmp_path, monkeypatch)
    captured = service.capture(record, 'paper.pdf')['result']
    calls = []

    def recognize(source, fingerprint, page, box, **options):
        calls.append(options)
        assert options == {'ocr': True}
        return {'raw_text': 'Latency 2O ms.', 'lines': [{'text': 'Latency 2O ms.', 'box_pdf': [10., 10., 300., 30.]}],
                'engine': 'synthetic-ocr-test', 'warning': 'Unverified numeric text'}

    monkeypatch.setattr('office_agent.research_page.render_page', recognize)
    assert not calls
    raw = service.ocr(captured['capture_id'], 1)
    assert raw['ok'], raw
    assert len(calls) == 1 and not raw['result']['ocr']['reviewed']
    raw = raw['result']
    assert raw['fragments'][0]['extraction_quality'] == 'ocr_unverified'
    assert not service.ocr(captured['capture_id'], 2)['ok'] and len(calls) == 1
    payload = {'capture_id': raw['capture_id'], 'corrections': [{'evidence_id': raw['fragments'][0]['id'], 'text': 'Latency 20 ms.'}], 'reviewed': False}
    assert service.review_ocr(payload)['reason'] == 'professional_ocr_review_required'
    payload['reviewed'] = True
    corrected = service.review_ocr(payload)
    assert corrected['ok'], corrected
    corrected = corrected['result']
    assert corrected['ocr']['raw']['raw_text'] == 'Latency 2O ms.'
    assert corrected['fragments'][0]['original_text'] == 'Latency 20 ms.'
    assert corrected['fragments'][0]['id'] != raw['fragments'][0]['id']
    assert corrected['fragments'][0]['extraction_quality'] == 'ocr_unverified'
    assert service._captures[raw['capture_id']]['proof']['ocr']['reviewed'] is False


def test_missing_ocr_stays_a_read_only_failure(tmp_path, monkeypatch):
    service, path, record, auth, busy = setup_service(tmp_path, monkeypatch)
    capture = service.capture(record, 'paper.pdf')['result']

    def missing(*args, **kwargs):
        raise ImportError('no-local-ocr')

    monkeypatch.setattr('office_agent.research_page.render_page', missing)
    assert service.ocr(capture['capture_id'], 1)['reason'] == 'professional_ocr_unavailable'
    assert len(service._captures) == 1 and len(list(tmp_path.glob('*.docx'))) == 1
