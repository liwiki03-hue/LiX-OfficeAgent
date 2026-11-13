from copy import deepcopy
import threading

from docx import Document
import pytest

from office_agent.grounded_edits import format_reference, validate_proof
from office_agent.literature import normalize
from office_agent.professional_edits import snapshot
from office_agent.webui.grounded import Grounded
from office_agent.webui.literature import Literature
from office_agent.writes import Editor, GuardError, digest
from test_fidelity import pdf_bytes
from test_literature import crossref
from test_professional_edits import fake_render, parts


def setup_service(tmp_path, monkeypatch, *, real_render=False):
    path = tmp_path / 'draft.docx'
    doc = Document()
    doc.add_paragraph('The experiment measured latency.')
    doc.add_paragraph('Reference placeholder')
    doc.add_paragraph('Unchanged paragraph.')
    doc.save(path)
    (tmp_path / 'paper.pdf').write_bytes(pdf_bytes(text='The measured latency was 20 ms under laboratory conditions.', width=600))
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    authenticated, busy = [True], [False]
    scope = lambda: (authenticated[0], busy[0], (path, editor.space.root))
    lock = threading.RLock()
    literature = Literature(lock, scope, lambda: editor, tmp_path)
    literature._bind()
    record = normalize(crossref(), 'crossref')
    literature.records[record['id']] = record
    literature.decisions[record['id']] = {'decision': 'include', 'reason': 'Synthetic technical experiment'}
    service = Grounded(lock, scope, lambda: editor, lambda: path, literature)
    if not real_render:
        monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    monkeypatch.setenv('TEMP', str(tmp_path))
    monkeypatch.setenv('TMP', str(tmp_path))
    return service, path, record['id'], authenticated, busy


def draft(service, path, record_id):
    capture = service.capture(record_id, 'paper.pdf')
    assert capture['ok'], capture
    structure = snapshot(service._editor(), path)
    return {'source_sha256': structure['document']['source_sha256'], 'capture_id': capture['result']['capture_id'],
            'body_locator': structure['items'][0]['locator_id'], 'reference_locator': structure['items'][1]['locator_id'],
            'identity_reviewed': True,
            'support': {'claim': 'The measured latency was 20 ms.', 'evidence_ids': [capture['result']['fragments'][0]['id']],
                        'relation': 'supports', 'applicability': 'Laboratory conditions only.',
                        'limitation': 'Synthetic source, not a real performance claim.', 'assessment': 'user_reviewed'}}


def test_actual_csl_structured_and_literal_names():
    metadata = {'id': 'fixture', 'type': 'article-journal', 'title': 'Thermal experiment',
                'author': [{'family': 'Lovelace', 'given': 'Ada'}], 'issued': {'date-parts': [[2024]]}, 'DOI': '10.1234/fixture'}
    inline, entry = format_reference(metadata)
    assert inline == '[1]' and entry.startswith('[1] ') and 'A. Lovelace' in entry and '2024' in entry and '10.1234/fixture' in entry
    assert '<i>' not in entry
    metadata['author'] = [{'literal': '研究团队'}]
    assert '研究团队' in format_reference(metadata)[1]


def test_evidence_plan_cancel_save_verify_and_rollback(tmp_path, monkeypatch):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    before, original = parts(path), digest(path)
    payload = draft(service, path, record_id)
    planned = service.plan(payload)
    assert planned['ok'], planned
    item = planned['result']['items'][0]
    assert item['evidence_fragments'][0]['page'] == 1 and item['evidence_fragments'][0]['bbox']
    assert item['claim_support']['assessment'] == 'user_reviewed'
    assert item['semantic_verification'] == 'user_reviewed_not_machine_proven'
    assert service.confirm(planned['result']['id'], False)['result']['status'] == 'cancelled'
    assert not service.confirm(planned['result']['id'], True)['ok']
    planned = service.plan(payload)
    saved = service.confirm(planned['result']['id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    output = saved['result']['output']
    from pathlib import Path
    output = Path(output)
    result = Document(output)
    assert result.paragraphs[0].text == payload['support']['claim'] + ' [1]'
    assert '10.1234/expt.01' in result.paragraphs[1].text
    assert result.paragraphs[2].text == 'Unchanged paragraph.'
    after = parts(output)
    assert all(after[key] == value for key, value in before.items() if key != 'word/document.xml')
    assert digest(path) == original
    assert service.rollback(planned['result']['id'], True)['result']['status'] == 'rolled_back'
    assert not output.exists() and digest(path) == original


@pytest.mark.parametrize('mutation,reason', [
    (lambda p: p.update(identity_reviewed=False), 'professional_identity_required'),
    (lambda p: p['support'].update(assessment='unverified'), 'professional_support_required'),
    (lambda p: p['support'].update(relation='insufficient'), 'professional_support_required'),
    (lambda p: p['support'].update(evidence_ids=['ev:' + '0' * 64]), 'professional_support_required'),
    (lambda p: p['support'].update(claim='The measured latency was 99 ms.'), 'professional_unsupported_fact'),
    (lambda p: p['support'].update(claim='The measured latency was 20 ms [2].'), 'professional_citation_invalid'),
    (lambda p: p.update(reference_locator=p['body_locator']), 'professional_overlapping_changes'),
    (lambda p: p.update(path='C:/outside.docx'), 'professional_invalid_request'),
])
def test_refuses_unsafe_or_unreviewed_claim(tmp_path, monkeypatch, mutation, reason):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    payload = draft(service, path, record_id)
    version = digest(path)
    mutation(payload)
    assert service.plan(payload) == {'ok': False, 'reason': reason}
    assert digest(path) == version and not service._editor().plans


def test_source_conflict_consumes_plan_before_write(tmp_path, monkeypatch):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    payload = draft(service, path, record_id)
    plan = service.plan(payload)['result']
    (tmp_path / 'paper.pdf').write_bytes(pdf_bytes(text='Changed source'))
    assert service.confirm(plan['id'], True) == {'ok': False, 'reason': 'professional_evidence_changed'}
    assert plan['id'] not in service._editor().plans
    assert not service._disabled
    assert len(list(tmp_path.glob('*.docx'))) == 1


def test_proof_tampering_and_deselection_refused(tmp_path, monkeypatch):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    planned = service.plan(draft(service, path, record_id))['result']
    request = service._editor().plans[planned['id']].request
    proof = deepcopy(request.proof)
    proof['fragments'][0]['original_text'] = 'Fake evidence'
    with pytest.raises(GuardError, match='professional_evidence_invalid'):
        validate_proof(request.model_copy(update={'proof': proof}))
    service.literature.decisions[record_id]['decision'] = 'exclude'
    assert service.confirm(planned['id'], True)['reason'] == 'professional_evidence_invalid'


@pytest.mark.parametrize('filename', ['../paper.pdf', 'sub/paper.pdf', 'C:\\paper.pdf', '.env', 'paper.pdf ', 'missing.pdf'])
def test_explicit_pdf_path_boundary(tmp_path, monkeypatch, filename):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    assert not service.capture(record_id, filename)['ok']
    assert not service._captures


def test_existing_citations_expiry_and_logout(tmp_path, monkeypatch):
    service, path, record_id, authenticated, busy = setup_service(tmp_path, monkeypatch)
    payload = draft(service, path, record_id)
    plan = service.plan(payload)['result']
    service._pending[plan['id']]['created'] -= 301
    assert service.confirm(plan['id'], True)['reason'] == 'professional_plan_expired'
    busy[0] = True
    assert service.sources()['reason'] == 'busy'
    busy[0] = False
    authenticated[0] = False
    assert service.sources()['reason'] == 'not_authenticated' and not service._captures
    authenticated[0] = True
    doc = Document(path)
    doc.paragraphs[2].text = 'Already cited [2].'
    doc.save(path)
    assert service.plan(draft(service, path, record_id))['reason'] == 'professional_existing_numbered_citations'


@pytest.mark.parametrize('target,text,reason', [
    ('body', 'The measured latency was 30 ms.', 'professional_protected_numbers'),
    ('reference', 'Reference placeholder 2024', 'professional_grounded_target'),
    ('header', 'Existing citation [5].', 'professional_existing_numbered_citations'),
])
def test_document_wide_citation_and_target_protection(tmp_path, monkeypatch, target, text, reason):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    document = Document(path)
    if target == 'header':
        document.sections[0].header.paragraphs[0].text = text
    else:
        document.paragraphs[0 if target == 'body' else 1].text = text
    document.save(path)
    original = digest(path)
    assert service.plan(draft(service, path, record_id))['reason'] == reason
    assert digest(path) == original and len(list(tmp_path.glob('*.docx'))) == 1


def test_explicit_local_association_requires_complete_nonretracted_metadata(tmp_path, monkeypatch):
    service, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    service.literature.records[record_id]['retraction_status'] = 'reported_retracted'
    assert service.capture(record_id, 'paper.pdf')['reason'] == 'professional_metadata_incomplete'
    service.literature.records[record_id]['retraction_status'] = 'unknown'
    service.literature.records[record_id]['doi'] = None
    assert service.capture(record_id, 'paper.pdf')['reason'] == 'professional_metadata_incomplete'
