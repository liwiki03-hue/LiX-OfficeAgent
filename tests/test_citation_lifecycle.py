from copy import deepcopy

from docx import Document
import pytest

from office_agent.literature import normalize
from office_agent.professional_edits import snapshot
from office_agent.webui.citations import Citations
from office_agent.writes import digest
from test_grounded import setup_service
from test_literature import crossref


def fixture(tmp_path, monkeypatch):
    grounded, path, first, authenticated, busy = setup_service(tmp_path, monkeypatch)
    second_row = normalize(crossref('Independent technical experiment', '10.1234/EXPT.02'), 'crossref')
    second = second_row['id']
    grounded.literature.records[second] = second_row
    grounded.literature.decisions[second] = {'decision': 'include', 'reason': 'Synthetic source'}
    doc = Document()
    doc.add_paragraph('The measured latency was 20 ms [4].')
    doc.add_paragraph('Independent result [2], repeated [7].')
    doc.add_paragraph('[2] First experiment. doi: 10.1234/expt.01')
    doc.add_paragraph('[4] Second experiment. doi: 10.1234/expt.02')
    doc.add_paragraph('[7] Duplicate first experiment. doi: 10.1234/expt.01')
    doc.add_paragraph('Reference placeholder')
    doc.save(path)
    service = Citations(grounded._lock, grounded._scope, grounded._editor, grounded._selected, grounded)
    source = snapshot(grounded._editor(), path)
    items = source['items']
    payload = {'source_sha256': source['document']['source_sha256'], 'authorized': [row['locator_id'] for row in items],
               'bibliography': [row['locator_id'] for row in items[2:]], 'number_records': {'2': first, '4': second, '7': first},
               'mapping_reviewed': True}
    return service, grounded, path, payload


def test_reorders_deduplicates_and_rolls_back_as_one_transaction(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    original = digest(path)
    plan = service.plan(payload)
    assert plan['ok'], plan
    manifest = plan['result']['manifest']
    assert manifest['number_to_key'] == {'1': 'doi:10.1234/expt.02', '2': 'doi:10.1234/expt.01'}
    assert manifest['number_changes'] == {'2': 2, '4': 1, '7': 2}
    assert manifest['previous_number_to_key']['7'] == 'doi:10.1234/expt.01'
    assert service.confirm(plan['result']['id'], False)['result']['status'] == 'cancelled'
    assert digest(path) == original and len(list(tmp_path.glob('*.docx'))) == 1
    plan = service.plan(payload)['result']
    saved = service.confirm(plan['id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    from pathlib import Path
    output = Path(saved['result']['output'])
    doc = Document(output)
    assert doc.paragraphs[0].text == 'The measured latency was 20 ms [1].'
    assert doc.paragraphs[1].text == 'Independent result [2], repeated [2].'
    assert '10.1234/expt.02' in doc.paragraphs[2].text
    assert '10.1234/expt.01' in doc.paragraphs[3].text
    assert all(row.text == '参考文献占位（未使用）' for row in doc.paragraphs[4:])
    assert digest(path) == original
    assert service.rollback(plan['id'], True)['result']['status'] == 'rolled_back'
    assert not output.exists()


def test_missing_authorized_renumbering_target_refuses_entire_plan(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    payload['authorized'].pop(0)
    version = digest(path)
    assert service.plan(payload)['reason'] == 'professional_citation_scope'
    assert digest(path) == version and not grounded._editor().plans


def test_remove_last_occurrences_removes_bibliography_but_preserves_claim_text(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    payload['removals'] = {payload['authorized'][1]: [2, 7]}
    plan = service.plan(payload)
    assert plan['ok'], plan
    assert plan['result']['manifest']['removed_keys'] == ['doi:10.1234/expt.01']
    assert 'Independent result , repeated .' in str(plan['result']['after'])


@pytest.mark.parametrize('mutation,reason', [
    (lambda p: p.update(mapping_reviewed=False), 'professional_citation_review'),
    (lambda p: p['number_records'].pop('7'), 'professional_bibliography_identity'),
    (lambda p: p['bibliography'].pop(2), 'professional_citation_scope'),
    (lambda p: p.update(source_sha256='0' * 64), 'professional_source_changed'),
])
def test_mapping_and_version_fail_closed(tmp_path, monkeypatch, mutation, reason):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    mutation(payload)
    assert service.plan(payload)['reason'] == reason


def test_cross_paragraph_new_evidence_claim_and_rebuilt_csl(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    record_id = payload['number_records']['2']
    capture = grounded.capture(record_id, 'paper.pdf')['result']
    payload['claims'] = [{'locator_id': payload['authorized'][0], 'text': 'The measured latency was 20 ms.',
                         'attributions': [{'capture_id': capture['capture_id'], 'evidence_ids': [capture['fragments'][0]['id']],
                                           'relation': 'supports', 'applicability': 'Laboratory only', 'limitation': 'Synthetic fixture',
                                           'assessment': 'user_reviewed', 'identity_reviewed': True}]}]
    reply = service.plan(payload)
    assert reply['ok'], reply
    assert reply['result']['manifest']['number_to_key'] == {'1': 'doi:10.1234/expt.01'}
    assert reply['result']['manifest']['removed_keys'] == ['doi:10.1234/expt.02']


def test_output_receipt_resume_remove_reference_then_rollback_descendant(tmp_path, monkeypatch):
    from pathlib import Path
    import threading
    from office_agent.webui.grounded import Grounded
    from office_agent.webui.literature import Literature
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    plan = service.plan(payload)['result']
    saved = service.confirm(plan['id'], True)['result']
    output = Path(saved['output'])
    editor = grounded._editor()
    scope = lambda: (True, False, (output, editor.space.root))
    lock = threading.RLock()
    literature = Literature(lock, scope, lambda: editor, tmp_path)
    new_grounded = Grounded(lock, scope, lambda: editor, lambda: output, literature)
    resumed = Citations(lock, scope, lambda: editor, lambda: output, new_grounded)
    read = resumed.read(plan['id'])
    assert read['ok'] and not read['result']['receipt_restores_write_approval'], read
    rows = read['result']['items']
    metadata = read['result']['manifest']['catalog']
    payload = {'source_sha256': digest(output), 'authorized': [row['locator_id'] for row in rows],
               'bibliography': [row['locator_id'] for row in rows[2:]],
               'number_records': {'1': metadata['doi:10.1234/expt.02']['id'], '2': metadata['doi:10.1234/expt.01']['id']},
               'removals': {rows[1]['locator_id']: [2]}, 'mapping_reviewed': True}
    update = resumed.plan(payload)
    assert update['ok'], update
    assert update['result']['manifest']['parent_receipt'] == plan['id']
    descendant = resumed.confirm(update['result']['id'], True)
    assert descendant['result']['status'] == 'saved', descendant
    assert resumed.rollback(update['result']['id'], True)['result']['status'] == 'rolled_back'
    assert output.exists()


def test_resume_receipt_external_change_expires_plan(tmp_path, monkeypatch):
    from pathlib import Path
    import json
    import threading
    from office_agent.webui.grounded import Grounded
    from office_agent.webui.literature import Literature
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    plan = service.plan(payload)['result']
    saved = service.confirm(plan['id'], True)['result']
    output = Path(saved['output'])
    editor = grounded._editor()
    scope = lambda: (True, False, (output, editor.space.root))
    lock = threading.RLock()
    literature = Literature(lock, scope, lambda: editor, tmp_path)
    resumed = Citations(lock, scope, lambda: editor, lambda: output, Grounded(lock, scope, lambda: editor, lambda: output, literature))
    read = resumed.read(plan['id'])['result']
    rows, metadata = read['items'], read['manifest']['catalog']
    payload = {'source_sha256': digest(output), 'authorized': [row['locator_id'] for row in rows],
               'bibliography': [row['locator_id'] for row in rows[2:]],
               'number_records': {'1': metadata['doi:10.1234/expt.02']['id'], '2': metadata['doi:10.1234/expt.01']['id']},
               'removals': {rows[1]['locator_id']: [2]}, 'mapping_reviewed': True}
    update = resumed.plan(payload)['result']
    audit = editor.audit / plan['id'] / 'audit.json'
    audit.write_text(audit.read_text(encoding='utf-8') + ' ', encoding='utf-8')
    assert resumed.confirm(update['id'], True)['reason'] == 'professional_citation_receipt'


def test_single_source_refusal_reproduced_before_bounded_migration(tmp_path, monkeypatch):
    from test_grounded import draft
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    assert grounded.plan(draft(grounded, path, payload['number_records']['2']))['reason'] == 'professional_existing_numbered_citations'
    assert service.plan(payload)['ok']


def test_same_doi_metadata_conflict_requires_explicit_version_resolution(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    other = normalize(crossref('Conflicting title', '10.1234/EXPT.01'), 'crossref')
    grounded.literature.records[other['id']] = other
    grounded.literature.decisions[other['id']] = {'decision': 'include', 'reason': 'Conflicting synthetic version'}
    payload['number_records']['7'] = other['id']
    assert service.plan(payload)['reason'] == 'professional_duplicate_metadata'


def test_one_claim_with_two_sources_including_archive_has_atomic_dependencies(tmp_path, monkeypatch):
    from test_archive_grounded import archive_source
    from test_fidelity import pdf_bytes
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    archived = archive_source(grounded, payload['number_records']['2'])
    first = grounded.capture_archive(archived['save_id'])['result']
    (tmp_path / 'independent.pdf').write_bytes(pdf_bytes(text='Independent tests measured latency of 20 ms.', width=600))
    second = grounded.capture(payload['number_records']['4'], 'independent.pdf')['result']
    attributions = [{'capture_id': captured['capture_id'], 'evidence_ids': [captured['fragments'][0]['id']],
                     'relation': 'supports', 'applicability': 'Each synthetic fixture only', 'limitation': 'Not actual scientific corroboration',
                     'identity_reviewed': True, 'assessment': 'user_reviewed'} for captured in (first, second)]
    payload['claims'] = [{'locator_id': payload['authorized'][0], 'text': 'The measured latency was 20 ms.', 'attributions': attributions}]
    plan = service.plan(payload)
    assert plan['ok'], plan
    assert plan['result']['manifest']['number_to_key'] == {'1': 'doi:10.1234/expt.01', '2': 'doi:10.1234/expt.02'}
    assert '[1], [2]' in plan['result']['items'][0]['text']
    assert plan['result']['manifest']['archive_dependencies'] == [archived['save_id']]
    assert grounded.literature.rollback(archived['save_id'], True)['reason'] == 'literature_fulltext_in_use'
    saved = service.confirm(plan['result']['id'], True)
    assert saved['result']['status'] == 'saved', saved
    assert grounded.literature.rollback(archived['save_id'], True)['reason'] == 'literature_fulltext_in_use'
    assert service.rollback(plan['result']['id'], True)['result']['status'] == 'rolled_back'
    assert grounded.literature.rollback(archived['save_id'], True)['result']['status'] == 'rolled_back'


def test_sensitive_support_note_is_refused_without_echo_or_audit(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    capture = grounded.capture(payload['number_records']['2'], 'paper.pdf')['result']
    synthetic_secret = 'sk-' + 'a' * 32
    payload['claims'] = [{'locator_id': payload['authorized'][0], 'text': 'The measured latency was 20 ms.',
                         'attributions': [{'capture_id': capture['capture_id'], 'evidence_ids': [capture['fragments'][0]['id']],
                                           'relation': 'supports', 'applicability': 'Synthetic fixture', 'limitation': synthetic_secret,
                                           'assessment': 'user_reviewed', 'identity_reviewed': True}]}]
    reply = service.plan(payload)
    assert reply == {'ok': False, 'reason': 'professional_sensitive_text'}
    assert synthetic_secret not in str(reply) and not grounded._editor().plans


def test_doi_prefix_is_not_treated_as_an_exact_bibliographic_identity(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    doc = Document(path)
    doc.paragraphs[2].text = '[2] Different paper. doi: 10.1234/expt.012'
    doc.save(path)
    source = snapshot(grounded._editor(), path)
    payload.update(source_sha256=source['document']['source_sha256'], authorized=[row['locator_id'] for row in source['items']],
                   bibliography=[row['locator_id'] for row in source['items'][2:]])
    assert service.plan(payload)['reason'] == 'professional_bibliography_identity'


def test_full_bibliography_can_remove_old_entries_and_add_a_new_source_atomically(tmp_path, monkeypatch):
    service, grounded, path, _ = fixture(tmp_path, monkeypatch)
    records = []
    for index in range(1, 10):
        row = normalize(crossref('Capacity fixture ' + str(index), f'10.1234/capacity.{index}'), 'crossref')
        grounded.literature.records[row['id']] = row
        grounded.literature.decisions[row['id']] = {'decision': 'include', 'reason': 'Synthetic capacity fixture'}
        records.append(row)
    document = Document()
    document.add_paragraph('Original discussion [1,2,3,4].')
    document.add_paragraph('Retained discussion [5,6,7,8].')
    for index, row in enumerate(records[:8], 1):
        document.add_paragraph(f'[{index}] Capacity fixture. doi: ' + row['doi'])
    document.save(path)
    source = snapshot(grounded._editor(), path)
    capture = grounded.capture(records[8]['id'], 'paper.pdf')['result']
    payload = {'source_sha256': source['document']['source_sha256'], 'authorized': [row['locator_id'] for row in source['items']],
               'bibliography': [row['locator_id'] for row in source['items'][2:]],
               'number_records': {str(index): row['id'] for index, row in enumerate(records[:8], 1)}, 'mapping_reviewed': True,
               'claims': [{'locator_id': source['items'][0]['locator_id'], 'text': 'The measured latency was 20 ms.',
                           'attributions': [{'capture_id': capture['capture_id'], 'evidence_ids': [capture['fragments'][0]['id']],
                             'relation': 'supports', 'applicability': 'Synthetic fixture', 'limitation': 'Not a real research conclusion',
                             'assessment': 'user_reviewed', 'identity_reviewed': True}]}]}
    plan = service.plan(payload)
    assert plan['ok'], plan
    assert len(plan['result']['manifest']['catalog']) == 5
    assert plan['result']['manifest']['number_to_key']['1'] == 'doi:10.1234/capacity.9'
    assert len(plan['result']['manifest']['removed_keys']) == 4
    assert not plan['result']['paragraphs_inserted_or_deleted'] and len(plan['result']['items']) <= 12


def test_numeric_vector_inside_protected_formula_is_not_renumbered_as_a_citation(tmp_path, monkeypatch):
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    document = Document(path)
    document.paragraphs[0].text = 'The measured latency was 20 ms [4]; formula $x=[4]$.'
    document.save(path)
    source = snapshot(grounded._editor(), path)
    payload.update(source_sha256=source['document']['source_sha256'], authorized=[row['locator_id'] for row in source['items']],
                   bibliography=[row['locator_id'] for row in source['items'][2:]])
    plan = service.plan(payload)
    assert plan['ok'], plan
    assert plan['result']['items'][0]['text'] == 'The measured latency was 20 ms [1]; formula $x=[4]$.'
    assert plan['result']['manifest']['number_changes'] == {'2': 2, '4': 1, '7': 2}


def test_core_plan_expiry_cannot_be_bypassed_by_direct_editor_execution(tmp_path, monkeypatch):
    from office_agent.citation_lifecycle import reconcile
    from office_agent.grounded_edits import seal
    from office_agent.advanced_writes import _parts
    from office_agent.writes import GuardError
    service, grounded, path, payload = fixture(tmp_path, monkeypatch)
    plan = service.plan(payload)['result']
    request = grounded._editor().plans[plan['id']].request
    intent = {**request.intent, 'deadline': 0.0}
    expired = request.model_copy(update={'intent': intent, 'intent_seal': seal(intent)})
    with pytest.raises(GuardError, match='professional_plan_expired'):
        reconcile(_parts(path), expired)


@pytest.mark.parametrize('text', ['0', '4-1', '1-30', '1,0'])
def test_invalid_or_excessive_numeric_cluster_is_refused(text):
    from office_agent.citation_lifecycle import numbers
    from office_agent.writes import GuardError
    with pytest.raises(GuardError):
        numbers(text)
