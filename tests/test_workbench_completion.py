from copy import deepcopy
import json

from docx import Document
import pytest

from office_agent.project_library import paper_families
from office_agent.research_capture_retrieval import retrieve_captures
from office_agent.writes import digest
from test_pdf_header_rows import header_pdf
from test_project_workbench import CSL, imported
from test_pdf_table import host
from test_workbench_research import prepared, review_comment


def test_explicit_paper_versions_share_budget_without_merging_metadata():
    entries = {key: {'item': {'doi': doi}, 'relations': []} for key, doi in [('a', 'doi-a'), ('b', 'doi-b'), ('c', 'doi-c'), ('d', 'doi-d')]}
    entries['a']['relations'] = [{'type': 'preprint_of', 'id': 'doi-b'}]
    entries['b']['relations'] = [{'type': 'version_of', 'id': 'c'}]
    entries['c']['relations'] = [{'type': 'possible_duplicate_of', 'id': 'doi-d'}]
    project = {'entries': entries}
    original = deepcopy(project)
    families = paper_families(project)
    assert families['a'] == families['b'] == families['c'] != families['d']
    assert project == original
    evidence = {f'ev-{i}': {'id': f'ev-{i}', 'original_text': 'latency under dry conditions',
                'source_sha256': str(i), 'location': str(i), 'diversity_key': families['a'] if i < 8 else families['d']} for i in range(10)}
    found, report = retrieve_captures(evidence, 'latency', ['latency'], [])
    assert sum(row['diversity_key'] == families['a'] for row in found.values()) == 3
    assert sum(row['diversity_key'] == families['d'] for row in found.values()) == 2
    assert report['diversity_policy'] == 'known_paper_max_3_except_explicit_context'


def test_empty_search_returns_registered_gaps_without_reading_or_model(tmp_path, monkeypatch):
    api, auth = host(tmp_path)
    key = imported(api)
    monkeypatch.setattr('office_agent.rag_sources.SourceExtractor.extract', lambda *args: pytest.fail('Unexpected body extraction'))
    question = 'Do not generalize 20 ms in 2024 beyond dry specimens'
    reply = api.workbench_action('search', {'question': question, 'focus': '', 'dense': False, 'read_ids': []})
    assert reply['ok'], reply
    data = reply['result']
    assert data['status'] == 'insufficient' and data['model_calls'] == 0
    assert data['original_question'] == question and data['coverage'][0]['question'] == question
    assert data['materials'][0]['id'] == key and data['materials'][0]['unprocessed_pages_unknown']
    assert not data['complete_fulltext_coverage'] and not data['absence_of_contradictions_proven']
    assert data['coverage'][0]['condition_review_required'] and data['coverage'][0]['pending_numeric_terms'] == ['20', '2024']


def test_queue_location_uses_current_host_anchor_beyond_first_page(tmp_path):
    api, auth = host(tmp_path)
    document = Document()
    for index in range(45):
        document.add_paragraph(f'Owned paragraph {index}.')
    document.save(tmp_path / 'report.docx')
    imported(api)
    original = digest(tmp_path / 'report.docx')
    anchor = api.read_structure_anchors(40)['result']['anchors'][-1]
    assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 40})['ok']
    queued = api.workbench_action('queue', {'operation': 'comment', 'text': 'Check this paragraph.', 'evidence_ids': []})
    assert queued['ok'], queued
    row = queued['result']['drafts'][0]
    first = api.read_structure_anchors()['result']['anchors'][0]
    assert api.workbench_action('target', {'locator_id': first['locator_id'], 'offset': 0})['ok']
    located = api.workbench_action('queue_target', {'id': row['id']})
    assert located['ok'], located
    assert located['result']['target']['locator_id'] == anchor['locator_id']
    assert located['result']['offset'] == 44 and not located['result']['write_approval_restored']
    assert api._workbench.project['drafts'][0]['status'] == 'pending'
    assert digest(tmp_path / 'report.docx') == original
    document.add_paragraph('External version change.')
    document.save(tmp_path / 'report.docx')
    assert not api.workbench_action('queue_target', {'id': row['id']})['ok']


def test_gap_material_navigation_reads_only_registered_metadata(tmp_path, monkeypatch):
    api, auth = host(tmp_path)
    rows = [{**CSL[0], 'title': f'Material {index:03}', 'DOI': f'10.1000/material-{index}'} for index in range(25)]
    imported_rows = api.workbench_action('import_csl', {'content': json.dumps(rows)})
    assert imported_rows['ok'], imported_rows
    key = next(key for key, entry in api._workbench.project['entries'].items() if entry['item']['title'] == 'Material 024')
    monkeypatch.setattr('office_agent.rag_sources.SourceExtractor.extract', lambda *args: pytest.fail('Unexpected body extraction'))
    result = api.workbench_action('record_page', {'id': key})
    assert result['ok'], result
    assert result['result']['files_read'] == 0 and result['result']['selected_record_id'] == key
    assert result['result']['snapshot']['entries'][0]['id'] == key
    assert not api.workbench_action('record_page', {'id': 'unknown'})['ok']


def test_reassociated_source_impact_preserves_old_identity_and_related_draft(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    row = review_comment(api, evidence)
    previous = deepcopy(api._workbench.project['entries'][key]['link'])
    (tmp_path / 'new.pdf').write_bytes(header_pdf([['Sample', 'Latency'], ['Identifier', '(ms)'], ['A', '35 ms']]))
    assert api.workbench_action('associate', {'id': key, 'filename': 'new.pdf'})['ok']
    result = api.workbench_action('impact', {'record_ids': [key]})
    assert result['ok'], result
    claims = [item for item in result['result']['impacts'] if item['kind'] == 'claim']
    assert claims and claims[0]['old_identity'] == previous['identity']
    assert claims[0]['old_sha256'] == previous['sha256'] != claims[0]['new_sha256']
    assert row['id'] in result['result']['invalidated_drafts']
    assert not result['result']['documents_rewritten'] and not result['result']['history_deleted']


def test_layout_impact_includes_evidence_draft_and_uncertain_historical_output(tmp_path):
    api, auth, key, evidence = prepared(tmp_path)
    stored = api._workbench.evidence[evidence['id']]
    stored['layout_stamp'] = 'a' * 16
    row = review_comment(api, stored)
    entry = api._workbench.project['entries'][key]
    api._workbench.project['outputs'].append({'receipt': 'f' * 32, 'record_ids': [key],
        'source_stamps': {key: entry['link']['sha256']}, 'target_name': 'report.docx', 'target_version': digest(tmp_path / 'report.docx'), 'kind': 'document'})
    api._workbench.layout_changed('table.pdf', entry['link']['sha256'], 'b' * 16, 'c' * 64)
    result = api.workbench_action('impact', {'record_ids': [key]})
    assert result['ok'], result
    impacts = result['result']['impacts']
    assert {'evidence', 'claim', 'draft', 'document'} <= {item['kind'] for item in impacts}
    output = next(item for item in impacts if item['kind'] == 'document')
    assert output['certainty'] == 'possible_registered_layout_dependency' and output['location'] == 'report.docx'
    assert row['id'] in result['result']['invalidated_drafts']
    assert api._workbench.project['drafts'][0]['status'] == 'recheck'
    assert not api.plan_workbench_queue()['ok']


def test_project_keeps_real_saved_evidence_registration_but_never_restores_grant(tmp_path):
    from office_agent.rag_workspace import WorkspaceRag
    from test_project_history import save
    from test_rag_workspace import SyntheticEmbedder
    api, auth, key, evidence = prepared(tmp_path)
    embedding = SyntheticEmbedder()
    api._rag._factory = lambda check: WorkspaceRag(api._workbench._editor().space, embedding, scope_check=check)
    plan = api.plan_pdf_ranges([[1, 1]], 'table.pdf')
    assert plan['ok'], plan
    saved = api.confirm_rag_index(plan['plan']['plan_id'], True)
    assert saved['ok'], saved
    query = api.query_pdf_ranges('20 ms', 5, 'table.pdf')
    assert query['ok'], query
    evidence_save = api.confirm_rag_query(query['plan']['plan_id'], True)
    assert evidence_save['ok'], evidence_save
    receipt = evidence_save['result']['save_id']
    bound = api.workbench_action('approved_evidence', {'id': key, 'receipt': receipt})
    assert bound['ok'], bound
    registration = deepcopy(api._workbench.project['evidence_records'])
    assert registration[0]['receipt'] == receipt and registration[0]['evidence_ids']
    calls = len(embedding.calls)
    assert api.workbench_action('approved_evidence', {'id': key, 'receipt': receipt})['ok']
    assert len(embedding.calls) == calls and len(api._workbench.project['evidence_records']) == 1
    project_receipt = save(api, 'Evidence registration')
    api._workbench.invalidate()
    api._rag.invalidate()
    restored = api.restore_project(project_receipt)
    assert restored['ok'], restored
    assert restored['result']['evidence_records'] == registration
    assert not api._workbench.evidence and not api._workbench.approved_evidence
    assert not api.workbench_action('approved_evidence', {'id': key, 'receipt': receipt})['ok']
    assert len(embedding.calls) == calls


def test_legacy_project_without_evidence_registry_is_readable_and_still_strict(tmp_path):
    from office_agent.project_library import validate
    api, auth = host(tmp_path)
    imported(api)
    legacy = deepcopy(api._workbench.project)
    legacy.pop('evidence_records')
    assert validate(legacy)['evidence_records'] == []
    legacy['evidence_records'] = [{'receipt': 'f' * 32, 'record_id': 'unknown', 'source_sha256': 'e' * 64, 'evidence_ids': ['ev:' + 'd' * 64]}]
    from office_agent.rag_evidence import EvidenceError
    with pytest.raises(EvidenceError):
        validate(legacy)
