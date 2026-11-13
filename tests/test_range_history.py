import json
import os
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.rag_workspace import WorkspaceRag
from office_agent.rag_evidence import EvidenceError
from office_agent.webui.bridge import Api
from office_agent.writes import Editor
from test_rag_workspace import SyntheticEmbedder
from test_research_reading import long_pdf


def fixture(tmp_path):
    path = tmp_path / 'report.docx'
    doc = Document()
    doc.add_paragraph('Owned experimental report.')
    doc.save(path)
    (tmp_path / 'paper.pdf').write_bytes(long_pdf())
    app = SimpleNamespace(path=path, editor=Editor(tmp_path, require_confirmation=True), busy=False)
    authenticated = [True]
    auth = SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': authenticated[0]})
    api = Api(app, auth=auth, settings_root=tmp_path)
    embedding = SyntheticEmbedder()
    api._rag._factory = lambda check: WorkspaceRag(app.editor.space, embedding, scope_check=check)
    failed = api.plan_pdf_ranges([[99, 100]], 'paper.pdf')
    assert not failed['ok'], failed
    attempt = api.get_pdf_ranges('paper.pdf')['ranges']['attempt']
    assert attempt['state'] == 'failed' and attempt['failed_pages'] == [99] and attempt['unprocessed_pages'] == [100]
    return api, app, auth, authenticated, embedding


def save(api):
    plan = api.plan_pdf_attempt_history('paper.pdf')
    assert plan['ok'], plan
    saved = api.confirm_pdf_attempt_history(plan['result']['plan_id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    return saved['result']['save_id']


def test_archive_cancel_save_real_fresh_host_restore_and_independent_removal(tmp_path):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    before = (tmp_path / 'paper.pdf').read_bytes()
    plan = api.plan_pdf_attempt_history('paper.pdf')['result']
    assert not list(tmp_path.glob('.office-rag-evidence-*')) and not embedding.calls
    cancelled = api.confirm_pdf_attempt_history(plan['plan_id'], False)
    assert cancelled['result']['wrote_nothing'] and not list(tmp_path.glob('.office-rag-evidence-*'))
    receipt = save(api)
    directory = tmp_path / ('.office-rag-evidence-' + receipt)
    assert {p.name for p in directory.iterdir()} == {'backup', 'audit-intent.json', 'audit-verified.json', 'pdf-range-attempt.json'}
    app.editor = Editor(tmp_path, require_confirmation=True)
    restored_api = Api(app, auth=auth, settings_root=tmp_path)
    restored_api._rag._factory = lambda check: pytest.fail('Restoring diagnostics must not configure or load a model')
    reply = restored_api.restore_pdf_attempt_history(receipt)
    assert reply['ok'], reply
    result = reply['result']
    assert result['source_state'] == 'matches' and result['processing_current']
    assert result['embedding_current'] is None and result['historical_only'] and result['wrote_nothing']
    assert not any(result[k] for k in ['restores_index', 'restores_attempt', 'restores_approval', 'restores_rollback_authority'])
    assert not restored_api._research_history._archive_backend._saved
    assert restored_api._rag._store is None
    (tmp_path / 'paper.pdf').write_bytes(before + b'\n')
    assert restored_api.restore_pdf_attempt_history(receipt)['result']['source_state'] == 'changed_or_moved'
    (tmp_path / 'paper.pdf').unlink()
    assert restored_api.restore_pdf_attempt_history(receipt)['result']['source_state'] == 'unavailable_or_refused'
    removal = restored_api.plan_remove_pdf_attempt_history(receipt)['result']
    assert restored_api.confirm_remove_pdf_attempt_history(removal['rollback_id'], False)['result']['wrote_nothing']
    removal = restored_api.plan_remove_pdf_attempt_history(receipt)['result']
    result = restored_api.confirm_remove_pdf_attempt_history(removal['rollback_id'], True)['result']
    assert result['status'] == 'rolled_back' and result['backups_audits_retained']
    assert not (directory / 'pdf-range-attempt.json').exists() and (directory / 'backup').exists()
    assert (directory / 'audit-removed.json').exists() and app.path.exists()
    assert not restored_api.restore_pdf_attempt_history(receipt)['ok'] and not embedding.calls


def test_changed_attempt_expired_plan_and_duplicate_never_write(tmp_path):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    plan = api.plan_pdf_attempt_history('paper.pdf')['result']
    (tmp_path / 'paper.pdf').write_bytes((tmp_path / 'paper.pdf').read_bytes() + b'\n')
    assert api.confirm_pdf_attempt_history(plan['plan_id'], True)['reason'] == 'professional_history_attempt_changed'
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    plan = api.plan_pdf_attempt_history('paper.pdf')['result']
    archive = api._research_history._archive_backend
    archive._plans[plan['plan_id']]['time'] -= 301
    result = api.confirm_pdf_attempt_history(plan['plan_id'], True)['result']
    assert result['code'] == 'evidence_expired' and result['wrote_nothing']
    assert api.confirm_pdf_attempt_history(plan['plan_id'], True)['reason'] == 'professional_plan_invalid'
    assert not list(tmp_path.glob('.office-rag-evidence-*')) and not embedding.calls


@pytest.mark.parametrize('tamper', ['extra', 'audit', 'hardlink'])
def test_restore_and_remove_reject_corrupt_receipts(tmp_path, tamper):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    receipt = save(api)
    directory = tmp_path / ('.office-rag-evidence-' + receipt)
    if tamper == 'extra':
        (directory / 'unexpected').write_text('Owned extra file', encoding='utf8')
    elif tamper == 'audit':
        (directory / 'audit-verified.json').write_text('{}', encoding='utf8')
    else:
        os.link(directory / 'pdf-range-attempt.json', tmp_path / 'owned-link.json')
    assert not api.restore_pdf_attempt_history(receipt)['ok']
    assert not api.plan_remove_pdf_attempt_history(receipt)['ok']
    assert (directory / 'pdf-range-attempt.json').exists()


def test_uncertain_write_fuses_and_logout_retires_plans(tmp_path, monkeypatch):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    plan = api.plan_pdf_attempt_history('paper.pdf')['result']
    authenticated[0] = False
    assert api.confirm_pdf_attempt_history(plan['plan_id'], True)['reason'] == 'not_authenticated'
    authenticated[0] = True
    assert api.confirm_pdf_attempt_history(plan['plan_id'], True)['reason'] == 'professional_plan_invalid'
    plan = api.plan_pdf_attempt_history('paper.pdf')['result']
    archive = api._research_history._archive_backend
    monkeypatch.setattr(archive, '_exclusive', lambda *args: (_ for _ in ()).throw(OSError('Owned injected failure')))
    result = api.confirm_pdf_attempt_history(plan['plan_id'], True)['result']
    assert result['status'] == 'uncertain' and not result['wrote_nothing']
    assert api.plan_pdf_attempt_history('paper.pdf')['reason'] == 'professional_uncertain'
    assert len(list(tmp_path.glob('.office-rag-evidence-*'))) == 1


def test_removal_plan_conflict_and_changed_parser_are_reported(tmp_path, monkeypatch):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    receipt = save(api)
    monkeypatch.setattr('office_agent.research_ranges.processing_stamp', lambda: 'f' * 64)
    restored = api.restore_pdf_attempt_history(receipt)['result']
    assert restored['source_state'] == 'matches' and not restored['processing_current']
    removal = api.plan_remove_pdf_attempt_history(receipt)['result']
    directory = tmp_path / ('.office-rag-evidence-' + receipt)
    artifact = directory / 'pdf-range-attempt.json'
    artifact.write_bytes(artifact.read_bytes() + b'\n')
    assert not api.confirm_remove_pdf_attempt_history(removal['rollback_id'], True)['ok']
    assert artifact.exists() and not (directory / 'audit-remove-intent.json').exists()


@pytest.mark.parametrize('field,value', [('parsed_pages', [True]), ('selected_pages', [99, 99]), ('unprocessed_pages', []), ('reason', 'fixture-secret-unapproved')])
def test_archive_validation_rejects_forged_diagnostic_fields(tmp_path, field, value):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    plan = api.plan_pdf_attempt_history('paper.pdf')['result']
    document = json.loads(json.dumps(plan['record']))
    document['attempt'][field] = value
    archive = api._research_history._archive_backend
    pending = set(archive._plans)
    with pytest.raises(EvidenceError):
        archive.plan_attempt(document, api._research_history._epoch)
    assert set(archive._plans) == pending and not list(tmp_path.glob('.office-rag-evidence-*'))


def test_changed_parse_version_never_counts_as_success(tmp_path, monkeypatch):
    api, app, auth, authenticated, embedding = fixture(tmp_path)
    owner = api._rag._store
    parsed = owner._pdf_source.extract('paper.pdf', start_page=1, end_page=1)
    parsed['fingerprint'] = 'a' * 64
    monkeypatch.setattr(owner._pdf_source, 'extract', lambda *args, **kwargs: parsed)
    assert not api.plan_pdf_ranges([[1, 2]], 'paper.pdf')['ok']
    attempt = api.get_pdf_ranges('paper.pdf')['ranges']['attempt']
    assert attempt['parsed_pages'] == [] and attempt['failed_pages'] == [1] and attempt['unprocessed_pages'] == [2]
    assert not owner.snapshot()['exists'] and not embedding.calls
