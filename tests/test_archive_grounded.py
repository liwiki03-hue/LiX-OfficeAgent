from copy import deepcopy
import json
from pathlib import Path

import pytest

from office_agent.literature import csl
from office_agent.professional_edits import snapshot
from test_fidelity import pdf_bytes
from test_grounded import setup_service


def archive_source(grounded, record_id, data=None):
    literature = grounded.literature
    record = literature.records[record_id]
    origin = {'record_id': record_id, 'source_url': 'https://arxiv.org/pdf/2401.12345', 'license': 'cc-by',
              'retrieved_at': record['retrieved_at'], 'source_record_sha256': record['source_record_sha256'], 'csl': csl(record)}
    plan = literature._plan(grounded._editor(), 'pdf', data or pdf_bytes(text='The measured latency was 20 ms.', width=600), origin)['plan']
    saved = literature.confirm(plan['plan_id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    return saved['result']


def archive_draft(grounded, path, saved):
    captured = grounded.capture_archive(saved['save_id'])
    assert captured['ok'], captured
    captured = captured['result']
    source = snapshot(grounded._editor(), path)
    return {'source_sha256': source['document']['source_sha256'], 'capture_id': captured['capture_id'],
            'body_locator': source['items'][0]['locator_id'], 'reference_locator': source['items'][1]['locator_id'],
            'identity_reviewed': True, 'support': {'claim': 'The measured latency was 20 ms.',
              'evidence_ids': [captured['fragments'][0]['id']], 'relation': 'supports', 'applicability': 'Synthetic laboratory conditions',
              'limitation': 'No real performance conclusion', 'assessment': 'user_reviewed'}}


def test_archive_receipt_to_grounded_plan_dependency_cancel_save_and_rollback(tmp_path, monkeypatch):
    grounded, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    saved = archive_source(grounded, record_id)
    payload = archive_draft(grounded, path, saved)
    plan = grounded.plan(payload)
    assert plan['ok'], plan
    literature = grounded.literature
    assert literature.rollback(saved['save_id'], True)['reason'] == 'literature_fulltext_in_use'
    assert grounded.confirm(plan['result']['id'], False)['result']['status'] == 'cancelled'
    assert not literature._dependencies
    plan = grounded.plan(payload)['result']
    result = grounded.confirm(plan['id'], True)
    assert result['result']['status'] == 'saved', result
    assert literature._dependencies[plan['id']]['state'] == 'output'
    assert literature.rollback(saved['save_id'], True)['reason'] == 'literature_fulltext_in_use'
    assert grounded.rollback(plan['id'], True)['result']['status'] == 'rolled_back'
    assert literature.rollback(saved['save_id'], True)['result']['status'] == 'rolled_back'
    assert not grounded.capture_archive(saved['save_id'])['ok']


@pytest.mark.parametrize('damage', ['pdf', 'audit', 'extra', 'directory'])
def test_archive_change_invalidates_write_before_admission(tmp_path, monkeypatch, damage):
    grounded, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    saved = archive_source(grounded, record_id)
    payload = archive_draft(grounded, path, saved)
    plan = grounded.plan(payload)['result']
    directory = (tmp_path / saved['evidence']).parent
    if damage == 'pdf':
        (directory / 'fulltext.pdf').write_bytes(pdf_bytes(text='Replacement'))
    elif damage == 'audit':
        (directory / 'audit-verified.json').write_bytes(b'{}')
    elif damage == 'extra':
        (directory / 'audit-remove-intent.json').write_bytes(b'{}')
    else:
        moved = directory.with_name(directory.name + '-moved')
        assert directory.resolve().is_relative_to(tmp_path.resolve()) and moved.resolve().is_relative_to(tmp_path.resolve())
        directory.rename(moved)
        directory.mkdir()
        for source in moved.iterdir():
            (directory / source.name).write_bytes(source.read_bytes())
    reply = grounded.confirm(plan['id'], True)
    assert reply == {'ok': False, 'reason': 'professional_evidence_changed'}
    assert len(list(tmp_path.glob('*.docx'))) == 1 and not grounded._disabled


def test_readonly_archive_restore_after_session_reset_and_truncation(tmp_path, monkeypatch):
    grounded, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    saved = archive_source(grounded, record_id, pdf_bytes(text='Multi-page synthetic text.', pages=6, width=600))
    grounded.invalidate()
    grounded.literature.invalidate()
    restored = grounded.capture_archive(saved['save_id'])
    assert restored['ok'], restored
    assert restored['result']['truncated'] and max(row['page'] for row in restored['result']['fragments']) == 5
    assert restored['result']['document']['access'] == 'public_fulltext'
    assert not restored['result']['doi_match_verified']
    assert grounded.literature.saves == {} and not grounded.literature.rollback(saved['save_id'], True)['ok']
    assert 'multicolumn' in str(restored['result']['limitations'])


@pytest.mark.parametrize('bad', ['../outside', 'f' * 31, 'C:/private', 'f' * 32])
def test_archive_receipt_never_accepts_arbitrary_path(tmp_path, monkeypatch, bad):
    grounded, path, record_id, _, _ = setup_service(tmp_path, monkeypatch)
    assert not grounded.capture_archive(bad)['ok']
