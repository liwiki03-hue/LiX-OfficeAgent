from pathlib import Path
from threading import RLock
from types import SimpleNamespace
from copy import deepcopy
import os

import pytest

from office_agent.render_identity import _stamp, _content, environment_identity
from office_agent.webui.range_review import RangeReview
from office_agent.writes import GuardError


def test_resource_content_cache_is_bound_to_file_identity_and_changes(tmp_path):
    font = tmp_path / 'fixture.ttf'
    font.write_bytes(b'first-font-bytes')
    first = _stamp(font)
    old = _content(first)
    font.write_bytes(b'updated-font-bytes')
    second = _stamp(font)
    assert first != second and _content(second) != old
    assert _content(second) == _content(second)


def test_same_size_resource_change_with_restored_mtime_still_rehashes_bytes(tmp_path):
    path = tmp_path / 'resource.ttf'
    path.write_bytes(b'first')
    first = _stamp(path)
    old = _content(first)
    times = path.stat()
    path.write_bytes(b'other')
    os.utime(path, ns=(times.st_atime_ns, times.st_mtime_ns))
    updated = _stamp(path)
    assert first[3:5] == updated[3:5]
    assert _content(updated) != old


@pytest.mark.parametrize('current', [{'complete': False, 'sha256': None}, {'complete': True, 'sha256': 'new'}])
def test_unknown_or_changed_environment_never_returns_cached_pages(monkeypatch, current):
    closed = []
    service = object.__new__(RangeReview)
    service.cache = {'key': ['review', {}], 'timer': SimpleNamespace(cancel=lambda: closed.append('timer')),
                     'directory': SimpleNamespace(cleanup=lambda: closed.append('directory'))}
    monkeypatch.setattr('office_agent.fidelity.runtime_path', lambda: Path('known-runtime'))
    monkeypatch.setattr('office_agent.render_identity.environment_identity', lambda runtime: current)
    with pytest.raises(GuardError, match='professional_review_environment_changed'):
        service._pdfs('review', {}, (), {'render_environment': {'complete': True, 'sha256': 'old'}})
    assert service.cache is None and closed == ['timer', 'directory']


def test_invalid_runtime_cannot_make_reusable_environment(tmp_path):
    result = environment_identity(tmp_path / 'missing/program/soffice.com')
    assert not result['complete'] and result['sha256'] is None


@pytest.mark.parametrize('unknown_environment', [False, True])
def test_legacy_metadata_restores_without_authorizing_old_page_cache(tmp_path, monkeypatch, unknown_environment):
    from tests.test_range_delivery import prepared
    review, writer, editor, source, receipt, state = prepared(tmp_path, monkeypatch)
    if unknown_environment:
        monkeypatch.setattr('office_agent.render_identity.environment_identity', lambda runtime: {'complete': False, 'sha256': None})
    identifier = review.start(receipt, [])['result']['review_id']
    opened = review.open(identifier, True)
    assert opened['ok'] and not opened['result']['page_environment_verified']
    assert review.mark(identifier, 0, 'needs_review', '仅历史记录')['ok']
    assert review.page(identifier, 0, 'before')['reason'] == 'professional_review_environment_changed'
    assert review.cache is None
    review.invalidate()


def test_environment_change_requires_read_consent_and_revalidation_without_write(tmp_path, monkeypatch):
    from docx import Document
    from office_agent.webui.professional import Professional
    from office_agent.output_dependencies import OutputDependencies
    from office_agent.writes import Editor, digest
    from office_agent.range_edits import view
    from office_agent.range_alignment import anchors_for
    from office_agent.advanced_writes import _parts
    path = tmp_path / 'source.docx'
    doc = Document()
    doc.add_paragraph('旧术语仅限室内。')
    doc.save(path)
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    identity = [{'complete': True, 'sha256': 'first'}]
    monkeypatch.setattr('office_agent.render_identity.environment_identity', lambda runtime: identity[0])
    reports = []
    def gate(root, source, output, request):
        result = {'inputs_unchanged': True, 'render_environment': deepcopy(identity[0]),
                  'anchor_contexts': anchors_for(_parts(source, range_edit=True), request),
                  'render': {'status': 'different', 'anchor_alignment': [{'location': request.operations[0].locator_id,
                    'reason': 'test', 'action': 'replace', 'before_pages': [1], 'after_pages': [1], 'before_status': 'unique', 'after_status': 'unique'}]}}
        reports.append(result)
        return result
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', gate)
    scope = lambda: (True, False, str(path))
    writer = Professional(RLock(), scope, lambda: editor, lambda: path)
    writer.dependencies = dependencies = OutputDependencies()
    row = view(editor, path)['items'][0]
    plan = writer.plan_ranges({'source_sha256': digest(path), 'operations': [{'locator_id': row['locator_id'], 'expected_text_sha256': row['text_sha256'],
        'island_id': row['islands'][0]['id'], 'action': 'replace', 'start': 0, 'end': 3, 'quote': '旧术语', 'text': '规范术语', 'reason': '术语一致'}]})
    assert plan['ok'], plan
    saved = writer.confirm(plan['result']['id'], True)
    assert saved['ok'], saved
    service = RangeReview(RLock(), scope, lambda: editor, lambda: path, dependencies, writer.rollback)
    started = service.start(saved['result']['save_id'], [])
    assert started['ok'], started
    identifier = started['result']['review_id']
    assert service.open(identifier, True)['ok']
    assert service.mark(identifier, 0, 'viewed')['ok']
    audits_before = sorted(editor.audit.iterdir())
    source_before = digest(path)
    identity[0] = {'complete': True, 'sha256': 'updated'}
    assert service.page(identifier, 0, 'before')['reason'] == 'professional_review_environment_changed'
    calls = []
    def revalidate(*args, **kwargs):
        calls.append(kwargs)
        return {**deepcopy(reports[0]), 'render_environment': deepcopy(identity[0])}
    monkeypatch.setattr('office_agent.fidelity.compare_documents', revalidate)
    assert service.open(identifier, False)['result']['status'] == 'cancelled'
    assert not calls
    opened = service.open(identifier, True)
    assert opened['ok'], opened
    assert opened['result']['items'][0]['state'] == 'needs_review'
    assert len(calls) == 1 and calls[0]['static_docx']
    assert digest(path) == source_before and sorted(editor.audit.iterdir()) == audits_before
    service.invalidate()
