from copy import deepcopy
from pathlib import Path

import pytest

from office_agent import render_handoff as handoff
from tests.test_fidelity import pdf_bytes


@pytest.fixture(autouse=True)
def isolated_handoff():
    handoff.clear()
    yield
    handoff.clear()


def prepared(tmp_path):
    pdfs = [tmp_path / 'before.pdf', tmp_path / 'after.pdf']
    for pdf in pdfs:
        pdf.write_bytes(pdf_bytes())
    report = {'inputs_unchanged': True, 'source_sha256': 'source', 'output_sha256': 'output',
              'static_projection_sha256': ['first', 'second'], 'runtime_sha256': 'runtime',
              'render_environment': {'complete': True, 'sha256': 'fonts'},
              'render': {'status': 'different'}, 'conversion_identity': handoff.implementation_identity()}
    return report, pdfs


def test_single_use_bounded_handoff_and_clear(tmp_path):
    report, pdfs = prepared(tmp_path)
    assert handoff.offer(tmp_path, report, pdfs)
    assert handoff.take(tmp_path, report) == tuple(path.read_bytes() for path in pdfs)
    assert handoff.take(tmp_path, report) is None
    assert handoff.offer(tmp_path, report, pdfs)
    handoff.clear()
    assert handoff.take(tmp_path, report) is None


@pytest.mark.parametrize('field,value', [('source_sha256', 'changed'), ('output_sha256', 'changed'),
    ('static_projection_sha256', ['changed', 'second']), ('runtime_sha256', 'changed'),
    ('render_environment', {'complete': True, 'sha256': 'changed'}), ('conversion_identity', 'changed'),
    ('inputs_unchanged', False), ('render', {'status': 'error'})])
def test_every_identity_change_refuses_and_retires(tmp_path, field, value):
    report, pdfs = prepared(tmp_path)
    assert handoff.offer(tmp_path, report, pdfs)
    modified = deepcopy(report)
    modified[field] = value
    assert handoff.take(tmp_path, modified) is None
    assert handoff.take(tmp_path, report) is None


def test_scope_expiry_corruption_capacity_unknown_and_old_timer(tmp_path, monkeypatch):
    report, pdfs = prepared(tmp_path)
    assert handoff.offer(tmp_path, report, pdfs)
    old_key, old_deadline = handoff._entry['key'], handoff._entry['deadline']
    assert handoff.offer(tmp_path, report, pdfs)
    handoff._expire(old_key, old_deadline)
    assert handoff._entry is not None
    assert handoff.take(tmp_path / 'other', report) is None
    assert handoff.offer(tmp_path, report, pdfs)
    monkeypatch.setattr(handoff.time, 'monotonic', lambda: handoff._entry['deadline'] + 1)
    assert handoff.take(tmp_path, report) is None
    monkeypatch.undo()
    assert handoff.offer(tmp_path, report, pdfs)
    handoff._entry['payloads'] = (b'corrupt', b'corrupt')
    assert handoff.take(tmp_path, report) is None
    monkeypatch.setattr(handoff, 'MAX_BYTES', 1)
    assert not handoff.offer(tmp_path, report, pdfs)
    monkeypatch.undo()
    report['render_environment']['complete'] = False
    assert not handoff.offer(tmp_path, report, pdfs)
    del report['conversion_identity']
    assert handoff.take(tmp_path, report) is None


def test_verified_gate_handoff_to_review_requires_consent_and_current_inputs(tmp_path, monkeypatch):
    from threading import RLock
    from docx import Document
    from office_agent.webui.professional import Professional
    from office_agent.webui.range_review import RangeReview
    from office_agent.output_dependencies import OutputDependencies
    from office_agent.writes import Editor, digest
    from tests.test_range_edits import operation
    from office_agent import fidelity
    from office_agent import research_page
    monkeypatch.setattr(research_page, '_run', fidelity._run)
    source = tmp_path / 'source.docx'
    doc = Document()
    doc.add_paragraph('独特条件 原结论 仅限室内。')
    doc.save(source)
    environment = {'complete': True, 'sha256': 'known'}
    runtime = tmp_path / 'runtime.com'
    runtime.write_bytes(b'owned-runtime')
    monkeypatch.setattr(fidelity, 'runtime_path', lambda *args: runtime)
    monkeypatch.setattr('office_agent.render_identity.environment_identity', lambda runtime: deepcopy(environment))
    conversions = []
    def convert(path, destination, *args):
        destination.mkdir()
        pdf = destination / 'result.pdf'
        pdf.write_bytes(pdf_bytes())
        conversions.append(True)
        return pdf
    monkeypatch.setattr(fidelity, '_convert', convert)
    def worker(command, *args):
        import json
        Path(command[4]).write_text(json.dumps({'status': 'different', 'anchor_alignment': [
            {'location': 'owned', 'reason': 'test', 'action': 'replace', 'before_pages': [1], 'after_pages': [1],
             'before_status': 'unique', 'after_status': 'unique'}]}), encoding='utf8')
    monkeypatch.setattr(fidelity, '_run', worker)
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    authenticated = [True]
    scope = lambda: (authenticated[0], False, str(source))
    dependencies = OutputDependencies()
    writer = Professional(RLock(), scope, lambda: editor, lambda: source)
    writer.dependencies = dependencies
    row = writer.read_ranges()['result']['items'][0]
    planned = writer.plan_ranges({'source_sha256': digest(source), 'operations': [operation(row, 5, 8, '新结论')]})
    saved = writer.confirm(planned['result']['id'], True)
    assert saved['ok'], saved
    assert len(conversions) == 2 and handoff._entry
    review = RangeReview(RLock(), scope, lambda: editor, lambda: source, dependencies, writer.rollback)
    identifier = review.start(saved['result']['save_id'], [])['result']['review_id']
    assert review.page(identifier, 0, 'before')['reason'] == 'professional_review_read_consent'
    assert handoff._entry
    assert review.open(identifier, True)['ok']
    first = review.page(identifier, 0, 'before')
    assert first['ok'], first
    assert first['result']['conversion_reused'] and len(conversions) == 2
    assert review.page(identifier, 0, 'after')['ok'] and len(conversions) == 2
    review.invalidate()
    assert handoff._entry is None and review.cache is None
    identifier = review.start(saved['result']['save_id'], [])['result']['review_id']
    assert review.open(identifier, True)['ok']
    assert review.page(identifier, 0, 'before')['ok'] and len(conversions) == 4
    source.write_bytes(source.read_bytes() + b'changed')
    assert not review.page(identifier, 0, 'after')['ok']
    authenticated[0] = False
    assert review.page(identifier, 0, 'after')['reason'] == 'not_authenticated'
    assert review.cache is None and handoff._entry is None
    review.invalidate()
