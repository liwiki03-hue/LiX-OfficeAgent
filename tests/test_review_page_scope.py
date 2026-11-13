from copy import deepcopy
import json
from pathlib import Path
from threading import RLock

from docx import Document
import pytest

from office_agent import fidelity, render_handoff
from office_agent.output_dependencies import OutputDependencies
from office_agent.webui.professional import Professional
from office_agent.webui.range_review import RangeReview
from office_agent.writes import Editor, GuardError, digest
from tests.test_fidelity import pdf_bytes
from tests.test_range_edits import operation


@pytest.fixture
def prepared_review(tmp_path, monkeypatch):
    source = tmp_path / 'source.docx'
    document = Document(); document.add_paragraph('独特条件 原结论 仅限室内。'); document.save(source)
    environment, observations = [{'complete': True, 'sha256': 'known'}], []
    runtime = tmp_path / 'runtime.com'; runtime.write_bytes(b'owned-runtime')
    monkeypatch.setattr(fidelity, 'runtime_path', lambda *args: runtime)
    def identity(runtime):
        observations.append(True)
        return deepcopy(environment[0])
    monkeypatch.setattr('office_agent.render_identity.environment_identity', identity)
    def convert(path, destination, *args):
        destination.mkdir(); pdf = destination / 'result.pdf'; pdf.write_bytes(pdf_bytes()); return pdf
    monkeypatch.setattr(fidelity, '_convert', convert)
    def worker(command, *args):
        Path(command[4]).write_text(json.dumps({'status': 'different', 'anchor_alignment': [
            {'location': 'owned', 'reason': 'test', 'action': 'replace', 'before_pages': [1], 'after_pages': [1],
             'before_status': 'unique', 'after_status': 'unique'}]}), encoding='utf8')
    monkeypatch.setattr(fidelity, '_run', worker)
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    scope = lambda: (True, False, str(source))
    dependencies = OutputDependencies()
    writer = Professional(RLock(), scope, lambda: editor, lambda: source); writer.dependencies = dependencies
    row = writer.read_ranges()['result']['items'][0]
    plan = writer.plan_ranges({'source_sha256': digest(source), 'operations': [operation(row, 5, 8, '新结论')]})
    saved = writer.confirm(plan['result']['id'], True)
    assert saved['ok'], saved
    render_handoff.clear()
    review = RangeReview(RLock(), scope, lambda: editor, lambda: source, dependencies, writer.rollback)
    identifier = review.start(saved['result']['save_id'], [])['result']['review_id']
    assert review.open(identifier, True)['ok']
    observations.clear()
    yield review, identifier, source, environment, observations
    review.invalidate()


@pytest.mark.parametrize('failure', ['none', 'environment', 'source', 'raster'])
def test_first_page_publishes_cache_only_after_final_validation(prepared_review, monkeypatch, failure):
    review, identifier, source, environment, observations = prepared_review
    directories = []
    def raster(pdf, *args):
        if not directories:
            assert review.cache is None
        directories.append(pdf.parent.parent)
        if failure == 'environment': environment[0] = {'complete': True, 'sha256': 'changed'}
        if failure == 'source': source.write_bytes(source.read_bytes()+b'changed')
        if failure == 'raster': raise GuardError('owned_raster_failure')
        return {'page': 1, 'image': 'owned-fixture'}
    monkeypatch.setattr('office_agent.research_page.render_page', raster)
    reply = review.page(identifier, 0, 'before')
    if failure == 'none':
        assert reply['ok'], reply
        assert len(observations) == 2 and review.cache is not None
        observations.clear()
        assert review.page(identifier, 0, 'after')['ok']
        assert len(observations) == 2
        review.invalidate()
    else:
        assert not reply['ok'] and review.cache is None
        assert 'image' not in reply.get('result', {})
    assert directories and all(not path.exists() for path in directories)
