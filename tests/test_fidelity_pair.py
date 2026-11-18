import json
from pathlib import Path

from docx import Document
import pytest

from office_agent import fidelity, verified_render_cache
from office_agent.diagram_process import OwnedRun
from tests.test_fidelity import pdf_bytes


@pytest.fixture
def rendering(tmp_path, monkeypatch):
    source, candidate, runtime = tmp_path / 'source.docx', tmp_path / 'candidate.docx', tmp_path / 'runtime.exe'
    book = Document()
    book.add_paragraph('Before 120 ms')
    book.save(source)
    book.paragraphs[0].text = 'After 120 ms'
    book.save(candidate)
    runtime.write_bytes(b'synthetic-runtime-not-executed')
    monkeypatch.setattr('office_agent.render_identity.environment_identity', lambda _: {'complete': True, 'sha256': 'fixed'})
    commands, offers = [], []
    actual_offer = verified_render_cache.offer
    def offer(identifier, data):
        offers.append(identifier)
        return actual_offer(identifier, data)
    monkeypatch.setattr(verified_render_cache, 'offer', offer)
    def run(command, **kwargs):
        commands.append(command)
        if command[2] == '--convert-pair':
            destination = Path(command[5])
            for source_name in command[3:5]:
                (destination / (Path(source_name).stem + '.pdf')).write_bytes(pdf_bytes())
        elif command[2] == '--convert':
            (Path(command[4]) / (Path(command[3]).stem + '.pdf')).write_bytes(pdf_bytes())
        else:
            Path(command[4]).write_text(json.dumps({'status': 'different', 'equivalent': False}))
        return OwnedRun('process_completed', 0, 42, .01, True)
    monkeypatch.setattr(fidelity, '_run_owned', run)
    verified_render_cache.clear()
    yield tmp_path, source, candidate, runtime, commands, offers, run
    verified_render_cache.clear()


def compare(rendering, **kwargs):
    root, source, candidate, runtime, *_ = rendering
    return fidelity.compare_documents(root, source, candidate, executable=runtime, static_docx=True, reuse_verified=True, **kwargs)


def test_default_stays_separate_and_candidate_has_isolated_cache_identity(rendering):
    original = compare(rendering)
    assert len(rendering[4]) == 3 and all(command[2] != '--convert-pair' for command in rendering[4])
    rendering[4].clear()
    result = compare(rendering, batch_conversion=True)
    assert len(rendering[4]) == 2 and rendering[4][0][2] == '--convert-pair'
    assert result['conversion_identity'] != original['conversion_identity']
    assert result['batch_conversion'] == {'requested': True, 'used': True, 'libreoffice_timeout_seconds': 25, 'automatic_fallback': False}
    assert result['inputs_unchanged'] and result['render']['status'] == 'different'


def test_candidate_refuses_all_per_input_cache_reads_and_offers(rendering, monkeypatch):
    def unexpected(identifier):
        raise AssertionError('candidate must not consume unproven per-input cache')
    monkeypatch.setattr(verified_render_cache, 'get', unexpected)
    result = compare(rendering, batch_conversion=True)
    assert result['batch_conversion']['used'] is True
    assert len(rendering[4]) == 2 and rendering[4][0][2] == '--convert-pair'
    assert result['conversion_cache'] == {'requested': True, 'effective': False, 'reason': 'candidate_partner_identity_unproven'}
    assert rendering[5] == []
    assert result['render']['status'] == 'different'


def test_batch_missing_second_pdf_never_compares_or_caches_partial_output(rendering, monkeypatch):
    def partial(command, **kwargs):
        rendering[4].append(command)
        (Path(command[5]) / (Path(command[3]).stem + '.pdf')).write_bytes(pdf_bytes())
        return OwnedRun('process_completed', 0, 42, .01, True)
    monkeypatch.setattr(fidelity, '_run_owned', partial)
    result = compare(rendering, batch_conversion=True)
    assert len(rendering[4]) == 1 and not rendering[5]
    assert result['render']['status'] == 'error' and result['inputs_unchanged']


@pytest.mark.parametrize('failure', ['timeout', 'cleanup'])
def test_batch_failure_has_no_fallback_or_cache_offer(rendering, monkeypatch, failure):
    def failed(command, **kwargs):
        rendering[4].append(command)
        if failure == 'cleanup':
            raise RuntimeError('owned cleanup not confirmed')
        return OwnedRun('render_timeout', None, 42, .1, True)
    monkeypatch.setattr(fidelity, '_run_owned', failed)
    result = compare(rendering, batch_conversion=True)
    assert len(rendering[4]) == 1 and not rendering[5]
    assert result['render']['status'] == 'error' and result['inputs_unchanged']


def test_changed_environment_rejects_pair_and_does_not_cache(rendering, monkeypatch):
    identities = iter([{'complete': True, 'sha256': 'before'}, {'complete': True, 'sha256': 'after'}])
    monkeypatch.setattr('office_agent.render_identity.environment_identity', lambda _: next(identities))
    result = compare(rendering, batch_conversion=True)
    assert result['render']['code'] == 'render_environment_changed' and not rendering[5]


def test_pair_cancels_owned_pdf_comparison_without_treating_cleanup_failure_as_cancel(rendering, monkeypatch):
    root, source, candidate, runtime, commands, offers, worker = rendering
    def run(command, **kwargs):
        assert callable(kwargs['cancelled'])
        if command[2] == '--convert-pair':
            return worker(command, **kwargs)
        commands.append(command)
        return OwnedRun('render_cancelled', None, 42, .1, False)
    monkeypatch.setattr(fidelity, '_run_owned', run)
    result = fidelity.compare_documents(root, source, candidate, executable=runtime, static_docx=True,
                                       reuse_verified=True, batch_conversion=True, cancelled=lambda: False)
    assert len(commands) == 2 and not offers and result['inputs_unchanged']
    assert result['render']['code'] == 'worker_launch_or_cleanup_failed'


def test_render_preflight_rejects_unsafe_package_before_batch_launch(rendering):
    import zipfile
    with zipfile.ZipFile(rendering[1], 'a') as archive:
        archive.writestr('custom/_rels/rejected.rels', '<Relationships><Relationship TargetMode="External" Target="https://example.invalid"/></Relationships>')
    with pytest.raises(ValueError, match='professional_unsafe_document'):
        compare(rendering, batch_conversion=True)
    assert not rendering[4] and not rendering[5]
