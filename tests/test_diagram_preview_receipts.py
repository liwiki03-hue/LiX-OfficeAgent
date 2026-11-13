"""Synthetic typed OwnedRun fixtures; not native/provider/GUI provenance proof."""
from dataclasses import FrozenInstanceError, replace
from pathlib import Path

import pytest

from office_agent.webui.bridge import Api
from office_agent.diagram_service import RenderedDiagram
from office_agent.diagrams import DiagramError, validate_source, validate_svg
from test_diagram_preview import App, Auth, SOURCE, rendered, start_blocked


@pytest.fixture
def api():
    value = Api(App(), auth=Auth())
    value._diagram._render = rendered
    return value


def seed(api, proposed=False, source=SOURCE):
    response = (api.propose_diagram('flowchart', 'A flow') if proposed else
                api.render_diagram_source('flowchart', source))
    assert response['ok'] is True
    digests = (response['metadata']['source_sha256'], response['metadata']['svg_sha256'])
    with api._lock:
        native = api._diagram._for_export(*digests)
    assert type(native) is RenderedDiagram
    return response, digests, native


@pytest.mark.parametrize('proposed', [False, True])
def test_manual_and_proposal_retain_canonical_native(api, proposed):
    response, digests, native = seed(api, proposed)
    assert set(response) == {'ok', 'kind', 'source', 'svg', 'metadata'}
    assert native.source == validate_source(SOURCE, 'flowchart')
    assert native.svg == validate_svg(native.svg) == response['svg']
    assert api._diagram._for_export(*digests) is native
    assert 'elapsed_seconds' not in response['metadata']
    assert not hasattr(api, '_for_export')
    with pytest.raises(FrozenInstanceError):
        native.svg = 'changed'


def test_mutated_client_response_cannot_change_snapshot(api):
    response, digests, native = seed(api)
    response['svg'] = '<script>untrusted</script>'
    response['source'] = 'client source'
    response['metadata'].clear()
    response['execution'] = {'empty_confirmed': True}
    assert api._diagram._for_export(*digests) is native
    assert native.source.source == SOURCE and native.execution.empty_confirmed is True


def test_both_digests_strict_and_mismatch_fail_closed(api):
    _, digests, native = seed(api)
    bad_values = [None, 1, {}, '0' * 64, 'A' * 64, 'g' * 64,
                  digests[0] + '\n', 'a' * 63, 'a' * 65, 'a' * 1048577]
    for value in bad_values:
        assert api._diagram._for_export(value, digests[1]) is None
        assert api._diagram._for_export(digests[0], value) is None
    assert api._diagram._for_export(*digests) is native


def test_scope_checks_without_explicit_invalidation(api):
    _, digests, native = seed(api)
    original = api._app.path
    api._auth.authenticated = False
    assert api._diagram._for_export(*digests) is None
    api._auth.authenticated = True
    api._app.busy = True
    assert api._diagram._for_export(*digests) is None
    api._app.busy = False
    api._app.path = Path('different.docx')
    assert api._diagram._for_export(*digests) is None
    api._app.path = original
    api._diagram._disabled = True
    assert api._diagram._for_export(*digests) is None
    api._diagram._disabled = False
    scope = api._diagram._scope
    def failed_scope():
        raise RuntimeError('private scope')
    api._diagram._scope = failed_scope
    assert api._diagram._for_export(*digests) is None
    api._diagram._scope = scope
    assert api._diagram._for_export(*digests) is native


@pytest.mark.parametrize('method,args', [
    ('cancel_diagram', ()), ('logout', ()), ('reset_accounts', (True,)),
    ('ask', ('new task',)), ('select_file', ()), ('load_session', ('history',)),
    ('answer_clarification', ('id', {})),
])
def test_existing_dispatch_immediately_invalidates_receipt(api, method, args):
    _, digests, _ = seed(api)
    getattr(api, method)(*args)
    assert api._diagram._receipt is None
    assert api._diagram._for_export(*digests) is None


def test_admitted_invalid_new_preview_retires_old_before_validation(api):
    _, digests, _ = seed(api)
    api._diagram._render = lambda *a, **k: pytest.fail('must not render')
    assert api.render_diagram_source('flowchart', 'invalid')['reason'] == 'unsupported_source'
    assert api._diagram._receipt is None
    assert api._diagram._for_export(*digests) is None


def test_admitted_renderer_failure_cannot_take_back_old_receipt(api):
    _, digests, _ = seed(api)
    def failed(*args, **kwargs):
        assert api._diagram._receipt is None
        raise DiagramError('renderer_failed')
    api._diagram._render = failed
    assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'renderer_failed'
    assert api._diagram._receipt is None
    assert api._diagram._for_export(*digests) is None


def test_inflight_denied_cancel_holds_slot_and_late_completion_never_resurrects(api):
    _, digests, _ = seed(api)
    worker, release, output = start_blocked(api, render=True)
    try:
        token = api._diagram._slot
        assert api._diagram._receipt is None
        assert api._diagram._for_export(*digests) is None
        assert api.cancel_diagram()['ok'] is True
        assert api._diagram._slot is token
        assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'busy'
        assert api._diagram._for_export(*digests) is None
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive()
    assert output == [{'ok': False, 'reason': 'render_cancelled'}]
    assert api._diagram._slot is None and api._diagram._receipt is None
    assert api._diagram._for_export(*digests) is None


def test_success_finally_rebinds_epoch_without_reviving_guards(api):
    guards = []
    def capture(source, **kwargs):
        guards.append(kwargs['cancelled'])
        assert not guards[-1]()
        return rendered(source, **kwargs)
    api._diagram._render = capture
    before = api._diagram._epoch
    _, digests, native = seed(api)
    assert api._diagram._epoch is not before
    assert api._diagram._receipt[0] is api._diagram._epoch
    assert guards[0]() is True
    assert api._diagram._for_export(*digests) is native
    seed(api)
    assert len(guards) == 2 and all(guard() for guard in guards)
    assert api._diagram._slot is None


def test_cleanup_failure_disables_and_clears_receipt(api):
    _, digests, _ = seed(api)
    def uncertain(source, **kwargs):
        native = rendered(source, **kwargs)
        return replace(native, execution=replace(native.execution, empty_confirmed=False))
    api._diagram._render = uncertain
    assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'renderer_cleanup_failed'
    assert api._diagram._disabled and api._diagram._receipt is None
    assert api._diagram._for_export(*digests) is None
    assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'renderer_cleanup_failed'


def test_only_last_immutable_receipt_retained(api):
    _, old_digests, first = seed(api)
    _, new_digests, second = seed(api, source='flowchart TD\nB --> C')
    assert first is not second and old_digests[0] != new_digests[0]
    assert api._diagram._for_export(*old_digests) is None
    assert api._diagram._for_export(*new_digests) is second
    assert api._diagram._receipt[3] is second
    assert len(second.svg.encode('utf-8')) <= 1024 * 1024


def test_unauthenticated_cancel_retires_receipt_before_auth_check(api):
    _, digests, _ = seed(api)
    api._auth.authenticated = False
    assert api.cancel_diagram() == {'ok': False, 'reason': 'not_authenticated'}
    assert api._diagram._receipt is None
    api._auth.authenticated = True
    assert api._diagram._for_export(*digests) is None


@pytest.mark.parametrize('expanded', ['x' * 1048577, '\u00e9' * 524289],
                         ids=['ascii-overflow', 'utf8-byte-overflow'])
def test_serialized_output_bound_rechecked_after_validation(api, monkeypatch, expanded):
    # Synthetic validator expansion only; not an SVG grammar or native test.
    _, digests, _ = seed(api)
    monkeypatch.setattr('office_agent.webui.diagram_preview.validate_svg',
                        lambda svg: expanded)
    assert api.render_diagram_source('flowchart', SOURCE) == {
        'ok': False, 'reason': 'unsafe_svg'}
    assert api._diagram._slot is None and api._diagram._receipt is None
    assert api._diagram._for_export(*digests) is None
