"""Backend API tests: synthetic models/renderers, no provider/native Office."""
from dataclasses import replace
from pathlib import Path
import threading

from langchain_core.messages import AIMessage
import pytest

from office_agent.webui.bridge import Api, WebviewApp
from office_agent.diagram_process import OwnedRun
from office_agent.diagram_service import RenderedDiagram
from office_agent.diagrams import DiagramError, validate_source, validate_svg

SOURCE = 'flowchart TD\nA --> B'
SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text x="1" y="10">A</text></svg>'


class Auth:
    def __init__(self, authenticated=True):
        self.authenticated = authenticated
    def status(self):
        return {'authenticated': self.authenticated}
    def login(self, *args):
        self.authenticated = True
        return {'ok': True, 'authenticated': True}
    def logout(self):
        self.authenticated = False
        return {'ok': True}
    def reset_accounts(self, confirmed=False):
        if confirmed:
            self.authenticated = False
        return {'ok': confirmed}


class Model:
    def invoke(self, messages):
        return AIMessage(content=SOURCE)


class App:
    def __init__(self):
        self.path = Path('selected.docx')
        self.busy = False
        self.factory_calls = 0
        self.model = Model()
    def _build_model(self):
        self.factory_calls += 1
        return self.model
    def _clear_private_state(self):
        self.path = None
    def select_file(self):
        return {'ok': True}  # Same-path selection still invalidates.
    def ask(self, question):
        return {'ok': True}
    def load_session(self, directory):
        return {'ok': True}
    def answer_clarification(self, *args):
        return {'ok': True}


def rendered(source, **kwargs):
    return RenderedDiagram(validate_source(source, kwargs['expected_kind']), SVG,
                           OwnedRun('process_completed', 0, 1, 0.1, True))


@pytest.fixture
def api():
    value = Api(App(), auth=Auth())
    value._diagram._render = rendered
    return value


def test_real_controller_receives_no_events_history_or_task_state():
    app = WebviewApp()
    app.path = Path('selected.docx')
    app._build_model = lambda: Model()
    api = Api(app, auth=Auth())
    api._diagram._render = rendered
    assert api._diagram._lock is api._lock
    assert api.propose_diagram('flowchart', 'A')['ok'] is True
    assert api.render_diagram_source('flowchart', SOURCE)['ok'] is True
    assert app.poll() == [] and app._history_events == []
    assert app.events.empty() and app._history_dir is None
    assert app.model_agent is None and app.pending_id is None and not app.busy


def test_proposed_tool_response_is_rejected_without_render(api):
    class ToolModel:
        def invoke(self, messages):
            return AIMessage(content=SOURCE, tool_calls=[
                {'name': 'export', 'args': {}, 'id': 'blocked'}])
    api._app.model = ToolModel()
    api._diagram._render = lambda *a, **k: pytest.fail('must not render tool response')
    assert api.propose_diagram('flowchart', 'A') == {
        'ok': False, 'reason': 'diagram_repair_exhausted'}
    assert api._diagram._slot is None


def test_normal_manual_and_proposed_preview(api):
    manual = api.render_diagram_source('flowchart', SOURCE)
    assert manual['ok'] is True
    assert set(manual) == {'ok', 'kind', 'source', 'svg', 'metadata'}
    assert manual['svg'] == validate_svg(SVG)
    assert manual['source'] == SOURCE and manual['kind'] == 'flowchart'
    assert manual['metadata']['empty_confirmed'] is True
    assert api._app.factory_calls == 0 and api._app.busy is False
    proposed = api.propose_diagram('flowchart', 'A flow')
    assert proposed == manual and api._app.factory_calls == 1
    assert not hasattr(api, 'diagram_call')


@pytest.mark.parametrize('operation', [lambda a: a.propose_diagram('flowchart', 'A'),
    lambda a: a.render_diagram_source('flowchart', SOURCE), lambda a: a.cancel_diagram()])
def test_auth_denied(api, operation):
    api._auth.authenticated = False
    assert operation(api) == {'ok': False, 'reason': 'not_authenticated'}
    assert api._app.factory_calls == 0


@pytest.mark.parametrize('kind,value', [(None, 'A'), ([], 'A'), ('unknown', 'A'),
    ('flowchart', None), ('flowchart', ''), ('flowchart', 'x' * 1001)])
def test_bad_proposal_before_factory(api, kind, value):
    assert api.propose_diagram(kind, value) == {'ok': False, 'reason': 'unsupported_source'}
    assert api._app.factory_calls == 0


@pytest.mark.parametrize('kind,source', [(None, SOURCE), ([], SOURCE), ('sequenceDiagram', SOURCE),
    ('flowchart', None), ('flowchart', 'flowchart TD\nclick A "https://evil"'),
    ('flowchart', 'x' * 8193)])
def test_bad_source_before_factory_or_renderer(api, kind, source):
    api._diagram._render = lambda *a, **k: pytest.fail('renderer must not run')
    assert api.render_diagram_source(kind, source) == {'ok': False, 'reason': 'unsupported_source'}
    assert api._app.factory_calls == 0


def start_blocked(api, *, render=False):
    entered, release = threading.Event(), threading.Event()
    output = []
    if render:
        def operation(source, **kwargs):
            entered.set()
            assert release.wait(5)
            return rendered(source, **kwargs)
        api._diagram._render = operation
        call = lambda: api.render_diagram_source('flowchart', SOURCE)
    else:
        class BlockedModel:
            def invoke(self, messages):
                entered.set()
                assert release.wait(5)
                return AIMessage(content=SOURCE)
        api._app.model = BlockedModel()
        call = lambda: api.propose_diagram('flowchart', 'A flow')
    worker = threading.Thread(target=lambda: output.append(call()))
    worker.start()
    assert entered.wait(5)
    return worker, release, output


@pytest.mark.parametrize('render', [False, True])
def test_cancel_keeps_actual_owned_slot_until_return(api, render):
    worker, release, output = start_blocked(api, render=render)
    try:
        token = api._diagram._slot
        assert api.cancel_diagram() == {'ok': True}
        assert api._diagram._slot is token and api._app.busy is False
        assert api.propose_diagram('flowchart', 'New') == {'ok': False, 'reason': 'busy'}
        assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'busy'
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive() and output == [{'ok': False, 'reason': 'render_cancelled'}]
    assert api._diagram._slot is None
    assert api.render_diagram_source('flowchart', SOURCE)['ok'] is True


@pytest.mark.parametrize('mutation', [lambda a: a.select_file(), lambda a: a.ask('task'),
    lambda a: a.load_session('history'), lambda a: a.answer_clarification('id', {}),
    lambda a: a._clear_private_state(), lambda a: setattr(a._app, 'path', Path('different.docx')),
    lambda a: setattr(a._app, 'busy', True), lambda a: a.reset_accounts(True)])
def test_file_task_clear_and_busy_fence_late_results(api, mutation):
    worker, release, output = start_blocked(api)
    try:
        mutation(api)
        assert api._diagram._slot is not None
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive() and output == [{'ok': False, 'reason': 'render_cancelled'}]
    assert api._diagram._slot is None


def test_logout_and_relogin_allowed_while_model_blocked(api):
    worker, release, output = start_blocked(api)
    try:
        assert api.logout()['ok'] is True
        assert api.login('local', 'test')['ok'] is True
        assert api.propose_diagram('flowchart', 'New')['reason'] == 'busy'
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive() and output == [{'ok': False, 'reason': 'render_cancelled'}]
    api._app.model = Model()
    assert api.propose_diagram('flowchart', 'New')['ok'] is True


def test_normal_busy_preserves_auth_behavior(api):
    api._app.busy = True
    assert api.propose_diagram('flowchart', 'A')['reason'] == 'busy'
    assert api.logout()['reason'] == 'busy'
    assert api._auth.authenticated is True
    assert api._app.factory_calls == 0


@pytest.mark.parametrize('change,reason', [
    (lambda r: {}, 'diagram_preview_mismatch'),
    (lambda r: replace(r, source=validate_source('flowchart TD\nB --> C')), 'diagram_preview_mismatch'),
    (lambda r: replace(r, svg='<script>secret</script>'), 'unsafe_svg'),
    (lambda r: replace(r, execution=replace(r.execution, exit_code=True)), 'diagram_preview_mismatch'),
    (lambda r: replace(r, execution=replace(r.execution, exit_code=1)), 'diagram_preview_mismatch'),
    (lambda r: replace(r, execution=replace(r.execution, code='process_failed')), 'diagram_preview_mismatch'),
    (lambda r: replace(r, execution=replace(r.execution, empty_confirmed=False)), 'renderer_cleanup_failed'),
    (lambda r: replace(r, execution={}), 'diagram_preview_mismatch'),
])
def test_malformed_manual_result_fail_closed_and_releases_slot(api, change, reason):
    api._diagram._render = lambda source, **kw: change(rendered(source, **kw))
    assert api.render_diagram_source('flowchart', SOURCE) == {'ok': False, 'reason': reason}
    assert api._diagram._slot is None
    api._diagram._render = rendered
    if reason == 'renderer_cleanup_failed':
        assert api.render_diagram_source('flowchart', SOURCE)['reason'] == reason
    else:
        assert api.render_diagram_source('flowchart', SOURCE)['ok'] is True


@pytest.mark.parametrize('error,reason', [(RuntimeError('SECRET /private/file'), 'diagram_failed'),
    (DiagramError('SECRET', 'private'), 'diagram_failed'),
    (DiagramError('renderer_cleanup_failed', 'private'), 'renderer_cleanup_failed')])
def test_errors_sanitized_and_slot_finalized(api, error, reason):
    def failed(*args, **kwargs):
        raise error
    api._diagram._render = failed
    assert api.render_diagram_source('flowchart', SOURCE) == {'ok': False, 'reason': reason}
    assert api._diagram._slot is None


def test_factory_error_sanitized(api):
    def failed():
        raise RuntimeError('SECRET provider')
    api._diagram._model_factory = failed
    assert api.propose_diagram('flowchart', 'A') == {'ok': False, 'reason': 'diagram_failed'}
    assert api._diagram._slot is None


def test_owned_cleanup_cannot_unlock_new_operation(api):
    worker, release, output = start_blocked(api, render=True)
    try:
        api.cancel_diagram()
        assert api._diagram._slot is not None
        assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'busy'
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive()
    second, release_second, output_second = start_blocked(api, render=True)
    try:
        token = api._diagram._slot
        api.cancel_diagram()
        assert api._diagram._slot is token
        assert api.render_diagram_source('flowchart', SOURCE)['reason'] == 'busy'
    finally:
        release_second.set()
        second.join(5)
    assert not second.is_alive()
    assert output == output_second == [{'ok': False, 'reason': 'render_cancelled'}]
    assert api._diagram._slot is None


@pytest.mark.parametrize('method,args', [
    ('execute', ()), ('undo', ()),
    ('confirm_image_generation', ('plan', True)),
    ('confirm_image_save', ('save', True)),
    ('confirm_image_rollback', ('rollback', True)),
])
def test_immediate_same_path_action_invalidates_without_busy_interval(api, method, args):
    # Check real production method names, then isolate their immediate dispatch.
    assert callable(getattr(WebviewApp, method))
    calls = []
    original_path = api._app.path
    def immediate(*received):
        assert api._app.path == original_path and api._app.busy is False
        calls.append(received)
        return {'ok': True}
    setattr(api._app, method, immediate)
    worker, release, output = start_blocked(api)
    try:
        slot = api._diagram._slot
        assert getattr(api, method)(*args) == {'ok': True}
        assert calls == [args]
        assert api._app.path == original_path and api._app.busy is False
        assert api._diagram._slot is slot
        assert api.propose_diagram('flowchart', 'Another')['reason'] == 'busy'
    finally:
        release.set()
        worker.join(5)
    assert not worker.is_alive()
    assert output == [{'ok': False, 'reason': 'render_cancelled'}]
    assert api._diagram._slot is None
