"""Private authenticated, read-only diagram facade; no events/history or exports.

Cancellation fences publication, not provider execution. The owned slot stays held
until the actual synchronous operation returns, including after logout/relogin.
"""
import re

from ..diagram_process import OwnedRun
from ..diagram_proposals import propose_diagram
from ..diagram_service import RenderedDiagram, render_diagram
from ..diagrams import DiagramError, diagram_prompt, validate_source, validate_svg

_CODES = frozenset({
    'unsupported_source', 'unsafe_svg', 'renderer_failed', 'renderer_unavailable',
    'render_timeout', 'render_cancelled', 'renderer_cleanup_failed',
    'diagram_preview_mismatch', 'diagram_model_failed', 'diagram_repair_exhausted',
})


class DiagramPreview:
    """scope() returns (authenticated, path, busy), only under the shared Api lock.

    Underscored model/render injection is internal test plumbing, never js_api.
    """
    def __init__(self, lock, scope, model_factory, *, _render=None):
        self._lock = lock
        self._scope = scope
        self._model_factory = model_factory
        self._render = render_diagram if _render is None else _render
        self._epoch = object()
        self._slot = None
        self._disabled = False
        self._receipt = None  # At most one immutable (epoch, owner, path, native) snapshot.

    def invalidate(self):
        with self._lock:
            self._receipt = None
            self._epoch = object()  # Never release an active operation's slot here.

    def cancel(self):
        with self._lock:
            self._receipt = None  # Also retire authority when authentication has already changed.
            if not self._scope()[0]:
                return {'ok': False, 'reason': 'not_authenticated'}
            self.invalidate()
            return {'ok': True}

    def _current(self, epoch, token, path):
        authenticated, current_path, busy = self._scope()
        return (authenticated and not busy and current_path == path
                and self._epoch is epoch and self._slot is token)

    def _cancelled(self, epoch, token, path):
        with self._lock:
            try:
                return not self._current(epoch, token, path)
            except Exception:
                return True

    def _for_export(self, expected_source_sha256, expected_svg_sha256):
        """Host-only lookup, NOT writer authority; caller binds save/task/account
        session under the shared Api RLock. No client SVG or execution claims.
        """
        with self._lock:
            try:
                if any(type(d) is not str or len(d) != 64 or
                       re.fullmatch(r'[0-9a-f]{64}', d) is None
                       for d in (expected_source_sha256, expected_svg_sha256)):
                    return None
                authenticated, path, busy = self._scope()
                receipt = self._receipt
                if (not authenticated or busy or self._disabled or self._slot is not None
                        or receipt is None or receipt[0] is not self._epoch or receipt[2] != path):
                    return None
                native = receipt[3]
                spec = validate_source(native.source.source, native.source.kind)
                run = native.execution
                if (type(native) is not RenderedDiagram or type(run) is not OwnedRun
                        or native.source != spec or run.empty_confirmed is not True
                        or run.code != 'process_completed' or type(run.exit_code) is not int
                        or run.exit_code != 0 or validate_svg(native.svg) != native.svg
                        or len(native.svg.encode('utf-8')) > 1024 * 1024):
                    return None
                evidence = native.evidence()
                if (spec.sha256 != expected_source_sha256
                        or evidence['svg_sha256'] != expected_svg_sha256):
                    return None
                return native
            except Exception:
                return None

    def propose(self, kind, goal):
        return self._run(kind, goal, proposed=True)

    def render(self, kind, source):
        return self._run(kind, source, proposed=False)

    def _run(self, kind, value, *, proposed):
        with self._lock:
            authenticated, path, busy = self._scope()
            if not authenticated:
                return {'ok': False, 'reason': 'not_authenticated'}
            if busy or self._slot is not None:
                return {'ok': False, 'reason': 'busy'}
            if self._disabled:
                return {'ok': False, 'reason': 'renderer_cleanup_failed'}
            self._receipt = None  # Admission retires the last preview even on validation failure.
            try:
                # Validate before even configuring a provider or invoking a renderer.
                diagram_prompt(kind, value if proposed else 'Manual preview')
                spec = None if proposed else validate_source(value, kind)
            except (DiagramError, TypeError, ValueError):
                return {'ok': False, 'reason': 'unsupported_source'}
            epoch, token = self._epoch, object()
            self._slot = token
        cancelled = lambda: self._cancelled(epoch, token, path)
        try:
            if cancelled():
                raise DiagramError('render_cancelled')
            if proposed:
                model = self._model_factory()
                if cancelled():
                    raise DiagramError('render_cancelled')
                proposal = propose_diagram(model, kind=kind, goal=value,
                                           cancelled=cancelled, _render=self._render)
                result = proposal.preview
                spec = validate_source(result.source.source, kind)
            else:
                result = self._render(spec.source, expected_kind=kind, cancelled=cancelled)
            # Manual previews must satisfy precisely the same execution/source checks
            # as model proposals; adapter claims are not a trust boundary.
            if type(result) is not RenderedDiagram or type(result.execution) is not OwnedRun:
                raise DiagramError('diagram_preview_mismatch')
            if result.execution.empty_confirmed is not True:
                raise DiagramError('renderer_cleanup_failed')
            if (result.source != spec or result.execution.code != 'process_completed'
                    or type(result.execution.exit_code) is not int or result.execution.exit_code != 0):
                raise DiagramError('diagram_preview_mismatch')
            safe_svg = validate_svg(result.svg)
            if len(safe_svg.encode('utf-8')) > 1024 * 1024:
                raise DiagramError('unsafe_svg')
            # Regenerate metadata from validated data; never return adapter extras.
            native = RenderedDiagram(spec, safe_svg, result.execution)
            metadata = native.evidence()
            metadata.pop('elapsed_seconds')  # Untrusted execution measurement is not display data.
            with self._lock:
                if not self._current(epoch, token, path):
                    return {'ok': False, 'reason': 'render_cancelled'}
                self._receipt = (epoch, token, path, native)
                return {'ok': True, 'kind': spec.kind, 'source': spec.source,
                        'svg': safe_svg, 'metadata': metadata}
        except DiagramError as exc:
            code = exc.code if type(exc.code) is str and exc.code in _CODES else 'diagram_failed'
            if code == 'renderer_cleanup_failed':
                with self._lock:
                    self._receipt = None
                    self._disabled = True  # Cleanup uncertainty requires a new facade/process.
            return {'ok': False, 'reason': code}
        except Exception:
            return {'ok': False, 'reason': 'diagram_failed'}
        finally:
            with self._lock:
                if self._slot is token:
                    self._slot = None
                    # Nothing else can release this slot, even scope invalidation.
                    # Advancing epoch also makes retained guards permanently stale.
                    retained = self._receipt
                    preserve = (self._epoch is epoch and retained is not None
                                and retained[0] is epoch and retained[1] is token)
                    self._epoch = object()
                    if preserve:
                        self._receipt = (self._epoch, token, path, retained[3])
