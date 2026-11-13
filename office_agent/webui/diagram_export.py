"""Authenticated current-receipt facade. Called only under Api's shared RLock.

Local IO is synchronous: cancellation/logout cannot preempt an admitted approved
write; they serialize after it. No model tool, client SVG, paths or execution claims.
"""
from pathlib import Path
import re
import uuid

from ..diagram_save import DiagramSaveError, DiagramSaveService
from ..diagram_raster import rasterize
from ..diagrams import DiagramError


class DiagramExport:
    def __init__(self, lock, scope, preview):
        self._lock, self._scope, self._preview = lock, scope, preview
        self._service = None
        self._native = None
        self._path = None
        self._digests = None
        self._raster = None
        self._disabled = False  # Sticky across preview changes, logout and login.

    def invalidate(self):
        with self._lock:
            if self._service is not None:
                self._service.invalidate()
            self._service = self._native = self._path = self._digests = self._raster = None

    def _current(self):
        authenticated, path, busy = self._scope()
        if not authenticated:
            self.invalidate()
            raise DiagramSaveError('not_authenticated')
        if busy:
            self.invalidate()
            raise DiagramSaveError('busy')
        if (self._service is None or path != self._path or self._digests is None
                or self._preview._for_export(*self._digests) is not self._native):
            self.invalidate()
            raise DiagramSaveError('export_operation_invalidated')
        return self._service

    @staticmethod
    def _id(value):
        if type(value) is not str or re.fullmatch(r'[0-9a-f]{32}', value) is None:
            raise DiagramSaveError('export_plan_invalid')
        return value

    def _call(self, operation):
        with self._lock:
            try:
                if self._disabled:
                    raise DiagramSaveError('export_service_disabled')
                result = operation()
                outcome = result.get('result', {})
                if (outcome.get('service_disabled') is True or outcome.get('code')
                        in {'export_rollback_partial', 'export_rollback_uncertain'}):
                    self._disabled = True
                return {'ok': True, **result}
            except DiagramSaveError as exc:
                if exc.code in {'export_cleanup_uncertain', 'export_service_disabled', 'renderer_cleanup_failed'}:
                    self._disabled = True
                return {'ok': False, 'reason': exc.code}
            except Exception:
                # An unknown exception may follow IO: retire all authority, never
                # claim that nothing was written, and never expose raw exception text.
                self._disabled = True
                self.invalidate()
                return {'ok': False, 'reason': 'export_failed', 'outcome_uncertain': True}

    def plan_save(self, source_sha256, svg_sha256):
        def operation():
            authenticated, path, busy = self._scope()
            if not authenticated:
                self.invalidate()
                raise DiagramSaveError('not_authenticated')
            if busy:
                self.invalidate()
                raise DiagramSaveError('busy')
            native = self._preview._for_export(source_sha256, svg_sha256)
            if native is None or path is None:
                raise DiagramSaveError('export_preview_invalid')
            if self._native is not native or self._path != path:
                self.invalidate()
                selected = Path(path)
                self._service = DiagramSaveService(selected.parent, selected, uuid.uuid4().hex)
                self._native, self._path = native, path
                self._digests = source_sha256, svg_sha256
            return {'plan': {**self._service.plan_save(native), 'selected_document': str(self._path)}}
        return self._call(operation)

    def plan_png(self, source_sha256, svg_sha256):
        def operation():
            authenticated, path, busy = self._scope()
            if not authenticated:
                self.invalidate()
                raise DiagramSaveError('not_authenticated')
            if busy:
                self.invalidate()
                raise DiagramSaveError('busy')
            native = self._preview._for_export(source_sha256, svg_sha256)
            if native is None or path is None:
                raise DiagramSaveError('export_preview_invalid')
            if self._native is not native or self._path != path:
                self.invalidate()
                selected = Path(path)
                self._service = DiagramSaveService(selected.parent, selected, uuid.uuid4().hex)
                self._native, self._path = native, path
                self._digests = source_sha256, svg_sha256
            if self._raster is None or self._raster.preview is not native:
                try:
                    self._raster = rasterize(native)
                except DiagramError as exc:
                    raise DiagramSaveError(exc.code) from None
            return {'plan': {**self._service.plan_png(native, self._raster),
                             'selected_document': str(self._path)}}
        return self._call(operation)

    def confirm_save(self, plan_id, confirmed=False):
        return self._call(lambda: {'result': self._current().confirm_save(self._id(plan_id), confirmed)})

    def plan_rollback(self, save_id):
        return self._call(lambda: {'plan': self._current().plan_rollback(self._id(save_id))})

    def confirm_rollback(self, rollback_id, confirmed=False):
        return self._call(lambda: {'result': self._current().confirm_rollback(self._id(rollback_id), confirmed)})
