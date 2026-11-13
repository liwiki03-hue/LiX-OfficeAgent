"""HOST-ONLY read-only SVG export approval capsule (saving NOT implemented).

Typed objects are trusted-host provenance, not authentication of client metadata.
A future bridge MUST use a server-held RenderedDiagram bound to task/account/session,
never reconstruct one from client SVG/source/execution fields. Nothing here mounts
an API, renders, invokes providers, creates directories, audits, saves or deletes.

Source guards observe snapshots, not mutation history or an OS-atomic transaction;
they do not prove safety against a same-privilege racing attacker. A successful
approval check still returns saved=False / export_saving_unimplemented. No approval
or bytes can be retrieved for reuse by a caller. Rollback authority does not exist.
"""
from __future__ import annotations

from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import re
import threading
import time
import uuid

from .diagram_process import OwnedRun
from .diagram_service import RenderedDiagram
from .diagrams import DiagramSource, validate_source, validate_svg
from .writes import Workspace, _audit_path

MAX_PLANS = 8
APPROVAL_TTL = 300.0
MAX_SOURCE_BYTES = 100_000_000
MAX_SVG_BYTES = 1_048_576


class DiagramExportError(ValueError):
    """Fixed safe codes only; never expose paths, raw SVG/source or exceptions."""
    def __init__(self, code: str):
        self.code = code
        super().__init__(code)


def _identity(info):
    return info.st_dev, info.st_ino


def _stamp(info):
    return (*_identity(info), info.st_size, info.st_mtime_ns, info.st_ctime_ns)


@dataclass(frozen=True)
class _Plan:
    metadata: dict
    expires: float


class DiagramExportService:
    """Selected document parent only; memory-only one-use bounded approvals.

    Constructor reads a bounded fingerprint but performs no filesystem writes.
    All public results are metadata. Confirm consumes the plan even on failure.
    Cancel/expiry/invalidate release approval slots; no history/restart authority.
    """
    def __init__(self, root: Path, source: Path, session_id: str):
        if type(session_id) is not str or not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", session_id):
            raise DiagramExportError("export_session_invalid")
        self._session = session_id
        self._lock = threading.RLock()
        self._active = True
        self._plans: dict[str, _Plan] = {}
        try:
            _audit_path(root, directory=True)
            self._space = Workspace(root)
            self._root = self._space.root
            self._root_identity = _identity(self._root.lstat())
            _audit_path(source, directory=False)
            self._source = self._space.file(str(source))
            if self._source.parent != self._root or self._source.suffix.lower() not in {".docx", ".xlsx"}:
                raise ValueError()
            self._source_stamp, self._source_sha = self._snapshot()
        except Exception:
            raise DiagramExportError("export_source_invalid") from None

    def _snapshot(self):
        _audit_path(self._root, directory=True)
        if _identity(self._root.lstat()) != self._root_identity:
            raise ValueError()
        _audit_path(self._source, directory=False)
        if self._space.file(str(self._source)) != self._source:
            raise ValueError()
        info = self._source.lstat()
        before = _stamp(info)
        if info.st_nlink != 1 or not 0 <= info.st_size <= MAX_SOURCE_BYTES:
            raise ValueError()
        digest, total = hashlib.sha256(), 0
        with self._source.open("rb") as stream:
            # Windows named/handle ctime semantics differ; named checks keep full stamp.
            opened = os.fstat(stream.fileno())
            if _stamp(opened)[:4] != before[:4] or opened.st_nlink != 1:
                raise ValueError()
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                total += len(block)
                if total > MAX_SOURCE_BYTES:
                    raise ValueError()
                digest.update(block)
            closed = os.fstat(stream.fileno())
            if _stamp(closed)[:4] != before[:4] or closed.st_nlink != 1:
                raise ValueError()
        _audit_path(self._source, directory=False)
        after = self._source.lstat()
        if _stamp(after) != before or after.st_nlink != 1 or total != before[2]:
            raise ValueError()
        return before, digest.hexdigest()

    def _source_guard(self):
        try:
            if self._snapshot() != (self._source_stamp, self._source_sha):
                raise ValueError()
        except Exception:
            raise DiagramExportError("export_source_changed") from None

    def _session_guard(self, session_id):
        if (not self._active or type(session_id) is not str or session_id != self._session):
            raise DiagramExportError("export_operation_invalidated")

    def _purge(self):
        now = time.monotonic()
        self._plans = {key: plan for key, plan in self._plans.items() if now < plan.expires}

    @staticmethod
    def _preview_metadata(preview):
        try:
            if type(preview) is not RenderedDiagram or type(preview.source) is not DiagramSource:
                raise ValueError()
            run = preview.execution
            if (type(run) is not OwnedRun or run.empty_confirmed is not True
                    or run.code != "process_completed" or type(run.exit_code) is not int
                    or run.exit_code != 0):
                raise ValueError()
            canonical = validate_source(preview.source.source, preview.source.kind)
            if canonical != preview.source or type(preview.svg) is not str:
                raise ValueError()
            safe = validate_svg(preview.svg)
            raw = safe.encode("utf-8")
            if not 0 < len(raw) <= MAX_SVG_BYTES:
                raise ValueError()
            return {"kind": canonical.kind, "source_sha256": canonical.sha256,
                    "svg_sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)}
        except Exception:
            raise DiagramExportError("export_preview_invalid") from None

    def plan_export(self, preview: RenderedDiagram) -> dict:
        with self._lock:
            self._session_guard(self._session)
            # Reject untrusted/invalid previews before any source I/O.
            metadata = self._preview_metadata(preview)
            self._purge()
            if len(self._plans) >= MAX_PLANS:
                raise DiagramExportError("export_plan_limit")
            self._source_guard()
            try:
                token = uuid.uuid4().hex
            except Exception:
                raise DiagramExportError("export_plan_invalid") from None
            if type(token) is not str or not re.fullmatch(r"[0-9a-f]{32}", token) or token in self._plans:
                raise DiagramExportError("export_plan_invalid")
            metadata.update(plan_id=token, destination_basename=f"diagram-{token}.svg",
                            document_sha256=self._source_sha, confirmation_required=True,
                            approval_ttl_seconds=APPROVAL_TTL, saving_implemented=False)
            self._plans[token] = _Plan(dict(metadata), time.monotonic() + APPROVAL_TTL)
            return dict(metadata)

    def confirm_export(self, plan_id: str, confirmed=False, *, session_id: str | None = None) -> dict:
        with self._lock:
            # Pop first: even a wrong-session, expired or nonboolean attempt consumes authority.
            plan = self._plans.pop(plan_id, None) if type(plan_id) is str else None
            self._session_guard(self._session if session_id is None else session_id)
            if plan is None:
                raise DiagramExportError("export_plan_invalid")
            if confirmed is not True:
                raise DiagramExportError("export_confirmation_required")
            if time.monotonic() >= plan.expires:
                raise DiagramExportError("export_plan_expired")
            self._source_guard()
            # Deliberately NOT a success result or transferable save authorization.
            return {"plan_id": plan_id, "saved": False, "code": "export_saving_unimplemented"}

    def cancel(self, plan_id: str) -> dict:
        with self._lock:
            self._session_guard(self._session)
            if type(plan_id) is not str or self._plans.pop(plan_id, None) is None:
                raise DiagramExportError("export_plan_invalid")
            return {"plan_id": plan_id, "cancelled": True}

    def invalidate(self) -> dict:
        with self._lock:
            self._active = False
            self._plans.clear()
            return {"invalidated": True}

    def plan_rollback(self, export_id: str) -> dict:
        # This capsule has never saved a file; no arbitrary path/history is accepted.
        raise DiagramExportError("export_rollback_unimplemented")

    def confirm_rollback(self, rollback_id: str, confirmed=False) -> dict:
        raise DiagramExportError("export_rollback_unimplemented")

    def __repr__(self):
        return "DiagramExportService(host_only=True, saving_implemented=False)"
