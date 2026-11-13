"""HOST-ONLY approved SVG new-copy transactions; deliberately NOT mounted.

A future bridge must obtain its private current RenderedDiagram receipt and bind
account/task/source-edit/logout/cancel epochs. Types/hashes/empty Jobs are not
client authentication or semantic proof. Never reconstruct authority from JSON.
Source is fingerprinted, never edited; backup is N/A for a newly created artifact.
Guards are snapshots, not journals, OS-atomic transactions or race exclusion.
Local bounded I/O is synchronous under an instance lock, not actively preemptible.
Restart/invalidation never reconstructs authority from retained audit files.
"""
from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
import threading
import time
import uuid

from .diagram_transactions import (
    APPROVAL_TTL, MAX_PLANS, MAX_SVG_BYTES, DiagramExportError,
    DiagramExportService, _identity, _stamp,
)
from .diagrams import validate_svg
from .diagram_raster import MAX_PNG_BYTES, RasterReceipt, verify_png
from .writes import _audit_path

MAX_AUDIT_BYTES = 16_384
MAX_SAVED = 8
MAX_IDENTIFIERS = 256  # Bounded lifetime anti-replay ledger, including consumed IDs.
DiagramSaveError = DiagramExportError


def _sha(data):
    return hashlib.sha256(data).hexdigest()


@dataclass(frozen=True)
class _Ticket:
    path: Path
    identity: tuple
    stamp: tuple
    digest: str
    size: int


@dataclass
class _Created:
    path: Path
    identity: tuple | None = None
    data: bytes = b""


@dataclass(frozen=True)
class _Approval:
    target: object
    expires: float


@dataclass(frozen=True)
class _Saved:
    output: _Ticket
    audit: _Ticket
    record: dict


class DiagramSaveService:
    """8 pending saves/rollbacks, 8 owned saves, 300s one-use approvals.

    At most 256 distinct identifiers per live instance; consumed IDs cannot be
    reissued. Public dictionaries contain metadata only. No directories, provider
    calls, Office automation, source overwrite or audit-history authorization.
    """
    def __init__(self, root: Path, source: Path, session_id: str):
        self._guard = DiagramExportService(root, source, session_id)
        self._root = self._guard._root
        self._lock = threading.RLock()
        self._disabled = False
        self._plans = {}
        self._rollbacks = {}
        self._saved = {}
        self._issued = set()

    def _ready(self, session_id=None):
        if self._disabled:
            raise DiagramSaveError("export_service_disabled")
        self._guard._session_guard(self._guard._session if session_id is None else session_id)

    def _disable(self):
        self._disabled = True
        self._plans.clear()
        self._rollbacks.clear()

    def _root_guard(self):
        _audit_path(self._root, directory=True)
        if _identity(self._root.lstat()) != self._guard._root_identity:
            raise ValueError()

    def _path(self, path, *, existing):
        self._root_guard()
        if path.parent != self._root or not path.is_absolute():
            raise ValueError()
        _audit_path(path, directory=False, missing=not existing)
        if self._guard._space.file(str(path), existing=existing) != path:
            raise ValueError()

    def _new_id(self):
        if len(self._issued) >= MAX_IDENTIFIERS:
            raise DiagramSaveError("export_id_limit")
        try:
            token = uuid.uuid4().hex
        except Exception:
            raise DiagramSaveError("export_plan_invalid") from None
        if (type(token) is not str or not re.fullmatch(r"[0-9a-f]{32}", token)
                or token in self._issued):
            raise DiagramSaveError("export_plan_invalid")
        self._issued.add(token)
        return token

    def _purge(self):
        now = time.monotonic()
        self._plans = {key: value for key, value in self._plans.items() if now < value.expires}
        self._rollbacks = {key: value for key, value in self._rollbacks.items() if now < value.expires}

    def _consume(self, table, token, confirmed, session_id):
        approval = table.pop(token, None) if type(token) is str else None
        self._ready(session_id)
        if approval is None:
            raise DiagramSaveError("export_plan_invalid")
        if confirmed is not True:
            raise DiagramSaveError("export_confirmation_required")
        if time.monotonic() >= approval.expires:
            raise DiagramSaveError("export_plan_expired")
        self._guard._source_guard()
        return approval.target

    def plan_save(self, preview):
        with self._lock:
            self._ready()
            metadata = self._guard._preview_metadata(preview)
            try:
                data = validate_svg(preview.svg).encode("utf-8")
                if not 0 < len(data) <= MAX_SVG_BYTES or _sha(data) != metadata["svg_sha256"]:
                    raise ValueError()
            except Exception:
                raise DiagramSaveError("export_preview_invalid") from None
            self._purge()
            if len(self._plans) >= MAX_PLANS:
                raise DiagramSaveError("export_plan_limit")
            if len(self._saved) >= MAX_SAVED:
                raise DiagramSaveError("export_saved_limit")
            self._guard._source_guard()
            token = self._new_id()
            metadata.update(plan_id=token, destination_basename=f"office-diagram-{token}.svg",
                            document_sha256=self._guard._source_sha,
                            confirmation_required=True, approval_ttl_seconds=APPROVAL_TTL,
                            backup="not_applicable_new_artifact")
            self._plans[token] = _Approval((data, dict(metadata)), time.monotonic() + APPROVAL_TTL)
            return dict(metadata)

    def plan_png(self, preview, raster):
        """PNG and SVG get distinct one-use approvals; client bytes are never accepted."""
        with self._lock:
            self._ready()
            metadata = self._guard._preview_metadata(preview)
            try:
                if (type(raster) is not RasterReceipt or raster.preview is not preview
                        or type(raster.png) is not bytes or raster.renderer != 'host-qt-svg-v1'
                        or not verify_png(raster.png, raster.width, raster.height)
                        or _sha(raster.png) != raster.sha256):
                    raise ValueError()
            except Exception:
                raise DiagramSaveError('export_preview_invalid') from None
            self._purge()
            if len(self._plans) >= MAX_PLANS:
                raise DiagramSaveError('export_plan_limit')
            if len(self._saved) >= MAX_SAVED:
                raise DiagramSaveError('export_saved_limit')
            self._guard._source_guard()
            token = self._new_id()
            metadata.update(plan_id=token, destination_basename=f'office-diagram-{token}.png',
                            document_sha256=self._guard._source_sha, png_sha256=raster.sha256,
                            width=raster.width, height=raster.height, renderer=raster.renderer,
                            confirmation_required=True, approval_ttl_seconds=APPROVAL_TTL,
                            backup='not_applicable_new_artifact')
            self._plans[token] = _Approval((raster.png, dict(metadata)), time.monotonic() + APPROVAL_TTL)
            return dict(metadata)

    def _read(self, path, limit, expected=None, *, svg=False, png=False):
        self._path(path, existing=True)
        before = path.lstat()
        if before.st_nlink != 1 or not 0 <= before.st_size <= limit:
            raise ValueError()
        if expected is not None and (_identity(before) != expected.identity or _stamp(before) != expected.stamp):
            raise ValueError()
        with path.open("rb") as stream:
            opened = os.fstat(stream.fileno())
            if opened.st_nlink != 1 or _stamp(opened)[:4] != _stamp(before)[:4]:
                raise ValueError()
            data = stream.read(limit + 1)
            after_handle = os.fstat(stream.fileno())
            if after_handle.st_nlink != 1 or _stamp(after_handle)[:4] != _stamp(before)[:4]:
                raise ValueError()
        self._path(path, existing=True)
        after = path.lstat()
        if after.st_nlink != 1 or _stamp(after) != _stamp(before) or len(data) != before.st_size:
            raise ValueError()
        digest = _sha(data)
        if expected is not None and (digest != expected.digest or len(data) != expected.size):
            raise ValueError()
        if svg and validate_svg(data.decode("utf-8", errors="strict")).encode("utf-8") != data:
            raise ValueError()
        if png:
            from .diagram_raster import normalize_png
            normalized, _, _ = normalize_png(data)
            if normalized != data:
                raise ValueError()
        return data, _Ticket(path, _identity(after), _stamp(after), digest, len(data))

    @staticmethod
    def _limit(path):
        return MAX_PNG_BYTES if path.suffix == '.png' else MAX_SVG_BYTES if path.suffix == '.svg' else MAX_AUDIT_BYTES

    def _create(self, path, data, created):
        self._path(path, existing=False)
        flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
        fd = os.open(str(path), flags, 0o600)
        ownership = _Created(path)
        created.append(ownership)  # Track immediately, even if fstat/fdopen fails.
        try:
            ownership.identity = _identity(os.fstat(fd))
            offset = 0
            while offset < len(data):
                written = os.write(fd, data[offset:])
                if type(written) is not int or not 0 < written <= len(data) - offset:
                    raise ValueError()
                offset += written
                ownership.data = data[:offset]
            os.fsync(fd)
        finally:
            os.close(fd)
        actual, ticket = self._read(path, self._limit(path),
                                    svg=path.suffix == '.svg', png=path.suffix == '.png')
        if ticket.identity != ownership.identity or actual != data:
            raise ValueError()
        return ticket

    def _cleanup(self, created):
        clean = True
        for owned in reversed(created):
            try:
                self._root_guard()
                if not owned.path.exists():
                    # Reject a dangling link/reparse entry rather than treating it as absent.
                    _audit_path(owned.path, directory=False, missing=True)
                    continue
                if owned.identity is None:
                    raise ValueError()
                data, ticket = self._read(owned.path, self._limit(owned.path), png=owned.path.suffix == '.png')
                if ticket.identity != owned.identity or data != owned.data:
                    raise ValueError()
                # Exact absolute owned path was just audited. Still a snapshot, not atomic.
                owned.path.unlink()
                if owned.path.exists():
                    raise ValueError()
            except Exception:
                clean = False
        if not clean:
            self._disable()
        return clean

    @staticmethod
    def _audit_bytes(record):
        raw = json.dumps(record, ensure_ascii=True, sort_keys=True, separators=(",", ":")).encode("utf-8")
        if len(raw) > MAX_AUDIT_BYTES:
            raise ValueError()
        return raw

    def confirm_save(self, plan_id, confirmed=False, *, session_id=None):
        with self._lock:
            data, metadata = self._consume(self._plans, plan_id, confirmed, session_id)
            if len(self._saved) >= MAX_SAVED:
                raise DiagramSaveError("export_saved_limit")
            output = self._root / metadata["destination_basename"]
            audit_path = self._root / f"office-diagram-{plan_id}.audit.json"
            created = []
            try:
                # Both destinations preflight before creation; existing files are never overwritten.
                self._path(output, existing=False)
                self._path(audit_path, existing=False)
                ticket = self._create(output, data, created)
                png = output.suffix == '.png'
                record = {"schema": 1, "kind": "diagram_png_save" if png else "diagram_svg_save", "save_id": plan_id,
                          "output_basename": output.name, "source_sha256": metadata["source_sha256"],
                          "document_sha256": metadata["document_sha256"], "svg_sha256": metadata['svg_sha256'] if png else ticket.digest,
                          "bytes": ticket.size, "saved": True, "output_removed": False,
                          "status": "saved", "backup": "not_applicable_new_artifact",
                          "independent_human_save_confirmed": True}
                if png:
                    record.update(png_sha256=ticket.digest, width=metadata['width'],
                                  height=metadata['height'], renderer=metadata['renderer'])
                audit = self._create(audit_path, self._audit_bytes(record), created)
                self._read(output, self._limit(output), ticket, svg=not png, png=png)
                self._read(audit_path, MAX_AUDIT_BYTES, audit)
                self._guard._source_guard()
                self._saved[plan_id] = _Saved(ticket, audit, record)
                return {"save_id": plan_id, "saved": True, "path": str(output), "audit_path": str(audit_path),
                        "sha256": ticket.digest, "bytes": ticket.size, "source_unchanged": True,
                        "backup": "not_applicable_new_artifact", "rollback_confirmation_required": True}
            except Exception:
                clean = self._cleanup(created)
                raise DiagramSaveError("export_save_failed" if clean else "export_cleanup_uncertain") from None

    def cancel(self, plan_id):
        with self._lock:
            self._ready()
            if type(plan_id) is not str or self._plans.pop(plan_id, None) is None:
                raise DiagramSaveError("export_plan_invalid")
            return {"plan_id": plan_id, "cancelled": True}

    def invalidate(self):
        with self._lock:
            self._guard.invalidate()
            self._plans.clear()
            self._rollbacks.clear()
            self._saved.clear()  # Keep files/audits; never restore authority from history.
            return {"invalidated": True}

    def plan_rollback(self, save_id):
        with self._lock:
            self._ready()
            self._purge()
            if type(save_id) is not str or save_id not in self._saved:
                raise DiagramSaveError("export_rollback_invalid")
            if len(self._rollbacks) >= MAX_PLANS:
                raise DiagramSaveError("export_plan_limit")
            saved = self._saved[save_id]
            try:
                self._guard._source_guard()
                self._read(saved.output.path, self._limit(saved.output.path), saved.output,
                           svg=saved.output.path.suffix == '.svg', png=saved.output.path.suffix == '.png')
                self._read(saved.audit.path, MAX_AUDIT_BYTES, saved.audit)
            except Exception:
                raise DiagramSaveError("export_rollback_changed") from None
            token = self._new_id()
            self._rollbacks[token] = _Approval(save_id, time.monotonic() + APPROVAL_TTL)
            return {"rollback_id": token, "save_id": save_id, "confirmation_required": True,
                    "deletes_only_unchanged_owned_output": True, "audit_retained": True,
                    "restores_history": False, "approval_ttl_seconds": APPROVAL_TTL}

    def _update_audit(self, saved):
        record = {**saved.record, "saved": False, "output_removed": True, "status": "rolled_back",
                  "independent_human_rollback_confirmed": True}
        payload = self._audit_bytes(record)
        self._read(saved.audit.path, MAX_AUDIT_BYTES, saved.audit)
        # Nontruncating open; prove opened AND named ownership/content before the first write.
        with saved.audit.path.open("r+b") as stream:
            opened = os.fstat(stream.fileno())
            if opened.st_nlink != 1 or _stamp(opened)[:4] != saved.audit.stamp[:4]:
                raise ValueError()
            old = stream.read(MAX_AUDIT_BYTES + 1)
            if _sha(old) != saved.audit.digest or len(old) != saved.audit.size:
                raise ValueError()
            self._read(saved.audit.path, MAX_AUDIT_BYTES, saved.audit)
            self._guard._source_guard()
            stream.seek(0)
            if stream.write(payload) != len(payload):
                raise ValueError()
            stream.flush()
            os.fsync(stream.fileno())
            stream.truncate(len(payload))
            stream.flush()
            os.fsync(stream.fileno())
            stream.seek(0)
            if stream.read(MAX_AUDIT_BYTES + 1) != payload:
                raise ValueError()
        actual, updated = self._read(saved.audit.path, MAX_AUDIT_BYTES)
        if updated.identity != saved.audit.identity or actual != payload:
            raise ValueError()

    def confirm_rollback(self, rollback_id, confirmed=False, *, session_id=None):
        with self._lock:
            save_id = self._consume(self._rollbacks, rollback_id, confirmed, session_id)
            saved = self._saved.get(save_id)
            try:
                if saved is None:
                    raise ValueError()
                self._read(saved.output.path, self._limit(saved.output.path), saved.output,
                           svg=saved.output.path.suffix == '.svg', png=saved.output.path.suffix == '.png')
                self._read(saved.audit.path, MAX_AUDIT_BYTES, saved.audit)
                self._guard._source_guard()
                # Recheck the exact absolute target after the potentially long source scan.
                self._read(saved.output.path, self._limit(saved.output.path), saved.output,
                           svg=saved.output.path.suffix == '.svg', png=saved.output.path.suffix == '.png')
                self._read(saved.audit.path, MAX_AUDIT_BYTES, saved.audit)
            except Exception:
                raise DiagramSaveError("export_rollback_changed") from None
            try:
                # Still not an atomic anti-race guarantee.
                saved.output.path.unlink()
            except Exception:
                try:
                    self._read(saved.output.path, self._limit(saved.output.path), saved.output,
                           svg=saved.output.path.suffix == '.svg', png=saved.output.path.suffix == '.png')
                except Exception:
                    self._disable()
                    return {"removed": None, "audit_updated": False, "code": "export_rollback_uncertain"}
                raise DiagramSaveError("export_rollback_failed") from None
            self._saved.pop(save_id, None)
            audit_updated = None
            try:
                self._update_audit(saved)
                audit_updated = True  # Verified update happened even if the later scope guard fails.
                self._guard._source_guard()
            except Exception:
                if audit_updated is None:
                    try:
                        self._read(saved.audit.path, MAX_AUDIT_BYTES, saved.audit)
                        audit_updated = False  # Proven original unchanged; otherwise unknown, not False.
                    except Exception:
                        pass
                self._disable()
                return {"save_id": save_id, "removed": True, "audit_updated": audit_updated,
                        "code": "export_rollback_partial", "service_disabled": True}
            return {"save_id": save_id, "removed": True, "audit_updated": True,
                    "audit_retained": True, "restores_history": False}

    def __repr__(self):
        return "DiagramSaveService(host_only=True, mounted=False)"
