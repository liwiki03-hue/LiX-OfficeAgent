"""Host-only P3 generation/save/rollback approval service; not registered with models or UI.

Every plan is read-only metadata. Network authorization, new-file save and rollback each
need their own one-use, session-bound, strictly-boolean approval. Generate/save/rollback
approvals are never interchangeable: save never invokes a provider, and rollback deletes
only a still-unchanged output this instance saved (never a file it cannot prove it owns).
Unsaved bytes remain in bounded memory and are never written by generation alone.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import threading
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from .imagegen import GeneratedImage, ImageGenConfig, ImageGenError, ImageRequest
from .imageprovider import OpenAIImageGenProvider, normalize_generated_image
from .writes import Workspace, _audit_path

MAX_SOURCE_BYTES = 100_000_000
MAX_AUDIT_BYTES = 16_384
MAX_OUTPUT_BYTES = 10_000_000
MAX_PLANS = 8
APPROVAL_TTL = 300.0


def _identity(info):
    return info.st_dev, info.st_ino


def _stamp(info):
    return (*_identity(info), info.st_size, info.st_mtime_ns, info.st_ctime_ns)


def _owned(info, identity):
    return _identity(info) == identity and info.st_nlink == 1


def _sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


@dataclass
class _Audit:
    path: Path
    lineage: tuple
    identity: tuple
    sha256: str
    record: dict


@dataclass
class _Plan:
    request: ImageRequest
    fingerprint: str = field(repr=False)
    public: dict
    expires: float
    status: str = "planned"
    image: GeneratedImage | None = field(default=None, repr=False)
    audit: _Audit | None = field(default=None, repr=False)


@dataclass
class _Approval:
    """One-use save/rollback authorization; expires even if never confirmed."""

    target: str
    expires: float
    status: str = "planned"


@dataclass
class _Saved:
    """Proof of an output this instance itself created; no history/restore authority."""

    plan_id: str
    path: Path
    identity: tuple = field(repr=False)
    sha256: str
    bytes: int
    metadata: dict
    audit: _Audit = field(repr=False)

    def public(self):
        return {"path": str(self.path), "sha256": self.sha256, "bytes": self.bytes,
                "image": dict(self.metadata)}


class ImageGenerationService:
    """No history/restart authority. Host must separately approve generation and saving.

    All public returns are metadata, except explicitly host-only preview(). There is
    no lock across provider calls. invalidate() fences late completions immediately.
    """
    def __init__(self, root: Path, source: Path, session_id: str, *,
                 config_loader: Callable = ImageGenConfig.from_env,
                 provider_factory: Callable = OpenAIImageGenProvider,
                 expected_version: str | None = None, standalone: bool = False,
                 reuse_metadata: bool = False):
        self._lock = threading.RLock()
        self._active = True
        self._running = None
        self._reuse_metadata = reuse_metadata is True
        self._plans: dict[str, _Plan] = {}
        self._saved: dict[str, _Saved] = {}
        self._save_plans: dict[str, _Approval] = {}
        self._rollback_plans: dict[str, _Approval] = {}
        if type(session_id) is not str or not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", session_id):
            raise ImageGenError("image_session_invalid")
        self._session = session_id
        self._config_loader, self._provider_factory = config_loader, provider_factory
        try:
            _audit_path(root, directory=True)
            self._space = Workspace(root)
            self.root = self._space.root
            self._root_identity = _identity(self.root.lstat())
            self._standalone = standalone is True and source is None
            self._output_space = self._space
            self._output_identity = self._root_identity
            if self._standalone:
                self.source = None
                self._source_sha = None
                self._source_stamp = None
                return
            _audit_path(source, directory=False)
            self.source = self._space.file(str(source))
            if self.source.parent != self.root or self.source.suffix.lower() not in {".docx", ".xlsx"}:
                raise ValueError()
            self._source_stamp, actual = self._source_snapshot()
            if expected_version is not None and (type(expected_version) is not str
                    or not re.fullmatch(r"[0-9a-f]{64}", expected_version)
                    or actual != expected_version):
                raise ValueError()
            self._source_sha = actual
        except Exception:
            raise ImageGenError("image_source_invalid") from None

    def _source_snapshot(self):
        """Observe current identity/metadata/SHA, not transient mutation history.

        Writes restored to an identical observable snapshot cannot be proved by
        these checks; they are not a filesystem journal or an atomic transaction.
        """
        _audit_path(self.root, directory=True)
        if _identity(self.root.lstat()) != self._root_identity:
            raise ImageGenError("image_source_changed")
        _audit_path(self.source, directory=False)
        if self._space.file(str(self.source)) != self.source:
            raise ImageGenError("image_source_changed")
        before = _stamp(self.source.lstat())
        # Windows lstat/fstat may disagree on ctime semantics. Handle comparisons below
        # retain identity/size/mtime; named pre/post and baseline checks retain full stamps.
        if before[2] > MAX_SOURCE_BYTES:
            raise ImageGenError("image_source_changed")
        digest = hashlib.sha256()
        total = 0
        with self.source.open("rb") as stream:
            if _stamp(os.fstat(stream.fileno()))[:4] != before[:4]:
                raise ImageGenError("image_source_changed")
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                total += len(block)
                if total > MAX_SOURCE_BYTES:
                    raise ImageGenError("image_source_changed")
                digest.update(block)
            if _stamp(os.fstat(stream.fileno()))[:4] != before[:4]:
                raise ImageGenError("image_source_changed")
        _audit_path(self.source, directory=False)
        if _stamp(self.source.lstat()) != before or total != before[2]:
            raise ImageGenError("image_source_changed")
        return before, digest.hexdigest()

    def _session_guard(self, session_id):
        if not self._active or type(session_id) is not str or session_id != self._session:
            raise ImageGenError("image_operation_invalidated")

    def _source_guard(self):
        if self._standalone:
            self._root_guard()
            return
        try:
            stamp, sha = self._source_snapshot()
            if stamp != self._source_stamp or sha != self._source_sha:
                raise ValueError()
        except Exception:
            raise ImageGenError("image_source_changed") from None

    def _config(self):
        try:
            config = self._config_loader()
            if not isinstance(config, ImageGenConfig):
                raise ValueError()
            return config
        except ImageGenError as error:
            if error.code == "image_configuration_missing":
                raise ImageGenError(error.code) from None
            raise ImageGenError("image_configuration_invalid") from None
        except Exception:
            raise ImageGenError("image_configuration_invalid") from None

    def _binding(self, plan, session_id):
        self._session_guard(session_id)
        self._source_guard()
        config = self._config()
        if config._binding_fingerprint() != plan.fingerprint:
            raise ImageGenError("image_configuration_changed")
        return config

    def _lookup(self, plan_id, session_id):
        self._session_guard(session_id)
        if type(plan_id) is not str or plan_id not in self._plans:
            raise ImageGenError("image_plan_invalid")
        return self._plans[plan_id]

    @staticmethod
    def _public(plan):
        # Deep copy: host/UI callers cannot mutate internal config/approval records.
        return json.loads(json.dumps({**plan.public, "status": plan.status}))

    def plan(self, request: ImageRequest) -> dict:
        with self._lock:
            self._session_guard(self._session)
            if self._running is not None:
                raise ImageGenError("image_operation_busy")
            if not isinstance(request, ImageRequest):
                raise ImageGenError("image_request_invalid")
            if len(self._plans) >= MAX_PLANS:
                raise ImageGenError("image_plan_limit")
            self._source_guard()
            config = self._config()
            plan_id = str(uuid.uuid4())
            if plan_id in self._plans:
                raise ImageGenError("image_plan_invalid")
            public = {"plan_id": plan_id, "kind": "image_generation", "source": str(self.source) if self.source is not None else None,
                      "source_sha256": self._source_sha, "effective_prompt": request.effective_prompt,
                      "size": request.size, "style": request.style, "n": request.n,
                      "configuration": config.public_config(), "confirmation_required": True,
                      "save_confirmation_required": True, "cost_unknown": True,
                      "provider_fees_irreversible": True, "approval_ttl_seconds": APPROVAL_TTL}
            plan = _Plan(request, config._binding_fingerprint(), public, time.monotonic() + APPROVAL_TTL)
            self._plans[plan_id] = plan
            return self._public(plan)

    def _lineage_guard(self, lineage):
        for path, identity in lineage:
            _audit_path(path, directory=True)
            if _identity(path.lstat()) != identity:
                raise ImageGenError("image_audit_changed")

    @staticmethod
    def _audit_bytes(record):
        payload = json.dumps(record, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        if len(payload) > MAX_AUDIT_BYTES:
            raise ImageGenError("image_audit_failed")
        return payload

    def _create_audit(self, plan_id, plan):
        lineage = [(self.root, self._root_identity)]
        parent = self.root
        for name in (".office_agent", "imagegen", plan_id):
            self._lineage_guard(lineage)
            parent = parent / name
            _audit_path(parent, directory=True, missing=True)
            if name == plan_id:
                parent.mkdir()  # owned unique directory MUST be exclusive
            elif not parent.exists():
                try:
                    parent.mkdir()
                except FileExistsError:
                    pass
            _audit_path(parent, directory=True)
            lineage.append((parent, _identity(parent.lstat())))
            self._lineage_guard(lineage)
        record = {"schema_version": 1, "kind": "image_generation", "transaction_id": plan_id,
                  "session_id": self._session, "source": str(self.source) if self.source is not None else None, "source_sha256": self._source_sha,
                  "prompt_sha256": _sha(plan.request.effective_prompt.encode("utf-8")),
                  "configuration": plan.public["configuration"], "size": plan.request.size,
                  "style": plan.request.style, "n": 1, "status": "generating",
                  "human_generation_confirmed": True, "cost_unknown": True,
                  "provider_fees_irreversible": True, "provider_request_state": "not_started", "saved": False}
        if self._reuse_metadata:
            record['reusable'] = {'version':1,'prompt':plan.request.prompt,'size':plan.request.size,'style':plan.request.style,'n':1}
        payload = self._audit_bytes(record)
        path = parent / "audit.json"
        _audit_path(path, directory=False, missing=True)
        with path.open("xb") as stream:
            identity = _identity(os.fstat(stream.fileno()))
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        ticket = _Audit(path, tuple(lineage), identity, _sha(payload), record)
        self._verify_audit(ticket)
        return ticket

    def _verify_audit(self, ticket):
        self._lineage_guard(ticket.lineage)
        _audit_path(ticket.path, directory=False)
        if not _owned(ticket.path.lstat(), ticket.identity):
            raise ImageGenError("image_audit_changed")
        with ticket.path.open("rb") as stream:
            if not _owned(os.fstat(stream.fileno()), ticket.identity):
                raise ImageGenError("image_audit_changed")
            payload = stream.read(MAX_AUDIT_BYTES + 1)
        if len(payload) > MAX_AUDIT_BYTES or _sha(payload) != ticket.sha256:
            raise ImageGenError("image_audit_changed")
        self._lineage_guard(ticket.lineage)
        _audit_path(ticket.path, directory=False)
        if not _owned(ticket.path.lstat(), ticket.identity):
            raise ImageGenError("image_audit_changed")

    def _update_audit(self, ticket, changes):
        self._verify_audit(ticket)
        record = {**ticket.record, **changes}
        payload = self._audit_bytes(record)
        # Open WITHOUT truncation. Verify opened object's ownership/content BEFORE writing.
        with ticket.path.open("r+b") as stream:
            if not _owned(os.fstat(stream.fileno()), ticket.identity):
                raise ImageGenError("image_audit_changed")
            old = stream.read(MAX_AUDIT_BYTES + 1)
            if _sha(old) != ticket.sha256:
                raise ImageGenError("image_audit_changed")
            self._lineage_guard(ticket.lineage)
            stream.seek(0)
            stream.write(payload)
            stream.truncate()
            stream.flush()
            os.fsync(stream.fileno())
            stream.seek(0)
            if stream.read(MAX_AUDIT_BYTES + 1) != payload:
                raise ImageGenError("image_audit_changed")
        self._lineage_guard(ticket.lineage)
        _audit_path(ticket.path, directory=False)
        if not _owned(ticket.path.lstat(), ticket.identity):
            raise ImageGenError("image_audit_changed")
        ticket.sha256, ticket.record = _sha(payload), record
        self._verify_audit(ticket)

    def confirm_generation(self, plan_id: str, session_id: str, confirmed: bool) -> dict:
        with self._lock:
            plan = self._lookup(plan_id, session_id)
            if type(confirmed) is not bool:
                raise ImageGenError("image_confirmation_required")
            if plan.status != "planned":
                raise ImageGenError("image_approval_consumed")
            if self._running is not None:
                raise ImageGenError("image_operation_busy")
            plan.status = "cancelled" if not confirmed else "preparing"
            if not confirmed:
                return self._public(plan)
            token = object()
            self._running = token
            try:
                if time.monotonic() > plan.expires:
                    raise ImageGenError("image_approval_expired")
                config = self._binding(plan, session_id)
                try:
                    plan.audit = self._create_audit(plan_id, plan)
                except Exception:
                    raise ImageGenError("image_audit_failed") from None
                self._binding(plan, session_id)
                self._verify_audit(plan.audit)
                provider = self._provider_factory()
                self._update_audit(plan.audit, {"provider_request_state": "attempt_committed"})
                self._binding(plan, session_id)
                # This is the authorization commit boundary, not proof of remote receipt/fees.
                plan.status = "generating"
            except Exception as error:
                self._generation_failure(plan, error)
                self._running = None
                raise ImageGenError(plan.public["error_code"],
                                    stage=plan.public.get("failure_stage"),
                                    detail_code=plan.public.get("failure_detail"),
                                    request_state=plan.public.get("provider_request_state"),
                                    provider_code=plan.public.get("provider_error")) from None
        try:
            # Deliberately OUTSIDE the lock: logout/selection invalidation must not wait for network.
            raw = provider.generate(plan.request, config)
            image = normalize_generated_image(raw)  # revalidate custom/injected providers too
            with self._lock:
                self._binding(plan, session_id)
                self._update_audit(plan.audit, {"status": "generated", "image": dict(image.metadata)})
                self._binding(plan, session_id)
                plan.image, plan.status = image, "generated"
                plan.public["transaction_id"] = plan_id
                plan.public["image"] = dict(image.metadata)
                return self._public(plan)
        except Exception as error:
            with self._lock:
                self._generation_failure(plan, error)
                raise ImageGenError(plan.public["error_code"],
                                    stage=plan.public.get("failure_stage"),
                                    detail_code=plan.public.get("failure_detail"),
                                    request_state=plan.public.get("provider_request_state"),
                                    provider_code=plan.public.get("provider_error")) from None
        finally:
            with self._lock:
                if self._running is token:
                    self._running = None

    def _generation_failure(self, plan, error):
        safe = {"image_operation_invalidated", "image_source_changed", "image_configuration_changed",
                "image_configuration_missing", "image_configuration_invalid", "image_approval_expired",
                "image_audit_failed", "image_audit_changed"}
        code = error.code if isinstance(error, ImageGenError) and error.code in safe else "image_generation_failed"
        detail_codes = {"image_response_timeout", "image_response_encoding", "image_response_too_large",
                        "image_response_invalid", "image_decoder_unavailable", "image_raster_too_large",
                        "image_raster_invalid", "image_provider_failed"}
        detail = error.code if isinstance(error, ImageGenError) and error.code in detail_codes else None
        if isinstance(error, ImageGenError) and error.provider_code in {'image_authentication_failed', 'image_model_not_found', 'image_rate_limited', 'image_service_failed', 'image_address_unreachable'}:
            plan.public['provider_error'] = error.provider_code
        committed = bool(plan.audit and plan.audit.record.get("provider_request_state") == "attempt_committed")
        stage = ("preflight" if not committed else
                 "response" if detail and detail.startswith("image_response_") else
                 "validation" if detail in {"image_decoder_unavailable", "image_raster_too_large", "image_raster_invalid"} else
                 "provider_request")
        plan.status, plan.image = "failed", None
        plan.public["error_code"] = code
        plan.public["failure_stage"] = stage
        plan.public["provider_request_state"] = "attempt_committed" if committed else "not_started"
        if detail:
            plan.public["failure_detail"] = detail
        if plan.audit is not None:
            try:
                self._update_audit(plan.audit, {"status": "failed", "error_code": code})
            except Exception:
                # Preserve foreign/tampered audit, NEVER overwrite to manufacture a failure record.
                plan.public["error_code"] = "image_audit_changed"

    def _root_guard(self):
        _audit_path(self.root, directory=True)
        if _identity(self.root.lstat()) != self._root_identity:
            raise ImageGenError("image_source_changed")

    def select_output_root(self, root):
        with self._lock:
            self._session_guard(self._session)
            if not self._standalone or self._running is not None or self._save_plans or self._saved:
                raise ImageGenError('image_output_invalid')
            _audit_path(root, directory=True)
            space = Workspace(root)
            self._output_space = space
            self._output_identity = _identity(space.root.lstat())
            self._output_selected = True

    def _target_guard(self, target):
        """Only a fixed new name directly inside the authorized selected parent is allowed."""
        _audit_path(self._output_space.root, directory=True)
        if _identity(self._output_space.root.lstat()) != self._output_identity:
            raise ImageGenError('image_output_invalid')
        _audit_path(target, directory=False, missing=True)
        try:
            resolved = self._output_space.file(str(target), existing=False)
        except Exception:
            raise ImageGenError("image_output_invalid") from None
        if resolved != target or resolved.parent != self._output_space.root:
            raise ImageGenError("image_output_invalid")
        return target

    def plan_save(self, plan_id: str, session_id: str) -> dict:
        """Read-only save plan for an unsaved generated image; must be approved separately."""
        with self._lock:
            plan = self._lookup(plan_id, session_id)
            if self._running is not None:
                raise ImageGenError("image_operation_busy")
            if plan.status != "generated" or plan.image is None:
                raise ImageGenError("image_preview_unavailable")
            self._root_guard()
            if self._standalone and not getattr(self, '_output_selected', False):
                raise ImageGenError('image_output_required')
            self._verify_audit(plan.audit)
            if len(self._save_plans) >= MAX_PLANS:
                raise ImageGenError("image_plan_limit")
            save_id = str(uuid.uuid4())
            target = self._target_guard(self._output_space.root / f"office-image-{save_id}.png")
            public = {"save_id": save_id, "kind": "image_save", "plan_id": plan_id,
                      "source": str(self.source) if self.source is not None else None, "source_sha256": self._source_sha,
                      "path": str(target), "filename": target.name, "parent": str(target.parent),
                      "bytes": len(plan.image.raw), "sha256": plan.image.metadata["sha256"],
                      "image": dict(plan.image.metadata), "confirmation_required": True,
                      "new_file": True, "overwrites_source": False, "overwrites_existing": False,
                      "replaces_generation_approval": False, "provider_fees_irreversible": True,
                      "saving_does_not_reverse_fees": True, "approval_ttl_seconds": APPROVAL_TTL}
            self._save_plans[save_id] = _Approval(plan_id, time.monotonic() + APPROVAL_TTL)
            return json.loads(json.dumps({**public, "status": "planned"}))

    def _owned_output(self, path, identity, sha256):
        """Prove the path is still exactly the single-linked file this instance created."""
        try:
            _audit_path(self._output_space.root, directory=True)
            if path.parent != self._output_space.root or _identity(self._output_space.root.lstat()) != self._output_identity:
                return False
            _audit_path(path, directory=False)
            if not _owned(path.lstat(), identity):
                return False
            with path.open("rb") as stream:
                if not _owned(os.fstat(stream.fileno()), identity):
                    return False
                payload = stream.read(MAX_OUTPUT_BYTES + 1)
            if len(payload) > MAX_OUTPUT_BYTES or _sha(payload) != sha256:
                return False
            _audit_path(path, directory=False)
            return _owned(path.lstat(), identity)
        except Exception:
            return False

    def _write_output(self, target, image):
        """Exclusive create of a NEW file; source is never opened for writing."""
        payload = image.raw
        if (len(payload) > MAX_OUTPUT_BYTES or len(payload) != image.metadata.get("bytes")
                or _sha(payload) != image.metadata.get("sha256")):
            raise ImageGenError("image_output_invalid")
        try:
            with target.open("xb") as stream:
                handle = _identity(os.fstat(stream.fileno()))
                stream.write(payload)
                stream.flush()
                os.fsync(stream.fileno())
                if _identity(os.fstat(stream.fileno())) != handle:
                    raise ValueError()
        except FileExistsError:
            raise ImageGenError("image_output_exists") from None
        except Exception:
            raise ImageGenError("image_output_failed") from None
        try:
            _audit_path(target, directory=False)
            if not _owned(target.lstat(), handle):
                raise ValueError()
            with target.open("rb") as stream:
                if not _owned(os.fstat(stream.fileno()), handle):
                    raise ValueError()
                if stream.read(MAX_OUTPUT_BYTES + 1) != payload:
                    raise ValueError()
        except Exception:
            raise ImageGenError("image_output_changed") from None
        return handle, {"path": str(target), "sha256": _sha(payload), "bytes": len(payload),
                        "file_identity": list(handle),
                        "image": dict(image.metadata)}

    def confirm_save(self, save_id: str, session_id: str, confirmed: bool) -> dict:
        """Independent one-use approval for writing the unsaved raster as a brand-new PNG."""
        with self._lock:
            self._session_guard(session_id)
            if type(confirmed) is not bool:
                raise ImageGenError("image_confirmation_required")
            if type(save_id) is not str or save_id not in self._save_plans:
                raise ImageGenError("image_save_invalid")
            approval = self._save_plans[save_id]
            plan = self._lookup(approval.target, session_id)
            if self._running is not None:
                raise ImageGenError("image_operation_busy")
            del self._save_plans[save_id]  # one use from here on, even on later failure
            if time.monotonic() > approval.expires:
                raise ImageGenError("image_approval_expired")
            if not confirmed:
                return {"save_id": save_id, "kind": "image_save", "status": "cancelled",
                        "written": False, "path": None}
            if plan.status != "generated" or plan.image is None:
                raise ImageGenError("image_preview_unavailable")
            self._root_guard()
            self._verify_audit(plan.audit)
            target = self._target_guard(self._output_space.root / f"office-image-{save_id}.png")
            token = object()
            self._running = token
            try:
                identity, record = self._write_output(target, plan.image)
                try:
                    self._update_audit(plan.audit, {"saved": True, "saved_output": record})
                except Exception:
                    # Never leave an unaudited artifact; remove only the exact owned bytes.
                    if self._discard_output(target, identity, record["sha256"]):
                        raise ImageGenError("image_audit_failed") from None
                    raise ImageGenError("image_audit_changed") from None
                self._root_guard()
                self._saved[save_id] = _Saved(approval.target, target, identity, record["sha256"],
                                              record["bytes"], dict(plan.image.metadata), plan.audit)
                plan.public["saved_output"] = {k: v for k, v in record.items() if k != "image"}
                public = {"save_id": save_id, "kind": "image_save", "status": "saved",
                          "written": True, "source_overwritten": False, "new_file": True,
                          **self._saved[save_id].public()}
                return json.loads(json.dumps(public))
            except ImageGenError:
                raise
            except Exception:
                raise ImageGenError("image_output_failed") from None
            finally:
                if self._running is token:
                    self._running = None

    def _discard_output(self, target, identity, sha256):
        """Best-effort removal of our own unaudited file; never touch anything else."""
        try:
            if not self._owned_output(target, identity, sha256):
                return False
            target.unlink()
            return True
        except Exception:
            return False

    def plan_rollback(self, save_id: str, session_id: str) -> dict:
        """Read-only plan to delete one saved output; requires a fresh separate approval."""
        with self._lock:
            self._session_guard(session_id)
            if self._running is not None:
                raise ImageGenError("image_operation_busy")
            if type(save_id) is not str or save_id not in self._saved:
                raise ImageGenError("image_rollback_invalid")
            record = self._saved[save_id]
            self._root_guard()
            if len(self._rollback_plans) >= MAX_PLANS:
                raise ImageGenError("image_plan_limit")
            rollback_id = str(uuid.uuid4())
            public = {"rollback_id": rollback_id, "kind": "image_rollback", "save_id": save_id,
                      "plan_id": record.plan_id, "path": str(record.path), "sha256": record.sha256,
                      "bytes": record.bytes, "confirmation_required": True,
                      "deletes_only_unchanged_owned_output": True,
                      "preserves_edited_or_replaced": True, "restores_history": False,
                      "restores_execution_authority": False, "provider_fees_irreversible": True,
                      "provider_fees_reversed": False, "approval_ttl_seconds": APPROVAL_TTL}
            self._rollback_plans[rollback_id] = _Approval(save_id, time.monotonic() + APPROVAL_TTL)
            return json.loads(json.dumps({**public, "status": "planned"}))

    def confirm_rollback(self, rollback_id: str, session_id: str, confirmed: bool) -> dict:
        """Fresh approval: delete only the still-unchanged output this instance created."""
        with self._lock:
            self._session_guard(session_id)
            if type(confirmed) is not bool:
                raise ImageGenError("image_confirmation_required")
            if type(rollback_id) is not str or rollback_id not in self._rollback_plans:
                raise ImageGenError("image_rollback_invalid")
            approval = self._rollback_plans[rollback_id]
            if self._running is not None:
                raise ImageGenError("image_operation_busy")
            del self._rollback_plans[rollback_id]
            if time.monotonic() > approval.expires:
                raise ImageGenError("image_approval_expired")
            record = self._saved.get(approval.target)
            if not confirmed:
                return {"rollback_id": rollback_id, "kind": "image_rollback", "status": "cancelled",
                        "removed": False, "path": None}
            self._root_guard()
            if record is None:
                raise ImageGenError("image_rollback_changed")
            # Audit proven before deletion; a tampered record never authorizes a removal.
            self._verify_audit(record.audit)
            if not self._owned_output(record.path, record.identity, record.sha256):
                raise ImageGenError("image_rollback_changed")
            try:
                record.path.unlink()
            except FileNotFoundError:
                raise ImageGenError("image_rollback_changed") from None
            except OSError:
                raise ImageGenError("image_rollback_failed") from None
            self._saved.pop(approval.target, None)
            try:
                self._update_audit(record.audit, {"saved": False, "output_removed": True,
                                                  "status": "rolled_back",
                                                  "provider_fees_irreversible": True})
            except Exception:
                # The removal is already durable; provenance failure is reported, not hidden.
                raise ImageGenError("image_audit_changed") from None
            return {"rollback_id": rollback_id, "kind": "image_rollback", "status": "rolled_back",
                    "removed": True, "path": str(record.path), "sha256": record.sha256,
                    "bytes": record.bytes, "restores_history": False,
                    "restores_execution_authority": False, "provider_fees_reversed": False}

    def preview(self, plan_id: str, session_id: str) -> GeneratedImage:
        with self._lock:
            plan = self._lookup(plan_id, session_id)
            self._source_guard()
            if plan.status != "generated" or plan.image is None:
                raise ImageGenError("image_preview_unavailable")
            self._verify_audit(plan.audit)
            return plan.image

    def invalidate(self) -> None:
        with self._lock:
            self._active = False
            for plan in self._plans.values():
                plan.image = None
            self._plans.clear()
            # Saved outputs already on disk are NOT deleted here, and no pending save/rollback
            # approval survives the scope fence. Audits are retained for provenance.
            self._saved.clear()
            self._save_plans.clear()
            self._rollback_plans.clear()
            # Do not delete audits, cancel remote work, or imply cloud fees can be undone.
            # In-flight calls retain their local plan but cannot publish after the scope fence.
            # _running is cleared only by that worker's matching token.
            return None

    # Rollback authority is limited to outputs this live instance saved; there is no audit
    # history replay that could re-grant write/execute authority after a restart or logout.
    # In-place audit/output guards detect identity/hash/link changes, not a filesystem
    # transaction against arbitrary concurrent writers: a hostile writer that wins the race
    # between the final check and the syscall is not fully excluded. Native handle ownership
    # hardening remains before this is exposed as a model tool or UI action.
    # Retention/maintenance is manual and is not part of model tools or this service.

__all__ = ["ImageGenerationService"]
