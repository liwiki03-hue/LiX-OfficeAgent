"""P3 host service: fake credentials, synthetic pixels, no Office/network, no writes outside tmp."""
import hashlib
import io
import json
import os
import threading
from dataclasses import replace
from pathlib import Path

import pytest
from PIL import Image

from office_agent import imagetransactions as service_module
from office_agent.imagegen import ImageGenConfig, ImageGenError, ImageRequest
from office_agent.imagetransactions import ImageGenerationService
from office_agent.writes import GuardError

KEY = "synthetic-independent-image-key"
SESSION = "synthetic-session"
PROMPT = "synthetic private prompt blue squares"


def png():
    stream = io.BytesIO()
    Image.new("RGB", (12, 8), "blue").save(stream, format="PNG")
    return stream.getvalue()


@pytest.fixture
def setup(tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"synthetic source bytes - not an Office input")
    config = [ImageGenConfig("openai_compatible", KEY, "https://image.example.test/v1", "synthetic-model")]
    calls = []
    action = [lambda: png()]
    class FakeProvider:
        def generate(self, request, actual_config):
            calls.append((request, actual_config))
            return action[0]()
    def make(**kwargs):
        return ImageGenerationService(tmp_path, source, SESSION, config_loader=lambda: config[0],
                                      provider_factory=FakeProvider, **kwargs)
    return tmp_path, source, config, calls, action, make


def fail(code, function, *args):
    with pytest.raises(ImageGenError) as error:
        function(*args)
    assert str(error.value) == code


def plan(service):
    return service.plan(ImageRequest(prompt=PROMPT))["plan_id"]


def record_path(root, plan_id):
    return root / ".office_agent" / "imagegen" / plan_id / "audit.json"


def test_plan_readonly_exact_prompt_public_copy_and_no_save_authority(setup):
    root, source, configs, calls, actions, make = setup
    service = make(expected_version=hashlib.sha256(source.read_bytes()).hexdigest())
    public = service.plan(ImageRequest(prompt=PROMPT, style="vivid"))
    assert public["effective_prompt"] == "Requested visual style: vivid.\n" + PROMPT
    assert public["confirmation_required"] is True and public["save_confirmation_required"] is True
    assert public["cost_unknown"] and public["provider_fees_irreversible"]
    assert KEY not in json.dumps(public) and "fingerprint" not in json.dumps(public)
    assert calls == [] and not (root / ".office_agent").exists()
    assert list(root.iterdir()) == [source]
    public["configuration"]["model"] = "caller-corruption"
    public["effective_prompt"] = "caller-corruption"
    service.confirm_generation(public["plan_id"], SESSION, True)
    assert calls[0][0].effective_prompt == "Requested visual style: vivid.\n" + PROMPT
    assert calls[0][1].model == "synthetic-model"
    assert not list(root.glob("*.png"))
    # Generation approval alone never saves: a separate save approval is still required.
    fail("image_save_invalid", service.confirm_save, public["plan_id"], SESSION, True)


def test_audit_exists_before_provider_no_prompt_key_pixels_saved(setup):
    root, source, configs, calls, actions, make = setup
    original = source.read_bytes()
    service = make()
    identifier = plan(service)
    path = record_path(root, identifier)
    def callback():
        audit = json.loads(path.read_text(encoding="utf-8"))
        assert audit["status"] == "generating" and audit["saved"] is False
        assert audit["provider_request_state"] == "attempt_committed"
        assert audit["prompt_sha256"] == hashlib.sha256(calls[0][0].effective_prompt.encode()).hexdigest()
        return png()
    actions[0] = callback
    result = service.confirm_generation(identifier, SESSION, True)
    image = service.preview(identifier, SESSION)
    assert result["status"] == "generated" and result["transaction_id"] == identifier
    assert image.raw.startswith(b"\x89PNG") and image.metadata["width"] == 12
    assert image.metadata["sha256"] == hashlib.sha256(image.raw).hexdigest()
    text = path.read_text(encoding="utf-8")
    assert KEY not in text and PROMPT not in text and "data:image/" not in text
    assert "raw" not in result and "b64_json" not in result
    assert json.loads(text)["image"] == dict(image.metadata)
    assert source.read_bytes() == original and not list(root.rglob("*.png"))
    assert len(calls) == 1
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)


@pytest.mark.parametrize("consent", [1, 0, "true", "false", None, [], {}])
def test_strict_boolean_consent(setup, consent):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    fail("image_confirmation_required", service.confirm_generation, identifier, SESSION, consent)
    assert calls == [] and not (root / ".office_agent").exists()


def test_cancel_consumes_without_any_files(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    assert service.confirm_generation(identifier, SESSION, False)["status"] == "cancelled"
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()


def test_wrong_session_unknown_plan_do_not_consume(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    fail("image_operation_invalidated", service.confirm_generation, identifier, "other-session", True)
    fail("image_plan_invalid", service.confirm_generation, "unknown", SESSION, True)
    service.confirm_generation(identifier, SESSION, True)
    fail("image_operation_invalidated", service.preview, identifier, "other-session")
    assert len(calls) == 1


@pytest.mark.parametrize("field,value", [("api_key", "synthetic-new-key"), ("model", "changed-model"),
                                       ("base_url", "https://changed.example.test/v1"),
                                       ("response_format", "b64_json"), ("timeout", 30)])
def test_changed_configuration_consumes_before_network(setup, field, value):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    configs[0] = replace(configs[0], **{field: value})
    fail("image_configuration_changed", service.confirm_generation, identifier, SESSION, True)
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()


@pytest.mark.parametrize("change", ["modify", "replace_same_bytes", "restore_bytes"])
def test_source_change_before_approval_no_call(setup, change):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    before = source.read_bytes()
    before_stat = source.stat()
    if change == "replace_same_bytes":
        replacement = root / "replacement.docx"
        replacement.write_bytes(before)
        replacement.replace(source)
    else:
        source.write_bytes(b"changed")
        if change == "restore_bytes":
            source.write_bytes(before)
            # This case exercises changed observable metadata with unchanged SHA.
            # Windows may coalesce immediate writes into the same timestamp tick;
            # snapshot comparison is NOT a journal of transient mutation history.
            os.utime(source, ns=(before_stat.st_atime_ns, before_stat.st_mtime_ns + 2_000_000_000))
    fail("image_source_changed", service.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()


def test_expiry_consumes_without_audit(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    expiry = service._plans[identifier].expires
    monkeypatch.setattr(service_module.time, "monotonic", lambda: expiry + 1)
    fail("image_approval_expired", service.confirm_generation, identifier, SESSION, True)
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()


@pytest.mark.parametrize("mode", ["parent_file", "unique_collision", "fsync_failure", "audit_cap", "guard_failure"])
def test_audit_failure_prevents_any_provider_call(setup, monkeypatch, mode):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    if mode == "parent_file":
        (root / ".office_agent").write_bytes(b"foreign file")
    elif mode == "unique_collision":
        foreign = record_path(root, identifier)
        foreign.parent.mkdir(parents=True)
        foreign.write_bytes(b"foreign audit")
    elif mode == "fsync_failure":
        def broken(fd):
            raise OSError(KEY)
        monkeypatch.setattr(service_module.os, "fsync", broken)
    elif mode == "audit_cap":
        monkeypatch.setattr(service_module, "MAX_AUDIT_BYTES", 16)
    else:
        original = service_module._audit_path
        def guarded(path, **kwargs):
            if path.name == "imagegen":
                raise GuardError(KEY)
            return original(path, **kwargs)
        monkeypatch.setattr(service_module, "_audit_path", guarded)
    fail("image_audit_failed", service.confirm_generation, identifier, SESSION, True)
    assert calls == []
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)
    if mode == "unique_collision":
        assert foreign.read_bytes() == b"foreign audit"
    if mode == "parent_file":
        assert (root / ".office_agent").read_bytes() == b"foreign file"


@pytest.mark.parametrize("mode", ["provider_error", "svg", "source_changed", "config_changed", "audit_edited", "audit_replaced"])
def test_inflight_failure_never_publishes_or_retries(setup, mode):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    path = record_path(root, identifier)
    foreign = []
    def action():
        if mode == "provider_error":
            raise RuntimeError(KEY + PROMPT)
        if mode == "svg":
            return b"<svg/>"
        if mode == "source_changed":
            source.write_bytes(b"external source edit")
        if mode == "config_changed":
            configs[0] = replace(configs[0], api_key="new-private-key")
        if mode == "audit_edited":
            path.write_bytes(b"foreign edit preserve")
            foreign.append(path.read_bytes())
        if mode == "audit_replaced":
            replacement = path.with_name("replacement.json")
            replacement.write_bytes(path.read_bytes())
            replacement.replace(path)
            foreign.append(path.read_bytes())
        return png()
    actions[0] = action
    expected = {"provider_error": "image_generation_failed", "svg": "image_generation_failed",
                "source_changed": "image_source_changed", "config_changed": "image_configuration_changed",
                "audit_edited": "image_audit_changed", "audit_replaced": "image_audit_changed"}[mode]
    fail(expected, service.confirm_generation, identifier, SESSION, True)
    assert service._plans[identifier].image is None and len(calls) == 1
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)
    if foreign:
        assert path.read_bytes() == foreign[0]
    else:
        audit = json.loads(path.read_text(encoding="utf-8"))
        assert audit["status"] == "failed" and KEY not in json.dumps(audit) and PROMPT not in json.dumps(audit)


def test_failure_diagnostics_are_safe_and_preserve_possible_fees(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    actions[0] = lambda: (_ for _ in ()).throw(ImageGenError("image_response_timeout"))
    with pytest.raises(ImageGenError) as captured:
        service.confirm_generation(identifier, SESSION, True)
    error = captured.value
    assert error.code == "image_generation_failed"
    assert (error.stage, error.detail_code, error.request_state) == (
        "response", "image_response_timeout", "attempt_committed")
    assert len(calls) == 1
    assert KEY not in repr(vars(error)) and PROMPT not in repr(vars(error))
    assert service._plans[identifier].public["failure_detail"] == "image_response_timeout"
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)

    blocked = make()
    identifier = plan(blocked)
    monkeypatch.setattr(blocked, "_create_audit", lambda *_: (_ for _ in ()).throw(OSError(KEY)))
    with pytest.raises(ImageGenError) as captured:
        blocked.confirm_generation(identifier, SESSION, True)
    assert (captured.value.code, captured.value.stage, captured.value.request_state) == (
        "image_audit_failed", "preflight", "not_started")
    assert len(calls) == 1


def test_invalidation_while_network_waits_does_not_block_or_publish(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    entered, release = threading.Event(), threading.Event()
    errors = []
    def action():
        entered.set()
        if not release.wait(timeout=5):
            raise RuntimeError("bounded synthetic wait expired")
        return png()
    actions[0] = action
    def worker():
        try:
            service.confirm_generation(identifier, SESSION, True)
        except ImageGenError as error:
            errors.append(error.code)
    thread = threading.Thread(target=worker)
    thread.start()
    try:
        assert entered.wait(timeout=5)
        fail("image_operation_busy", service.plan, ImageRequest(prompt=PROMPT))
        service.invalidate()  # would block until fake timeout if lock were held over network
    finally:
        release.set()
        thread.join(timeout=6)
    assert not thread.is_alive()
    assert errors == ["image_operation_invalidated"] and service._plans == {} and service._running is None
    assert json.loads(record_path(root, identifier).read_text(encoding="utf-8"))["status"] == "failed"
    fail("image_operation_invalidated", service.preview, identifier, SESSION)
    assert len(calls) == 1 and not list(root.rglob("*.png"))


def test_invalidate_prior_approval_and_no_restart_authority(setup):
    root, source, configs, calls, actions, make = setup
    old = make()
    identifier = plan(old)
    old.invalidate()
    fail("image_operation_invalidated", old.confirm_generation, identifier, SESSION, True)
    new = make()
    fail("image_plan_invalid", new.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()


def test_finite_retention_and_preview_tamper_guard(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifiers = [plan(service) for _ in range(service_module.MAX_PLANS)]
    fail("image_plan_limit", service.plan, ImageRequest(prompt=PROMPT))
    service.confirm_generation(identifiers[0], SESSION, True)
    record_path(root, identifiers[0]).write_bytes(b"external audit change")
    fail("image_audit_changed", service.preview, identifiers[0], SESSION)


@pytest.mark.parametrize("bad", ["", "../escape", "a b", "x" * 129, 3, None])
def test_session_format_is_strict(setup, bad):
    root, source, configs, calls, actions, make = setup
    fail("image_session_invalid", ImageGenerationService, root, source, bad)


@pytest.mark.parametrize("expected", ["0" * 64, "invalid", 1])
def test_expected_version_is_not_rebound(setup, expected):
    root, source, configs, calls, actions, make = setup
    with pytest.raises(ImageGenError, match="image_source_invalid"):
        make(expected_version=expected)
    assert calls == []


def test_constructor_limits_and_source_directory(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    nested = root / "nested"
    nested.mkdir()
    document = nested / "other.docx"
    document.write_bytes(b"synthetic")
    fail("image_source_invalid", ImageGenerationService, root, document, SESSION)
    fail("image_source_invalid", ImageGenerationService, root, Path("relative.docx"), SESSION)
    monkeypatch.setattr(service_module, "MAX_SOURCE_BYTES", 4)
    fail("image_source_invalid", make)


def test_opened_ctime_difference_never_weakens_identity_mtime_sha_checks(setup, monkeypatch):
    from types import SimpleNamespace
    root, source, configs, calls, actions, make = setup
    original_fstat = service_module.os.fstat
    def different_ctime(fd):
        info = original_fstat(fd)
        return SimpleNamespace(st_dev=info.st_dev, st_ino=info.st_ino, st_size=info.st_size,
                               st_mtime_ns=info.st_mtime_ns, st_ctime_ns=info.st_ctime_ns + 123456789,
                               st_nlink=info.st_nlink)
    monkeypatch.setattr(service_module.os, "fstat", different_ctime)
    service = make()
    identifier = plan(service)
    assert service.confirm_generation(identifier, SESSION, True)["status"] == "generated"
    assert len(calls) == 1
    source.write_bytes(b"external modification")
    fail("image_source_changed", service.preview, identifier, SESSION)


def test_added_audit_hardlink_blocks_updates_and_preserves_both_aliases(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    path = record_path(root, identifier)
    alias = root / "external-audit-link.json"
    before = []
    def add_alias():
        before.append(path.read_bytes())
        alias.hardlink_to(path)
        return png()
    actions[0] = add_alias
    fail("image_audit_changed", service.confirm_generation, identifier, SESSION, True)
    assert path.read_bytes() == before[0] == alias.read_bytes()
    assert path.stat().st_nlink == 2
    assert service._plans[identifier].image is None and len(calls) == 1


def test_missing_invalid_config_and_public_exception_sanitizing(setup):
    root, source, configs, calls, actions, make = setup
    def missing():
        raise ImageGenError("image_configuration_missing")
    def private_failure():
        raise RuntimeError(KEY)
    for loader, code in [(missing, "image_configuration_missing"), (private_failure, "image_configuration_invalid"),
                         (lambda: None, "image_configuration_invalid")]:
        service = ImageGenerationService(root, source, SESSION, config_loader=loader)
        fail(code, service.plan, ImageRequest(prompt=PROMPT))
    assert calls == [] and not (root / ".office_agent").exists()


def generated(service):
    identifier = plan(service)
    assert service.confirm_generation(identifier, SESSION, True)["status"] == "generated"
    return identifier


def outputs(root):
    return sorted(path.name for path in root.glob("office-image-*.png"))


def test_save_plan_is_readonly_and_needs_its_own_approval(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    before = outputs(root)
    approval = service.plan_save(identifier, SESSION)
    assert outputs(root) == before and calls and len(calls) == 1
    assert approval["new_file"] and approval["overwrites_source"] is False
    assert approval["overwrites_existing"] is False and approval["replaces_generation_approval"] is False
    assert approval["saving_does_not_reverse_fees"] is True and approval["confirmation_required"] is True
    assert Path(approval["path"]).parent == root and approval["filename"] == Path(approval["path"]).name
    assert approval["sha256"] == service.preview(identifier, SESSION).metadata["sha256"]
    assert "data:image/" not in json.dumps(approval) and KEY not in json.dumps(approval)
    fail("image_plan_invalid", service.plan_save, "unknown-plan", SESSION)
    fail("image_operation_invalidated", service.plan_save, identifier, "other-session")


def test_confirm_save_creates_one_new_validated_file_with_audit(setup):
    root, source, configs, calls, actions, make = setup
    source_before = source.read_bytes()
    service = make()
    identifier = generated(service)
    approval = service.plan_save(identifier, SESSION)
    result = service.confirm_save(approval["save_id"], SESSION, True)
    target = Path(result["path"])
    assert result["status"] == "saved" and result["written"] and result["new_file"]
    assert result["source_overwritten"] is False and outputs(root) == [target.name]
    raw = target.read_bytes()
    assert raw == service.preview(identifier, SESSION).raw and hashlib.sha256(raw).hexdigest() == result["sha256"]
    with Image.open(target) as image:
        assert image.format == "PNG" and image.size == (12, 8)
    audit = json.loads(record_path(root, identifier).read_text(encoding="utf-8"))
    assert audit["saved"] is True and audit["saved_output"]["sha256"] == result["sha256"]
    assert audit["saved_output"]["path"] == str(target)
    assert audit["saved_output"]["image"]["width"] == 12 and "raw" not in json.dumps(audit)
    assert "data:image/" not in json.dumps(audit) and PROMPT not in json.dumps(audit)
    assert source.read_bytes() == source_before and len(calls) == 1
    fail("image_save_invalid", service.confirm_save, approval["save_id"], SESSION, True)
    assert outputs(root) == [target.name]


@pytest.mark.parametrize("consent", [1, 0, "true", "false", None, [], {}])
def test_save_strict_boolean_does_not_consume(setup, consent):
    root, source, configs, calls, actions, make = setup
    service = make()
    approval = service.plan_save(generated(service), SESSION)
    fail("image_confirmation_required", service.confirm_save, approval["save_id"], SESSION, consent)
    assert outputs(root) == []
    assert service.confirm_save(approval["save_id"], SESSION, False)["status"] == "cancelled"
    fail("image_save_invalid", service.confirm_save, approval["save_id"], SESSION, True)
    assert outputs(root) == [] and len(calls) == 1


def test_save_cancel_expiry_and_existing_target_never_overwrite(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    cancelled = service.plan_save(identifier, SESSION)
    assert service.confirm_save(cancelled["save_id"], SESSION, False) == {
        "save_id": cancelled["save_id"], "kind": "image_save", "status": "cancelled",
        "written": False, "path": None}
    assert outputs(root) == []
    expired = service.plan_save(identifier, SESSION)
    approval = service._save_plans[expired["save_id"]]
    monkeypatch.setattr(service_module.time, "monotonic", lambda: approval.expires + 1)
    fail("image_approval_expired", service.confirm_save, expired["save_id"], SESSION, True)
    fail("image_save_invalid", service.confirm_save, expired["save_id"], SESSION, True)
    monkeypatch.undo()
    # A file that appeared before the final guard is reported, never opened for writing.
    guarded = service.plan_save(identifier, SESSION)
    foreign = Path(guarded["path"])
    foreign.write_bytes(b"foreign pre-existing file")
    fail("image_output_invalid", service.confirm_save, guarded["save_id"], SESSION, True)
    assert foreign.read_bytes() == b"foreign pre-existing file" and outputs(root) == [foreign.name]
    # A file that appears after the guard is refused by the exclusive create itself.
    raced = service.plan_save(identifier, SESSION)
    raced_path = Path(raced["path"])
    raced_path.write_bytes(b"foreign race winner")
    monkeypatch.setattr(service, "_target_guard", lambda target: target)
    fail("image_output_exists", service.confirm_save, raced["save_id"], SESSION, True)
    assert raced_path.read_bytes() == b"foreign race winner"
    assert sorted(outputs(root)) == sorted([foreign.name, raced_path.name])
    audit = json.loads(record_path(root, identifier).read_text(encoding="utf-8"))
    assert audit["saved"] is False and "saved_output" not in audit


def test_save_requires_generated_state_and_root_identity(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    planned = plan(service)
    fail("image_preview_unavailable", service.plan_save, planned, SESSION)
    identifier = generated(service)
    monkeypatch.setattr(service, "_root_identity", (1, 2))
    fail("image_source_changed", service.plan_save, identifier, SESSION)
    assert outputs(root) == []


def test_save_audit_failure_removes_only_owned_unaudited_file(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    approval = service.plan_save(identifier, SESSION)
    def broken(ticket, changes):
        raise RuntimeError(KEY)
    monkeypatch.setattr(service, "_update_audit", broken)
    fail("image_audit_failed", service.confirm_save, approval["save_id"], SESSION, True)
    assert outputs(root) == [] and not (root / ".office_agent").is_symlink()
    audit = json.loads(record_path(root, identifier).read_text(encoding="utf-8"))
    assert audit["saved"] is False


def test_save_audit_failure_preserves_file_changed_concurrently(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    approval = service.plan_save(identifier, SESSION)
    def racy(ticket, changes):
        Path(approval["path"]).write_bytes(b"foreign concurrent rewrite")
        raise RuntimeError(KEY)
    monkeypatch.setattr(service, "_update_audit", racy)
    fail("image_audit_changed", service.confirm_save, approval["save_id"], SESSION, True)
    assert Path(approval["path"]).read_bytes() == b"foreign concurrent rewrite"
    assert json.loads(record_path(root, identifier).read_text(encoding="utf-8"))["saved"] is False


def test_rollback_requires_fresh_approval_and_deletes_unchanged_output(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    saved = service.confirm_save(service.plan_save(identifier, SESSION)["save_id"], SESSION, True)
    target = Path(saved["path"])
    rollback = service.plan_rollback(saved["save_id"], SESSION)
    assert target.exists() and rollback["restores_history"] is False
    assert rollback["deletes_only_unchanged_owned_output"] and rollback["preserves_edited_or_replaced"]
    assert rollback["provider_fees_reversed"] is False and rollback["confirmation_required"] is True
    cancelled = service.confirm_rollback(rollback["rollback_id"], SESSION, False)
    assert cancelled["removed"] is False and target.exists()
    fresh = service.plan_rollback(saved["save_id"], SESSION)
    result = service.confirm_rollback(fresh["rollback_id"], SESSION, True)
    assert result["status"] == "rolled_back" and result["removed"] is True
    assert result["restores_execution_authority"] is False and not target.exists()
    audit = json.loads(record_path(root, identifier).read_text(encoding="utf-8"))
    assert audit["saved"] is False and audit["output_removed"] is True and audit["status"] == "rolled_back"
    fail("image_rollback_invalid", service.confirm_rollback, fresh["rollback_id"], SESSION, True)
    fail("image_rollback_invalid", service.plan_rollback, saved["save_id"], SESSION)


@pytest.mark.parametrize("consent", [1, 0, "true", None, [], {}])
def test_rollback_strict_boolean_does_not_delete(setup, consent):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    saved = service.confirm_save(service.plan_save(identifier, SESSION)["save_id"], SESSION, True)
    rollback = service.plan_rollback(saved["save_id"], SESSION)
    fail("image_confirmation_required", service.confirm_rollback, rollback["rollback_id"], SESSION, consent)
    assert Path(saved["path"]).exists()


@pytest.mark.parametrize("mode", ["edited", "replaced_same_bytes", "missing", "tampered_audit", "wrong_session"])
def test_rollback_preserves_anything_it_cannot_prove_it_owns(setup, mode):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    saved = service.confirm_save(service.plan_save(identifier, SESSION)["save_id"], SESSION, True)
    target = Path(saved["path"])
    original = target.read_bytes()
    rollback = service.plan_rollback(saved["save_id"], SESSION)
    session = "other-session" if mode == "wrong_session" else SESSION
    if mode == "edited":
        target.write_bytes(b"human edited the saved image")
    elif mode == "replaced_same_bytes":
        replacement = root / "replacement.png"
        replacement.write_bytes(original)
        replacement.replace(target)
    elif mode == "missing":
        target.unlink()
    elif mode == "tampered_audit":
        record_path(root, identifier).write_bytes(b"foreign audit")
    expected = {"wrong_session": "image_operation_invalidated", "tampered_audit": "image_audit_changed"}.get(
        mode, "image_rollback_changed")
    fail(expected, service.confirm_rollback, rollback["rollback_id"], session, True)
    if mode == "edited":
        assert target.read_bytes() == b"human edited the saved image"
    elif mode == "replaced_same_bytes":
        assert target.read_bytes() == original
    elif mode in ("wrong_session", "tampered_audit"):
        assert target.read_bytes() == original
    else:
        assert not target.exists()


def test_invalidate_drops_approvals_but_keeps_saved_output(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    saved = service.confirm_save(service.plan_save(identifier, SESSION)["save_id"], SESSION, True)
    pending = service.plan_save(identifier, SESSION)
    rollback = service.plan_rollback(saved["save_id"], SESSION)
    service.invalidate()
    target = Path(saved["path"])
    assert target.exists()
    fail("image_operation_invalidated", service.plan_rollback, saved["save_id"], SESSION)
    fail("image_operation_invalidated", service.confirm_save, pending["save_id"], SESSION, True)
    fail("image_operation_invalidated", service.confirm_rollback, rollback["rollback_id"], SESSION, True)
    assert target.exists() and outputs(root) == [target.name]


def test_save_and_rollback_plan_limits_and_busy_fence(setup):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = generated(service)
    approvals = [service.plan_save(identifier, SESSION) for _ in range(service_module.MAX_PLANS)]
    fail("image_plan_limit", service.plan_save, identifier, SESSION)
    saved = service.confirm_save(approvals[0]["save_id"], SESSION, True)
    rollbacks = [service.plan_rollback(saved["save_id"], SESSION) for _ in range(service_module.MAX_PLANS)]
    fail("image_plan_limit", service.plan_rollback, saved["save_id"], SESSION)
    service._running = object()
    try:
        fail("image_operation_busy", service.plan_save, identifier, SESSION)
        fail("image_operation_busy", service.plan_rollback, saved["save_id"], SESSION)
        fail("image_operation_busy", service.confirm_rollback, rollbacks[0]["rollback_id"], SESSION, True)
    finally:
        service._running = None
    assert Path(saved["path"]).exists()


def test_source_guard_rejects_changed_hash_even_with_same_stamp(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    baseline_stamp, baseline_sha = service._source_stamp, service._source_sha
    changed_sha = hashlib.sha256(b"different snapshot bytes").hexdigest()
    assert changed_sha != baseline_sha
    monkeypatch.setattr(service, "_source_snapshot", lambda: (baseline_stamp, changed_sha))
    fail("image_source_changed", service.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)


def test_source_guard_rejects_changed_stamp_even_with_same_hash(setup, monkeypatch):
    root, source, configs, calls, actions, make = setup
    service = make()
    identifier = plan(service)
    baseline_stamp, baseline_sha = service._source_stamp, service._source_sha
    changed_stamp = (*baseline_stamp[:3], baseline_stamp[3] + 2_000_000_000, baseline_stamp[4])
    assert changed_stamp != baseline_stamp
    monkeypatch.setattr(service, "_source_snapshot", lambda: (changed_stamp, baseline_sha))
    fail("image_source_changed", service.confirm_generation, identifier, SESSION, True)
    assert calls == [] and not (root / ".office_agent").exists()
    fail("image_approval_consumed", service.confirm_generation, identifier, SESSION, True)
