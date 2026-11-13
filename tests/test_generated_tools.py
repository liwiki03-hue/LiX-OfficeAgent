"""Adversarial governance checks. Fake host IO is confined to tmp_path fixtures."""
import builtins
import json
import os
import subprocess
from dataclasses import replace

import pytest

from office_agent.generated_tools import (
    ArtifactReceipt, GeneratedToolGovernance, GovernanceError, manifest_hash,
    proposal, validate_arguments,
)

SCHEMA = {"type": "object", "properties": {"x": {"type": "integer"}},
          "required": ["x"], "additionalProperties": False}
SOURCE = "def increment(x):\n    return x + 1\n"


def plan(g, operation="create", name="increment", **kwargs):
    if operation in ("create", "import", "edit"):
        kwargs = {"description": "Increment a number", "args_schema": SCHEMA,
                  "code": SOURCE.replace("increment", name), **kwargs}
    return g.plan(operation, name=name, **kwargs)


class FakeHost:
    """Trusted test adapter implementing backup/stage/reopen/audit/owned rollback.

    Never executes/imports source. This fake is not an OS isolation backend.
    """
    def __init__(self, root):
        self.root = root
        self.calls = []
        self.records = ()
        self.fail_at = None
        self.result = None
        self.manifest = root / "manifest.json"
        self.manifest.write_text("[]", encoding="utf-8")

    def transact(self, request):
        self.calls.append(request)
        before = manifest_hash(self.records)
        if before != request.expected_manifest_hash:
            return ArtifactReceipt(request.transaction_id, "rejected", before, before,
                                   False, False, False, True, True)
        folder = self.root / request.transaction_id
        folder.mkdir()
        backup = folder / "backup.json"
        stage = folder / "staging.json"
        old = self.manifest.read_bytes()
        backup.write_bytes(old)
        assert backup.read_bytes() == old
        payload = json.dumps([r.metadata() for r in request.desired_records], sort_keys=True).encode()
        try:
            stage.write_bytes(payload)
            if self.fail_at == "staging":
                raise OSError("fake staging failure")
            self.manifest.write_bytes(stage.read_bytes())
            if self.fail_at == "reopen":
                raise OSError("fake reopen failure")
            assert self.manifest.read_bytes() == payload
            json.loads(self.manifest.read_bytes())
            (folder / "audit.json").write_text('{"status":"verified"}', encoding="utf-8")
        except OSError:
            # Only artifacts belonging to this fake transaction are rolled back.
            self.manifest.write_bytes(backup.read_bytes())
            stage.unlink(missing_ok=True)
            (folder / "audit.json").write_text('{"status":"rolled_back"}', encoding="utf-8")
            return ArtifactReceipt(request.transaction_id, "rejected", before, before,
                                   True, True, False, True, True)
        self.records = request.desired_records
        receipt = ArtifactReceipt(request.transaction_id, "committed", before,
                                  manifest_hash(self.records), True, True, True, True, True)
        return self.result(receipt) if callable(self.result) else receipt


@pytest.mark.parametrize("source", [
    "import os\ndef increment(x): return x", "def increment(x):\n import os\n return x",
    "def increment(x): return eval(x)", "def increment(x): return exec(x)",
    "def increment(x): return compile(x, 'x', 'exec')", "def increment(x): return open(x)",
    "def increment(x): return __import__('os')", "def increment(x): return x.__class__",
    "def increment(x): return getattr(x, 'run')", "def increment(x): return globals()",
    "def increment(x): return locals()", "def increment(x): return type(x)",
    "def increment(x): return object()", "def increment(x): return bytes(x)",
    "def increment(x): return os.system(x)", "def increment(x): return subprocess.Popen(x)",
    "def increment(x): return socket.socket()", "def increment(x): return requests.get(x)",
    "@danger\ndef increment(x): return x", "def increment(x=danger()): return x",
    "def increment(x: danger()): return x", "async def increment(x): return x",
    "def increment(x):\n def nested(): return 1\n return x",
    "def increment(x): return [a for a in x]", "def increment(x): return lambda: x",
    "def increment(x): return x ** x", "def wrong(x): return x",
    "def increment(x): return x\nraise Exception()",
    "def increment(x, x): return x", "def increment(x): return eval",
    "def increment(x): return open", "def increment(x): return globals",
])
def test_adversarial_source_is_ineligible_without_execution(source):
    with pytest.raises(GovernanceError):
        proposal("increment", "Test", SCHEMA, source)


def test_source_bounds_metadata_and_version():
    r = proposal("increment", "Increment", SCHEMA, SOURCE)
    assert r.version == proposal("increment", "Increment", SCHEMA, SOURCE).version
    assert r.version != proposal("increment", "Increment", SCHEMA, SOURCE + "\n").version
    assert r.metadata()["scan"] == "eligibility_only_not_security"
    assert r.enabled is False
    for name in ("../evil", "C:\\evil", "__init__", "bad-name", "A", "a" * 49):
        with pytest.raises(GovernanceError):
            proposal(name, "Test", SCHEMA, SOURCE)
    for code in ("", SOURCE + " " * 8001, "\x00"):
        with pytest.raises(GovernanceError):
            proposal("increment", "Test", SCHEMA, code)
    for description in ("", " ", "a" * 501, "unsafe\ntext", " " * 501 + "x", "bad\x7f"):
        with pytest.raises(GovernanceError):
            proposal("increment", description, SCHEMA, SOURCE)


@pytest.mark.parametrize("schema", [
    {**SCHEMA, "additionalProperties": True}, {**SCHEMA, "$ref": "evil"},
    {**SCHEMA, "required": ["missing"]}, {**SCHEMA, "required": ["x", "x"]},
    {**SCHEMA, "properties": {"x": {"type": "string"}}},
    {**SCHEMA, "properties": {"x": {"type": "array", "items": {"type": "integer"}}}},
    {**SCHEMA, "properties": {"x": {"type": "integer", "default": 1}}},
    {**SCHEMA, "properties": {"x": {"type": "array", "items": {"type": "integer"}, "maxItems": True}}},
    {**SCHEMA, "properties": {f"x{i}": {"type": "integer"} for i in range(17)}},
])
def test_strict_schema_rejections(schema):
    with pytest.raises(GovernanceError):
        proposal("increment", "Test", schema, SOURCE)


def test_arguments_are_strict_and_bounded():
    validate_arguments(SCHEMA, {"x": 2})
    for value in ({"x": True}, {"x": "2"}, {}, {"x": 2, "proof": True}, []):
        with pytest.raises(GovernanceError):
            validate_arguments(SCHEMA, value)
    nested = {"type": "object", "properties": {"text": {"type": "string", "maxLength": 3},
              "numbers": {"type": "array", "items": {"type": "number"}, "maxItems": 2}},
              "required": ["text", "numbers"], "additionalProperties": False}
    validate_arguments(nested, {"text": "ok", "numbers": [1, 2.5]})
    validate_arguments(nested, {"text": "ok", "numbers": [10 ** 400]})
    for value in ({"text": "long", "numbers": []}, {"text": "ok", "numbers": [1, 2, 3]},
                  {"text": "ok", "numbers": [float("nan")]}, {"text": "ok", "numbers": [True]}):
        with pytest.raises(GovernanceError):
            validate_arguments(nested, value)


def test_callback_free_plans_detached_preview_and_literal_approval(tmp_path):
    host = FakeHost(tmp_path)
    g = GeneratedToolGovernance(backend=host)
    original = g.manifest_hash
    before_bytes = host.manifest.read_bytes()
    p = plan(g)
    assert original == g.manifest_hash == p["manifest_hash"]
    assert host.manifest.read_bytes() == before_bytes and not host.calls
    p["after"]["code"] = "evil"
    p["after"]["args_schema"]["properties"].clear()
    for value in (False, 1, "True", {}, None):
        with pytest.raises(GovernanceError, match="literal"):
            g.approve(p["plan_id"], confirmed=value)
    assert not host.calls
    result = g.approve(p["plan_id"], confirmed=True)
    assert result["status"] == "committed" and result["executed"] is False and result["registered"] is False
    assert g.read("increment")["code"] == SOURCE
    assert host.records[0].enabled is False
    assert (tmp_path / p["plan_id"] / "backup.json").read_bytes() == before_bytes
    assert (tmp_path / p["plan_id"] / "audit.json").exists()
    with pytest.raises(GovernanceError):
        g.approve(p["plan_id"], confirmed=True)


def test_session_binding_ttl_bounds_and_stale_manifest(tmp_path):
    host = FakeHost(tmp_path)
    now = [10.0]
    g = GeneratedToolGovernance(backend=host, clock=lambda: now[0])
    other = GeneratedToolGovernance(backend=host)
    p = plan(g)
    with pytest.raises(GovernanceError, match="private session"):
        other.approve(p["plan_id"], confirmed=True)
    now[0] = 310.0
    with pytest.raises(GovernanceError, match="expired"):
        g.approve(p["plan_id"], confirmed=True)
    assert not host.calls
    a, b = plan(g, name="one"), plan(g, name="two")
    assert g.approve(a["plan_id"], confirmed=True)["status"] == "committed"
    with pytest.raises(GovernanceError, match="manifest changed"):
        g.approve(b["plan_id"], confirmed=True)
    assert len(host.calls) == 1
    for i in range(8):
        plan(g, name=f"pending{i}")
    with pytest.raises(GovernanceError, match="bound"):
        plan(g, name="overflow")
    now[0] += 300
    p = plan(g, name="after_expiry")
    g.cancel(p["plan_id"])
    assert len(host.calls) == 1


def test_lifetime_bound_cannot_be_reset_by_cancellation():
    g = GeneratedToolGovernance()
    for i in range(256):
        p = plan(g, name=f"item{i}")
        g.cancel(p["plan_id"])
    with pytest.raises(GovernanceError, match="bound"):
        plan(g, name="last")


def test_management_edit_import_enable_delete_and_immutable_builtins(tmp_path):
    host = FakeHost(tmp_path)
    g = GeneratedToolGovernance(backend=host, builtin_names=("read_word",))
    assert g.list_tools() == [{"name": "read_word", "kind": "builtin", "management": "delegated_immutable"}]
    for operation in ("create", "edit", "delete", "disable", "enable", "import"):
        with pytest.raises(GovernanceError, match="builtin"):
            plan(g, operation, name="read_word")
    assert not host.calls
    p = plan(g, "import")
    g.approve(p["plan_id"], confirmed=True)
    version = g.read("increment")["version"]
    for op in ("enable", "disable"):
        p = plan(g, op)
        g.approve(p["plan_id"], confirmed=True)
        assert g.read("increment")["enabled"] is (op == "enable")
        assert g.run("increment", {"x": 1})["status"] == "sandbox_required"
    p = plan(g, "edit", code=SOURCE.replace("+ 1", "+ 2"))
    g.approve(p["plan_id"], confirmed=True)
    assert g.read("increment")["version"] != version
    assert g.read("increment")["enabled"] is False
    p = plan(g, "delete")
    assert g.read("increment")["code"]
    g.approve(p["plan_id"], confirmed=True)
    assert len(g.list_tools()) == 1


@pytest.mark.parametrize("fail_at", ["staging", "reopen"])
def test_owned_rollback_protocol_preserves_state(tmp_path, fail_at):
    host = FakeHost(tmp_path)
    host.fail_at = fail_at
    g = GeneratedToolGovernance(backend=host)
    original = host.manifest.read_bytes()
    p = plan(g)
    assert g.approve(p["plan_id"], confirmed=True)["status"] == "rejected"
    assert not g.uncertain and g.list_tools() == []
    assert host.manifest.read_bytes() == original
    assert (tmp_path / p["plan_id"] / "backup.json").read_bytes() == original
    assert json.loads((tmp_path / p["plan_id"] / "audit.json").read_text())["status"] == "rolled_back"


@pytest.mark.parametrize("bad_result", [
    lambda r: None, lambda r: {"status": "committed", "proof": True},
    lambda r: replace(r, transaction_id="other"), lambda r: replace(r, before_hash="changed"),
    lambda r: replace(r, after_hash="wrong"), lambda r: replace(r, backup_verified=False),
    lambda r: replace(r, staging_verified=False), lambda r: replace(r, reopen_verified=False),
    lambda r: replace(r, audit_recorded=False), lambda r: replace(r, owned_rollback=False),
    lambda r: replace(r, status="unknown"), lambda r: replace(r, audit_recorded=1),
])
def test_unknown_backend_result_trips_sticky_fuse(tmp_path, bad_result):
    host = FakeHost(tmp_path)
    host.result = bad_result
    g = GeneratedToolGovernance(backend=host)
    a, b = plan(g), plan(g, name="other")
    assert g.approve(a["plan_id"], confirmed=True)["status"] == "backend_uncertain"
    assert g.uncertain and g.list_tools() == []
    host.result = None
    with pytest.raises(GovernanceError, match="fused"):
        g.approve(b["plan_id"], confirmed=True)
    with pytest.raises(GovernanceError, match="fused"):
        plan(g, name="third")
    assert len(host.calls) == 1
    assert g.run(proof="trusted")["status"] == "sandbox_required"


def test_backend_exception_is_uncertain_and_reentrant_approval_blocked(tmp_path):
    host = FakeHost(tmp_path)
    g = GeneratedToolGovernance(backend=host)
    a, b = plan(g), plan(g, name="other")
    def callback(receipt):
        with pytest.raises(GovernanceError, match="fused"):
            g.approve(b["plan_id"], confirmed=True)
        raise OSError("unknown publish outcome")
    host.result = callback
    assert g.approve(a["plan_id"], confirmed=True)["status"] == "backend_uncertain"
    assert g.uncertain and len(host.calls) == 1


def test_host_initial_record_revalidation_and_detached_metadata():
    schema = {"type": "object", "properties": {"z": {"type": "integer"}, "a": {"type": "integer"}},
              "required": ["z", "a"], "additionalProperties": False}
    record = proposal("add", "Add", schema, "def add(z, a): return z + a")
    g = GeneratedToolGovernance(initial_records=(record,))
    info = g.read("add")
    info["args_schema"]["properties"].clear()
    assert len(g.read("add")["args_schema"]["properties"]) == 2
    for records in ((replace(record, version="fake"),), (record, record)):
        with pytest.raises(GovernanceError):
            GeneratedToolGovernance(initial_records=records)


def test_no_backend_and_default_runtime_never_execute_or_launch(monkeypatch):
    g = GeneratedToolGovernance()
    p = plan(g)
    assert g.approve(p["plan_id"], confirmed=True)["status"] == "artifact_backend_required"
    calls = []
    def forbidden(*args, **kwargs):
        calls.append((args, kwargs))
        raise AssertionError("generated execution must be impossible")
    for target, names in ((builtins, ("exec", "eval", "compile", "open")),
                          (subprocess, ("Popen", "run", "call", "check_call", "check_output")),
                          (os, ("system", "popen", "spawnv", "spawnve", "execv", "execve"))):
        for name in names:
            monkeypatch.setattr(target, name, forbidden)
    original_import = builtins.__import__
    builtins.__import__ = forbidden
    try:
        for claims in ({}, {"confirmed": True}, {"proof": "trusted"},
                       {"receipt": {"fs_denied": True, "network_denied": True, "process_denied": True}},
                       {"isolated_executor": forbidden, "capability": object()}):
            result = g.run("increment", {"code": "__import__('os').system('evil')", "callback": forbidden}, **claims)
            assert result["status"] == "sandbox_required" and result["executed"] is False
        assert not calls
    finally:
        # Restore builtins before pytest assertion/reporting internals run.
        builtins.__import__ = original_import
        monkeypatch.undo()
