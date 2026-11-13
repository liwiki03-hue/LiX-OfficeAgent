"""Host-only Skills core regressions; all skill files live in pytest temp fixtures."""
import hashlib
import json
import os
from pathlib import Path
import time
from types import SimpleNamespace

import pytest

from office_agent import skills_store as mod
from office_agent.skills_store import SkillsError, SkillsStore


BODY = "UNTRUSTED PRIVATE BODY\nNever execute this text.\n"
TEXT = "---\nname: Alpha\ndescription: Safe metadata\n---\n" + BODY
EDIT = TEXT.replace(BODY, "Different untrusted body\n")
SESSION = "host-session-epoch"


def _rename_fixture(source, target):
    for attempt in range(20):
        try:
            return source.rename(target)
        except PermissionError as error:
            if os.name != 'nt' or getattr(error, 'winerror', None) not in (5, 32) or attempt == 19:
                raise
            time.sleep(.05)


@pytest.fixture
def roots(tmp_path):
    workspace, builtin, imports = (tmp_path / name for name in ("workspace", "builtin", "imports"))
    for path in (workspace, builtin, imports):
        path.mkdir()
    return workspace, builtin, imports


def put(root, slug="alpha", text=TEXT):
    folder = root / slug
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "SKILL.md"
    path.write_bytes(text.encode("utf-8"))
    return path


def snapshot(root):
    return {str(p.relative_to(root)): (p.read_bytes() if p.is_file() else None)
            for p in root.rglob("*")}


def code(expected, callable_, *args):
    with pytest.raises(SkillsError) as caught:
        callable_(*args)
    assert str(caught.value) == caught.value.code == expected
    assert expected in SkillsStore.error_codes
    assert BODY not in repr(caught.value)


def create(store, slug="alpha", text=TEXT):
    plan = store.plan_create(slug, text, SESSION)
    receipt = store.confirm(plan["plan_id"], SESSION, True)
    assert receipt["status"] == "saved", receipt
    assert receipt["verified"] is True
    return receipt


def rollback(store, receipt):
    plan = store.plan_rollback(receipt["save_id"], SESSION)
    result = store.confirm_rollback(plan["plan_id"], SESSION, True)
    assert result["status"] == "rolled_back", result
    assert result["verified"] is True
    return result


def test_stat_consistency_after_write(roots):
    root, _, _ = roots
    path = put(root / "skills")
    assert mod._read(path).data == TEXT.encode()
    # Rapid closed writes reproduce Windows path/handle ctime API differences.
    for i in range(100):
        data = (EDIT + str(i)).encode()
        path.write_bytes(data)
        assert mod._read(path).data == data


def test_plans_write_nothing_and_absent_directory_creation(roots):
    root, builtin, imports = roots
    store = SkillsStore(root, builtin)
    before = snapshot(root)
    plan = store.plan_create("alpha", TEXT, SESSION)
    assert snapshot(root) == before == {}
    assert plan["wrote_nothing"] and plan["confirmation_required"]
    assert BODY not in repr(store) + repr(plan)
    assert not (root / "skills").exists()
    receipt = store.confirm(plan["plan_id"], SESSION, True)
    assert receipt["status"] == "saved"
    assert (root / "skills/alpha/SKILL.md").read_text(encoding="utf-8") == TEXT
    assert store.load("user:alpha")["body"] == BODY
    assert store.catalog() == [{"id": "user:alpha", "name": "Alpha", "description": "Safe metadata", "source": "user", "enabled": True}]
    stage = root / (".office-skills-" + receipt["save_id"])
    assert (stage / "backup").read_bytes() == b""
    audit = json.loads((stage / "audit-verified.json").read_bytes())
    assert audit["status"] == "verified" and audit["operation"] == "create"
    assert audit["target"] == str(Path("skills/alpha/SKILL.md"))
    assert BODY not in json.dumps(audit)
    assert receipt["applied_sha256"] == hashlib.sha256(TEXT.encode()).hexdigest()
    assert receipt["audit_sha256"] == hashlib.sha256((stage / "audit-verified.json").read_bytes()).hexdigest()
    result = rollback(store, receipt)
    assert result["applied_sha256"] is None
    assert not (root / "skills/alpha/SKILL.md").exists()
    assert (root / "skills/alpha").is_dir() and stage.is_dir()
    assert store.catalog() == []


def test_edit_delete_and_rollback_retain_resources(roots):
    root, _, _ = roots
    target = put(root / "skills")
    resources = target.parent / "resources"
    resources.mkdir()
    (resources / "manual.txt").write_text("original resource")
    store = SkillsStore(root)
    before = snapshot(root)
    plan = store.plan_edit("user:alpha", EDIT, SESSION)
    assert snapshot(root) == before
    saved = store.confirm(plan["plan_id"], SESSION, True)
    assert target.read_text(encoding="utf-8") == EDIT
    rollback(store, saved)
    assert target.read_text(encoding="utf-8") == TEXT
    plan = store.plan_delete("user:alpha", SESSION)
    assert plan["resources_deleted"] is False and plan["directory_deleted"] is False
    saved = store.confirm(plan["plan_id"], SESSION, True)
    assert not target.exists() and store.catalog() == []
    assert (resources / "manual.txt").read_text() == "original resource"
    rollback(store, saved)
    assert target.read_text(encoding="utf-8") == TEXT
    assert (resources / "manual.txt").read_text() == "original resource"


@pytest.mark.parametrize("approval,expected", [(1, "confirmation_required"), ("true", "confirmation_required"), (None, "confirmation_required"), (False, "skills_denied")])
def test_literal_true_and_denial_one_use(roots, approval, expected):
    root, _, _ = roots
    store = SkillsStore(root)
    plan = store.plan_create("alpha", TEXT, SESSION)
    result = store.confirm(plan["plan_id"], SESSION, approval)
    assert result["code"] == expected and result["wrote_nothing"] is True
    assert snapshot(root) == {}
    assert store.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"


def test_foreign_session_and_wrong_api_cannot_consume_owner_plan(roots):
    root, _, _ = roots
    store = SkillsStore(root)
    plan = store.plan_create("alpha", TEXT, SESSION)
    assert store.confirm(plan["plan_id"], "foreign", True)["code"] == "skills_session_invalid"
    assert store.confirm(plan["plan_id"], "", True)["code"] == "skills_session_invalid"
    assert store.confirm_rollback(plan["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"
    receipt = store.confirm(plan["plan_id"], SESSION, True)
    assert receipt["status"] == "saved"
    assert store.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"
    code("skills_plan_invalid", store.plan_rollback, receipt["save_id"], "foreign")
    rp = store.plan_rollback(receipt["save_id"], SESSION)
    assert store.confirm_rollback(rp["plan_id"], "foreign", True)["code"] == "skills_session_invalid"
    assert store.confirm(rp["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"
    assert store.confirm_rollback(rp["plan_id"], SESSION, True)["status"] == "rolled_back"
    assert store.confirm_rollback(rp["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"


def test_ttl_capacity_collision_and_invalidate(roots, monkeypatch):
    root, _, _ = roots
    now = [10.0]
    monkeypatch.setattr(mod.time, "monotonic", lambda: now[0])
    store = SkillsStore(root)
    plan = store.plan_create("alpha", TEXT, SESSION)
    now[0] += mod.TTL_SECONDS
    assert store.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_plan_expired"
    assert snapshot(root) == {}
    monkeypatch.setattr(mod, "MAX_PENDING", 2)
    first = store.plan_create("alpha", TEXT, SESSION)
    store.plan_create("beta", TEXT, SESSION)
    code("skills_capacity", store.plan_create, "gamma", TEXT, SESSION)
    store.invalidate()
    assert store.confirm(first["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"
    monkeypatch.setattr(mod.secrets, "token_hex", lambda n: plan["plan_id"])
    code("skills_id_collision", store.plan_create, "gamma", TEXT, SESSION)
    monkeypatch.setattr(mod, "MAX_IDS", len(store._issued))
    code("skills_capacity", store.plan_create, "gamma", TEXT, SESSION)


def test_disabled_read_enabled_load_and_state_rollback(roots):
    root, builtin, _ = roots
    user = put(root / "skills")
    built = put(builtin, "reference", TEXT.replace("Alpha", "Reference"))
    originals = (user.read_bytes(), built.read_bytes())
    store = SkillsStore(root, builtin)
    for identifier in ("user:alpha", "builtin:reference"):
        before = snapshot(root)
        plan = store.plan_enabled(identifier, False, SESSION)
        assert snapshot(root) == before
        saved = store.confirm(plan["plan_id"], SESSION, True)
        assert saved["status"] == "saved"
        assert store.read(identifier)["body"] == BODY and not store.read(identifier)["enabled"]
        code("skills_disabled", store.load, identifier)
        assert json.loads((root / ".office-skills-state.json").read_bytes())[identifier] is False
        rollback(store, saved)
        assert store.load(identifier)["enabled"]
        assert not (root / ".office-skills-state.json").exists()
    assert (user.read_bytes(), built.read_bytes()) == originals
    plan = store.plan_enabled("user:alpha", False, SESSION)
    assert store.confirm(plan["plan_id"], SESSION, True)["status"] == "saved"
    plan = store.plan_enabled("user:alpha", True, SESSION)
    assert store.confirm(plan["plan_id"], SESSION, True)["status"] == "saved"
    assert store.load("Alpha")["enabled"]
    code("skills_enabled_invalid", store.plan_enabled, "user:alpha", 1, SESSION)


def test_builtin_immutable_ambiguity_and_untrusted_resource_listing(roots):
    root, builtin, _ = roots
    put(root / "skills")
    target = put(builtin)
    scripts = target.parent / "scripts"
    scripts.mkdir()
    (scripts / "never.py").write_text("raise AssertionError('must not execute')")
    store = SkillsStore(root, builtin)
    before = snapshot(builtin)
    code("skills_builtin_immutable", store.plan_edit, "builtin:alpha", EDIT, SESSION)
    code("skills_builtin_immutable", store.plan_delete, "builtin:alpha", SESSION)
    code("skills_name_ambiguous", store.read, "Alpha")
    assert store.read("builtin:alpha")["untrusted_instructions"] is True
    assert store.resources("builtin:alpha") == [{"kind": "scripts", "name": "never.py", "size_bytes": (scripts / "never.py").stat().st_size, "executable": False}]
    assert snapshot(builtin) == before


def test_import_only_skill_and_source_stability(roots):
    root, _, imports = roots
    source = imports / "SKILL.md"
    source.write_text(TEXT, encoding="utf-8")
    (imports / "scripts").mkdir()
    (imports / "scripts/run.py").write_text("raise AssertionError('never')")
    (imports / "resources").mkdir()
    (imports / "resources/secret").write_text("not imported")
    store = SkillsStore(root)
    before = snapshot(root)
    plan = store.plan_import("alpha", imports, SESSION)
    assert snapshot(root) == before
    assert plan["copied_files"] == ["SKILL.md"] and not plan["scripts_copied"]
    assert store.confirm(plan["plan_id"], SESSION, True)["status"] == "saved"
    assert list((root / "skills/alpha").iterdir()) == [root / "skills/alpha/SKILL.md"]
    plan = store.plan_import("beta", imports, SESSION)
    source.write_text(EDIT, encoding="utf-8")
    before = snapshot(root)
    result = store.confirm(plan["plan_id"], SESSION, True)
    assert result["code"] == "skills_source_changed" and result["wrote_nothing"]
    assert snapshot(root) == before and store.available


@pytest.mark.parametrize("frontmatter", [
    "name: Alpha\ndescription: Okay\nextra: bad", "name: Alpha\nname: Again\ndescription: Okay",
    "name: Alpha", "name: true\ndescription: Okay", "name: 123\ndescription: Okay",
    "name: [Alpha]\ndescription: Okay", "name: &alias Alpha\ndescription: Okay",
    "name: !tag Alpha\ndescription: Okay", "name: *alias\ndescription: Okay",
    "name: Alpha # comment\ndescription: Okay", "name: Alpha\ndescription: |\n  multiline",
    "name: Alpha\ndescription: {nested: bad}", "name: \"\\n\"\ndescription: Okay",
    " name: Alpha\ndescription: Okay", "name: Alpha\ndescription: ''",
    "name: .nan\ndescription: Okay", "name: +.inf\ndescription: Okay",
    "name: y\ndescription: Okay",
])
def test_restricted_yaml_rejects_unsupported_forms_without_writes(roots, frontmatter):
    root, _, _ = roots
    store = SkillsStore(root)
    code("skills_frontmatter_invalid", store.plan_create, "alpha", "---\n" + frontmatter + "\n---\n" + BODY, SESSION)
    assert snapshot(root) == {}


@pytest.mark.parametrize("name,description", [("'Alpha''s'", "'Plain: description # as quoted text'"), ('"Alpha"', '"Unicode \\u4e2d description"')])
def test_restricted_yaml_quoted_strings(roots, name, description):
    root, _, _ = roots
    store = SkillsStore(root)
    text = f"---\r\nname: {name}\r\ndescription: {description}\r\n---\r\n{BODY}"
    create(store, text=text)
    assert store.read("user:alpha")["body"] == BODY


@pytest.mark.parametrize("slug", ["../escape", "Alpha", "a/b", "con", "nul", "a.", "a ", "", "a" * 65])
def test_invalid_slugs(roots, slug):
    root, _, _ = roots
    store = SkillsStore(root)
    code("skills_slug_invalid", store.plan_create, slug, TEXT, SESSION)
    assert snapshot(root) == {}


@pytest.mark.parametrize("state", ['{"user:alpha": 1}', '{"user:alpha": false, "user:alpha": true}', '{"other:alpha": true}', '[]', '{broken'])
def test_invalid_state_is_fail_closed(roots, state):
    root, _, _ = roots
    put(root / "skills")
    (root / ".office-skills-state.json").write_text(state)
    store = SkillsStore(root)
    code("skills_state_invalid", store.catalog)
    code("skills_state_invalid", store.read, "user:alpha")


def test_stale_target_and_rollback_artifacts_no_write(roots):
    root, _, _ = roots
    path = put(root / "skills")
    store = SkillsStore(root)
    plan = store.plan_edit("user:alpha", EDIT, SESSION)
    path.write_text(TEXT + "externally modified", encoding="utf-8")
    before = snapshot(root)
    assert store.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_source_changed"
    assert snapshot(root) == before and store.available
    path.write_text(TEXT, encoding="utf-8")
    plan = store.plan_edit("user:alpha", EDIT, SESSION)
    receipt = store.confirm(plan["plan_id"], SESSION, True)
    rp = store.plan_rollback(receipt["save_id"], SESSION)
    stage = root / (".office-skills-" + receipt["save_id"])
    (stage / "backup").write_text("corrupt")
    before = snapshot(root)
    assert store.confirm_rollback(rp["plan_id"], SESSION, True)["code"] == "skills_source_changed"
    assert snapshot(root) == before and path.read_text(encoding="utf-8") == EDIT
    code("skills_source_changed", store.plan_rollback, receipt["save_id"], SESSION)


def test_root_and_parent_replacement_are_stale(roots):
    root, _, _ = roots
    path = put(root / "skills")
    store = SkillsStore(root)
    plan = store.plan_edit("user:alpha", EDIT, SESSION)
    path.parent.rename(path.parent.with_name("old"))
    put(root / "skills")
    result = store.confirm(plan["plan_id"], SESSION, True)
    assert result["code"] == "skills_source_changed" and result["wrote_nothing"]
    _rename_fixture(root, root.with_name("retired"))
    root.mkdir()
    code("skills_path_changed", store.catalog)


def test_hardlink_rejected_and_reparse_guard_without_platform_privileges(roots, monkeypatch):
    root, _, _ = roots
    path = put(root / "skills")
    os.link(path, root / "other-link")
    store = SkillsStore(root)
    code("skills_file_invalid", store.catalog)
    (root / "other-link").unlink()
    original = Path.lstat
    def reparse(self, *args, **kwargs):
        st = original(self, *args, **kwargs)
        if self == path.parent:
            return SimpleNamespace(st_mode=st.st_mode, st_dev=st.st_dev, st_ino=st.st_ino, st_file_attributes=0x400)
        return st
    monkeypatch.setattr(Path, "lstat", reparse)
    code("skills_path_changed", store.read, "user:alpha")


def test_file_limits_catalog_capacity_and_resources_not_recursive(roots, monkeypatch):
    root, _, _ = roots
    path = put(root / "skills")
    store = SkillsStore(root)
    path.write_bytes(b"x" * (mod.MAX_FILE_BYTES + 1))
    code("skills_file_invalid", store.catalog)
    path.write_text(TEXT, encoding="utf-8")
    put(root / "skills", "beta", TEXT.replace("Alpha", "Beta"))
    monkeypatch.setattr(mod, "MAX_SKILLS", 1)
    code("skills_capacity", store.catalog)
    monkeypatch.setattr(mod, "MAX_SKILLS", 64)
    scripts = path.parent / "scripts"
    (scripts / "nested").mkdir(parents=True)
    (scripts / "nested/never.py").write_text("never execute")
    code("skills_file_invalid", store.resources, "user:alpha")


def test_errors_and_constructor_redact_arbitrary_os_text(roots, monkeypatch):
    root, _, _ = roots
    assert SkillsStore.error_type is SkillsError
    assert len(SkillsStore.error_codes) == 25
    assert str(SkillsError(BODY)) == "skills_invalid"
    code("skills_roots_overlap", SkillsStore, root, root)
    def fail(*args, **kwargs):
        raise PermissionError(BODY + str(root))
    monkeypatch.setattr(Path, "lstat", fail)
    code("skills_io_failed", SkillsStore, root)


def test_public_read_and_import_redact_bad_data(roots):
    root, _, imports = roots
    (imports / "SKILL.md").write_bytes(b"\xff" + BODY.encode())
    store = SkillsStore(root)
    code("skills_frontmatter_invalid", store.plan_import, "alpha", imports, SESSION)
    code("skills_id_invalid", store.read, None)
    code("skills_not_found", store.read, "missing")
    assert snapshot(root) == {}


def test_uncertainty_sticky_and_restart_does_not_restore_grants(roots, monkeypatch):
    root, _, _ = roots
    path = put(root / "skills")
    store = SkillsStore(root)
    plan = store.plan_edit("user:alpha", EDIT, SESSION)
    original = store._audit
    def fail_verified(directory, planned, status):
        if status == "verified":
            raise OSError(BODY)
        return original(directory, planned, status)
    monkeypatch.setattr(store, "_audit", fail_verified)
    result = store.confirm(plan["plan_id"], SESSION, True)
    assert result == {"status": "uncertain", "code": "skills_uncertain", "applied": None, "wrote_nothing": False}
    assert path.read_text(encoding="utf-8") == EDIT and not store.available
    stage = root / (".office-skills-" + plan["plan_id"])
    assert (stage / "backup").read_text(encoding="utf-8") == TEXT
    store.invalidate()
    assert not store.available
    code("skills_uncertain", store.plan_edit, "user:alpha", TEXT, SESSION)
    restarted = SkillsStore(root)
    assert restarted.read("user:alpha")["body"] == "Different untrusted body\n"
    code("skills_plan_invalid", restarted.plan_rollback, plan["plan_id"], SESSION)
    assert restarted.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_plan_invalid"


def test_saved_capacity_does_not_block_rollback_and_is_rechecked(roots, monkeypatch):
    root, _, _ = roots
    put(root / "skills", "alpha")
    put(root / "skills", "beta", TEXT.replace("Alpha", "Beta"))
    put(root / "skills", "gamma", TEXT.replace("Alpha", "Gamma"))
    store = SkillsStore(root)
    monkeypatch.setattr(mod, "MAX_PENDING", 2)
    p1 = store.plan_edit("user:alpha", EDIT, SESSION)
    p2 = store.plan_edit("user:beta", EDIT.replace("Alpha", "Beta"), SESSION)
    s1 = store.confirm(p1["plan_id"], SESSION, True)
    p3 = store.plan_edit("user:gamma", EDIT.replace("Alpha", "Gamma"), SESSION)
    s2 = store.confirm(p2["plan_id"], SESSION, True)
    assert s1["status"] == s2["status"] == "saved"
    before = snapshot(root)
    rejected = store.confirm(p3["plan_id"], SESSION, True)
    assert rejected["code"] == "skills_capacity" and rejected["wrote_nothing"]
    assert snapshot(root) == before
    code("skills_capacity", store.plan_edit, "user:gamma", EDIT, SESSION)
    rollback(store, s2)
    rollback(store, s1)
    assert len(store._saved) == 0


def test_import_directory_replaced_and_enabled_skill_changed(roots):
    root, _, imports = roots
    (imports / "SKILL.md").write_bytes(TEXT.encode())
    store = SkillsStore(root)
    plan = store.plan_import("alpha", imports, SESSION)
    imports.rename(imports.with_name("old-imports"))
    imports.mkdir()
    (imports / "SKILL.md").write_bytes(TEXT.encode())
    before = snapshot(root)
    assert store.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_source_changed"
    assert snapshot(root) == before
    path = put(root / "skills")
    # An externally-created skills root retires the constructor's missing-root snapshot.
    code("skills_path_changed", store.catalog)
    store = SkillsStore(root)
    plan = store.plan_enabled("user:alpha", False, SESSION)
    path.write_bytes(EDIT.encode())
    before = snapshot(root)
    assert store.confirm(plan["plan_id"], SESSION, True)["code"] == "skills_source_changed"
    assert snapshot(root) == before and not (root / ".office-skills-state.json").exists()


def test_stage_collision_retains_existing_contents_and_fuses_conservatively(roots):
    root, _, _ = roots
    store = SkillsStore(root)
    plan = store.plan_create("alpha", TEXT, SESSION)
    stage = root / (".office-skills-" + plan["plan_id"])
    stage.mkdir()
    (stage / "backup").write_bytes(b"existing")
    before = snapshot(root)
    result = store.confirm(plan["plan_id"], SESSION, True)
    assert result["status"] == "uncertain" and result["applied"] is None
    assert snapshot(root) == before and not store.available


def test_read_metadata_changed_between_catalog_and_open_is_rejected(roots, monkeypatch):
    root, _, _ = roots
    path = put(root / "skills")
    store = SkillsStore(root)
    find = store._find
    def mutate(identifier):
        item = find(identifier)
        path.write_bytes(TEXT.replace("Alpha", "Changed").encode())
        return item
    monkeypatch.setattr(store, "_find", mutate)
    code("skills_source_changed", store.read, "user:alpha")


def test_rollback_stage_backup_tamper_is_uncertain_not_false_success(roots, monkeypatch):
    root, _, _ = roots
    store = SkillsStore(root)
    receipt = create(store)
    rp = store.plan_rollback(receipt["save_id"], SESSION)
    original = store._create
    def tamper(path, data):
        result = original(path, data)
        if path.name == "backup":
            path.write_bytes(b"tampered")
        return result
    monkeypatch.setattr(store, "_create", tamper)
    result = store.confirm_rollback(rp["plan_id"], SESSION, True)
    assert result["status"] == "uncertain" and result["restored"] is None
    assert not store.available and (root / "skills/alpha/SKILL.md").exists()
