"""Real sqlite-vec tests. Vectors are controlled fixtures, not embedding quality claims."""
from contextlib import closing
import os
from pathlib import Path
import sqlite3
import time

import pytest
import sqlite_vec

import office_agent.vectorstore as core
from office_agent.vectorstore import VectorStore, VectorStoreError


@pytest.fixture
def store(tmp_path):
    return VectorStore(tmp_path, model="fixture-model", dimension=3)


def records(text="The vehicle needs maintenance", vector=None, kind="document"):
    return [{"source": "trusted-label.docx", "location": "paragraph[1]", "text": text,
             "embedding": vector or [1, 0, 0], "kind": kind}]


def replace(store, doc="doc", text="The vehicle needs maintenance", vector=None, kind="document", session="host-session"):
    plan = store.plan_replace(doc, "fp-" + doc, records(text, vector, kind), session)
    return store.confirm_replace(plan["plan_id"], session, True)


def error(code, action):
    with pytest.raises(VectorStoreError) as exc:
        action()
    assert exc.value.code == code
    assert str(exc.value) == code
    assert repr(exc.value) == f"VectorStoreError('{code}')"
    return exc.value


def test_missing_query_and_plans_are_read_only(store):
    assert store.query([1, 0, 0]) == []
    assert store.snapshot()["exists"] is False
    p = store.plan_replace("doc", "fingerprint", records(), "session")
    assert p["confirmation_required"] is True
    assert list(store._root.iterdir()) == []
    error("confirmation_required", lambda: store.confirm_replace(p["plan_id"], "session", 1))
    error("invalid_authority", lambda: store.confirm_replace(p["plan_id"], "session", True))
    assert list(store._root.iterdir()) == []


def test_actual_extension_paraphrase_persistence_kind_and_bounds(store):
    replace(store)
    replace(store, "other", "The ledger shows a loss", [0, 1, 0])
    replace(store, "memory", "private recollection", [1, 0, 0], "memory")
    # 'automobile repair' occurs nowhere; a supplied semantic fixture still matches.
    hits = store.query([0.99, 0.01, 0], keyword="automobile repair", top_k=2)
    assert hits[0]["doc_id"] == "doc"
    assert hits[0]["distance"] < 0.001
    assert all(h["doc_id"] != "memory" for h in hits)
    assert store.query([1, 0, 0], kind="memory")[0]["doc_id"] == "memory"
    assert hits[0]["model"] == "fixture-model"
    assert hits[0]["source_fingerprint"] == "fp-doc"
    assert hits[0]["rank"] == 1
    with closing(sqlite3.connect(store._db)) as conn:
        conn.enable_load_extension(True)
        sqlite_vec.load(conn)
        conn.enable_load_extension(False)
        assert conn.execute("SELECT vec_version()").fetchone()[0].startswith("v0.1.")
        assert conn.execute("PRAGMA journal_mode").fetchone()[0] == "delete"
        assert conn.execute("SELECT count(*) FROM vectors").fetchone()[0] == 3
    reopened = VectorStore(store._root, model=store.model, dimension=3)
    assert reopened.query([1, 0, 0])[0]["text"] == "The vehicle needs maintenance"
    assert not Path(str(store._db) + "-wal").exists()
    error("model_mismatch", lambda: VectorStore(store._root, model="changed", dimension=3))
    error("model_mismatch", lambda: VectorStore(store._root, model=store.model, dimension=2))
    with pytest.raises(AttributeError):
        store.model = "changed"


def test_precise_literal_keywords_fuse_with_rrf(store):
    replace(store, "a", "semantic neighbor", [1, 0, 0])
    replace(store, "b", "contains exact needle", [0, 1, 0])
    replace(store, "c", "another neighbor", [0.9, 0.1, 0])
    assert store.query([1, 0, 0], keyword="exact needle", top_k=3)[0]["doc_id"] == "b"
    assert store.query([1, 0, 0], keyword="%' OR 1=1 --", top_k=3)[0]["doc_id"] == "a"
    error("invalid_input", lambda: store.query([1, 0, 0], top_k=11))
    error("invalid_input", lambda: store.query([1, 0, 0], top_k=True))
    error("invalid_input", lambda: store.query([1, 0, 0], kind="everything"))


def test_full_text_bounded_host_evidence_and_stale_source(store):
    text = "X" * 8000
    replace(store, text=text)
    hit = store.query([1, 0, 0])[0]
    assert hit["text"] == text
    assert len(hit["excerpt"]) == 200
    evidence = store.read_evidence(hit["id"], limit=2000, source_fingerprint="fp-doc")
    assert len(evidence["text"]) == 2000 and evidence["has_more"]
    assert store.read_evidence(hit["id"], offset=7000, source_fingerprint="fp-doc")["has_more"] is False
    error("stale_source", lambda: store.read_evidence(hit["id"], source_fingerprint="different"))
    error("invalid_input", lambda: store.read_evidence(hit["id"], limit=2001, source_fingerprint="fp-doc"))
    assert set(store.snapshot()) == {"exists", "model", "dimension", "revision", "chunk_count", "audit_count", "uncertain"}


@pytest.mark.parametrize("vector", [[0, 0, 0], [1, 2], [float('nan'), 0, 0], [float('inf'), 0, 0], [1e100, 0, 0], [True, 0, 0], [1e-100, 0, 0]])
def test_invalid_vectors_do_not_write(store, vector):
    error("invalid_vector", lambda: store.plan_replace("doc", "fp", records(vector=vector), "s"))
    error("invalid_vector", lambda: store.query(vector))
    assert list(store._root.iterdir()) == []


def test_float32_extreme_norms_are_safe(store):
    replace(store, vector=[3e38, 3e38, 0])
    assert store.query([3e38, 3e38, 0])[0]["distance"] == pytest.approx(0, abs=1e-6)


@pytest.mark.parametrize("change", [lambda r: r.update(source=1), lambda r: r.update(location=""), lambda r: r.update(text="a" * 8001), lambda r: r.update(kind="bad"), lambda r: r.update(path="secret"), lambda r: r.pop("source")])
def test_record_contract(store, change):
    r = records()[0]
    change(r)
    with pytest.raises(VectorStoreError):
        store.plan_replace("doc", "fp", [r], "s")
    assert list(store._root.iterdir()) == []


def test_session_stale_and_invalidate_consume_authority(store):
    p = store.plan_replace("doc", "fp", records(), "s")
    error("session_mismatch", lambda: store.confirm_replace(p["plan_id"], "other", True))
    error("invalid_authority", lambda: store.confirm_replace(p["plan_id"], "s", True))
    p = store.plan_replace("doc", "fp", records(), "s")
    replace(store)
    error("stale_database", lambda: store.confirm_replace(p["plan_id"], "s", True))
    p = store.plan_delete("doc", "s")
    store.invalidate()
    error("invalid_authority", lambda: store.confirm_replace(p["plan_id"], "s", True))
    assert store.snapshot()["chunk_count"] == 1


def test_replace_delete_and_independently_confirmed_rollback(store):
    first = replace(store)
    p = store.plan_rollback(first["save_id"], "host-session")
    assert p["first_absent_restored_empty"] is True
    error("confirmation_required", lambda: store.confirm_rollback(p["rollback_id"], "host-session", False))
    p = store.plan_rollback(first["save_id"], "host-session")
    restored = store.confirm_rollback(p["rollback_id"], "host-session", True)
    assert restored["first_absent_restored_empty"] is True
    assert store.snapshot()["exists"] is True
    assert store.query([1, 0, 0]) == []
    replace(store, text="old")
    saved = replace(store, text="new")
    rollback = store.plan_rollback(saved["save_id"], "host-session")
    store.confirm_rollback(rollback["rollback_id"], "host-session", True)
    assert store.query([1, 0, 0])[0]["text"] == "old"
    p = store.plan_delete("doc", "host-session")
    saved = store.confirm_replace(p["plan_id"], "host-session", True)
    assert store.query([1, 0, 0]) == []
    p = store.plan_rollback(saved["save_id"], "host-session")
    store.confirm_rollback(p["rollback_id"], "host-session", True)
    assert store.query([1, 0, 0])[0]["text"] == "old"
    assert store.snapshot()["revision"] == 7
    assert store.snapshot()["audit_count"] == 7
    error("document_missing", lambda: store.plan_delete("absent", "s"))


def test_rollback_stale_backup_and_current_database(store):
    save = replace(store)
    error("session_mismatch", lambda: store.plan_rollback(save["save_id"], "wrong"))
    backup = store._saves[save["save_id"]]["backup"]
    backup.write_bytes(b"changed")
    error("stale_database", lambda: store.plan_rollback(save["save_id"], "host-session"))
    save = replace(store, text="latest")
    replace(store, "second")
    error("stale_database", lambda: store.plan_rollback(save["save_id"], "host-session"))


def test_expiry_and_authority_capacity(store, monkeypatch):
    p = store.plan_replace("doc", "fp", records(), "s")
    deadline = store._pending[p["plan_id"]]["expiry"]
    monkeypatch.setattr(core.time, "monotonic", lambda: deadline + 1)
    error("expired_authority", lambda: store.confirm_replace(p["plan_id"], "s", True))


time_original = time.monotonic()


def test_pending_and_chunk_capacity(store, monkeypatch):
    for _ in range(16):
        store.plan_replace("doc", "fp", records(), "s")
    error("authority_limit", lambda: store.plan_replace("doc", "fp", records(), "s"))
    store.invalidate()
    error("document_limit", lambda: store.plan_replace("doc", "fp", records() * 501, "s"))
    # Exercise the identical production bound branch with a smaller deterministic cap.
    monkeypatch.setattr(core, "MAX_CHUNKS", 2)
    replace(store, "a")
    replace(store, "b")
    error("capacity_limit", lambda: replace(store, "c"))
    replace(store, "a", text="replacement stays within capacity")
    assert store.snapshot()["chunk_count"] == 2


def test_root_changed_database_links_and_sidecars(tmp_path):
    root = tmp_path / "root"
    root.mkdir()
    store = VectorStore(root, model="fixture-model", dimension=3)
    moved = tmp_path / "moved"
    root.rename(moved)
    root.mkdir()
    error("stale_root", lambda: store.query([1, 0, 0]))
    store = VectorStore(root, model="fixture-model", dimension=3)
    replace(store)
    os.link(store._db, tmp_path / "hardlink")
    error("unsafe_path", lambda: store.snapshot())
    (tmp_path / "hardlink").unlink()
    Path(str(store._db) + "-wal").write_bytes(b"")
    error("unsafe_sidecar", lambda: store.snapshot())
    error("unsafe_path", lambda: VectorStore(str(root) + "/../root", model="m", dimension=3))


def test_linked_backup_refused(store, tmp_path):
    saved = replace(store)
    backup = store._saves[saved["save_id"]]["backup"]
    os.link(backup, tmp_path / "backup-link")
    error("unsafe_path", lambda: store.plan_rollback(saved["save_id"], "host-session"))


def test_fault_after_publication_audit_fuses_and_redacts(store, monkeypatch):
    replace(store, text="private original")
    before = store._db.read_bytes()
    p = store.plan_replace("doc", "fp", records("secret replacement"), "s")
    def failure(*args):
        raise RuntimeError("secret content and credential")
    monkeypatch.setattr(store, "_audit", failure)
    failure = error("uncertain_state", lambda: store.confirm_replace(p["plan_id"], "s", True))
    assert len(failure.save_id) == 32
    assert store._db.read_bytes() != before
    assert store.snapshot()["uncertain"] is True
    reopened = VectorStore(store._root, model=store.model, dimension=3)
    error("uncertain_state", lambda: reopened.query([1, 0, 0]))
    assert not list(store._root.glob("*.stage"))


def test_uncertain_publication_fuse_survives_invalidate(store, monkeypatch):
    replace(store)
    original = store._verify
    def failure(path):
        if path == store._db:
            raise VectorStoreError("verification_failed")
        return original(path)
    monkeypatch.setattr(store, "_verify", failure)
    p = store.plan_replace("doc", "fp", records("new"), "s")
    with pytest.raises(VectorStoreError):
        store.confirm_replace(p["plan_id"], "s", True)
    assert store.snapshot()["uncertain"] is True
    assert store.invalidate()["uncertain"] is True
    error("uncertain_state", lambda: store.query([1, 0, 0]))
    error("uncertain_state", lambda: store.plan_replace("doc", "fp", records(), "s"))


def test_extension_loading_disabled_and_audit_metadata_only(store):
    replace(store, text="sentinel private text")
    with store._connection(store._db) as conn:
        with pytest.raises(sqlite3.OperationalError, match="not authorized"):
            conn.load_extension("not-an-official-package")
        audit = conn.execute("SELECT * FROM audit").fetchall()
        assert audit == [(1, "replace", 1, 1, 0)]
    assert "sentinel" not in str(store.snapshot())


def test_uncertain_write_exposes_safe_receipt_and_independent_rollback(store, monkeypatch):
    replace(store, text="original")
    original_verify = store._verify
    def fail_reopen(path):
        if path == store._db:
            raise VectorStoreError("verification_failed")
        return original_verify(path)
    monkeypatch.setattr(store, "_verify", fail_reopen)
    plan = store.plan_replace("doc", "fp-new", records("changed"), "host-session")
    failure = error("uncertain_state", lambda: store.confirm_replace(plan["plan_id"], "host-session", True))
    assert len(failure.save_id) == 32
    monkeypatch.setattr(store, "_verify", original_verify)
    rollback = store.plan_rollback(failure.save_id, "host-session")
    store.confirm_rollback(rollback["rollback_id"], "host-session", True)
    assert store.snapshot()["uncertain"] is False
    assert store.query([1, 0, 0])[0]["text"] == "original"


def test_rejected_rollback_session_consumes_plan(store):
    saved = replace(store)
    plan = store.plan_rollback(saved["save_id"], "host-session")
    error("session_mismatch", lambda: store.confirm_rollback(plan["rollback_id"], "wrong", True))
    error("invalid_authority", lambda: store.confirm_rollback(plan["rollback_id"], "host-session", True))
    assert store.query([1, 0, 0])


def test_stale_rollback_plan_and_invalidate_saved_authority(store):
    saved = replace(store)
    plan = store.plan_rollback(saved["save_id"], "host-session")
    replace(store, "second")
    error("stale_database", lambda: store.confirm_rollback(plan["rollback_id"], "host-session", True))
    store.invalidate()
    error("invalid_authority", lambda: store.plan_rollback(saved["save_id"], "host-session"))


def test_external_database_identity_change_is_rejected(store):
    replace(store)
    replacement = store._root / "replacement"
    replacement.write_bytes(store._db.read_bytes())
    os.replace(replacement, store._db)
    error("stale_database", lambda: store.query([1, 0, 0]))
    error("stale_database", lambda: store.snapshot())


def test_reparse_and_symlink_attributes_refused(tmp_path, monkeypatch):
    original_lstat = Path.lstat
    class ReparseStat:
        def __init__(self, s):
            self._stat = s
            self.st_file_attributes = 0x400
        def __getattr__(self, name):
            return getattr(self._stat, name)
    def lstat(path):
        s = original_lstat(path)
        return ReparseStat(s) if path == tmp_path else s
    monkeypatch.setattr(Path, "lstat", lstat)
    error("unsafe_path", lambda: VectorStore(tmp_path, model="fixture-model", dimension=3))
    import stat
    class SymlinkStat(ReparseStat):
        def __init__(self, s):
            super().__init__(s)
            self.st_file_attributes = 0
            self.st_mode = stat.S_IFLNK | 0o777
    monkeypatch.setattr(Path, "lstat", lambda p: SymlinkStat(original_lstat(p)) if p == tmp_path else original_lstat(p))
    error("unsafe_path", lambda: VectorStore(tmp_path, model="fixture-model", dimension=3))


def test_owned_stage_replacement_is_not_removed(store, monkeypatch):
    replace(store, text="original")
    original_copy = store._copy
    replacement_paths = []
    def tamper(source, target, fingerprint):
        original_copy(source, target, fingerprint)
        if target.suffix == ".stage":
            intruder = store._root / "intruder"
            intruder.write_bytes(b"external file must survive")
            os.replace(intruder, target)
            replacement_paths.append(target)
    monkeypatch.setattr(store, "_copy", tamper)
    plan = store.plan_replace("doc", "fp", records("changed"), "s")
    error("stale_owned_file", lambda: store.confirm_replace(plan["plan_id"], "s", True))
    assert replacement_paths[0].read_bytes() == b"external file must survive"
    assert store.query([1, 0, 0])[0]["text"] == "original"


def test_atomic_replace_failure_leaves_original(store, monkeypatch):
    replace(store)
    original = store._db.read_bytes()
    def fail(*args):
        raise OSError("redacted private path")
    monkeypatch.setattr(core.os, "replace", fail)
    plan = store.plan_delete("doc", "s")
    error("write_failed", lambda: store.confirm_replace(plan["plan_id"], "s", True))
    assert store._db.read_bytes() == original
    assert store.snapshot()["uncertain"] is False


def test_audit_and_byte_capacity_prevent_publication(store, monkeypatch):
    replace(store)
    before = store._db.read_bytes()
    monkeypatch.setattr(core, "MAX_AUDIT", 1)
    plan = store.plan_delete("doc", "s")
    error("audit_limit", lambda: store.confirm_replace(plan["plan_id"], "s", True))
    assert store._db.read_bytes() == before
    monkeypatch.setattr(core, "MAX_DB_BYTES", len(before) - 1)
    error("database_limit", lambda: store.snapshot())


def test_authorities_expire_and_save_capacity(store, monkeypatch):
    saved = replace(store)
    expiry = store._saves[saved["save_id"]]["expiry"]
    monkeypatch.setattr(core.time, "monotonic", lambda: expiry + 1)
    error("invalid_authority", lambda: store.plan_rollback(saved["save_id"], "host-session"))
    assert not store._saves
    assert not store._owned
    monkeypatch.setattr(core, "MAX_AUTHORITIES", 1)
    replace(store)
    plan = store.plan_replace("doc", "fp", records(), "s")
    error("authority_limit", lambda: store.confirm_replace(plan["plan_id"], "s", True))


def test_initialization_input_and_metadata_failures(tmp_path):
    error("invalid_dimension", lambda: VectorStore(tmp_path, model="m", dimension=4097))
    error("invalid_dimension", lambda: VectorStore(tmp_path, model="m", dimension=True))
    error("invalid_input", lambda: VectorStore(tmp_path, model="", dimension=3))
    error("unsafe_path", lambda: VectorStore("relative", model="m", dimension=3))
    db = tmp_path / ".office-knowledge.sqlite3"
    db.write_bytes(b"not a database: private data")
    error("database_error", lambda: VectorStore(tmp_path, model="m", dimension=3))


def test_rollback_prepublication_failure_can_be_replanned(store, monkeypatch):
    replace(store, text="original")
    saved = replace(store, text="changed")
    copy = store._copy
    def fail(*args):
        raise VectorStoreError("verification_failed")
    monkeypatch.setattr(store, "_copy", fail)
    plan = store.plan_rollback(saved["save_id"], "host-session")
    error("verification_failed", lambda: store.confirm_rollback(plan["rollback_id"], "host-session", True))
    assert store.query([1, 0, 0])[0]["text"] == "changed"
    monkeypatch.setattr(store, "_copy", copy)
    plan = store.plan_rollback(saved["save_id"], "host-session")
    store.confirm_rollback(plan["rollback_id"], "host-session", True)
    assert store.query([1, 0, 0])[0]["text"] == "original"


def test_post_audit_reopen_fault_persists_uncertainty(store, monkeypatch):
    replace(store)
    verify = store._verify
    calls = [0]
    def failure(path):
        if path == store._db:
            calls[0] += 1
            if calls[0] == 2:
                raise VectorStoreError("verification_failed")
        return verify(path)
    monkeypatch.setattr(store, "_verify", failure)
    p = store.plan_replace("doc", "fp-new", records("new"), "s")
    error("uncertain_state", lambda: store.confirm_replace(p["plan_id"], "s", True))
    reopened = VectorStore(store._root, model=store.model, dimension=3)
    assert reopened.snapshot()["uncertain"] is True
    error("uncertain_state", lambda: reopened.query([1, 0, 0]))


def test_real_ten_thousand_chunk_capacity(store):
    assert core.MAX_CHUNKS == 10_000
    assert core.MAX_DOCUMENT_CHUNKS == 500
    for i in range(20):
        plan = store.plan_replace(f"doc-{i}", f"fp-{i}", records() * 500, "s")
        store.confirm_replace(plan["plan_id"], "s", True)
        store.invalidate()  # Retires bounded rollback IDs, never modifies backup files.
    assert store.snapshot()["chunk_count"] == 10_000
    error("capacity_limit", lambda: store.plan_replace("overflow", "fp", records(), "s"))
    assert len(store.query([1, 0, 0], top_k=10)) == 10


def test_maximum_dimension_actual_extension_roundtrip(tmp_path):
    store = VectorStore(tmp_path, model="fixture-model", dimension=4096)
    p = store.plan_replace("doc", "fp", records(vector=[1] + [0] * 4095), "s")
    store.confirm_replace(p["plan_id"], "s", True)
    assert store.query([1] + [0] * 4095)[0]["distance"] == pytest.approx(0)
    assert store.snapshot()["dimension"] == 4096


def test_ambiguous_atomic_replace_failure_sets_fuse_and_can_rollback(store, monkeypatch):
    replace(store, text="original")
    atomic_replace = core.os.replace
    def ambiguous(stage, target):
        atomic_replace(stage, target)
        raise OSError("private unknown publication outcome")
    monkeypatch.setattr(core.os, "replace", ambiguous)
    p = store.plan_replace("doc", "fp", records("changed"), "s")
    failure = error("uncertain_state", lambda: store.confirm_replace(p["plan_id"], "s", True))
    assert store.snapshot()["uncertain"] is True
    monkeypatch.setattr(core.os, "replace", atomic_replace)
    rollback = store.plan_rollback(failure.save_id, "s")
    store.confirm_rollback(rollback["rollback_id"], "s", True)
    assert store.query([1, 0, 0])[0]["text"] == "original"


def test_persistent_wal_header_refused_before_connection_creates_sidecars(store):
    replace(store)
    with closing(sqlite3.connect(store._db)) as conn:
        assert conn.execute("PRAGMA journal_mode=WAL").fetchone()[0] == "wal"
    conn.close()
    assert not Path(str(store._db) + "-wal").exists()
    assert not Path(str(store._db) + "-shm").exists()
    error("unsafe_sidecar", lambda: VectorStore(store._root, model=store.model, dimension=3))
    assert not Path(str(store._db) + "-wal").exists()
    assert not Path(str(store._db) + "-shm").exists()
