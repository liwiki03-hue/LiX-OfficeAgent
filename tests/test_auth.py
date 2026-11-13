"""Small deterministic tests for the local gate; no GUI or Office automation."""
import json
import os
from pathlib import Path
import stat
from types import SimpleNamespace

import pytest

from office_agent.auth import LocalAuth


class Clock:
    def __init__(self):
        self.now = 1000.0

    def __call__(self):
        return self.now


def assert_safe(result, password="Secret123", digest=None):
    text = json.dumps(result, ensure_ascii=False)
    assert password not in text
    if digest:
        assert digest not in text
    assert set(("ok", "authenticated", "user", "error", "message")) <= result.keys()


def test_registration_login_logout_and_shared_store(tmp_path, caplog):
    auth = LocalAuth(tmp_path / "app", clock=Clock())
    assert auth.status()["authenticated"] is False
    assert not auth.root.exists()  # status must not create storage
    assert auth.register("alice", "Secret123")["ok"]
    assert auth.status()["authenticated"] is False  # register is not login
    original = auth.accounts_path.read_bytes()
    duplicate = auth.register("alice", "Secret123")
    assert duplicate["error"] == "registration_failed"
    assert auth.accounts_path.read_bytes() == original
    second = LocalAuth(auth.root, clock=Clock())
    assert second.register("bobby", "OtherPass8")["ok"]
    document = json.loads(auth.accounts_path.read_text(encoding="utf-8"))
    assert [r["username"] for r in document["accounts"]] == ["alice", "bobby"]
    record = document["accounts"][0]
    assert record["iterations"] >= 200_000
    assert len(record["salt"]) == 32 and len(record["hash"]) == 64
    assert b"Secret123" not in auth.accounts_path.read_bytes()
    for result in (duplicate, auth.login("alice", "Secret123"), auth.status(), auth.logout()):
        assert_safe(result, digest=record["hash"])
    assert auth.status()["authenticated"] is False
    assert second.login("bobby", "OtherPass8")["authenticated"]
    assert second.status()["user"] == "bobby"
    assert second.login("bobby", "WrongPass9")["authenticated"] is False
    assert second.status()["authenticated"] is False
    assert "Secret123" not in caplog.text and record["hash"] not in caplog.text
    assert not list(auth.accounts_path.parent.glob("*.tmp"))


def test_global_failure_lock_does_not_depend_on_username(tmp_path):
    clock = Clock()
    auth = LocalAuth(tmp_path, clock=clock)
    other = LocalAuth(tmp_path, clock=clock)
    assert auth.register("alice", "Secret123")["ok"]
    for username in ("alice", "none1", "none2", "none3"):
        assert other.login(username, "WrongPass9")["error"] == "invalid_credentials"
    locked = auth.login("none4", "WrongPass9")
    assert locked["error"] == "locked" and locked["retry_after"] == 300
    assert auth.login("alice", "Secret123")["error"] == "locked"
    clock.now += 299
    assert other.login("none5", "WrongPass9")["retry_after"] == 1
    clock.now += 1
    assert auth.login("alice", "Secret123")["authenticated"]
    assert auth.logout()["authenticated"] is False
    assert other.login("alice", "WrongPass9")["error"] == "invalid_credentials"


def test_validation_is_bounded_and_generic(tmp_path):
    auth = LocalAuth(tmp_path)
    for username in ("abc", "x" * 33, "ab/cd", "ab\\cd", "ab\x00cd", "ab:cd"):
        assert auth.register(username, "Secret123")["error"] == "invalid_username"
    for password in ("short1", "lettersOnly", "12345678", "A1" + "x" * 1023):
        response = auth.register("alice", password)
        assert response["error"] == "invalid_password"
        assert_safe(response, password=password)
    assert not auth.root.joinpath("auth").exists()
    assert auth.login("alice", "A1" + "x" * 2000)["error"] == "invalid_credentials"


@pytest.mark.parametrize("length", [7, 8, 16, 17])
def test_registration_password_length_8_to_16(tmp_path, length):
    auth = LocalAuth(tmp_path)
    password = "A1" + "x" * (length - 2)
    result = auth.register("alice", password)
    assert result["ok"] is (8 <= length <= 16)
    if result["ok"]:
        assert auth.login("alice", password)["authenticated"]
    else:
        assert result["error"] == "invalid_password"
        assert not auth.accounts_path.exists()
    assert_safe(result, password=password)


@pytest.mark.parametrize("supplementary_count,accepted", [(3, True), (7, True), (8, False)])
def test_registration_matches_frontend_utf16_length(tmp_path, supplementary_count, accepted):
    auth = LocalAuth(tmp_path)
    password = "A1" + "\U0001f600" * supplementary_count
    result = auth.register("unicode_user", password)
    assert result["ok"] is accepted
    if accepted:
        assert auth.login("unicode_user", password)["ok"]
    else:
        assert result["error"] == "invalid_password"
        assert not auth.accounts_path.exists()


def test_existing_long_password_login_remains_compatible(tmp_path, monkeypatch):
    auth = LocalAuth(tmp_path)
    password = "Legacy1" + "x" * 20
    # Synthesize a historical account only in the test's temporary store.
    with monkeypatch.context() as old_policy:
        old_policy.setattr(LocalAuth, "MAX_PASSWORD_LENGTH", 1024)
        assert auth.register("alice", password)["ok"]
    assert auth.register("bobby", password)["error"] == "invalid_password"
    assert auth.login("alice", password)["authenticated"]
    assert auth.logout()["authenticated"] is False
    assert auth.login("alice", "A1" + "x" * 1023)["error"] == "invalid_credentials"


def test_corruption_is_closed_reset_is_confirmed_and_business_safe(tmp_path):
    auth = LocalAuth(tmp_path / "app", clock=Clock())
    auth.accounts_path.parent.mkdir(parents=True)
    broken = b'{"version":1,"accounts":[broken]}'
    auth.accounts_path.write_bytes(broken)
    business = auth.root / "business.docx"
    business.write_bytes(b"keep me")
    for response in (auth.register("alice", "Secret123"), auth.login("alice", "Secret123")):
        assert response["error"] == "storage_unavailable"
        assert str(tmp_path) not in json.dumps(response)
        assert_safe(response)
    assert auth.accounts_path.read_bytes() == broken
    assert auth.reset_accounts()["error"] == "confirmation_required"
    assert auth.reset_accounts(confirmed=1)["error"] == "confirmation_required"
    assert auth.accounts_path.read_bytes() == broken
    assert auth.reset_accounts(confirmed=True)["ok"]
    assert not auth.accounts_path.exists() and business.read_bytes() == b"keep me"
    assert auth.register("alice", "Secret123")["ok"]
    other = LocalAuth(auth.root)
    assert other.login("alice", "Secret123")["ok"]
    assert auth.reset_accounts(confirmed=True)["ok"]
    assert other.status()["authenticated"] is False
    assert business.read_bytes() == b"keep me"


@pytest.mark.parametrize("bad", ["iterations", "hash", "salt", "schema", "count", "size"])
def test_untrusted_store_limits(tmp_path, bad):
    auth = LocalAuth(tmp_path, clock=Clock())
    auth.accounts_path.parent.mkdir()
    record = {"username": "alice", "iterations": 200_000, "salt": "00" * 16,
              "hash": "00" * 32, "created_at": 1000}
    document = {"version": 1, "accounts": [record]}
    if bad == "iterations":
        record["iterations"] = 1_000_001
    elif bad == "hash":
        record["hash"] = "not-a-hash"
    elif bad == "salt":
        record["salt"] = "zz" * 16
    elif bad == "schema":
        document["extra"] = "unexpected"
    elif bad == "count":
        document["accounts"] = [record] * (auth.MAX_ACCOUNTS + 1)
    data = (b" " * (auth.MAX_FILE_SIZE + 1) if bad == "size"
            else json.dumps(document).encode())
    auth.accounts_path.write_bytes(data)
    assert auth.login("alice", "Secret123")["error"] == "storage_unavailable"
    assert auth.register("bobby", "OtherPass8")["error"] == "storage_unavailable"
    assert auth.accounts_path.read_bytes() == data


@pytest.mark.parametrize("link_kind", ["symlink", "junction"])
@pytest.mark.parametrize("location", ["ancestor", "file"])
def test_every_ancestor_and_file_link_is_rejected(tmp_path, monkeypatch, link_kind, location):
    """Inject lstat metadata: deterministic even without Windows symlink privilege."""
    auth = LocalAuth(tmp_path / "app", clock=Clock())
    auth.accounts_path.parent.mkdir(parents=True)
    auth.accounts_path.write_bytes(b"outside account content")
    target = tmp_path if location == "ancestor" else auth.accounts_path
    original = Path.lstat

    def linked_lstat(path, *args, **kwargs):
        info = original(path, *args, **kwargs)
        if path != target:
            return info
        return SimpleNamespace(st_mode=stat.S_IFLNK if link_kind == "symlink" else info.st_mode,
                               st_file_attributes=0x400 if link_kind == "junction" else 0)

    monkeypatch.setattr(Path, "lstat", linked_lstat)
    assert auth.status()["ok"]  # does not touch disk
    for result in (auth.register("alice", "Secret123"), auth.login("alice", "Secret123"),
                   auth.reset_accounts(confirmed=True)):
        assert result["error"] == "storage_unavailable"
        assert str(tmp_path) not in json.dumps(result)
    assert auth.accounts_path.read_bytes() == b"outside account content"


@pytest.mark.skipif(os.name == "nt", reason="Windows symlink privilege is not assumed; reparse checks tested with lstat metadata")
def test_real_symlink_ancestor_never_followed(tmp_path):
    outside = tmp_path / "outside"
    outside.mkdir()
    linked = tmp_path / "linked"
    linked.symlink_to(outside, target_is_directory=True)
    auth = LocalAuth(linked / "app")
    assert auth.register("alice", "Secret123")["error"] == "storage_unavailable"
    assert auth.reset_accounts(confirmed=True)["error"] == "storage_unavailable"
    assert list(outside.iterdir()) == []


def test_concurrent_registration_keeps_both_accounts(tmp_path):
    from concurrent.futures import ThreadPoolExecutor

    first, second = LocalAuth(tmp_path), LocalAuth(tmp_path)
    with ThreadPoolExecutor(max_workers=2) as executor:
        one = executor.submit(first.register, "alice", "Secret123")
        two = executor.submit(second.register, "bobby", "OtherPass8")
        assert one.result(timeout=10)["ok"]
        assert two.result(timeout=10)["ok"]
    document = json.loads(first.accounts_path.read_text(encoding="utf-8"))
    assert {record["username"] for record in document["accounts"]} == {"alice", "bobby"}


def test_atomic_write_failure_preserves_existing_accounts(tmp_path, monkeypatch):
    auth = LocalAuth(tmp_path)
    assert auth.register("alice", "Secret123")["ok"]
    original = auth.accounts_path.read_bytes()

    def denied_replace(*args):
        raise OSError("secret storage path and Secret123 must not leak")

    monkeypatch.setattr(os, "replace", denied_replace)
    result = auth.register("bobby", "OtherPass8")
    assert result["error"] == "storage_unavailable"
    assert_safe(result)
    assert "secret storage path" not in json.dumps(result)
    assert auth.accounts_path.read_bytes() == original
    assert list(auth.accounts_path.parent.iterdir()) == [auth.accounts_path]


def test_default_location_is_application_data(monkeypatch, tmp_path):
    if os.name == "nt":
        monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
        expected = tmp_path / "OfficeAgent"
    else:
        monkeypatch.setattr(Path, "home", lambda: tmp_path)
        expected = tmp_path / ".local" / "share" / "OfficeAgent"
    auth = LocalAuth()
    assert auth.accounts_path == expected / "auth" / "accounts.json"
    assert auth.status()["authenticated"] is False
    assert not expected.exists()


def test_malformed_json_duplicate_fields_are_rejected(tmp_path):
    auth = LocalAuth(tmp_path)
    auth.accounts_path.parent.mkdir()
    payload = b'{"version":1,"version":1,"accounts":[]}'
    auth.accounts_path.write_bytes(payload)
    assert auth.register("alice", "Secret123")["error"] == "storage_unavailable"
    assert auth.accounts_path.read_bytes() == payload
