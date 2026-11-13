"""Local desktop gate, not OS security, encryption, or per-user data isolation.

Accounts live in application data, never in an authorized document workspace.
Sessions and the shared per-store login throttle are process memory only (a
restart clears them). All accounts share business data. No credentials are logged.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import math
import os
from pathlib import Path
import secrets
import stat
import threading
import time
import unicodedata
from typing import Callable


class _StorageError(Exception):
    """Internal exception; never expose its detail to the UI."""


class _StoreState:
    def __init__(self) -> None:
        self.lock = threading.RLock()
        self.failures = 0
        self.locked_until = 0.0
        self.generation = 0


class LocalAuth:
    """Standard-library local authentication.

    ``root`` is the application-data directory; records are in
    ``root/auth/accounts.json``. It must not be a business/document directory.
    ``clock`` supplies seconds and can be injected for deterministic tests.
    Register does not log in. Reset needs literal ``confirmed=True`` and removes
    only accounts.json, including a corrupt file; it never recursively deletes.

    Symlinks and Windows reparse points (including junctions) in every existing
    ancestor are refused. These checks are not an OS sandbox: an adversary with
    access to the same OS account can alter files/process memory, and this store
    is not an encrypted vault. Locks coordinate instances in this process, not
    separate application processes.
    """

    ITERATIONS = 200_000
    MAX_ITERATIONS = 1_000_000
    MAX_PASSWORD_LENGTH = 16
    # Previously registered accounts remain usable; this does not permit new long passwords.
    MAX_LOGIN_PASSWORD_LENGTH = 1024
    MAX_ACCOUNTS = 1000
    MAX_FILE_SIZE = 1024 * 1024
    LOCK_SECONDS = 300
    LIMITATIONS = (
        "本地门禁不是加密库；同机账户共享工作区数据，不提供用户数据隔离。"
        "登录状态与失败锁定仅在当前进程内有效，重启后清除。"
    )
    _states_lock = threading.Lock()
    _states: dict[str, _StoreState] = {}

    def __init__(self, root: Path | None = None,
                 clock: Callable[[], float] = time.time) -> None:
        if root is None:
            if os.name == "nt":
                base = os.environ.get("LOCALAPPDATA")
                # Missing LOCALAPPDATA is not permission to use the workspace.
                root = (Path(base) if base else Path.home() / "AppData" / "Local") / "OfficeAgent"
            else:
                root = Path.home() / ".local" / "share" / "OfficeAgent"
        # abspath is lexical, unlike resolve(), which would follow links.
        self.root = Path(os.path.abspath(os.fspath(root)))
        self.accounts_path = self.root / "auth" / "accounts.json"
        self._clock = clock
        self._user: str | None = None
        key = os.path.normcase(str(self.accounts_path))
        with self._states_lock:
            self._state = self._states.setdefault(key, _StoreState())
        self._generation = self._state.generation

    @staticmethod
    def _valid_username(value: object) -> bool:
        return (isinstance(value, str) and 4 <= len(value) <= 32
                and not any(c in '/\\:' or unicodedata.category(c).startswith('C')
                            for c in value)
                and value.strip() == value and value not in {".", ".."})

    @classmethod
    def _valid_password(cls, value: object, *, for_login: bool = False) -> bool:
        maximum = cls.MAX_LOGIN_PASSWORD_LENGTH if for_login else cls.MAX_PASSWORD_LENGTH
        if not isinstance(value, str) or len(value) > maximum:
            return False
        # Registration matches HTML maxlength/JS length (UTF-16). Login retains
        # the historical code-point upper bound and accepts every valid new password.
        length = sum(2 if ord(c) > 0xFFFF else 1 for c in value)
        return (8 <= length and (len(value) if for_login else length) <= maximum
                and any(c in "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" for c in value)
                and any(c in "0123456789" for c in value)
                and not any(unicodedata.category(c) == "Cs" for c in value))

    def _result(self, ok: bool, error: str | None = None, **extra: object) -> dict:
        if self._generation != self._state.generation:
            self._user = None
            self._generation = self._state.generation
        messages = {
            "invalid_username": "用户名需为4–32字符，不能包含控制字符或路径字符。",
            "invalid_password": "密码需为8–16字符且包含英文字母和数字。",
            "registration_failed": "无法注册，请检查输入或稍后重试。",
            "invalid_credentials": "用户名或密码不正确。",
            "locked": "尝试次数过多，请稍后再试。",
            "storage_unavailable": "账户数据不可用；请检查本机存储或确认重置本地账户。",
            "confirmation_required": "重置全部本地账户需要独立确认，不会删除业务文件。",
        }
        return {"ok": ok, "authenticated": self._user is not None,
                "user": self._user, "error": error,
                "message": messages.get(error, ""), **extra}

    @staticmethod
    def _check_path(path: Path) -> None:
        """lstat every component without resolving or following a reparse point."""
        for component in reversed((path, *path.parents)):
            try:
                info = component.lstat()
            except FileNotFoundError:
                continue
            if (stat.S_ISLNK(info.st_mode)
                    or getattr(info, "st_file_attributes", 0) & 0x400):
                raise _StorageError()
            if component != path and not stat.S_ISDIR(info.st_mode):
                raise _StorageError()

    @staticmethod
    def _unique_object(pairs: list[tuple[str, object]]) -> dict:
        result = {}
        for key, value in pairs:
            if key in result:
                raise _StorageError()
            result[key] = value
        return result

    def _read_accounts(self) -> list[dict]:
        path = self.accounts_path
        self._check_path(path)
        try:
            before = path.lstat()
        except FileNotFoundError:
            return []
        if not stat.S_ISREG(before.st_mode) or before.st_size > self.MAX_FILE_SIZE:
            raise _StorageError()
        fd = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
                     | getattr(os, "O_BINARY", 0))
        with os.fdopen(fd, "rb") as stream:
            opened = os.fstat(stream.fileno())
            if not stat.S_ISREG(opened.st_mode) or (before.st_dev, before.st_ino) != (opened.st_dev, opened.st_ino):
                raise _StorageError()
            data = stream.read(self.MAX_FILE_SIZE + 1)
        self._check_path(path)
        after = path.lstat()
        if (opened.st_dev, opened.st_ino) != (after.st_dev, after.st_ino) or len(data) > self.MAX_FILE_SIZE:
            raise _StorageError()
        document = json.loads(data.decode("utf-8"), object_pairs_hook=self._unique_object)
        if (not isinstance(document, dict) or set(document) != {"version", "accounts"}
                or type(document["version"]) is not int or document["version"] != 1
                or not isinstance(document["accounts"], list)
                or len(document["accounts"]) > self.MAX_ACCOUNTS):
            raise _StorageError()
        seen = set()
        for record in document["accounts"]:
            if (not isinstance(record, dict)
                    or set(record) != {"username", "salt", "hash", "iterations", "created_at"}
                    or not self._valid_username(record["username"])
                    or record["username"] in seen
                    or type(record["iterations"]) is not int
                    or not self.ITERATIONS <= record["iterations"] <= self.MAX_ITERATIONS
                    or type(record["created_at"]) not in (int, float)
                    or not math.isfinite(record["created_at"]) or record["created_at"] < 0):
                raise _StorageError()
            for field, length in (("salt", 32), ("hash", 64)):
                value = record[field]
                if (not isinstance(value, str) or len(value) != length
                        or any(c not in "0123456789abcdef" for c in value)):
                    raise _StorageError()
            seen.add(record["username"])
        return document["accounts"]

    def _write_accounts(self, records: list[dict]) -> None:
        path = self.accounts_path
        self._check_path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self._check_path(path)
        payload = json.dumps({"version": 1, "accounts": records}, ensure_ascii=False).encode("utf-8")
        if len(payload) > self.MAX_FILE_SIZE:
            raise _StorageError()
        temporary = path.parent / (".accounts-" + secrets.token_hex(16) + ".tmp")
        created = False
        try:
            fd = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_EXCL
                         | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_BINARY", 0), 0o600)
            created = True
            with os.fdopen(fd, "wb") as stream:
                stream.write(payload)
                stream.flush()
                os.fsync(stream.fileno())
            self._check_path(path)
            self._check_path(temporary)
            os.replace(temporary, path)
            created = False
        finally:
            if created:
                # Do not traverse a newly substituted parent during cleanup.
                self._check_path(temporary)
                temporary.unlink(missing_ok=True)

    def register(self, username: str, password: str) -> dict:
        with self._state.lock:
            if not self._valid_username(username):
                return self._result(False, "invalid_username")
            if not self._valid_password(password):
                return self._result(False, "invalid_password")
            try:
                records = self._read_accounts()
                if len(records) >= self.MAX_ACCOUNTS or any(r["username"] == username for r in records):
                    return self._result(False, "registration_failed")
                salt = secrets.token_bytes(16)
                digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, self.ITERATIONS)
                records.append({"username": username, "salt": salt.hex(), "hash": digest.hex(),
                                "iterations": self.ITERATIONS, "created_at": self._clock()})
                self._write_accounts(records)
                return self._result(True)
            except (OSError, ValueError, TypeError, RecursionError, OverflowError, _StorageError):
                return self._result(False, "storage_unavailable")

    def login(self, username: str, password: str) -> dict:
        with self._state.lock:
            # A failed authentication attempt cannot retain an older session.
            self._user = None
            now = self._clock()
            if now < self._state.locked_until:
                return self._result(False, "locked", retry_after=math.ceil(self._state.locked_until - now))
            if self._state.locked_until:
                self._state.failures = 0
                self._state.locked_until = 0.0
            try:
                records = self._read_accounts()
                valid_input = self._valid_username(username) and self._valid_password(password, for_login=True)
                record = next((r for r in records if valid_input and r["username"] == username), None)
                # Unknown accounts use the same PBKDF cost. Oversized inputs are
                # never hashed; they still count toward the store-wide throttle.
                supplied = password.encode("utf-8") if valid_input else b"invalid-input"
                salt = bytes.fromhex(record["salt"]) if record else bytes(16)
                rounds = record["iterations"] if record else self.ITERATIONS
                digest = hashlib.pbkdf2_hmac("sha256", supplied, salt, rounds)
                expected = bytes.fromhex(record["hash"]) if record else bytes(32)
                matches = hmac.compare_digest(digest, expected)
                if record is not None and matches:
                    self._state.failures = 0
                    self._user = record["username"]
                    self._generation = self._state.generation
                    return self._result(True)
                self._state.failures += 1
                if self._state.failures >= 5:
                    self._state.locked_until = now + self.LOCK_SECONDS
                    return self._result(False, "locked", retry_after=self.LOCK_SECONDS)
                return self._result(False, "invalid_credentials")
            except (OSError, ValueError, TypeError, RecursionError, OverflowError, _StorageError):
                return self._result(False, "storage_unavailable")

    def account_key(self) -> str | None:
        """Host-only registration identity for opt-in preferences (never browser-facing).

        A reset followed by reuse of the same username gets a new random salt and
        cannot retrieve the previous registration's orphaned memory file.
        """
        with self._state.lock:
            self._result(True)  # Apply cross-instance account-reset generation fence.
            if self._user is None:
                return None
            try:
                record = next((r for r in self._read_accounts() if r['username'] == self._user), None)
                if record is None:
                    return None
                return hashlib.sha256(('office-memory-v1:' + record['salt'] + ':' +
                                       record['username'] + ':' + str(record['created_at'])).encode('utf-8')).hexdigest()
            except (OSError, ValueError, TypeError, RecursionError, OverflowError, _StorageError):
                return None

    def logout(self) -> dict:
        with self._state.lock:
            self._user = None
            return self._result(True)

    def status(self) -> dict:
        # In particular, this must not mkdir or persist an anonymous session.
        with self._state.lock:
            return self._result(True, limitations=self.LIMITATIONS)

    def reset_accounts(self, confirmed: bool = False) -> dict:
        with self._state.lock:
            if confirmed is not True:
                return self._result(False, "confirmation_required")
            try:
                self._check_path(self.accounts_path)
                try:
                    info = self.accounts_path.lstat()
                except FileNotFoundError:
                    info = None
                if info is not None:
                    if not stat.S_ISREG(info.st_mode):
                        raise _StorageError()
                    self.accounts_path.unlink()
                self._state.failures = 0
                self._state.locked_until = 0.0
                self._state.generation += 1
                self._user = None
                return self._result(True)
            except (OSError, ValueError, _StorageError):
                return self._result(False, "storage_unavailable")


__all__ = ["LocalAuth"]
