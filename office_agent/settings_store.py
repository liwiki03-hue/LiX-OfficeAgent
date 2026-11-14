"""Host-only .env settings; no settings.json import, no model-tool exposure.
Caller binds authentication/account epoch and serializes calls under its API lock.
Snapshots are not a change journal. Final check/replace is not an OS transaction
against other same-permission writers. Backups are private; audit is metadata only.
"""
from __future__ import annotations
import hashlib
import io
import json
import math
import os
from pathlib import Path
import re
import secrets
import stat
import time
from dataclasses import dataclass, field
from urllib.parse import urlsplit
from dotenv import dotenv_values
from .settings_definition import PROVIDER_DEFAULTS, LIMITS, MODEL_PATTERN

MAX_ENV_BYTES = 65536
MAX_PENDING = 8
MAX_IDS = 256
TTL_SECONDS = 300
PROVIDERS = ("deepseek", "dashscope", "ollama")
EMBEDDING_NUMBERS = {
    "embedding_dimension": ("OFFICE_EMBEDDING_DIMENSION", 4096, None),
    "embedding_timeout": ("OFFICE_EMBEDDING_TIMEOUT", 20, 20),
    "embedding_budget": ("OFFICE_EMBEDDING_BUDGET", 60, 30),
}
FIELDS = frozenset({"provider", "api_key", "base_url", "model", "context_window", "embedding_channel", "embedding_model", *EMBEDDING_NUMBERS})
MANAGED = frozenset({"MODEL_PROVIDER", "OFFICE_CONTEXT_WINDOW", "OFFICE_EMBEDDING_CHANNEL", "OFFICE_EMBEDDING_MODEL", *(spec[0] for spec in EMBEDDING_NUMBERS.values()), *(f"{p.upper()}_{s}" for p in PROVIDERS for s in ("MODEL", "BASE_URL")), "DEEPSEEK_API_KEY", "DASHSCOPE_API_KEY"})
ASSIGNMENT = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=")

MANAGED |= frozenset({'OFFICE_RERANKER_MODEL', 'OFFICE_RERANKER_TIMEOUT'})
MANAGED |= frozenset({'OPENALEX_API_KEY', 'OFFICE_LITERATURE_KEY_REQUESTS', 'OFFICE_LITERATURE_TIMEOUT', 'OFFICE_LITERATURE_REQUEST_LIMIT'})
MANAGED |= frozenset({'OFFICE_MODEL_CONNECTIONS', 'OFFICE_CHAT_HISTORY'})
FIELDS |= frozenset({'chat_history'})
FIELDS |= frozenset({'chat_reply_tokens', 'image_reuse_metadata'})
MANAGED |= frozenset({'OFFICE_CHAT_REPLY_TOKENS', 'OFFICE_IMAGE_REUSE_METADATA'})
FIELDS |= frozenset({'network_mode', 'network_proxy'})
MANAGED |= frozenset({'OFFICE_NETWORK_MODE', 'OFFICE_NETWORK_PROXY'})


class SettingsError(ValueError):
    def __init__(self, code="settings_invalid", field=None):
        self.code = code
        self.field = field
        super().__init__(code)

@dataclass(repr=False)
class _File:
    data: bytes = field(repr=False)
    stamp: tuple

@dataclass(repr=False)
class _Plan:
    session: str
    created: float
    source: _File
    desired: bytes = field(repr=False)
    saved_id: str | None = None

@dataclass(repr=False)
class _Saved:
    session: str
    source: _File
    applied: _File
    backup: _File
    audit: _File
    directory: Path
    identity: tuple

def _identity(st):
    return (st.st_dev, st.st_ino)

def _stamp(st):
    return (*_identity(st), st.st_size, st.st_mtime_ns, st.st_ctime_ns)

def _unsafe(st):
    return stat.S_ISLNK(st.st_mode) or bool(getattr(st, "st_file_attributes", 0) & 0x400)


def _private_values(values):
    private = [v for k, v in values.items() if v and any(marker in k.upper() for marker in ('KEY', 'TOKEN', 'SECRET', 'PASSWORD'))]
    if values.get('OFFICE_NETWORK_PROXY'):
        from urllib.parse import unquote
        parsed = urlsplit(values['OFFICE_NETWORK_PROXY'])
        private.extend([values['OFFICE_NETWORK_PROXY'], *[unquote(value) for value in (parsed.username, parsed.password) if value]])
    if values.get('OFFICE_MODEL_CONNECTIONS'):
        from .model_connections import from_values
        private.extend(row['api_key'] for row in from_values(values)['connections'] if row['api_key'])
    return private

def _text(value, limit, empty=False):
    if type(value) is not str or len(value) > limit or (not empty and not value) or any(ord(c) < 32 or ord(c) == 127 or 0xD800 <= ord(c) <= 0xDFFF for c in value):
        raise SettingsError()
    return value

def _validate(changes):
    if type(changes) is not dict or not changes or not changes.keys() <= FIELDS:
        raise SettingsError()
    for key, value in changes.items():
        if key == 'network_mode':
            if type(value) is not str or value not in {'auto', 'direct', 'manual'}:
                raise SettingsError('settings_invalid', field=key)
        elif key == 'network_proxy':
            _text(value, 2048, True)
            if value:
                from .network import NetworkOptions
                try:
                    NetworkOptions('manual', value)
                except ValueError:
                    raise SettingsError('settings_invalid', field=key) from None
        elif key in ('chat_history', 'image_reuse_metadata'):
            if type(value) is not bool:
                raise SettingsError()
        elif key == 'chat_reply_tokens':
            if type(value) is not int or not 64 <= value <= 16384:
                raise SettingsError('settings_invalid', field=key)
        elif key == "provider":
            if type(value) is not str or value not in PROVIDERS:
                raise SettingsError()
        elif key == "context_window":
            if type(value) is not int or not LIMITS[key][0] <= value <= LIMITS[key][1]:
                raise SettingsError()
        elif key in EMBEDDING_NUMBERS:
            numeric_types = (int,) if key == "embedding_dimension" else (int, float)
            if type(value) not in numeric_types or not LIMITS[key][0] <= value <= LIMITS[key][1] or not math.isfinite(value):
                raise SettingsError()
        elif key == "embedding_channel":
            if type(value) is not str or value not in ("api", "local"):
                raise SettingsError()
        elif key == "api_key":
            _text(value, 4096, True)
            if any(c.isspace() or c in "'\"\\$#=" for c in value):
                raise SettingsError()
        elif key == "base_url":
            _text(value, 2048)
            try:
                p = urlsplit(value)
                if p.scheme not in ("http", "https") or not p.hostname or p.username is not None or p.password is not None or p.query or p.fragment or "\\" in value or any(c.isspace() for c in value) or p.port == 0:
                    raise SettingsError()
            except Exception:
                raise SettingsError() from None
        else:
            _text(value, 200)
            local_path = key == 'embedding_model' and changes.get('embedding_channel') != 'api' and Path(value).is_absolute() and not value.startswith(('\\\\', '//')) and not any(c in value for c in "'\"$#=\\")
            if not re.fullmatch(MODEL_PATTERN, value) and not local_path:
                raise SettingsError()
    return dict(changes)

class SettingsStore:
    """One-shot approvals, bounded in memory, never recovered from historical audit."""
    def __init__(self, root):
        self.root = Path(root).absolute()
        self._pending = {}
        self._saved = {}
        self._issued = set()
        self._owned_dirs = {}
        self._restart_required = False
        self._fused = False
        self._snapshots = {}
        self._connection_snapshots = {}
        self._root_identity = self._root_check(True)

    def __repr__(self):
        return "<SettingsStore host-only>"

    def _root_check(self, initial=False):
        try:
            for path in (self.root, *self.root.parents):
                st = path.lstat()
                if _unsafe(st) or not stat.S_ISDIR(st.st_mode):
                    raise SettingsError("settings_path_changed")
            identity = _identity(self.root.lstat())
            if not initial and identity != self._root_identity:
                raise SettingsError("settings_path_changed")
            return identity
        except SettingsError:
            raise
        except Exception:
            raise SettingsError("settings_path_changed") from None

    def _read(self, path, missing=False):
        self._root_check()
        try:
            for parent in path.parents:
                if parent == self.root:
                    break
                st = parent.lstat()
                if (_unsafe(st) or not stat.S_ISDIR(st.st_mode) or
                    (parent in self._owned_dirs and _identity(st) != self._owned_dirs[parent])):
                    raise SettingsError("settings_path_changed")
            try:
                before = path.lstat()
            except FileNotFoundError:
                if missing:
                    return _File(b"", ("missing",))
                raise
            if _unsafe(before) or not stat.S_ISREG(before.st_mode) or before.st_nlink != 1 or before.st_size > MAX_ENV_BYTES:
                raise SettingsError("settings_file_invalid")
            with path.open("rb") as handle:
                opened = os.fstat(handle.fileno())
                if _identity(opened) != _identity(before) or opened.st_nlink != 1 or opened.st_size != before.st_size or opened.st_mtime_ns != before.st_mtime_ns:
                    raise SettingsError("settings_source_changed")
                data = handle.read(MAX_ENV_BYTES + 1)
                end = os.fstat(handle.fileno())
            after = path.lstat()
            if len(data) > MAX_ENV_BYTES or _stamp(after) != _stamp(before) or _unsafe(after) or after.st_nlink != 1 or _stamp(end) != _stamp(opened):
                raise SettingsError("settings_source_changed")
            return _File(data, (*_stamp(after), hashlib.sha256(data).hexdigest()))
        except SettingsError:
            raise
        except Exception:
            raise SettingsError("settings_io_failed") from None

    def _same(self, path, expected, missing=False):
        current = self._read(path, missing)
        if current.stamp != expected.stamp:
            raise SettingsError("settings_source_changed")
        return current

    def _values(self, source):
        try:
            text = source.data.decode("utf-8")
            if "\x00" in text:
                raise SettingsError("settings_file_invalid")
            seen = set()
            for line in text.splitlines():
                match = ASSIGNMENT.match(line)
                if not match:
                    bare = re.fullmatch(r"\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*(?:#.*)?", line)
                    if bare and bare[1] in MANAGED:
                        if bare[1] in seen:
                            raise SettingsError("settings_file_invalid")
                        seen.add(bare[1])
                if match and match[1] in MANAGED:
                    if match[1] in seen:
                        raise SettingsError("settings_file_invalid")
                    seen.add(match[1])
                    rhs = line[match.end():].strip()
                    if rhs.startswith(('"', "'")):
                        pattern = r'^"(?:[^"\\]|\\.)*"(?:\s+#.*)?$' if rhs[0] == '"' else r"^'(?:[^'\\]|\\.)*'(?:\s+#.*)?$"
                        if not re.fullmatch(pattern, rhs):
                            raise SettingsError("settings_file_invalid")
            return dict(dotenv_values(stream=io.StringIO(text), interpolate=False))
        except SettingsError:
            raise
        except Exception:
            raise SettingsError("settings_file_invalid") from None

    def _public(self, values, extra_private=()):
        provider = values.get("MODEL_PROVIDER") or "deepseek"
        if provider not in PROVIDERS:
            raise SettingsError("settings_file_invalid")
        prefix = provider.upper()
        result = {"provider": provider, "model": values.get(prefix + "_MODEL") or "", "base_url": values.get(prefix + "_BASE_URL") or "", "context_window": None, "embedding_channel": values.get("OFFICE_EMBEDDING_CHANNEL") or "local", "embedding_model": values.get("OFFICE_EMBEDDING_MODEL") or ""}
        try:
            if 'OFFICE_CHAT_REPLY_TOKENS' in values:
                raw_reply = values['OFFICE_CHAT_REPLY_TOKENS']
                if not re.fullmatch(r'[0-9]{1,5}', raw_reply):
                    raise SettingsError()
                result['chat_reply_tokens'] = int(raw_reply)
            if 'OFFICE_IMAGE_REUSE_METADATA' in values:
                if values['OFFICE_IMAGE_REUSE_METADATA'].lower() not in ('true', 'false'):
                    raise SettingsError()
                result['image_reuse_metadata'] = values['OFFICE_IMAGE_REUSE_METADATA'].lower() == 'true'
            if 'OFFICE_CHAT_HISTORY' in values:
                if values['OFFICE_CHAT_HISTORY'].lower() not in ('true', 'false'):
                    raise SettingsError()
                result['chat_history'] = values['OFFICE_CHAT_HISTORY'].lower() == 'true'
            raw = values.get("OFFICE_CONTEXT_WINDOW")
            if raw:
                if not re.fullmatch(r"[0-9]{1,7}", raw):
                    raise SettingsError()
                result["context_window"] = int(raw)
            for name, (env, upper, default) in EMBEDDING_NUMBERS.items():
                if env not in values:
                    result[name] = default
                    continue
                raw = values[env]
                pattern = r"[0-9]{1,4}" if name == "embedding_dimension" else r"[0-9]{1,2}(?:\.[0-9]{1,17})?"
                if type(raw) is not str or not re.fullmatch(pattern, raw):
                    raise SettingsError()
                result[name] = int(raw) if name == "embedding_dimension" else float(raw)
            _validate({k: v for k, v in result.items() if v is not None and v != ""})
        except Exception:
            raise SettingsError("settings_file_invalid") from None
        # A malicious/accidental config can copy a credential into a model or URL.
        # Redact every known secret value before returning any public string field.
        private_values = list(extra_private) + _private_values(values)
        for name, value in result.items():
            if isinstance(value, str) and any(secret in value for secret in private_values):
                result[name] = "[redacted]"
        configured = bool(values.get(prefix + "_API_KEY")) if provider != "ollama" else False
        result.update(api_key="********" if configured else "", api_key_configured=configured)
        result.update(network_mode=values.get('OFFICE_NETWORK_MODE') or 'direct',
                      network_proxy='********' if values.get('OFFICE_NETWORK_PROXY') else '',
                      network_proxy_configured=bool(values.get('OFFICE_NETWORK_PROXY')))
        return result

    def snapshot(self):
        source = self._read(self.root / ".env", True)
        prior_secrets = []
        for saved in self._saved.values():
            prior_secrets.extend(_private_values(self._values(saved.source)))
        values = self._values(source)
        revision = secrets.token_hex(16)
        self._snapshots[revision] = source
        self._snapshots = dict(list(self._snapshots.items())[-8:])
        profiles = {}
        for provider in PROVIDERS:
            public = self._public({**values, 'MODEL_PROVIDER': provider}, prior_secrets)
            profiles[provider] = {k: public[k] for k in ('model', 'base_url', 'api_key_configured')}
            profiles[provider].update({name + '_present': values.get(provider.upper() + '_' + env) is not None for name, env in [('model', 'MODEL'), ('base_url', 'BASE_URL')]})
        import importlib.util
        return {**self._public(values, prior_secrets), "profiles": profiles,
                "revision": revision, "defaults": PROVIDER_DEFAULTS, "limits": LIMITS,
                "model_pattern": MODEL_PATTERN, "runtime": {"available": False},
                "local_embedding_available": importlib.util.find_spec('sentence_transformers') is not None or all(importlib.util.find_spec(name) is not None for name in ('onnx', 'onnxruntime', 'tokenizers', 'numpy')),
                "disk_provider_explicit": bool(values.get('MODEL_PROVIDER')),
                "providers": list(PROVIDERS), "restart_required": self._restart_required, "storage": ".env", "migration": "manual_reentry_only", "available": not self._fused}

    def plan_edit(self, changes, session_id, revision):
        expected = self._snapshots.get(revision) if type(revision) is str else None
        if expected is None:
            raise SettingsError('settings_source_changed')
        self._same(self.root / '.env', expected, True)
        if type(changes) is not dict or not changes or not changes.keys() <= FIELDS:
            raise SettingsError()
        for name, value in changes.items():
            try:
                _validate({name: value})
            except SettingsError:
                raise SettingsError(field=name) from None
        if changes.get('base_url', '').startswith('http:') and changes.get('provider') != 'ollama':
            raise SettingsError(field='base_url')
        return self.plan_update(changes, session_id)

    def _session(self, session):
        if type(session) is not str or not 1 <= len(session) <= 200:
            raise SettingsError("settings_session_invalid")

    def _id(self):
        if self._fused:
            raise SettingsError("settings_uncertain")
        now = time.monotonic()
        self._pending = {k: v for k, v in self._pending.items() if now - v.created < TTL_SECONDS}
        if len(self._pending) >= MAX_PENDING or len(self._issued) >= MAX_IDS:
            raise SettingsError("settings_capacity")
        for _ in range(4):
            value = secrets.token_hex(16)
            if type(value) is str and re.fullmatch(r"[0-9a-f]{32}", value) and value not in self._issued:
                self._issued.add(value)
                return value
        raise SettingsError("settings_id_collision")

    def plan_update(self, changes, session_id):
        if self._fused:
            raise SettingsError('settings_uncertain')
        self._session(session_id)
        changes = _validate(changes)
        if len(self._saved) >= MAX_PENDING:
            raise SettingsError("settings_capacity")
        source = self._read(self.root / ".env", True)
        values = self._values(source)
        current = self._public(values)
        provider = changes.get("provider", values.get("MODEL_PROVIDER") or "deepseek")
        if 'network_mode' in changes or 'network_proxy' in changes:
            from .network import NetworkOptions
            try:
                NetworkOptions(changes.get('network_mode', values.get('OFFICE_NETWORK_MODE') or 'direct'),
                               changes.get('network_proxy', values.get('OFFICE_NETWORK_PROXY') or ''))
            except ValueError:
                raise SettingsError('settings_invalid', field='network_proxy') from None
        if provider == "ollama" and "api_key" in changes:
            raise SettingsError()
        if "base_url" in changes and changes["base_url"].startswith("http:") and provider != "ollama":
            raise SettingsError()
        mapping = {"provider": "MODEL_PROVIDER", "model": provider.upper() + "_MODEL", "base_url": provider.upper() + "_BASE_URL", "api_key": provider.upper() + "_API_KEY", "context_window": "OFFICE_CONTEXT_WINDOW", "embedding_channel": "OFFICE_EMBEDDING_CHANNEL", "embedding_model": "OFFICE_EMBEDDING_MODEL"}
        mapping['chat_history'] = 'OFFICE_CHAT_HISTORY'
        mapping['chat_reply_tokens'] = 'OFFICE_CHAT_REPLY_TOKENS'
        mapping['image_reuse_metadata'] = 'OFFICE_IMAGE_REUSE_METADATA'
        mapping.update(network_mode='OFFICE_NETWORK_MODE', network_proxy='OFFICE_NETWORK_PROXY')
        mapping.update({name: spec[0] for name, spec in EMBEDDING_NUMBERS.items()})
        actual = {k: v for k, v in changes.items() if
                  (current.get(k) != v if k in LIMITS and mapping[k] in values else str(v) != str(values.get(mapping[k], '')))}
        if not actual:
            raise SettingsError('settings_no_changes')
        before = self._public({**values, 'MODEL_PROVIDER': provider})
        remaining = {mapping[k]: str(v) for k, v in actual.items()}
        output = []
        newline = "\r\n" if b"\r\n" in source.data else "\n"
        for line in source.data.decode("utf-8").splitlines(keepends=True):
            match = ASSIGNMENT.match(line)
            if match and match[1] in remaining:
                value = remaining.pop(match[1])
                suffix = re.search(r"\s+#.*$", line.rstrip("\r\n"))
                output.append(match[1] + "=" + value + (suffix[0] if suffix else "") + newline)
            else:
                output.append(line)
        if remaining and output and not output[-1].endswith(("\n", "\r")):
            output.append(newline)
        output.extend(k + "=" + v + newline for k, v in remaining.items())
        desired = "".join(output).encode("utf-8")
        if len(desired) > MAX_ENV_BYTES:
            raise SettingsError("settings_capacity")
        prior_secrets = _private_values(values)
        public = self._public(self._values(_File(desired, ())), prior_secrets)
        plan_id = self._id()
        self._pending[plan_id] = _Plan(session_id, time.monotonic(), source, desired)
        differences = []
        for key in sorted(actual):
            if key in ('api_key', 'network_proxy'):
                differences.append({'field': key, 'before': '已配置' if before['network_proxy_configured' if key == 'network_proxy' else 'api_key_configured'] else '未配置', 'after': '替换' if changes[key] else '清除'})
            else:
                differences.append({'field': key, 'before': (current['provider'] if values.get('MODEL_PROVIDER') else None) if key == 'provider' else before.get(key), 'after': public.get(key)})
        return {"plan_id": plan_id, "settings": public, "changed_fields": sorted(actual), "differences": differences, "confirmation_required": True, "wrote_nothing": True, "expires_in_seconds": TTL_SECONDS, "restart_required": True}

    def _consume(self, plan_id, session, confirmed):
        if type(plan_id) is not str or len(plan_id) != 32:
            raise SettingsError("settings_plan_invalid")
        plan = self._pending.pop(plan_id, None)
        if plan is None:
            raise SettingsError("settings_plan_invalid")
        self._session(session)
        if session != plan.session:
            raise SettingsError("settings_session_invalid")
        if type(confirmed) is not bool:
            raise SettingsError("confirmation_required")
        if confirmed is not True:
            raise SettingsError("settings_denied")
        if self._fused:
            raise SettingsError("settings_uncertain")
        if time.monotonic() - plan.created >= TTL_SECONDS:
            raise SettingsError("settings_plan_expired")
        return plan

    def _create(self, path, data):
        self._root_check()
        parent = path.parent
        st = parent.lstat()
        if (_unsafe(st) or not stat.S_ISDIR(st.st_mode) or
            parent not in self._owned_dirs or _identity(st) != self._owned_dirs[parent]):
            raise SettingsError("settings_path_changed")
        with path.open("xb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        result = self._read(path)
        if result.data != data:
            raise SettingsError("settings_verification_failed")
        return result

    def _audit_data(self, status):
        return json.dumps({"version": 1, "operation": "settings_update", "storage": ".env", "status": status, "restart_required": True}, sort_keys=True).encode("utf-8")

    def _replace(self, desired, expected, directory):
        temp = directory / "replacement.env"
        owned = self._create(temp, desired)
        self._same(self.root / ".env", expected, True)
        self._same(temp, owned)
        self._root_check()
        os.replace(temp, self.root / ".env")
        result = self._read(self.root / ".env")
        if result.data != desired or result.stamp[:2] != owned.stamp[:2]:
            raise SettingsError("settings_verification_failed")
        return result

    def confirm_update(self, plan_id, session_id, confirmed):
        wrote = False
        try:
            plan = self._consume(plan_id, session_id, confirmed)
            if plan.saved_id is not None:
                raise SettingsError("settings_plan_invalid")
            self._same(self.root / ".env", plan.source, True)
            directory = self.root / (".office-settings-" + plan_id)
            wrote = True
            directory.mkdir(mode=0o700)
            self._owned_dirs[directory] = _identity(directory.lstat())
            backup = self._create(directory / "backup.env", plan.source.data)
            audit = self._create(directory / "audit.json", self._audit_data("intent"))
            applied = self._replace(plan.desired, plan.source, directory)
            self._restart_required = True
            self._same(directory / "audit.json", audit)
            finished = self._create(directory / "audit-final.json", self._audit_data("verified"))
            self._saved[plan_id] = _Saved(session_id, plan.source, applied, backup, finished, directory, _identity(directory.lstat()))
            return {"status": "saved", "save_id": plan_id, "applied": True, "wrote_nothing": False, "restart_required": True}
        except Exception as exc:
            if wrote:
                self._fused = True
                self._restart_required = True
            return {"status": "uncertain" if wrote else "rejected", "code": "settings_uncertain" if wrote else (exc.code if isinstance(exc, SettingsError) else "settings_io_failed"), "applied": None if wrote else False, "wrote_nothing": not wrote, "restart_required": self._restart_required}

    def _saved_check(self, saved):
        self._root_check()
        st = saved.directory.lstat()
        if _unsafe(st) or not stat.S_ISDIR(st.st_mode) or _identity(st) != saved.identity:
            raise SettingsError("settings_path_changed")
        self._same(self.root / ".env", saved.applied)
        self._same(saved.directory / "backup.env", saved.backup)
        self._same(saved.directory / "audit-final.json", saved.audit)

    def plan_rollback(self, save_id, session_id):
        self._session(session_id)
        if type(save_id) is not str:
            raise SettingsError("settings_plan_invalid")
        saved = self._saved.get(save_id)
        if saved is None or saved.session != session_id:
            raise SettingsError("settings_plan_invalid")
        try:
            self._saved_check(saved)
        except SettingsError:
            raise
        except Exception:
            raise SettingsError("settings_io_failed") from None
        rollback_id = self._id()
        self._pending[rollback_id] = _Plan(session_id, time.monotonic(), saved.applied, saved.source.data, save_id)
        return {"rollback_id": rollback_id, "confirmation_required": True, "restores_absent_source_as_empty_env": saved.source.stamp == ("missing",), "wrote_nothing": True, "restart_required": True, "expires_in_seconds": TTL_SECONDS}

    def confirm_rollback(self, rollback_id, session_id, confirmed):
        wrote = False
        try:
            plan = self._consume(rollback_id, session_id, confirmed)
            saved = self._saved.get(plan.saved_id)
            if saved is None or saved.session != session_id:
                raise SettingsError("settings_plan_invalid")
            self._saved_check(saved)
            directory = self.root / (".office-settings-" + rollback_id)
            wrote = True
            directory.mkdir(mode=0o700)
            self._owned_dirs[directory] = _identity(directory.lstat())
            self._create(directory / "audit.json", self._audit_data("rollback_intent"))
            self._replace(saved.source.data, saved.applied, directory)
            self._create(directory / "audit-final.json", self._audit_data("rolled_back"))
            self._saved.pop(plan.saved_id, None)
            self._restart_required = True
            return {"status": "rolled_back", "restored": True, "wrote_nothing": False, "restart_required": True}
        except Exception as exc:
            if wrote:
                self._fused = True
                self._restart_required = True
            return {"status": "uncertain" if wrote else "rejected", "code": "settings_uncertain" if wrote else (exc.code if isinstance(exc, SettingsError) else "settings_io_failed"), "restored": None if wrote else False, "wrote_nothing": not wrote, "restart_required": self._restart_required}

    def invalidate(self):
        self._pending.clear()
        self._saved.clear()
        self._snapshots.clear()
        self._connection_snapshots.clear()
