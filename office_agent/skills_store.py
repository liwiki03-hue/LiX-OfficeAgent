"""Standalone, host-only Skills store. Not a model tool or an authentication gate.

Host supplies roots/session epoch and serializes ALL calls under its lock. Literal
True is a host approval input, not proof of a human click. Files are untrusted
instructions, never executable authority. No script execution or recursive scan.
Final check/replace is NOT an OS transaction against same-permission writers.
Backups contain skill text; audit contains metadata only. Uncertain writes fuse
this instance; invalidate does not clear that fuse. Restart never restores grants.
Imports copy SKILL.md only; delete leaves directories/resources intact.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from functools import wraps
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import stat
import time

MAX_SKILLS = 64
MAX_FILE_BYTES = 65536
MAX_ENTRIES = 64
MAX_PENDING = 8
MAX_IDS = 256
TTL_SECONDS = 300
SLUG = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*\Z")
RESERVED = frozenset({"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))})


ERROR_CODES = frozenset({
    "skills_invalid", "skills_slug_invalid", "skills_path_changed",
    "skills_not_found", "skills_file_invalid", "skills_source_changed",
    "skills_io_failed", "skills_frontmatter_invalid", "skills_roots_overlap",
    "skills_state_invalid", "skills_id_invalid", "skills_capacity",
    "skills_name_ambiguous", "skills_disabled", "skills_session_invalid",
    "skills_uncertain", "skills_id_collision", "skills_exists",
    "skills_builtin_immutable", "skills_enabled_invalid", "skills_plan_invalid",
    "confirmation_required", "skills_denied", "skills_plan_expired",
    "skills_verification_failed",
})


class SkillsError(ValueError):
    def __init__(self, code="skills_invalid"):
        self.code = code if type(code) is str and code in ERROR_CODES else "skills_invalid"
        super().__init__(self.code)


def _public(method):
    """Never expose OS/path/content exception text to a host facade."""
    @wraps(method)
    def safe(*args, **kwargs):
        try:
            return method(*args, **kwargs)
        except SkillsError:
            raise
        except Exception:
            raise SkillsError("skills_io_failed") from None
    return safe


def _slug(value):
    if type(value) is not str or not 1 <= len(value) <= 64 or not SLUG.fullmatch(value) or value in RESERVED:
        raise SkillsError("skills_slug_invalid")
    return value


def _unsafe(st):
    return stat.S_ISLNK(st.st_mode) or bool(getattr(st, "st_file_attributes", 0) & 0x400)


def _identity(st):
    return st.st_dev, st.st_ino


def _stamp(st):
    return (*_identity(st), st.st_size, st.st_mtime_ns, st.st_ctime_ns)


def _directory(path, missing=False):
    """Check lexical ancestry before opening: no resolve/follow of reparse paths."""
    observations = []
    for parent in reversed((path, *path.parents)):
        try:
            st = parent.lstat()
        except FileNotFoundError:
            if missing:
                observations.append((str(parent), None))
                continue
            raise SkillsError("skills_path_changed") from None
        if _unsafe(st) or not stat.S_ISDIR(st.st_mode):
            raise SkillsError("skills_path_changed")
        observations.append((str(parent), _identity(st)))
    return tuple(observations)


@dataclass(repr=False)
class _File:
    data: bytes = field(repr=False)
    stamp: tuple
    parents: tuple


def _read(path, missing=False):
    parents = _directory(path.parent, missing)
    try:
        before = path.lstat()
    except FileNotFoundError:
        if missing:
            return _File(b"", ("missing",), parents)
        raise SkillsError("skills_not_found") from None
    if _unsafe(before) or not stat.S_ISREG(before.st_mode) or before.st_nlink != 1 or before.st_size > MAX_FILE_BYTES:
        raise SkillsError("skills_file_invalid")
    try:
        with path.open("rb") as handle:
            opened = os.fstat(handle.fileno())
            # Windows path-stat ctime may mean creation time while handle-stat
            # ctime means change time. Compare identity/size/mtime across APIs;
            # retain full timestamps for handle/handle and path/path comparisons.
            cross_api = slice(None, -1) if os.name == "nt" else slice(None)
            if _stamp(opened)[cross_api] != _stamp(before)[cross_api] or opened.st_nlink != 1:
                raise SkillsError("skills_source_changed")
            data = handle.read(MAX_FILE_BYTES + 1)
            end = os.fstat(handle.fileno())
        after = path.lstat()
        if (len(data) > MAX_FILE_BYTES or _stamp(end) != _stamp(opened) or
                _stamp(after) != _stamp(before) or _unsafe(after) or after.st_nlink != 1 or
                _directory(path.parent) != parents):
            raise SkillsError("skills_source_changed")
        return _File(data, (*_stamp(after), hashlib.sha256(data).hexdigest()), parents)
    except SkillsError:
        raise
    except Exception:
        raise SkillsError("skills_io_failed") from None


def _parse(data):
    """Restricted safe YAML subset: exactly name/description single-line scalars.

    Plain strings, YAML single quoted (doubled apostrophe), and JSON-compatible
    double quoted strings supported. Multiline, tags, aliases, anchors, mappings,
    sequences, comments, extra keys and implicit typed scalars explicitly rejected.
    No reliance on a transitive YAML dependency or permissive implicit coercion.
    """
    try:
        text = data.decode("utf-8")
        if len(data) > MAX_FILE_BYTES or "\x00" in text or any(0xD800 <= ord(c) <= 0xDFFF for c in text):
            raise ValueError()
        lines = text.splitlines(keepends=True)
        if not lines or lines[0].rstrip("\r\n") != "---":
            raise ValueError()
        end = next(i for i in range(1, min(len(lines), 10)) if lines[i].rstrip("\r\n") == "---")
        values = {}
        for line in lines[1:end]:
            match = re.fullmatch(r"(name|description):[ \t]*(.+)", line.rstrip("\r\n"))
            if not match or match[1] in values:
                raise ValueError()
            key, raw = match[1], match[2].strip()
            if raw.startswith('"'):
                value = json.loads(raw)
            elif raw.startswith("'"):
                if not re.fullmatch(r"'(?:[^']|'')*'", raw):
                    raise ValueError()
                value = raw[1:-1].replace("''", "'")
            else:
                if (raw[0] in "!&*|>{[?%@`-" or any(c in raw for c in "#:{}[]\"'") or
                        raw.lower() in ("null", "true", "false", "yes", "no", "on", "off", "~", "y", "n", ".inf", "+.inf", "-.inf", ".nan") or
                        re.fullmatch(r"[+-]?[0-9].*", raw)):
                    raise ValueError()
                value = raw
            limit = 100 if key == "name" else 500
            if type(value) is not str or not value.strip() or len(value) > limit or any(ord(c) < 32 or ord(c) == 127 or 0xD800 <= ord(c) <= 0xDFFF for c in value):
                raise ValueError()
            values[key] = value
        if values.keys() != {"name", "description"}:
            raise ValueError()
        return values, "".join(lines[end + 1:])
    except Exception:
        raise SkillsError("skills_frontmatter_invalid") from None


def _content(text):
    if type(text) is not str or len(text) > MAX_FILE_BYTES:
        raise SkillsError()
    try:
        data = text.encode("utf-8")
    except UnicodeError:
        raise SkillsError() from None
    _parse(data)
    return data


@dataclass(repr=False)
class _Plan:
    session: str
    created: float
    operation: str
    target: Path
    source: _File
    desired: bytes | None = field(repr=False)
    import_path: Path | None = None
    imported: _File | None = None
    save_id: str | None = None


@dataclass(repr=False)
class _Saved:
    plan: _Plan
    applied: _File
    directory: Path
    backup: _File
    audit: _File


class SkillsStore:
    """Root is trusted workspace; user files live at skills/{slug}/SKILL.md.

    builtin_root is a separately host-supplied directory with same slug layout.
    All builtin FILES are immutable. Enable/disable is workspace user state, even
    for builtins. Catalog has only id/name/description/source/enabled, never body.
    Catalog still reads/decodes each bounded full SKILL.md to validate it; this is
    metadata-only output, not metadata-only filesystem access. Resources likewise
    read bounded file bytes for validation, but never return or execute contents.
    Unsafe/invalid entries fail the whole catalog, not silently skip or follow.
    """
    error_type = SkillsError
    error_codes = ERROR_CODES

    @_public
    def __init__(self, workspace_root, builtin_root=None):
        self.root = Path(os.path.abspath(workspace_root))
        self.skills_root = self.root / "skills"
        self.builtin_root = Path(os.path.abspath(builtin_root)) if builtin_root is not None else None
        self._root_identity = _directory(self.root)
        self._builtin_identity = _directory(self.builtin_root) if self.builtin_root is not None else None
        if self.builtin_root is not None and (self.builtin_root == self.root or self.builtin_root in self.root.parents or self.root in self.builtin_root.parents):
            raise SkillsError("skills_roots_overlap")
        self._skills_identity = _directory(self.skills_root, True)
        self._pending, self._saved, self._issued = {}, {}, set()
        self._fused = False

    def __repr__(self):
        return "<SkillsStore host-only>"

    def _check(self):
        if _directory(self.root) != self._root_identity or _directory(self.skills_root, True) != self._skills_identity:
            raise SkillsError("skills_path_changed")
        if self.builtin_root is not None and _directory(self.builtin_root) != self._builtin_identity:
            raise SkillsError("skills_path_changed")

    def _same(self, path, expected, parents=True):
        self._check()
        current = _read(path, True)
        if current.stamp != expected.stamp or (parents and current.parents != expected.parents):
            raise SkillsError("skills_source_changed")
        return current

    def _state(self):
        self._check()
        source = _read(self.root / ".office-skills-state.json", True)
        if source.stamp == ("missing",):
            return {}, source
        try:
            def pairs(items):
                result = {}
                for key, value in items:
                    if key in result:
                        raise ValueError()
                    result[key] = value
                return result
            values = json.loads(source.data, object_pairs_hook=pairs)
            if type(values) is not dict or len(values) > MAX_SKILLS * 2:
                raise ValueError()
            for key, enabled in values.items():
                self._identifier(key)
                if type(enabled) is not bool:
                    raise ValueError()
            return values, source
        except Exception:
            raise SkillsError("skills_state_invalid") from None

    def _identifier(self, identifier):
        if type(identifier) is not str or identifier.count(":") != 1:
            raise SkillsError("skills_id_invalid")
        source, slug = identifier.split(":")
        _slug(slug)
        if source not in ("user", "builtin") or (source == "builtin" and self.builtin_root is None):
            raise SkillsError("skills_id_invalid")
        return source, slug

    def _path(self, identifier):
        source, slug = self._identifier(identifier)
        return (self.skills_root if source == "user" else self.builtin_root) / slug / "SKILL.md"

    def _entries(self, path):
        _directory(path)
        entries = []
        with os.scandir(path) as scan:
            for entry in scan:
                if len(entries) >= MAX_ENTRIES:
                    raise SkillsError("skills_capacity")
                entries.append(Path(entry.path))
        return sorted(entries)

    @_public
    def catalog(self):
        self._check()
        enabled, _ = self._state()
        result = []
        for source, root in (("user", self.skills_root), ("builtin", self.builtin_root)):
            if root is None or not root.exists():
                continue
            for directory in self._entries(root):
                _slug(directory.name)
                _directory(directory)
                path = directory / "SKILL.md"
                # Deleted skills leave resources untouched and are no longer catalogued.
                if not path.exists():
                    _read(path, True)  # still rejects dangling links/reparse parents
                    continue
                metadata, _ = _parse(_read(path).data)
                identifier = source + ":" + directory.name
                result.append({"id": identifier, **metadata, "source": source, "enabled": enabled.get(identifier, True)})
                if len(result) > MAX_SKILLS:
                    raise SkillsError("skills_capacity")
        self._check()
        return result

    def _find(self, identifier_or_name):
        if type(identifier_or_name) is not str or len(identifier_or_name) > 100:
            raise SkillsError("skills_id_invalid")
        matches = [item for item in self.catalog() if identifier_or_name in (item["id"], item["name"])]
        if len(matches) != 1:
            raise SkillsError("skills_not_found" if not matches else "skills_name_ambiguous")
        return matches[0]

    @_public
    def read(self, identifier_or_name):
        """Host management view, including disabled untrusted instruction bodies."""
        return self._view(identifier_or_name, require_enabled=False)

    @_public
    def load(self, identifier_or_name):
        return self._view(identifier_or_name, require_enabled=True)

    def _view(self, identifier_or_name, require_enabled):
        item = self._find(identifier_or_name)
        if require_enabled and not item["enabled"]:
            raise SkillsError("skills_disabled")
        file = _read(self._path(item["id"]))
        metadata, body = _parse(file.data)
        if metadata != {key: item[key] for key in ("name", "description")}:
            raise SkillsError("skills_source_changed")
        self._check()
        return {**item, **metadata, "body": body, "sha256": file.stamp[-1], "untrusted_instructions": True}

    @_public
    def resources(self, identifier_or_name):
        """One level of optional resources/scripts, names/sizes only, never executes."""
        item = self._find(identifier_or_name)
        directory = self._path(item["id"]).parent
        result = []
        for kind in ("resources", "scripts"):
            folder = directory / kind
            if not folder.exists():
                _directory(folder, True)
                continue
            for path in self._entries(folder):
                # Read guard validates single link and bounded bytes; no body is returned.
                file = _read(path)
                if len(result) >= MAX_ENTRIES:
                    raise SkillsError("skills_capacity")
                result.append({"kind": kind, "name": path.name, "size_bytes": len(file.data), "executable": False})
        self._check()
        return result

    def _session(self, session):
        if type(session) is not str or not 1 <= len(session) <= 200 or any(ord(c) < 32 for c in session):
            raise SkillsError("skills_session_invalid")

    def _id(self, rollback=False):
        if self._fused:
            raise SkillsError("skills_uncertain")
        now = time.monotonic()
        self._pending = {key: plan for key, plan in self._pending.items() if now - plan.created < TTL_SECONDS}
        if (len(self._pending) >= MAX_PENDING or
                (not rollback and len(self._saved) >= MAX_PENDING) or len(self._issued) >= MAX_IDS):
            raise SkillsError("skills_capacity")
        for _ in range(4):
            identifier = secrets.token_hex(16)
            if type(identifier) is str and re.fullmatch(r"[0-9a-f]{32}", identifier) and identifier not in self._issued:
                self._issued.add(identifier)
                return identifier
        raise SkillsError("skills_id_collision")

    def _plan(self, operation, target, source, desired, session, **extra):
        self._session(session)
        self._check()
        identifier = self._id(rollback=operation == "rollback")
        self._pending[identifier] = _Plan(session, time.monotonic(), operation, target, source, desired, **extra)
        return {"plan_id": identifier, "operation": operation, "confirmation_required": True, "wrote_nothing": True,
                "expires_in_seconds": TTL_SECONDS, "backup_contains_skill_text": True,
                "source_sha256": None if source.stamp == ("missing",) else source.stamp[-1]}

    @_public
    def plan_create(self, slug, content, session_id):
        _slug(slug)
        desired = _content(content)
        if len(self.catalog()) >= MAX_SKILLS:
            raise SkillsError("skills_capacity")
        target = self._path("user:" + slug)
        source = _read(target, True)
        if source.stamp != ("missing",):
            raise SkillsError("skills_exists")
        return {**self._plan("create", target, source, desired, session_id), "id": "user:" + slug}

    def _user(self, identifier):
        source, _ = self._identifier(identifier)
        if source != "user":
            raise SkillsError("skills_builtin_immutable")
        return self._path(identifier)

    @_public
    def plan_edit(self, identifier, content, session_id):
        target = self._user(identifier)
        self._check()
        source = _read(target)
        _, previous = _parse(source.data)
        if previous.startswith(('LiX-Public-Skill-Package/1\n', 'LiX-Public-Skill-Package/2\n', 'LiX-Public-Skill-Package/3\n', 'LiX-Public-Skill-Package/4\n', 'LiX-Public-Skill-Package/5\n')):
            from .public_skill_package import repack_edit
            content = repack_edit(previous, content)
        desired = _content(content)
        return {**self._plan("edit", target, source, desired, session_id), "id": identifier}

    @_public
    def plan_public_update(self, identifier, content, expected_sha256, session_id):
        target = self._user(identifier)
        self._check()
        source = _read(target)
        if source.stamp[-1] != expected_sha256:
            raise SkillsError('skills_source_changed')
        from .public_skill_package import unpack
        _, previous = unpack(_parse(source.data)[1])
        _, following = unpack(_parse(_content(content))[1])
        if previous is None or following is None or any(previous[key] != following[key] for key in ('repository', 'path')):
            raise SkillsError('skills_invalid')
        return {**self._plan('edit', target, source, _content(content), session_id), 'id':identifier}

    @_public
    def plan_delete(self, identifier, session_id):
        target = self._user(identifier)
        self._check()
        source = _read(target)
        _parse(source.data)
        return {**self._plan("delete", target, source, None, session_id), "id": identifier, "resources_deleted": False, "directory_deleted": False}

    @_public
    def plan_import(self, slug, trusted_directory, session_id):
        """HOST ONLY: directory selected by host, never accept a client path."""
        _slug(slug)
        directory = Path(os.path.abspath(trusted_directory))
        _directory(directory)
        imported = _read(directory / "SKILL.md")
        _parse(imported.data)
        preview = self.plan_create(slug, imported.data.decode("utf-8"), session_id)
        plan = self._pending[preview["plan_id"]]
        plan.operation, plan.import_path, plan.imported = "import", directory / "SKILL.md", imported
        return {**preview, "operation": "import", "copied_files": ["SKILL.md"], "resources_copied": False, "scripts_copied": False}

    @_public
    def plan_enabled(self, identifier, enabled, session_id):
        if type(enabled) is not bool:
            raise SkillsError("skills_enabled_invalid")
        item = self._find(identifier)
        values, source = self._state()
        values[item["id"]] = enabled
        desired = json.dumps(values, sort_keys=True).encode("utf-8")
        if len(values) > MAX_SKILLS * 2:
            raise SkillsError("skills_capacity")
        skill_path = self._path(item["id"])
        skill_source = _read(skill_path)
        return {**self._plan("enabled", self.root / ".office-skills-state.json", source, desired, session_id,
                            import_path=skill_path, imported=skill_source), "id": item["id"], "enabled": enabled}

    def _consume(self, identifier, session, confirmed, rollback):
        if type(identifier) is not str:
            raise SkillsError("skills_plan_invalid")
        self._session(session)
        plan = self._pending.get(identifier)
        if plan is None or (plan.save_id is not None) != rollback:
            raise SkillsError("skills_plan_invalid")
        if plan.session != session:
            raise SkillsError("skills_session_invalid")
        # Only the owner, through the correct confirmation API, can consume it.
        # Once owned, even denial/expiry is one-use (no replay of approval).
        self._pending.pop(identifier)
        if type(confirmed) is not bool or confirmed is not True:
            raise SkillsError("confirmation_required" if type(confirmed) is not bool else "skills_denied")
        if self._fused:
            raise SkillsError("skills_uncertain")
        if time.monotonic() - plan.created >= TTL_SECONDS:
            raise SkillsError("skills_plan_expired")
        return plan

    def _create(self, path, data):
        self._check()
        _directory(path.parent)
        with path.open("xb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        result = _read(path)
        if result.data != data:
            raise SkillsError("skills_verification_failed")
        return result

    def _audit(self, directory, plan, status):
        data = json.dumps({"version": 1, "operation": plan.operation, "status": status,
                           "target": str(plan.target.relative_to(self.root)),
                           "before_sha256": None if plan.source.stamp == ("missing",) else plan.source.stamp[-1],
                           "desired_sha256": None if plan.desired is None else hashlib.sha256(plan.desired).hexdigest()}, sort_keys=True).encode("utf-8")
        return self._create(directory / ("audit-" + status + ".json"), data)

    def _apply(self, plan, directory):
        self._same(plan.target, plan.source)
        if plan.desired is None:
            # Exact target/stamp checked; backup and fsynced intent already present.
            plan.target.unlink()
            result = _read(plan.target, True)
            if result.stamp != ("missing",) or result.parents != plan.source.parents:
                raise SkillsError("skills_verification_failed")
            self._check()
            return result
        replacement = self._create(directory / "replacement", plan.desired)
        self._same(plan.target, plan.source)
        if not plan.target.parent.exists():
            if not self.skills_root.exists():
                self.skills_root.mkdir()
                self._skills_identity = _directory(self.skills_root)
            plan.target.parent.mkdir()
        checked_target = self._same(plan.target, plan.source, parents=False)
        self._same(directory / "replacement", replacement)
        os.replace(directory / "replacement", plan.target)
        result = _read(plan.target)
        if (result.data != plan.desired or result.stamp[:2] != replacement.stamp[:2] or
                result.parents != checked_target.parents):
            raise SkillsError("skills_verification_failed")
        if plan.target.name == "SKILL.md":
            _parse(result.data)
        else:
            self._state()
        self._check()
        return result

    def confirm(self, plan_id, session_id, confirmed):
        wrote = False
        try:
            plan = self._consume(plan_id, session_id, confirmed, False)
            if len(self._saved) >= MAX_PENDING:
                raise SkillsError("skills_capacity")
            self._same(plan.target, plan.source)
            if plan.operation in ("create", "import") and len(self.catalog()) >= MAX_SKILLS:
                raise SkillsError("skills_capacity")
            if plan.import_path is not None:
                if _read(plan.import_path).stamp != plan.imported.stamp or _directory(plan.import_path.parent) != plan.imported.parents:
                    raise SkillsError("skills_source_changed")
            directory = self.root / (".office-skills-" + plan_id)
            # Collision cannot overwrite anything; fuse conservatively after write attempt.
            wrote = True
            directory.mkdir(mode=0o700)
            backup = self._create(directory / "backup", plan.source.data)
            intent = self._audit(directory, plan, "intent")
            self._same(directory / "backup", backup)
            self._same(directory / "audit-intent.json", intent)
            applied = self._apply(plan, directory)
            self._same(directory / "audit-intent.json", intent)
            audit = self._audit(directory, plan, "verified")
            saved = _Saved(plan, applied, directory, backup, audit)
            self._saved_check(saved)
            self._saved[plan_id] = saved
            return {"status": "saved", "save_id": plan_id, "applied": True, "wrote_nothing": False,
                    "verified": True, "applied_sha256": None if applied.stamp == ("missing",) else applied.stamp[-1],
                    "audit_sha256": audit.stamp[-1]}
        except Exception as exc:
            if wrote:
                self._fused = True
            return {"status": "uncertain" if wrote else "rejected", "code": "skills_uncertain" if wrote else (exc.code if isinstance(exc, SkillsError) else "skills_io_failed"), "applied": None if wrote else False, "wrote_nothing": not wrote}

    def _saved_check(self, saved):
        self._same(saved.plan.target, saved.applied)
        self._same(saved.directory / "backup", saved.backup)
        self._same(saved.directory / "audit-verified.json", saved.audit)

    @_public
    def plan_rollback(self, save_id, session_id):
        self._session(session_id)
        if type(save_id) is not str or save_id not in self._saved or self._saved[save_id].plan.session != session_id:
            raise SkillsError("skills_plan_invalid")
        saved = self._saved[save_id]
        self._saved_check(saved)
        desired = None if saved.plan.source.stamp == ("missing",) else saved.plan.source.data
        return self._plan("rollback", saved.plan.target, saved.applied, desired, session_id, save_id=save_id)

    def confirm_rollback(self, rollback_id, session_id, confirmed):
        wrote = False
        try:
            plan = self._consume(rollback_id, session_id, confirmed, True)
            saved = self._saved.get(plan.save_id)
            if saved is None or saved.plan.session != session_id:
                raise SkillsError("skills_plan_invalid")
            self._saved_check(saved)
            directory = self.root / (".office-skills-" + rollback_id)
            wrote = True
            directory.mkdir(mode=0o700)
            backup = self._create(directory / "backup", plan.source.data)
            intent = self._audit(directory, plan, "intent")
            self._same(directory / "backup", backup)
            self._same(directory / "audit-intent.json", intent)
            self._saved_check(saved)
            applied = self._apply(plan, directory)
            self._same(directory / "audit-intent.json", intent)
            audit = self._audit(directory, plan, "verified")
            self._saved_check(_Saved(plan, applied, directory, backup, audit))
            self._saved.pop(plan.save_id)
            return {"status": "rolled_back", "restored": True, "wrote_nothing": False,
                    "verified": True, "applied_sha256": None if applied.stamp == ("missing",) else applied.stamp[-1],
                    "audit_sha256": audit.stamp[-1]}
        except Exception as exc:
            if wrote:
                self._fused = True
            return {"status": "uncertain" if wrote else "rejected", "code": "skills_uncertain" if wrote else (exc.code if isinstance(exc, SkillsError) else "skills_io_failed"), "restored": None if wrote else False, "wrote_nothing": not wrote}

    def invalidate(self):
        """Retire approvals/rollback grants only; preserve files, audit and fuse."""
        self._pending.clear()
        self._saved.clear()

    @property
    def available(self):
        return not self._fused
