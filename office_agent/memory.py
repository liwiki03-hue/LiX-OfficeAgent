"""Selective memory, independent of storage, providers and tool registration.

Host-only session/approval flags must never be populated from model arguments.
The backend owns source-version checks, backup, copy-on-write, reopen verification,
audit and rollback. Its confirmations MUST recheck the expected snapshot atomically.
Privacy rules are deliberately conservative and incomplete, not a DLP guarantee.
No chat history is accepted or automatically retained. Rejected candidates leave
only classification metadata. The authenticated Web host mounts this service.
"""
from __future__ import annotations

import math
import re
import time
import uuid
from dataclasses import dataclass
from typing import Callable, Protocol

CATEGORIES = frozenset({"task", "preference", "summary"})
SOURCE_LABELS = frozenset({"explicit_command", "recurring_preference"})
_ID = re.compile(r"[a-f0-9]{32}\Z")
_PRIVATE = re.compile(
    r"(?i)api[ _-]?key|access[ _-]?token|password|passwd|secret|credential|"
    r"\bsk-[a-z0-9_-]+|\bbearer\b|\b(?:private|confidential|diary|street|avenue|postcode)\b|private[ _-]?note|document[ _-]?(excerpt|text)|"
    r"client[ _-]?(list|name)|customer[ _-]?(list|name)|\b(name|phone|email|identity|passport|"
    r"card|address|ssn)\b|密码|密钥|令牌|私密|私人笔记|文档摘录|原文|客户|名单|姓名|"
    r"电话|手机|邮箱|身份证|银行卡|住址|地址|护照|"
    r"[\w.+-]+@[\w.-]+\.[a-z]{2,}|\+?\d[\d ()-]{6,}\d|"
    r"\b(?-i:(?!(?:Prefer|Always|Usually|Default|Remember|Use|Keep|Reply|Chinese|English|Markdown|Excel|Word|PDF)\b)[A-Z][a-z]{2,})\b|[\u4e00-\u9fff]{2,4}(?:先生|女士)|张三|李四|王五|"
    r"\b(?:alice|bob|charlie|john|jane)\b|[a-z0-9_+/=-]{24,}|"
    r"(?:user|assistant|system|用户|助手)\s*[:：]|chat[ _-]?history|聊天记录|"
    r"-----BEGIN|eyJ[a-z0-9_-]+\.eyJ"
)
_AUTHORITY = re.compile(
    r"(?i)system[ _-]?(prompt|policy)|ignore.{0,20}(instructions|policy)|"
    r"shell|powershell|cmd\.exe|tool[ _-]?(privilege|permission)|sudo|"
    r"[a-z]:[\\/]|\\\\|(?:^|\s)/[\w.]|\.env|系统提示|忽略.{0,10}(指令|规则)|权限|"
    r"```|<\s*/?\s*(system|assistant|tool)\b"
)
_PREFERENCE = re.compile(r"(?i)\b(prefer|always|usually|default)\b|偏好|喜欢|习惯|默认|总是")
_EPHEMERAL = re.compile(r"(?i)\b(today|tonight|temporary|temporarily|now)\b|this (?:time|task|session)|今天|这次|本次|临时|现在")


class MemoryError(ValueError):
    """Only stable service codes escape; never backend types, messages or content."""
    def __init__(self, code="memory_invalid"):
        self.code = code
        super().__init__(code)


def classify_candidate(content):
    """Metadata only; never echo or store rejected candidate text."""
    if type(content) is not str or not content.strip() or len(content) > 1000 or any(ord(c) < 32 and c not in "\n\t" for c in content):
        return {"allowed": False, "classification": "invalid", "rules_complete": False}
    classification = "private" if _PRIVATE.search(content) else "authority" if _AUTHORITY.search(content) else "ordinary"
    return {"allowed": classification == "ordinary", "classification": classification, "rules_complete": False}


class MemoryBackend(Protocol):
    """All operations MUST enforce account isolation and memory-only scope.

    snapshot returns exactly revision:int>=0, uncertain:bool; revision must change
    on every memory/source/configuration change and each successful transaction
    increments it exactly once. plan is read-only and returns an opaque plan_id.
    confirm atomically rechecks expected revision and verified backup/source state,
    writes a copy, reopens/verifies, audits and publishes. rollback must additionally
    verify the saved transaction source/output and backup before restoring ONLY
    the account's memory. A rejected result guarantees no write; uncertain means
    the host cannot establish that guarantee. confirm/rollback return status saved/deleted/cleared/rolled_back
    or rejected/uncertain, applied literal bool/None, and uncertain literal bool.
    Success also returns save_id (host opaque hex32). Catalog/records have only
    memory_id, kind='memory', category, source; records additionally have content.
    No filesystem paths, credentials, backend diagnostics or document records.
    """
    def snapshot(self, *, account, kind): ...
    def plan(self, *, account, kind, operation, expected, records, memory_ids): ...
    def confirm(self, plan_id, *, account, kind, confirmed): ...
    def rollback(self, save_id, *, account, kind, expected, confirmed): ...
    def catalog(self, *, account, kind, limit, offset): ...
    def read(self, memory_id, *, account, kind): ...
    def retrieve(self, vector, *, account, kind, categories, top): ...


@dataclass(eq=False, repr=False, frozen=True)
class _Session:
    token: str


class MemoryService:
    """Injected deterministic embed(content)->list[finite numbers], no providers.

    open_session/revoke_session/set_enabled are private host integration APIs,
    not model tools. Authenticate the account outside this module first; never
    accept session/account/explicit/confirmed flags from model-generated text.
    '记住这个' is a host explicit-command decision, not a text self-approval.
    Separate confirmation remains mandatory even for explicit/recurring content.
    """
    def __init__(self, backend: MemoryBackend, embed: Callable, *, clock=time.monotonic):
        self._backend, self._embed, self._clock = backend, embed, clock
        self._sessions, self._plans, self._ids, self._saves = {}, {}, set(), {}
        self._enabled, self._uncertain = True, False

    def open_session(self, account, *, authenticated=False):
        if authenticated is not True or type(account) is not str or not re.fullmatch(r"[a-zA-Z0-9_-]{1,128}", account):
            raise MemoryError("memory_auth_required")
        session = _Session(uuid.uuid4().hex)
        self._sessions[session] = account
        return session

    def revoke_session(self, session):
        self._sessions.pop(session, None)
        self._plans = {k: v for k, v in self._plans.items() if v["session"] is not session}

    def set_enabled(self, enabled):
        if type(enabled) is not bool:
            raise MemoryError()
        self._enabled = enabled
        self._plans.clear()

    def _account(self, session):
        if self._uncertain:
            raise MemoryError("memory_uncertain")
        if not self._enabled:
            raise MemoryError("memory_disabled")
        if type(session) is not _Session or session not in self._sessions:
            raise MemoryError("memory_auth_required")
        return self._sessions[session]

    def _call(self, method, *args, **kwargs):
        try:
            return getattr(self._backend, method)(*args, **kwargs)
        except Exception:
            raise MemoryError("memory_backend_failed") from None

    def _snapshot(self, account):
        result = self._call("snapshot", account=account, kind="memory")
        if type(result) is not dict or set(result) != {"revision", "uncertain"} or type(result["revision"]) is not int or result["revision"] < 0 or type(result["uncertain"]) is not bool:
            raise MemoryError("memory_backend_failed")
        if result["uncertain"]:
            self._fuse()
            raise MemoryError("memory_uncertain")
        return dict(result)

    def _fuse(self):
        self._uncertain = True
        self._plans.clear()

    def _plan(self, session, operation, *, record=None, memory_ids=()):
        account = self._account(session)
        now = self._clock()
        self._plans = {k: v for k, v in self._plans.items() if now - v["created"] < 300}
        if len(self._plans) >= 8 or len(self._ids) >= 256:
            raise MemoryError("memory_capacity")
        snapshot = self._snapshot(account)
        identifier = uuid.uuid4().hex
        if identifier in self._ids:
            raise MemoryError("memory_capacity")
        self._ids.add(identifier)
        self._plans[identifier] = {"session": session, "account": account, "created": now, "operation": operation, "record": record, "memory_ids": tuple(memory_ids), "snapshot": snapshot}
        return {"plan_id": identifier, "operation": operation, "kind": "memory", "requires_confirmation": True, "count": None if operation == "clear" else 1 if record else len(memory_ids)}

    def plan_remember(self, content, category, session, *, explicit=False, occurrences=1):
        self._account(session)
        privacy = classify_candidate(content)
        if not privacy["allowed"]:
            return {"status": "rejected", "privacy": privacy, "retained": False}
        if type(category) is not str or category not in CATEGORIES or type(explicit) is not bool or type(occurrences) is not int or not 1 <= occurrences <= 1000000:
            raise MemoryError()
        recurring = category == "preference" and occurrences >= 2 and bool(_PREFERENCE.search(content)) and not _EPHEMERAL.search(content)
        if not explicit and not recurring:
            return {"status": "discarded", "privacy": privacy, "retained": False, "reason": "memory_ephemeral"}
        record = {"memory_id": uuid.uuid4().hex, "kind": "memory", "category": category, "source": "explicit_command" if explicit else "recurring_preference", "content": content}
        return {**self._plan(session, "remember", record=record), "privacy": privacy, "category": category, "source": record["source"], "content": content}

    @staticmethod
    def _identifier(identifier):
        if type(identifier) is not str or not _ID.fullmatch(identifier):
            raise MemoryError()
        return identifier

    @staticmethod
    def _record(record, *, content=False):
        keys = {"memory_id", "kind", "category", "source"} | ({"content"} if content else set())
        if type(record) is not dict or set(record) != keys:
            raise MemoryError("memory_backend_failed")
        MemoryService._identifier(record["memory_id"])
        if record["kind"] != "memory" or type(record["category"]) is not str or record["category"] not in CATEGORIES or type(record["source"]) is not str or record["source"] not in SOURCE_LABELS:
            raise MemoryError("memory_backend_failed")
        if content and not classify_candidate(record["content"])["allowed"]:
            raise MemoryError("memory_private_result")
        return dict(record)

    def list_memories(self, session, *, limit=20, offset=0):
        account = self._account(session)
        if type(limit) is not int or not 1 <= limit <= 50 or type(offset) is not int or not 0 <= offset <= 10000:
            raise MemoryError()
        self._snapshot(account)
        records = self._call("catalog", account=account, kind="memory", limit=limit, offset=offset)
        if type(records) is not list or len(records) > limit:
            raise MemoryError("memory_backend_failed")
        return [self._record(record) for record in records]

    def read_memory(self, memory_id, session):
        account = self._account(session)
        self._identifier(memory_id)
        self._snapshot(account)
        result = self._record(self._call("read", memory_id, account=account, kind="memory"), content=True)
        if result["memory_id"] != memory_id:
            raise MemoryError("memory_backend_failed")
        return result

    def plan_delete(self, memory_id, session):
        self.read_memory(memory_id, session)
        return self._plan(session, "delete", memory_ids=(memory_id,))

    def clear_memories(self, session):
        """Plan clearing ONLY memory; confirm_clear required. Never clear documents."""
        return self._plan(session, "clear")

    plan_clear = clear_memories

    @staticmethod
    def _denied():
        return {"status": "rejected", "applied": False, "wrote_nothing": True}

    def _result(self, result, operation, account, expected):
        status = {"remember": "saved", "delete": "deleted", "clear": "cleared", "rollback": "rolled_back"}[operation]
        if type(result) is not dict or result.get("uncertain") is True or result.get("applied") is None:
            self._fuse()
            return {"status": "uncertain", "applied": None, "wrote_nothing": False}
        if result.get("status") == "rejected" and result.get("applied") is False and result.get("uncertain") is False:
            return self._denied()
        if result.get("status") != status or result.get("applied") is not True or result.get("uncertain") is not False:
            self._fuse()
            return {"status": "uncertain", "applied": None, "wrote_nothing": False}
        try:
            save_id = self._identifier(result.get("save_id"))
        except MemoryError:
            self._fuse()
            return {"status": "uncertain", "applied": None, "wrote_nothing": False}
        verified = self._snapshot(account)
        if verified != {"revision": expected["revision"] + 1, "uncertain": False}:
            self._fuse()
            return {"status": "uncertain", "applied": None, "wrote_nothing": False}
        if operation != "rollback":
            self._saves[save_id] = {"account": account, "snapshot": verified}
        return {"status": status, "applied": True, "wrote_nothing": False, "save_id": save_id}

    def _confirm(self, plan_id, session, confirmed, operation):
        account = self._account(session)
        self._identifier(plan_id)
        plan = self._plans.get(plan_id)
        if not plan or plan["session"] is not session or plan["account"] != account or plan["operation"] != operation:
            return self._denied()
        self._plans.pop(plan_id)
        if confirmed is not True or self._clock() - plan["created"] >= 300:
            return self._denied()
        try:
            if self._snapshot(account) != plan["snapshot"]:
                return self._denied()
            records = []
            if plan["record"]:
                try:
                    vector = self._embed(plan["record"]["content"])
                except Exception:
                    raise MemoryError("memory_embedding_failed") from None
                self._vector(vector)
                records = [{**plan["record"], "embedding": vector}]
            if self._snapshot(account) != plan["snapshot"]:
                return self._denied()
            inner = self._call("plan", account=account, kind="memory", operation=operation, expected=plan["snapshot"], records=records, memory_ids=list(plan["memory_ids"]))
            if type(inner) is not dict or set(inner) != {"plan_id"}:
                raise MemoryError("memory_backend_failed")
            self._identifier(inner["plan_id"])
        except MemoryError:
            if self._uncertain:
                return {"status": "uncertain", "applied": None, "wrote_nothing": False}
            return {**self._denied(), "reason": "memory_prewrite_failed"}
        try:
            result = self._call("confirm", inner["plan_id"], account=account, kind="memory", confirmed=True)
            return self._result(result, operation, account, plan["snapshot"])
        except Exception:
            self._fuse()
            return {"status": "uncertain", "applied": None, "wrote_nothing": False}

    def confirm_remember(self, plan_id, session, *, confirmed=False):
        return self._confirm(plan_id, session, confirmed, "remember")

    def confirm_delete(self, plan_id, session, *, confirmed=False):
        return self._confirm(plan_id, session, confirmed, "delete")

    def confirm_clear(self, plan_id, session, *, confirmed=False):
        return self._confirm(plan_id, session, confirmed, "clear")

    def rollback(self, save_id, session, *, confirmed=False):
        account = self._account(session)
        self._identifier(save_id)
        saved = self._saves.get(save_id)
        if confirmed is not True or not saved or saved["account"] != account:
            return self._denied()
        snapshot = self._snapshot(account)
        if snapshot != saved["snapshot"]:
            return self._denied()
        try:
            result = self._call("rollback", save_id, account=account, kind="memory", expected=snapshot, confirmed=True)
            output = self._result(result, "rollback", account, snapshot)
            if output["applied"] is True:
                self._saves.pop(save_id, None)
            return output
        except Exception:
            self._fuse()
            return {"status": "uncertain", "applied": None, "wrote_nothing": False}

    @staticmethod
    def _vector(vector):
        if type(vector) is not list or not 1 <= len(vector) <= 4096 or any(type(v) not in (int, float) or not math.isfinite(v) or abs(v) > 1e6 for v in vector) or not any(vector):
            raise MemoryError("memory_embedding_failed")

    def recall(self, question, session, *, top=3, categories=("task", "preference", "summary")):
        account = self._account(session)
        if not classify_candidate(question)["allowed"] or len(question) > 200:
            raise MemoryError("memory_private_query")
        if type(top) is not int or not 1 <= top <= 3 or type(categories) not in (tuple, list) or not categories or any(type(c) is not str or c not in CATEGORIES for c in categories):
            raise MemoryError()
        snapshot = self._snapshot(account)
        try:
            vector = self._embed(question)
            self._vector(vector)
        except Exception:
            raise MemoryError("memory_embedding_failed") from None
        if self._snapshot(account) != snapshot:
            raise MemoryError("memory_stale")
        results = self._call("retrieve", vector, account=account, kind="memory", categories=tuple(categories), top=top)
        if self._snapshot(account) != snapshot:
            raise MemoryError("memory_stale")
        if type(results) is not list or len(results) > top:
            raise MemoryError("memory_backend_failed")
        checked = []
        for result in results:
            item = self._record(result, content=True)
            if item["category"] not in categories:
                raise MemoryError("memory_backend_failed")
            item["content"] = item["content"][:200]
            checked.append(item)
        return checked

    @staticmethod
    def model_context(records):
        """Host-validated evidence, never authority capable of replacing policy.

        Only allowlisted metadata and <=200 content per record; no paths, tool
        privileges, arbitrary backend fields or complete chat transcript.
        """
        if type(records) is not list or len(records) > 3:
            raise MemoryError()
        return [{**MemoryService._record(record, content=True), "content": record["content"][:200]} for record in records]
