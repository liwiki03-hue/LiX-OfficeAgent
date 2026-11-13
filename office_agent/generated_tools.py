"""Independent generated-tool governance; deliberately no generated-code runtime.

AST eligibility is only a review aid, never isolation. This module does not own
filesystem roots, register tools, or execute source. A trusted host may inject a
transactional artifact backend at bootstrap; never accept one from model args.
The host authenticates the human separately before calling approve(). Python
objects/private attributes are not a security boundary against hostile host code.
"""
from __future__ import annotations

import ast
import hashlib
import json
import math
import re
import secrets
import threading
import time
from dataclasses import asdict, dataclass, replace
from typing import Protocol


class GovernanceError(ValueError):
    """Deterministic proposal/approval rejection, without code execution."""


def _json(value: object) -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True, allow_nan=False)
    except (TypeError, ValueError, RecursionError) as exc:
        raise GovernanceError("strict bounded JSON required") from exc


def _hash(value: object) -> str:
    return hashlib.sha256(_json(value).encode("utf-8")).hexdigest()


def _name(value: object) -> str:
    if type(value) is not str or not re.fullmatch(r"[a-z][a-z0-9_]{0,47}", value) or "__" in value:
        raise GovernanceError("name must be a lowercase slug (1..48), without magic names")
    return value


def _schema(schema: object, depth: int = 0) -> None:
    if depth > 4 or type(schema) is not dict:
        raise GovernanceError("schema must be an object, depth <=4")
    if depth == 0 and len(_json(schema)) > 8000:
        raise GovernanceError("schema exceeds JSON bound")
    kind = schema.get("type")
    fields = {"type"}
    if kind == "object":
        fields |= {"properties", "required", "additionalProperties"}
        props, required = schema.get("properties"), schema.get("required")
        if type(props) is not dict or len(props) > 16 or type(required) is not list or len(required) > 16:
            raise GovernanceError("object schema requires <=16 properties and required list")
        if schema.get("additionalProperties") is not False:
            raise GovernanceError("additionalProperties must be literal false")
        if any(type(x) is not str for x in required) or len(set(required)) != len(required) or not set(required) <= set(props):
            raise GovernanceError("invalid required fields")
        for key, sub in props.items():
            _name(key)
            _schema(sub, depth + 1)
    elif kind == "array":
        fields |= {"items", "maxItems"}
        if type(schema.get("maxItems")) is not int or not 0 <= schema["maxItems"] <= 64:
            raise GovernanceError("arrays require maxItems <=64")
        _schema(schema.get("items"), depth + 1)
    elif kind == "string":
        fields |= {"maxLength"}
        if type(schema.get("maxLength")) is not int or not 0 <= schema["maxLength"] <= 2000:
            raise GovernanceError("strings require maxLength <=2000")
    elif kind in ("integer", "number", "boolean", "null"):
        pass
    else:
        raise GovernanceError("unsupported JSON schema type")
    if set(schema) != fields:
        raise GovernanceError("unknown or missing JSON schema keyword")


def validate_arguments(schema: dict, arguments: object) -> None:
    """Validate the supported strict JSON-schema subset, without coercion."""
    if len(_json(arguments)) > 16000:
        raise GovernanceError("arguments exceed JSON bound")
    _schema(schema)

    def visit(spec: dict, value: object) -> None:
        kind = spec["type"]
        if kind == "object":
            if type(value) is not dict or not set(spec["required"]) <= set(value) or not set(value) <= set(spec["properties"]):
                raise GovernanceError("argument object fields mismatch")
            for key, item in value.items():
                visit(spec["properties"][key], item)
        elif kind == "array":
            if type(value) is not list or len(value) > spec["maxItems"]:
                raise GovernanceError("invalid array argument")
            for item in value:
                visit(spec["items"], item)
        elif kind == "string":
            if type(value) is not str or len(value) > spec["maxLength"]:
                raise GovernanceError("invalid string argument")
        elif kind == "number":
            if type(value) not in (int, float) or (type(value) is float and not math.isfinite(value)):
                raise GovernanceError("invalid number argument")
        elif type(value) is not {"integer": int, "boolean": bool, "null": type(None)}[kind]:
            raise GovernanceError("argument type mismatch")
    visit(schema, arguments)


@dataclass(frozen=True)
class GeneratedTool:
    name: str
    description: str
    schema_json: str
    code: str
    version: str
    enabled: bool = False

    def metadata(self) -> dict:
        return {**asdict(self), "args_schema": json.loads(self.schema_json), "kind": "generated",
                "runtime_status": "sandbox_required", "scan": "eligibility_only_not_security"}


def proposal(name: str, description: str, args_schema: dict, code: str) -> GeneratedTool:
    """Review host-supplied source text; never ask a model to access files."""
    _name(name)
    if type(description) is not str or not 1 <= len(description) <= 500 or not description.strip() or not description.isprintable():
        raise GovernanceError("description must be printable text, 1..500 characters")
    if type(code) is not str or not 1 <= len(code) <= 8000 or "\x00" in code:
        raise GovernanceError("source text must be 1..8000 characters")
    _schema(args_schema)
    encoded = _json(args_schema)
    if len(encoded) > 8000 or args_schema["type"] != "object":
        raise GovernanceError("argument schema must be a bounded object")
    try:
        tree = ast.parse(code)
    except (SyntaxError, ValueError, RecursionError, MemoryError) as exc:
        raise GovernanceError("source cannot be reviewed") from exc
    if len(tree.body) != 1 or type(tree.body[0]) is not ast.FunctionDef:
        raise GovernanceError("exactly one ordinary function required")
    fn = tree.body[0]
    args = fn.args
    if (fn.name != name or fn.decorator_list or fn.returns or args.posonlyargs or args.kwonlyargs
            or args.vararg or args.kwarg or args.defaults or args.kw_defaults
            or any(a.annotation for a in args.args)
            or len(args.args) != len(args_schema["properties"])
            or {a.arg for a in args.args} != set(args_schema["properties"])):
        raise GovernanceError("function name/signature must match schema, no evaluation at definition")
    # Conservative eligibility grammar: calls, attributes, imports, IO,
    # constructors, introspection, lambdas, comprehensions, loops and magic
    # names are rejected. Even this grammar is NOT an OS security boundary.
    allowed = (ast.Module, ast.FunctionDef, ast.arguments, ast.arg, ast.Return,
               ast.Assign, ast.Name, ast.Load, ast.Store, ast.Constant, ast.BinOp,
               ast.UnaryOp, ast.BoolOp, ast.Compare, ast.If, ast.IfExp, ast.List,
               ast.Tuple, ast.Dict, ast.Subscript, ast.Slice, ast.Add, ast.Sub,
               ast.Mult, ast.Div, ast.FloorDiv, ast.Mod, ast.USub, ast.UAdd,
               ast.Not, ast.And, ast.Or, ast.Eq, ast.NotEq, ast.Lt, ast.LtE,
               ast.Gt, ast.GtE, ast.Is, ast.IsNot, ast.In, ast.NotIn, ast.Pass)
    nodes = list(ast.walk(tree))
    if len(nodes) > 1000:
        raise GovernanceError("source AST exceeds review bound")
    local_names = {a.arg for a in args.args} | {
        node.id for node in nodes if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store)
    }
    for node in nodes:
        if (type(node) not in allowed or (isinstance(node, ast.Name)
                and ("__" in node.id or node.id not in local_names))):
            raise GovernanceError("source outside informational eligibility policy")
        if isinstance(node, ast.FunctionDef) and node is not fn:
            raise GovernanceError("nested functions are ineligible")
    version = _hash({"name": name, "description": description, "schema": encoded, "code": code})
    return GeneratedTool(name, description, encoded, code, version)


def manifest_hash(records: tuple[GeneratedTool, ...]) -> str:
    return _hash([asdict(record) for record in sorted(records, key=lambda r: r.name)])


@dataclass(frozen=True)
class ArtifactTransaction:
    """Immutable compare-and-swap request; contains text, never filesystem roots."""
    transaction_id: str
    operation: str
    name: str
    expected_manifest_hash: str
    desired_records: tuple[GeneratedTool, ...]


@dataclass(frozen=True)
class ArtifactReceipt:
    """Trusted backend result, NOT an isolation capability or caller proof.

    committed means backup verified -> staging -> publish -> reopen verified ->
    durable audit, with rollback restricted to transaction-owned artifacts.
    rejected means confirmed zero change or completed owned rollback + audit.
    Any unrecognized/malformed result is uncertain and trips a sticky fuse.
    """
    transaction_id: str
    status: str
    before_hash: str
    after_hash: str
    backup_verified: bool
    staging_verified: bool
    reopen_verified: bool
    audit_recorded: bool
    owned_rollback: bool


class TransactionalArtifactBackend(Protocol):
    """Authenticated host bootstrap dependency, not a model-callable backend.

    Must enforce filesystem authorization/link guards itself, atomically compare
    the physical manifest hash, preserve backups, stage/reopen artifacts without
    importing them, durably audit, and only roll back artifacts it owns. Return
    rejected only when no mutation remains; raise/unknown outcomes are uncertain.
    This module supplies no concrete disk backend and does not assert OS isolation.
    """
    def transact(self, request: ArtifactTransaction) -> ArtifactReceipt: ...


@dataclass(frozen=True)
class _Pending:
    request: ArtifactTransaction
    expires_at: float


class GeneratedToolGovernance:
    """Host-owned session. Approval mutates artifacts only, never runs/registers."""
    TTL = 300
    MAX_PENDING = 8
    MAX_LIFETIME = 256

    def __init__(self, *, backend: TransactionalArtifactBackend | None = None,
                 initial_records: tuple[GeneratedTool, ...] = (),
                 builtin_names: tuple[str, ...] = (), clock=time.monotonic):
        self._backend = backend
        self._clock = clock
        self._lock = threading.RLock()
        self._records: dict[str, GeneratedTool] = {}
        self._builtins = frozenset(_name(n) for n in builtin_names)
        if len(initial_records) > 256:
            raise GovernanceError("catalogue bound exceeded")
        for record in initial_records:
            if type(record) is not GeneratedTool or type(record.enabled) is not bool:
                raise GovernanceError("invalid host initial record")
            checked = proposal(record.name, record.description, json.loads(record.schema_json), record.code)
            if record.version != checked.version or record.name in self._records or record.name in self._builtins:
                raise GovernanceError("invalid/duplicate initial record")
            self._records[record.name] = record
        self._pending: dict[str, _Pending] = {}
        self._lifetime = 0
        self._uncertain = False

    @property
    def manifest_hash(self) -> str:
        with self._lock:
            return manifest_hash(tuple(self._records.values()))

    @property
    def uncertain(self) -> bool:
        return self._uncertain

    def list_tools(self) -> list[dict]:
        with self._lock:
            return ([{"name": n, "kind": "builtin", "management": "delegated_immutable"} for n in sorted(self._builtins)]
                    + [r.metadata() for r in sorted(self._records.values(), key=lambda r: r.name)])

    def read(self, name: str) -> dict:
        with self._lock:
            if name in self._builtins:
                return {"name": name, "kind": "builtin", "management": "delegated_immutable"}
            if name not in self._records:
                raise GovernanceError("unknown generated tool")
            return self._records[name].metadata()

    def plan(self, operation: str, *, name: str, description: str | None = None,
             args_schema: dict | None = None, code: str | None = None) -> dict:
        """Create/import/edit source is supplied as text, never a model path."""
        with self._lock:
            if self._uncertain:
                raise GovernanceError("backend uncertainty: session fused; host reconciliation required")
            _name(name)
            if name in self._builtins:
                raise GovernanceError("builtin management is delegated and immutable")
            if operation not in ("create", "import", "edit", "enable", "disable", "delete"):
                raise GovernanceError("unsupported management operation")
            before = self._records.get(name)
            desired = dict(self._records)
            if operation in ("create", "import", "edit"):
                if (operation == "edit") != (before is not None):
                    raise GovernanceError("create/import needs new name; edit needs existing tool")
                desired[name] = proposal(name, description, args_schema, code)
                if len(desired) > 256:
                    raise GovernanceError("catalogue bound exceeded")
            else:
                if before is None:
                    raise GovernanceError("unknown generated tool")
                if any(x is not None for x in (description, args_schema, code)):
                    raise GovernanceError("unexpected management arguments")
                if operation == "delete":
                    del desired[name]
                else:
                    desired[name] = replace(before, enabled=operation == "enable")
            now = self._clock()
            self._pending = {key: item for key, item in self._pending.items() if now < item.expires_at}
            if len(self._pending) >= self.MAX_PENDING or self._lifetime >= self.MAX_LIFETIME:
                raise GovernanceError("private session plan bound exceeded")
            token = secrets.token_hex(24)
            request = ArtifactTransaction(token, operation, name, self.manifest_hash,
                                          tuple(sorted(desired.values(), key=lambda r: r.name)))
            self._pending[token] = _Pending(request, now + self.TTL)
            self._lifetime += 1
            return {"plan_id": token, "operation": operation, "name": name,
                    "before": None if before is None else before.metadata(),
                    "after": None if name not in desired else desired[name].metadata(),
                    "manifest_hash": request.expected_manifest_hash,
                    "desired_manifest_hash": manifest_hash(request.desired_records),
                    "confirmation_required": True, "expires_in_seconds": self.TTL,
                    "runtime_status": "sandbox_required"}

    def cancel(self, plan_id: str) -> None:
        with self._lock:
            if type(plan_id) is not str or plan_id not in self._pending:
                raise GovernanceError("unknown private session plan")
            del self._pending[plan_id]

    def approve(self, plan_id: str, *, confirmed: bool = False) -> dict:
        """Only authenticated human host code should call this method.

        Literal True is necessary, not human authentication in itself. This API
        is deliberately not mounted as a ToolSpec/API/UI endpoint.
        """
        with self._lock:
            if confirmed is not True:
                raise GovernanceError("literal human confirmed True required")
            if self._uncertain:
                raise GovernanceError("backend uncertainty: session fused; host reconciliation required")
            if type(plan_id) is not str or plan_id not in self._pending:
                raise GovernanceError("unknown private session plan")
            pending = self._pending.pop(plan_id)  # consume once, including expiry/failure
            request = pending.request
            if self._clock() >= pending.expires_at:
                raise GovernanceError("private session plan expired")
            if request.expected_manifest_hash != self.manifest_hash:
                raise GovernanceError("manifest changed; replan and reconfirm")
            if self._backend is None:
                return {"status": "artifact_backend_required", "runtime_status": "sandbox_required"}
            # Mark uncertain before callback: exceptions/reentrancy cannot leave
            # the session open while an artifact mutation may be in flight.
            self._uncertain = True
            try:
                receipt = self._backend.transact(request)
                valid = (type(receipt) is ArtifactReceipt
                         and receipt.transaction_id == request.transaction_id
                         and receipt.before_hash == request.expected_manifest_hash
                         and receipt.audit_recorded is True and receipt.owned_rollback is True)
                if valid and receipt.status == "rejected" and receipt.after_hash == receipt.before_hash:
                    self._uncertain = False
                    return {"status": "rejected", "runtime_status": "sandbox_required"}
                if not (valid and receipt.status == "committed"
                        and receipt.backup_verified is True and receipt.staging_verified is True
                        and receipt.reopen_verified is True
                        and receipt.after_hash == manifest_hash(request.desired_records)):
                    return {"status": "backend_uncertain", "runtime_status": "sandbox_required"}
            except Exception:
                return {"status": "backend_uncertain", "runtime_status": "sandbox_required"}
            self._records = {r.name: r for r in request.desired_records}
            self._uncertain = False
            return {"status": "committed", "manifest_hash": self.manifest_hash,
                    "runtime_status": "sandbox_required", "registered": False, "executed": False}

    def run(self, name: object = None, arguments: object = None, **caller_claims: object) -> dict:
        """Unconditionally blocked. Caller receipts/proofs/args authorize nothing.

        No isolated executor is available. A future host integration would need
        independently verified OS denial of filesystem/network/process access
        and resource enforcement; AST and resource caps cannot supply that proof.
        """
        return {"status": "sandbox_required", "executed": False,
                "reason": "No trusted real OS isolation capability/executor is installed."}
