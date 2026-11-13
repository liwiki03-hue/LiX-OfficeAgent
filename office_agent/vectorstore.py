"""Trusted-host, bounded sqlite-vec storage; never a model-facing filesystem tool.

The host supplies extracted records and independently collects literal human approval.
No source files, providers, environment configuration, or Office are accessed here.
Backups are retained, not silently garbage-collected. Authority is process-local.
A persistent uncertainty marker blocks crash-interrupted publications on reopening.
Only the approved audit/fuse metadata is updated in-place after content publication.
"""
from __future__ import annotations

import hashlib
from functools import wraps
import math
import os
from pathlib import Path
import re
import sqlite3
import stat
import struct
import threading
import time
import uuid
from contextlib import contextmanager

import sqlite_vec


class VectorStoreError(Exception):
    """Safe, content-free public failure. Only the stable code may be rendered."""

    def __init__(self, code: str):
        self.code = code
        super().__init__(code)


MAX_CHUNKS = 10_000
MAX_DOCUMENT_CHUNKS = 500
MAX_TEXT = 8_000
MAX_DIMENSION = 4_096
MAX_DB_BYTES = 256 * 1024 * 1024
MAX_AUTHORITIES = 16
MAX_AUDIT = 1_000
AUTHORITY_SECONDS = 300
_BASENAME = ".office-knowledge.sqlite3"


def _fail(code):
    raise VectorStoreError(code)


def _safe(method):
    @wraps(method)
    def call(*args, **kwargs):
        try:
            return method(*args, **kwargs)
        except VectorStoreError:
            raise
        except Exception:
            raise VectorStoreError("host_error") from None
    return call


def _string(value, maximum, *, empty=False):
    if type(value) is not str or len(value) > maximum or (not empty and not value.strip()) or "\x00" in value:
        _fail("invalid_input")
    return value


def _identity(path, *, directory=False):
    try:
        s = path.lstat()
    except OSError:
        _fail("unsafe_path")
    if stat.S_ISLNK(s.st_mode) or getattr(s, "st_file_attributes", 0) & 0x400:
        _fail("unsafe_path")
    if directory:
        if not stat.S_ISDIR(s.st_mode):
            _fail("unsafe_path")
    elif not stat.S_ISREG(s.st_mode) or s.st_nlink != 1:
        _fail("unsafe_path")
    return (s.st_dev, s.st_ino)


def _ancestors(root):
    return tuple((str(p), _identity(p, directory=True)) for p in reversed((root, *root.parents)))


class VectorStore:
    @_safe
    def __init__(self, root, *, model: str, dimension: int, profile=None):
        self._model = _string(model, 200)
        if type(dimension) is not int or not 1 <= dimension <= MAX_DIMENSION:
            _fail("invalid_dimension")
        self._dimension = dimension
        raw = os.fspath(root)
        if isinstance(raw, str) and any(p == ".." for p in raw.replace("\\", "/").split("/")):
            _fail("unsafe_path")
        if not isinstance(raw, str) or not Path(raw).is_absolute() or raw.startswith(("\\\\?\\", "\\\\.\\")):
            _fail("unsafe_path")
        self._root = Path(raw)
        if any(p in (".", "..") or p.endswith((" ", ".")) or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)", p, re.I) for p in self._root.parts[1:]):
            _fail("unsafe_path")
        self._roots = _ancestors(self._root)
        if profile is not None and (type(profile) is not str or not re.fullmatch('[a-f0-9]{64}', profile)):
            _fail('invalid_input')
        self._db = self._root / (_BASENAME if profile is None else '.office-knowledge-' + profile + '.sqlite3')
        self._pending = {}
        self._saves = {}
        self._owned = {}
        self._fused = False
        self._bound = False
        self._baseline = None
        self._lock = threading.RLock()
        self._lexical_cache = None
        self._schema = self._reference_schema()
        self.snapshot()  # Missing DB is deliberately not initialized.

    @property
    def model(self):
        return self._model

    @property
    def dimension(self):
        return self._dimension

    def _guard(self):
        if _ancestors(self._root) != self._roots:
            _fail("stale_root")
        for suffix in ("-wal", "-shm", "-journal"):
            if os.path.lexists(str(self._db) + suffix):
                _fail("unsafe_sidecar")

    def _fingerprint(self, path):
        self._guard()
        if not os.path.lexists(path):
            if path == self._db and self._bound and self._baseline is not None:
                _fail("stale_database")
            return None
        ident = _identity(path)
        before = path.stat()
        if before.st_size > MAX_DB_BYTES:
            _fail("database_limit")
        hasher = hashlib.sha256()
        size = 0
        with path.open("rb") as stream:
            while data := stream.read(1024 * 1024):
                size += len(data)
                if size > MAX_DB_BYTES:
                    _fail("database_limit")
                hasher.update(data)
        digest = hasher.hexdigest()
        after = path.stat()
        if ident != _identity(path) or (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
            _fail("stale_database")
        fingerprint = (ident, after.st_size, after.st_mtime_ns, digest)
        if path == self._db and self._bound and fingerprint != self._baseline:
            _fail("stale_database")
        return fingerprint

    @contextmanager
    def _connection(self, path, *, create=False):
        self._guard()
        self._check_owned(path)
        # Reject a persistent WAL header before SQLite can create WAL/SHM sidecars.
        with path.open("rb") as stream:
            header = stream.read(100)
        if header and (len(header) < 20 or header[:16] != b"SQLite format 3\x00"):
            _fail("database_error")
        if header and (header[18] != 1 or header[19] != 1):
            _fail("unsafe_sidecar")
        conn = None
        try:
            conn = sqlite3.connect(path.as_uri() + ("?mode=rw" if create else "?mode=ro"), uri=True, timeout=2)
            conn.enable_load_extension(True)
            try:
                sqlite_vec.load(conn)
            finally:
                conn.enable_load_extension(False)
            conn.execute("PRAGMA busy_timeout=2000")
            conn.execute("PRAGMA trusted_schema=OFF")
            if create:
                conn.execute("PRAGMA journal_mode=DELETE")
                conn.execute("PRAGMA synchronous=FULL")
            else:
                conn.execute("PRAGMA query_only=ON")
            # Bound query/verification execution, including malformed host DBs.
            steps = [0]
            def budget():
                steps[0] += 1
                return int(steps[0] > 200_000)
            conn.set_progress_handler(budget, 1000)
            yield conn
        except VectorStoreError:
            raise
        except Exception:
            raise VectorStoreError("database_error") from None
        finally:
            if conn is not None:
                conn.close()

    def _initialize(self, conn):
        conn.executescript(f"""
            CREATE TABLE metadata(model TEXT NOT NULL, dimension INTEGER NOT NULL, revision INTEGER NOT NULL, uncertain INTEGER NOT NULL);
            CREATE TABLE chunks(id INTEGER PRIMARY KEY AUTOINCREMENT, doc_id TEXT NOT NULL, source_fingerprint TEXT NOT NULL,
                source TEXT NOT NULL, location TEXT NOT NULL, text TEXT NOT NULL, kind TEXT NOT NULL);
            CREATE VIRTUAL TABLE vectors USING vec0(embedding float[{self.dimension}] distance_metric=cosine, kind TEXT partition key);
            CREATE TABLE audit(seq INTEGER PRIMARY KEY, operation TEXT NOT NULL, revision INTEGER NOT NULL,
                chunk_count INTEGER NOT NULL, first_absent_restored_empty INTEGER NOT NULL);
        """)
        conn.execute("INSERT INTO metadata VALUES(?,?,0,0)", (self.model, self.dimension))
        conn.commit()

    @staticmethod
    def _schema_entries(conn):
        return tuple(conn.execute("SELECT type,substr(name,1,128),substr(tbl_name,1,128), CASE WHEN sql IS NULL THEN NULL WHEN length(sql)<=4096 THEN sql ELSE 'invalid_schema' END FROM sqlite_schema ORDER BY type,name LIMIT 64"))

    def _reference_schema(self):
        # An in-memory official-extension database supplies shadow-table definitions.
        # This never opens a caller DB, writes a file, or accepts caller SQL.
        conn = sqlite3.connect(':memory:')
        try:
            conn.enable_load_extension(True)
            try:
                sqlite_vec.load(conn)
            finally:
                conn.enable_load_extension(False)
            conn.execute('PRAGMA trusted_schema=OFF')
            self._initialize(conn)
            return self._schema_entries(conn)
        finally:
            conn.close()

    def _reject_profile_mismatch(self, conn):
        from copy import copy

        rows = conn.execute('SELECT model,dimension FROM metadata LIMIT 2').fetchall()
        if len(rows) != 1:
            _fail('invalid_database')
        model, dimension = rows[0]
        _string(model, 200)
        if type(dimension) is not int or not 1 <= dimension <= MAX_DIMENSION or (model, dimension) == (self.model, self.dimension):
            _fail('invalid_database')
        probe = copy(self)
        probe._model, probe._dimension = model, dimension
        probe._schema = probe._reference_schema()
        metadata = probe._metadata(conn, verify=True)
        self._fingerprint(self._db)
        if metadata['uncertain']:
            _fail('uncertain_state')
        _fail('model_mismatch')

    def _metadata(self, conn, *, verify=False):
        schema = self._schema_entries(conn)
        if schema != self._schema:
            normalized = []
            for kind,name,table,sql in schema:
                if kind == 'table' and name == 'vectors' and type(sql) is str:
                    sql = re.sub(r'float\[\d+\]',f'float[{self.dimension}]',sql)
                normalized.append((kind,name,table,sql))
            if tuple(normalized) == self._schema:
                self._reject_profile_mismatch(conn)
            _fail('invalid_database')
        rows = conn.execute("SELECT model,dimension,revision,uncertain FROM metadata LIMIT 2").fetchall()
        if len(rows) != 1 or rows[0][:2] != (self.model, self.dimension):
            self._reject_profile_mismatch(conn)
        revision = rows[0][2]
        if type(revision) is not int or revision < 0 or rows[0][3] not in (0, 1):
            _fail("invalid_database")
        count = conn.execute("SELECT count(*) FROM chunks").fetchone()[0]
        audit_count = conn.execute("SELECT count(*) FROM audit").fetchone()[0]
        if count > MAX_CHUNKS or audit_count > MAX_AUDIT:
            _fail("capacity_limit")
        if audit_count:
            if conn.execute("SELECT max(revision) FROM audit").fetchone()[0] != revision:
                _fail("invalid_database")
        elif revision != 0:
            _fail("invalid_database")
        if conn.execute("SELECT 1 FROM audit WHERE operation NOT IN ('replace','delete','rollback') OR revision<1 OR chunk_count<0 OR chunk_count>? OR first_absent_restored_empty NOT IN (0,1) LIMIT 1", (MAX_CHUNKS,)).fetchone():
            _fail("invalid_database")
        if conn.execute("SELECT 1 FROM chunks GROUP BY doc_id HAVING count(*)>? LIMIT 1", (MAX_DOCUMENT_CHUNKS,)).fetchone():
            _fail("capacity_limit")
        if verify:
            if conn.execute("PRAGMA quick_check").fetchone() != ("ok",):
                _fail("verification_failed")
            if conn.execute("SELECT count(*) FROM vectors").fetchone()[0] != count:
                _fail("verification_failed")
            if conn.execute("SELECT 1 FROM chunks c LEFT JOIN vectors v ON c.id=v.rowid WHERE v.rowid IS NULL OR c.kind!=v.kind LIMIT 1").fetchone():
                _fail("verification_failed")
            for row in conn.execute("SELECT c.doc_id,c.source_fingerprint,c.source,c.location,c.text,c.kind,v.embedding FROM chunks c JOIN vectors v ON c.id=v.rowid"):
                for value, bound in zip(row[:5], (128, 256, 512, 512, MAX_TEXT)):
                    _string(value, bound)
                if row[5] not in ("document", "memory") or len(row[6]) != 4 * self.dimension:
                    _fail("invalid_database")
                self._vector(struct.unpack(f"<{self.dimension}f", row[6]))
        return {"model": self.model, "dimension": self.dimension, "revision": revision, "chunk_count": count, "audit_count": audit_count, "uncertain": bool(rows[0][3])}

    @_safe
    def snapshot(self):
        with self._lock:
            fingerprint = self._fingerprint(self._db)
            if not self._bound:
                self._baseline = fingerprint
                self._bound = True
            if fingerprint is None:
                return {"exists": False, "model": self.model, "dimension": self.dimension, "revision": 0, "chunk_count": 0, "audit_count": 0, "uncertain": self._fused}
            with self._connection(self._db) as conn:
                result = self._metadata(conn, verify=True)
            if self._fingerprint(self._db) != fingerprint:
                _fail("stale_database")
            self._fused = self._fused or result["uncertain"]
            return dict(result, exists=True, uncertain=self._fused)

    def _vector(self, values):
        if type(values) not in (list, tuple) or len(values) != self.dimension:
            _fail("invalid_vector")
        if any(type(v) not in (float, int) or not math.isfinite(v) for v in values):
            _fail("invalid_vector")
        try:
            packed = struct.pack(f"<{self.dimension}f", *values)
            floats = struct.unpack(f"<{self.dimension}f", packed)
        except (OverflowError, struct.error):
            _fail("invalid_vector")
        if not all(math.isfinite(v) for v in floats) or not any(v != 0 for v in floats):
            _fail("invalid_vector")
        # Normalize in float64 first: protects native cosine from float32 norm overflow.
        norm = math.sqrt(sum(v * v for v in floats))
        return struct.pack(f"<{self.dimension}f", *(v / norm for v in floats))

    def _prune(self):
        now = time.monotonic()
        for mapping in (self._pending, self._saves):
            for key in list(mapping):
                if mapping[key]["expiry"] <= now:
                    if "backup" in mapping[key]:
                        self._owned.pop(str(mapping[key]["backup"]), None)
                    del mapping[key]

    def _plan(self, operation, session, **payload):
        _string(session, 128)
        self._guard()
        self._prune()
        if self._fused and operation != "rollback":
            _fail("uncertain_state")
        if len(self._pending) >= MAX_AUTHORITIES:
            _fail("authority_limit")
        token = uuid.uuid4().hex
        fingerprint = self._fingerprint(self._db)
        metadata = self.snapshot()
        self._pending[token] = dict(operation=operation, session=session, fingerprint=fingerprint,
            expiry=time.monotonic() + AUTHORITY_SECONDS, revision=metadata["revision"], **payload)
        return {"rollback_id" if operation == "rollback" else "plan_id": token,
                "operation": operation, "confirmation_required": True, "revision": metadata["revision"],
                "chunk_count": payload.get("count", metadata["chunk_count"]),
                "first_absent_restored_empty": payload.get("first_absent", False)}

    @_safe
    def plan_replace(self, doc_id: str, source_fingerprint: str, records: list[dict], session_id: str):
        with self._lock:
            _string(doc_id, 128)
            _string(source_fingerprint, 256)
            if type(records) is not list or not 1 <= len(records) <= MAX_DOCUMENT_CHUNKS:
                _fail("document_limit")
            prepared = []
            for record in records:
                if type(record) is not dict or set(record) - {"source", "location", "text", "embedding", "kind"} or not {"source", "location", "text", "embedding"} <= set(record):
                    _fail("invalid_input")
                kind = record.get("kind", "document")
                if kind not in ("document", "memory"):
                    _fail("invalid_input")
                prepared.append((_string(record["source"], 512), _string(record["location"], 512),
                    _string(record["text"], MAX_TEXT), kind, self._vector(record["embedding"])))
            metadata = self.snapshot()
            old = 0
            if metadata["exists"]:
                with self._connection(self._db) as conn:
                    old = conn.execute("SELECT count(*) FROM chunks WHERE doc_id=?", (doc_id,)).fetchone()[0]
            count = metadata["chunk_count"] - old + len(prepared)
            if count > MAX_CHUNKS:
                _fail("capacity_limit")
            return self._plan("replace", session_id, doc_id=doc_id, source_fingerprint=source_fingerprint, records=tuple(prepared), count=count)

    @_safe
    def corpus_manifest(self, scope):
        import json

        with self._lock:
            _string(scope, 64)
            fingerprint = self._fingerprint(self._db)
            if fingerprint is None:
                return {}
            result = {}
            with self._connection(self._db) as conn:
                metadata = self._metadata(conn, verify=True)
                if self._fused or metadata['uncertain']:
                    _fail('uncertain_state')
                for doc_id, version, source, location in conn.execute("SELECT doc_id,source_fingerprint,source,location FROM chunks WHERE kind='document'"):
                    try:
                        where = json.loads(location)
                    except ValueError:
                        continue
                    if type(where) is dict and where.get('corpus_scope') == scope:
                        row = result.setdefault(doc_id, {'source': source, 'fingerprint': version, 'locations': []})
                        if row['source'] != source or row['fingerprint'] != version:
                            _fail('verification_failed')
                        row['locations'].append(where)
            if self._fingerprint(self._db) != fingerprint:
                _fail('stale_database')
            return result

    @_safe
    def plan_sync(self, scope, documents, removals, session_id):
        import json

        with self._lock:
            _string(scope, 64)
            if type(documents) is not dict or type(removals) is not list or len(documents) + len(removals) > 40 or set(documents) & set(removals):
                _fail('invalid_input')
            old = self.corpus_manifest(scope)
            if any(key not in old for key in removals):
                _fail('invalid_input')
            prepared = {}
            for doc_id, document in documents.items():
                _string(doc_id, 128)
                _string(document['fingerprint'], 256)
                records = document['records']
                if type(records) is not list or not 1 <= len(records) <= MAX_DOCUMENT_CHUNKS:
                    _fail('document_limit')
                rows = []
                for record in records:
                    where = json.loads(record['location'])
                    if where.get('corpus_scope') != scope or record.get('kind') != 'document':
                        _fail('invalid_input')
                    rows.append((_string(record['source'], 128), _string(record['location'], 512),
                                 _string(record['text'], MAX_TEXT), 'document', self._vector(record['embedding'])))
                prepared[doc_id] = (document['fingerprint'], tuple(rows))
            snapshot = self.snapshot()
            count = snapshot['chunk_count'] - sum(len(old[key]['locations']) for key in set(prepared) | set(removals) if key in old) + sum(len(row[1]) for row in prepared.values())
            if count > MAX_CHUNKS:
                _fail('capacity_limit')
            if snapshot['exists']:
                with self._connection(self._db) as conn:
                    for key in prepared:
                        if key not in old and conn.execute('SELECT 1 FROM chunks WHERE doc_id=?', (key,)).fetchone():
                            _fail('invalid_input')
            return self._plan('replace', session_id, documents=prepared, removals=tuple(removals), count=count)

    @_safe
    def plan_delete(self, doc_id: str, session_id: str):
        with self._lock:
            _string(doc_id, 128)
            metadata = self.snapshot()
            old = 0
            if metadata["exists"]:
                with self._connection(self._db) as conn:
                    old = conn.execute("SELECT count(*) FROM chunks WHERE doc_id=?", (doc_id,)).fetchone()[0]
            if not old:
                _fail("document_missing")
            return self._plan("delete", session_id, doc_id=doc_id, count=metadata["chunk_count"] - old)

    def _consume(self, token, session, confirmed, operations):
        if type(token) is not str:
            _fail("invalid_authority")
        plan = self._pending.pop(token, None)  # ALL rejected confirmations consume authority.
        if plan is None:
            _fail("invalid_authority")
        if plan["expiry"] <= time.monotonic():
            _fail("expired_authority")
        if type(session) is not str or session != plan["session"]:
            _fail("session_mismatch")
        if confirmed is not True:
            _fail("confirmation_required")
        if plan["operation"] not in operations:
            _fail("invalid_authority")
        self._guard()
        if self._fingerprint(self._db) != plan["fingerprint"]:
            _fail("stale_database")
        if self._fused and plan["operation"] != "rollback":
            _fail("uncertain_state")
        return plan

    def _owned_file(self, suffix):
        self._guard()
        path = self._root / (".office-vector-" + uuid.uuid4().hex + suffix)
        fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        os.close(fd)
        self._owned[str(path)] = _identity(path)
        return path

    def _check_owned(self, path):
        identity = _identity(path)
        if str(path) in self._owned and self._owned[str(path)] != identity:
            _fail("stale_owned_file")
        return identity

    def _copy(self, source, target, expected):
        if self._fingerprint(source) != expected:
            _fail("stale_database")
        self._check_owned(target)
        self._check_owned(source)
        copied = 0
        with source.open("rb") as src, target.open("wb") as dst:
            while data := src.read(1024 * 1024):
                copied += len(data)
                if copied > MAX_DB_BYTES:
                    _fail("database_limit")
                dst.write(data)
            dst.flush()
            os.fsync(dst.fileno())
        if self._fingerprint(source) != expected or self._fingerprint(target)[-1] != expected[-1]:
            _fail("verification_failed")

    def _audit(self, conn, operation, revision, first_absent):
        if conn.execute("SELECT count(*) FROM audit").fetchone()[0] >= MAX_AUDIT:
            _fail("audit_limit")
        count = conn.execute("SELECT count(*) FROM chunks").fetchone()[0]
        conn.execute("UPDATE metadata SET revision=?,uncertain=0", (revision,))
        conn.execute("INSERT INTO audit(operation,revision,chunk_count,first_absent_restored_empty) VALUES(?,?,?,?)", (operation, revision, count, int(first_absent)))
        conn.commit()

    def _prepare_audit(self, conn):
        if conn.execute("SELECT count(*) FROM audit").fetchone()[0] >= MAX_AUDIT:
            _fail("audit_limit")
        conn.execute("UPDATE metadata SET uncertain=1")
        conn.commit()

    def _refresh_baseline(self):
        self._bound = False
        self._baseline = self._fingerprint(self._db)
        self._bound = True

    def _finish_audit(self, operation, revision, first_absent):
        try:
            with self._connection(self._db, create=True) as conn:
                self._audit(conn, operation, revision, first_absent)
            self._refresh_baseline()
            result = self._verify(self._db)
            if result["uncertain"] or result["revision"] != revision:
                _fail("verification_failed")
            return result
        except Exception:
            # Preserve the persistent fuse even if audit committed but final reopen failed.
            # Never touch content here; an unsafe/missing root still refuses this write.
            try:
                with self._connection(self._db, create=True) as conn:
                    conn.execute("UPDATE metadata SET uncertain=1")
                    conn.commit()
            finally:
                self._refresh_baseline()
            raise

    def _verify(self, path):
        with self._connection(path) as conn:
            return self._metadata(conn, verify=True)

    def _publish(self, stage, expected):
        self._guard()
        self._check_owned(stage)
        stage_fingerprint = self._fingerprint(stage)
        if self._fingerprint(self._db) != expected:
            _fail("stale_database")
        previous_fuse = self._fused
        self._fused = True  # Set BEFORE a potentially ambiguous publication failure.
        try:
            os.replace(stage, self._db)
        except Exception:
            self._bound = False
            current = self._fingerprint(self._db)
            self._baseline = current
            self._bound = True
            if current == expected:
                self._fused = previous_fuse  # Proved no database change, not uncertainty.
            raise
        self._refresh_baseline()
        if self._baseline[-1] != stage_fingerprint[-1]:
            _fail("verification_failed")
        return self._verify(self._db)

    @_safe
    def confirm_replace(self, plan_id, session_id, confirmed):
        with self._lock:
            plan = self._consume(plan_id, session_id, confirmed, {"replace", "delete"})
            self._prune()
            if len(self._saves) >= MAX_AUTHORITIES:
                _fail("authority_limit")
            backup = self._owned_file(".backup")
            stage = None
            save_id = None
            try:
                if plan["fingerprint"] is None:
                    with self._connection(backup, create=True) as conn:
                        self._initialize(conn)
                else:
                    self._copy(self._db, backup, plan["fingerprint"])
                backup_fingerprint = self._fingerprint(backup)
                self._verify(backup)
                stage = self._owned_file(".stage")
                self._copy(backup, stage, backup_fingerprint)
                with self._connection(stage, create=True) as conn:
                    documents = plan.get('documents')
                    if documents is None:
                        documents = {plan['doc_id']: (plan.get('source_fingerprint'), plan.get('records', ()))}
                    affected = set(documents) | set(plan.get('removals', ()))
                    for doc_id in affected:
                        ids = conn.execute('SELECT id FROM chunks WHERE doc_id=?', (doc_id,)).fetchall()
                        conn.executemany('DELETE FROM vectors WHERE rowid=?', ids)
                        conn.execute('DELETE FROM chunks WHERE doc_id=?', (doc_id,))
                    for doc_id, (version, records) in documents.items():
                        for source, location, text, kind, embedding in records:
                            cursor = conn.execute("INSERT INTO chunks(doc_id,source_fingerprint,source,location,text,kind) VALUES(?,?,?,?,?,?)",
                                (doc_id, version, source, location, text, kind))
                            conn.execute("INSERT INTO vectors(rowid,embedding,kind) VALUES(?,?,?)", (cursor.lastrowid, embedding, kind))
                    conn.commit()
                self._verify(stage)
                with self._connection(stage, create=True) as conn:
                    self._prepare_audit(conn)  # persistent fuse, not a success audit
                self._verify(stage)
                save_id = uuid.uuid4().hex
                save = dict(backup=backup, backup_fingerprint=backup_fingerprint, session=session_id,
                    expiry=time.monotonic() + AUTHORITY_SECONDS, first_absent=plan["fingerprint"] is None,
                    fingerprint=None, revision=plan["revision"] + 1)
                self._saves[save_id] = save
                try:
                    self._publish(stage, plan["fingerprint"])
                    metadata = self._finish_audit(plan["operation"], plan["revision"] + 1, False)
                finally:
                    if self._fused:
                        save["fingerprint"] = self._fingerprint(self._db)
                self._fused = False
                return dict(metadata, save_id=save_id, first_absent_restored_empty=save["first_absent"])
            except Exception as exc:
                if self._fused:
                    failure = VectorStoreError("uncertain_state")
                    failure.save_id = save_id
                    raise failure from None
                if save_id is not None:
                    self._saves.pop(save_id, None)
                if isinstance(exc, VectorStoreError):
                    raise
                raise VectorStoreError("write_failed") from None
            finally:
                if stage is not None and os.path.lexists(stage):
                    self._guard()
                    self._check_owned(stage)
                    stage.unlink()  # only this invocation's exclusive owned stage
                if stage is not None:
                    self._owned.pop(str(stage), None)
                if not any(s["backup"] == backup for s in self._saves.values()):
                    self._owned.pop(str(backup), None)

    @_safe
    def plan_rollback(self, save_id, session_id):
        with self._lock:
            self._prune()
            save = self._saves.get(save_id) if type(save_id) is str else None
            if save is None:
                _fail("invalid_authority")
            if save["session"] != session_id:
                _fail("session_mismatch")
            if self._fingerprint(self._db) != save["fingerprint"] or self._fingerprint(save["backup"]) != save["backup_fingerprint"]:
                _fail("stale_database")
            return self._plan("rollback", session_id, save_id=save_id, first_absent=save["first_absent"])

    @_safe
    def confirm_rollback(self, rollback_id, session_id, confirmed):
        with self._lock:
            plan = self._consume(rollback_id, session_id, confirmed, {"rollback"})
            save = self._saves.get(plan["save_id"])
            if save is None or save["expiry"] <= time.monotonic():
                _fail("invalid_authority")
            if self._fingerprint(save["backup"]) != save["backup_fingerprint"] or self._fingerprint(self._db) != save["fingerprint"]:
                _fail("stale_database")
            stage = self._owned_file(".stage")
            try:
                self._copy(save["backup"], stage, save["backup_fingerprint"])
                with self._connection(self._db) as conn:
                    audit = conn.execute("SELECT operation,revision,chunk_count,first_absent_restored_empty FROM audit ORDER BY seq LIMIT ?", (MAX_AUDIT + 1,)).fetchall()
                with self._connection(stage, create=True) as conn:
                    conn.execute("DELETE FROM audit")
                    conn.executemany("INSERT INTO audit(operation,revision,chunk_count,first_absent_restored_empty) VALUES(?,?,?,?)", audit)
                    conn.execute("UPDATE metadata SET revision=?", (plan["revision"],))
                    conn.commit()
                self._verify(stage)
                with self._connection(stage, create=True) as conn:
                    self._prepare_audit(conn)
                self._verify(stage)
                self._publish(stage, plan["fingerprint"])
                result = self._finish_audit("rollback", plan["revision"] + 1, save["first_absent"])
                self._fused = False
                self._saves.pop(plan["save_id"], None)
                return dict(result, first_absent_restored_empty=save["first_absent"])
            except Exception as exc:
                if self._fused:
                    save["fingerprint"] = self._fingerprint(self._db)
                    failure = VectorStoreError("uncertain_state")
                    failure.save_id = plan["save_id"]
                    raise failure from None
                if isinstance(exc, VectorStoreError):
                    raise
                raise VectorStoreError("write_failed") from None
            finally:
                if os.path.lexists(stage):
                    self._guard()
                    self._check_owned(stage)
                    stage.unlink()
                self._owned.pop(str(stage), None)
                self._owned.pop(str(save["backup"]), None)

    @_safe
    def query(self, embedding, *, keyword="", top_k=5, kind="document"):
        with self._lock:
            vector = self._vector(embedding)
            _string(keyword, 200, empty=True)
            if type(top_k) is not int or not 1 <= top_k <= 10 or kind not in ("document", "memory"):
                _fail("invalid_input")
            if self._fused:
                _fail("uncertain_state")
            fingerprint = self._fingerprint(self._db)
            if fingerprint is None:
                return []
            with self._connection(self._db) as conn:
                self._metadata(conn, verify=True)
                semantic = conn.execute("SELECT rowid,distance FROM vectors WHERE embedding MATCH ? AND k=? AND kind=? ORDER BY distance", (vector, top_k, kind)).fetchall()
                keyword_ids = []
                if keyword.strip():
                    keyword_ids = [r[0] for r in conn.execute("SELECT id FROM chunks WHERE kind=? AND instr(lower(text),lower(?))>0 ORDER BY id LIMIT 100", (kind, keyword.strip()))]
                distances = dict(semantic)
                scores = {}
                for ranking in ([r[0] for r in semantic], keyword_ids):
                    for rank, chunk_id in enumerate(ranking, 1):
                        scores[chunk_id] = scores.get(chunk_id, 0) + 1 / (60 + rank)
                result = []
                for rank, chunk_id in enumerate(sorted(scores, key=lambda i: (-scores[i], i))[:top_k], 1):
                    row = conn.execute("SELECT doc_id,source,location,text,source_fingerprint FROM chunks WHERE id=? AND kind=?", (chunk_id, kind)).fetchone()
                    if row is None:
                        _fail("verification_failed")
                    distance = distances.get(chunk_id)
                    if distance is None:
                        distance = conn.execute("SELECT vec_distance_cosine(embedding,?) FROM vectors WHERE rowid=?", (vector, chunk_id)).fetchone()[0]
                    if not math.isfinite(distance):
                        _fail("verification_failed")
                    result.append(dict(id=chunk_id, doc_id=row[0], source=row[1], location=row[2], text=row[3], excerpt=row[3][:200],
                        model=self.model, source_fingerprint=row[4], distance=distance, rank=rank))
            if self._fingerprint(self._db) != fingerprint:
                _fail("stale_database")
            return result

    @_safe
    def research_profile(self, doc_id, source_fingerprint, *, chunk_version, tokenizer_version, parser_version):
        import json

        with self._lock:
            for value, bound in ((doc_id, 128), (source_fingerprint, 256), (chunk_version, 64),
                                 (tokenizer_version, 64), (parser_version, 64)):
                _string(value, bound)
            if self._fused:
                _fail('uncertain_state')
            fingerprint = self._fingerprint(self._db)
            if fingerprint is None:
                return {'status': 'missing', 'chunk_count': 0, 'revision': 0}
            with self._connection(self._db) as conn:
                metadata = self._metadata(conn, verify=True)
                if metadata['uncertain']:
                    self._fused = True
                    _fail('uncertain_state')
                rows = conn.execute("SELECT source_fingerprint,location FROM chunks WHERE doc_id=? AND kind='document'", (doc_id,)).fetchall()
                status = 'current' if rows else 'missing'
                for version, location in rows:
                    if version != source_fingerprint:
                        status = 'stale'
                        break
                    try:
                        profile = json.loads(location)
                        matches = (type(profile) is dict and profile.get('chunker') == chunk_version and
                                   profile.get('tokenizer') == tokenizer_version and profile.get('parser') == parser_version)
                    except (ValueError, TypeError):
                        matches = False
                    if not matches:
                        status = 'profile_mismatch'
                result = {'status': status, 'chunk_count': len(rows), 'revision': metadata['revision']}
            if self._fingerprint(self._db) != fingerprint:
                _fail('stale_database')
            return result

    @_safe
    def query_scoped(self, embedding, *, keyword, sources, candidates=50, top_k=5, chunk_version, candidate_pool=False, ranking='hybrid'):
        import json
        from collections import Counter
        from .hybrid_ranking import rankings, fuse, terms, VERSION

        with self._lock:
            vector = self._vector(embedding)
            _string(keyword, 200)
            _string(chunk_version, 64)
            if (type(sources) is not dict or not 1 <= len(sources) <= 20 or
                    type(candidates) is not int or not 10 <= candidates <= 100 or
                    type(candidate_pool) is not bool or ranking not in {'hybrid', 'dense', 'lexical'} or
                    type(top_k) is not int or not 1 <= top_k <= (100 if candidate_pool else 10) or candidates < top_k):
                _fail('invalid_input')
            admitted = []
            for doc_id, source_version in sources.items():
                _string(doc_id, 128)
                _string(source_version, 256)
                admitted.extend((doc_id, source_version))
            if self._fused:
                _fail('uncertain_state')
            fingerprint = self._fingerprint(self._db)
            if fingerprint is None:
                return []
            with self._connection(self._db) as conn:
                metadata = self._metadata(conn, verify=True)
                if metadata['uncertain']:
                    self._fused = True
                    _fail('uncertain_state')
                for doc_id, source_version in sources.items():
                    versions = conn.execute("SELECT DISTINCT source_fingerprint FROM chunks WHERE doc_id=? AND kind='document'", (doc_id,)).fetchall()
                    if not versions:
                        _fail('document_missing')
                    if versions != [(source_version,)]:
                        _fail('stale_source')
                predicate = ' OR '.join('(c.doc_id=? AND c.source_fingerprint=?)' for _ in sources)
                rows = conn.execute('SELECT c.id,c.doc_id,c.source,c.location,c.text,c.source_fingerprint,'
                                    'vec_distance_cosine(v.embedding,?) FROM chunks c JOIN vectors v ON c.id=v.rowid '
                                    "WHERE c.kind='document' AND v.kind='document' AND (" + predicate + ') ORDER BY c.id',
                                    (vector, *admitted)).fetchall()
                records = []
                for row in rows:
                    location = json.loads(row[3])
                    if location.get('chunker') != chunk_version:
                        _fail('index_profile_mismatch')
                    if not math.isfinite(row[6]):
                        _fail('verification_failed')
                    records.append(dict(id=row[0], doc_id=row[1], source=row[2], location=row[3], text=row[4],
                                        source_fingerprint=row[5], distance=row[6], model=self.model, excerpt=row[4][:200]))
                cache_key = (fingerprint, tuple(sorted(sources.items())), VERSION)
                if self._lexical_cache is None or self._lexical_cache[0] != cache_key:
                    self._lexical_cache = (cache_key, {row['id']: Counter(terms(row['text'])) for row in records})
                lists = rankings(records, keyword, candidates, term_counts=self._lexical_cache[1])
                selected = fuse(lists if ranking == 'hybrid' else [lists[2]] if ranking == 'dense' else lists[:2], top_k)
                by_id = {row['id']: row for row in records}
                result = [dict(by_id[key], rank=rank) for rank, key in enumerate(selected, 1)]
            if self._fingerprint(self._db) != fingerprint:
                _fail('stale_database')
            return result

    @_safe
    def read_evidence(self, chunk_id, *, offset=0, limit=2000, source_fingerprint: str):
        """Host-only bounded reread, with mandatory stale-source identity supplied by host."""
        with self._lock:
            if type(chunk_id) is not int or chunk_id < 1 or type(offset) is not int or not 0 <= offset <= MAX_TEXT or type(limit) is not int or not 1 <= limit <= 2000:
                _fail("invalid_input")
            _string(source_fingerprint, 256)
            if self._fused:
                _fail("uncertain_state")
            fingerprint = self._fingerprint(self._db)
            if fingerprint is None:
                _fail("evidence_missing")
            with self._connection(self._db) as conn:
                self._metadata(conn, verify=True)
                row = conn.execute("SELECT text,source,location,source_fingerprint FROM chunks WHERE id=?", (chunk_id,)).fetchone()
            if self._fingerprint(self._db) != fingerprint:
                _fail("stale_database")
            if row is None:
                _fail("evidence_missing")
            if row[3] != source_fingerprint:
                _fail("stale_source")
            return {"id": chunk_id, "source": row[1], "location": row[2], "text": row[0][offset:offset + limit],
                    "offset": offset, "has_more": offset + limit < len(row[0]), "source_fingerprint": row[3], "model": self.model}

    def invalidate(self):
        with self._lock:
            self._lexical_cache = None
            self._pending.clear()
            self._saves.clear()
            self._owned.clear()
            # Deliberately NOT resetting the uncertainty fuse or deleting backups.
            return {"uncertain": self._fused}
