"""Independently approved immutable retrieval evidence; no providers or client paths.

Full hit text is retained locally before any model excerpt is returned. Approval
and rollback grants are host-session-bound and memory-only; OS races/crash
transactions are not claimed. Unknown post-admission outcomes fuse the instance.
"""
from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import re
import secrets
import stat
import threading
import time

from .rag import SECRET

MAX_BYTES = 512 * 1024
MAX_PENDING = 8
MAX_IDS = 256
TTL = 300
CODES = frozenset({'evidence_invalid', 'evidence_sensitive', 'evidence_source_changed',
                   'evidence_path_changed', 'evidence_plan_invalid', 'evidence_session_invalid',
                   'evidence_capacity', 'evidence_expired', 'evidence_denied',
                   'evidence_uncertain', 'evidence_io_failed', 'evidence_verification_failed'})


class EvidenceError(ValueError):
    def __init__(self, code='evidence_invalid'):
        self.code = code if type(code) is str and code in CODES else 'evidence_invalid'
        super().__init__(self.code)


def _fail(code='evidence_invalid'):
    raise EvidenceError(code)


def _text(value, maximum, *, empty=False):
    if type(value) is not str or len(value) > maximum or (not empty and not value.strip()) or '\x00' in value:
        _fail()
    if SECRET.search(value):
        _fail('evidence_sensitive')
    try:
        value.encode('utf-8')
    except UnicodeError:
        _fail()
    return value


def _hash(data):
    return hashlib.sha256(data).hexdigest()


def _json(value):
    data = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'), allow_nan=False).encode('utf-8')
    if len(data) > MAX_BYTES:
        _fail('evidence_capacity')
    return data


def _dir(path):
    identity = None
    for item in reversed((path, *path.parents)):
        info = item.lstat()
        if (not stat.S_ISDIR(info.st_mode) or item.is_symlink() or
                getattr(info, 'st_file_attributes', 0) & 0x400):
            _fail('evidence_path_changed')
        if item == path:
            identity = (info.st_dev, info.st_ino)
    return identity


def _stamp(info):
    return (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_nlink)


def _read(path):
    parent = _dir(path.parent)
    before = path.lstat()
    if (not stat.S_ISREG(before.st_mode) or before.st_nlink != 1 or
            getattr(before, 'st_file_attributes', 0) & 0x400 or before.st_size > MAX_BYTES):
        _fail('evidence_path_changed')
    with path.open('rb') as handle:
        first = os.fstat(handle.fileno())
        if _stamp(first) != _stamp(before):
            _fail('evidence_path_changed')
        data = handle.read(MAX_BYTES + 1)
        last = os.fstat(handle.fileno())
        if _stamp(last) != _stamp(first) or last.st_ctime_ns != first.st_ctime_ns:
            _fail('evidence_path_changed')
    after = path.lstat()
    if (_stamp(after) != _stamp(before) or after.st_ctime_ns != before.st_ctime_ns or
            len(data) != before.st_size or _dir(path.parent) != parent):
        _fail('evidence_path_changed')
    return data, (_stamp(after), after.st_ctime_ns, parent, _hash(data))


class RagEvidence:
    artifact_name = 'evidence.json'
    directory_prefix = '.office-rag-evidence-'
    error_type = EvidenceError
    error_codes = CODES

    def __init__(self, workspace, source_version, *, clock=time.monotonic):
        self.root = Path(workspace.root)
        self._version = source_version
        self._clock = clock
        self._lock = threading.RLock()
        self._plans = {}
        self._saved = {}
        self._issued = set()
        self._fused = False
        try:
            self._identity = _dir(self.root)
        except Exception:
            raise EvidenceError('evidence_path_changed') from None

    def __repr__(self):
        return '<RagEvidence redacted>'

    def _guard(self):
        if self._fused:
            _fail('evidence_uncertain')
        if _dir(self.root) != self._identity:
            _fail('evidence_path_changed')

    def _session(self, session):
        if type(session) is not str or not re.fullmatch('[a-f0-9]{32}', session):
            _fail('evidence_session_invalid')

    def _fresh(self, hits):
        versions = {}
        for hit in hits:
            previous = versions.setdefault(hit['doc_id'], hit['source_fingerprint'])
            if previous != hit['source_fingerprint']:
                _fail('evidence_source_changed')
        for doc_id, fingerprint in versions.items():
            if self._version(doc_id) != fingerprint:
                _fail('evidence_source_changed')

    def ensure_available(self, session):
        """Host preflight before provider reads; no new authority and no disk writes."""
        with self._lock:
            try:
                self._session(session); self._guard()
                self._plans = {key: value for key, value in self._plans.items() if self._clock() - value['time'] < TTL}
                if len(self._plans) >= MAX_PENDING or len(self._saved) >= MAX_PENDING or len(self._issued) >= MAX_IDS:
                    _fail('evidence_capacity')
            except EvidenceError:
                raise
            except Exception:
                raise EvidenceError('evidence_io_failed') from None

    def _hits(self, hits):
        if type(hits) is not list or len(hits) > 10:
            _fail()
        checked, ids = [], set()
        fields = {'id', 'doc_id', 'source', 'location', 'text', 'model',
                  'source_fingerprint', 'distance', 'rank', 'excerpt'}
        for rank, hit in enumerate(hits, 1):
            if type(hit) is not dict or set(hit) != fields:
                _fail()
            if (type(hit['id']) is not int or hit['id'] < 1 or hit['id'] in ids or
                    type(hit['rank']) is not int or hit['rank'] != rank or
                    type(hit['distance']) not in (int, float) or not math.isfinite(hit['distance']) or
                    not -1e-6 <= hit['distance'] <= 2.000001):
                _fail()
            ids.add(hit['id'])
            item = dict(hit)
            for key, maximum in (('doc_id', 128), ('source', 128), ('location', 512), ('text', 8000), ('model', 256)):
                _text(item[key], maximum)
            if (type(item['source_fingerprint']) is not str or
                    not re.fullmatch('[a-f0-9]{64}', item['source_fingerprint']) or
                    type(item['excerpt']) is not str or item['excerpt'] != item['text'][:200]):
                _fail()
            if item['source'].replace('\\', '/').split('/')[-1].lower().startswith('.env'):
                _fail('evidence_sensitive')
            checked.append(item)
        return checked

    def _id(self):
        self._plans = {key: value for key, value in self._plans.items() if self._clock() - value['time'] < TTL}
        if len(self._plans) >= MAX_PENDING or len(self._issued) >= MAX_IDS:
            _fail('evidence_capacity')
        for _ in range(4):
            value = secrets.token_hex(16)
            if type(value) is str and re.fullmatch('[a-f0-9]{32}', value) and value not in self._issued:
                self._issued.add(value)
                return value
        _fail('evidence_capacity')

    def _contexts(self, contexts, hits):
        from .research_context import CONTEXT_VERSION, CONTEXT_BUDGET
        from .token_chunks import TOKENIZER_VERSION, token_count

        if type(contexts) is not list or len(contexts) != len(hits):
            _fail()
        for context, hit in zip(contexts, hits):
            if (type(context) is not dict or set(context) != {'hit_id', 'source', 'source_fingerprint', 'retrieved_location',
                    'context_version', 'tokenizer', 'token_budget', 'tokens', 'sections', 'writable'} or
                    type(context['hit_id']) is not int or context['hit_id'] != hit['id'] or context['source'] != hit['source'] or
                    context['source_fingerprint'] != hit['source_fingerprint'] or context['retrieved_location'] != hit['location'] or
                    context['context_version'] != CONTEXT_VERSION or context['tokenizer'] != TOKENIZER_VERSION or
                    type(context['token_budget']) is not int or context['token_budget'] != CONTEXT_BUDGET or
                    context['writable'] is not False or type(context['sections']) is not list or not 1 <= len(context['sections']) <= 4):
                _fail()
            total = 0
            for section in context['sections']:
                if (type(section) is not dict or set(section) != {'role', 'location', 'char_start', 'char_end', 'text', 'tokens', 'truncated'} or
                        section['role'] not in {'parent', 'heading', 'table_first_row', 'adjacent'} or
                        type(section['char_start']) is not int or type(section['char_end']) is not int or
                        not 0 <= section['char_start'] < section['char_end'] <= 8000 or type(section['truncated']) is not bool):
                    _fail()
                _text(section['text'], 8000)
                _text(section['location'], 128)
                actual = token_count(section['text'])
                if type(section['tokens']) is not int or section['tokens'] != actual or len(section['text']) != section['char_end'] - section['char_start']:
                    _fail()
                total += actual
            if type(context['tokens']) is not int or context['tokens'] != total or total > CONTEXT_BUDGET:
                _fail()
        return json.loads(_json(contexts))

    def plan(self, query, hits, session, *, contexts=None, retrieval=None):
        with self._lock:
            try:
                self._session(session); self._guard()
                _text(query, 200)
                if len(self._saved) >= MAX_PENDING:
                    _fail('evidence_capacity')
                checked = self._hits(hits)
                self._fresh(checked)
                document = {'version': 1, 'query': query, 'hits': checked}
                if contexts is not None:
                    document['version'] = 2
                    document['contexts'] = self._contexts(contexts, checked)
                if retrieval is not None:
                    if type(retrieval) is not dict or len(_json(retrieval)) > 4096:
                        _fail()
                    document['retrieval'] = json.loads(_json(retrieval))
                payload = _json(document)
                identifier = self._id()
                self._plans[identifier] = {'session': session, 'time': self._clock(), 'operation': 'save',
                                           'payload': payload, 'hits': checked}
                return {'plan_id': identifier, 'operation': 'save_evidence', 'hit_count': len(checked),
                        'byte_count': len(payload), 'confirmation_required': True, 'wrote_nothing': True,
                        'payload_sha256': _hash(payload), 'expires_in_seconds': TTL}
            except EvidenceError:
                raise
            except Exception:
                raise EvidenceError('evidence_io_failed') from None

    def _consume(self, identifier, session, confirmed, operation):
        self._session(session)
        plan = self._plans.get(identifier) if type(identifier) is str else None
        if plan is None or plan['operation'] != operation:
            _fail('evidence_plan_invalid')
        if plan['session'] != session:
            _fail('evidence_session_invalid')
        self._plans.pop(identifier)
        if type(confirmed) is not bool or confirmed is not True:
            _fail('evidence_denied')
        if self._clock() - plan['time'] >= TTL:
            _fail('evidence_expired')
        self._guard()
        return plan

    def _exclusive(self, path, data):
        self._guard()
        _dir(path.parent)
        with path.open('xb') as handle:
            handle.write(data); handle.flush(); os.fsync(handle.fileno())
            identity = _stamp(os.fstat(handle.fileno()))
        result, stamp = _read(path)
        if result != data or stamp[0] != identity:
            _fail('evidence_verification_failed')
        return stamp

    def _same(self, path, stamp):
        if _read(path)[1] != stamp:
            _fail('evidence_path_changed')

    def _saved_check(self, saved):
        self._guard()
        if _dir(saved['directory']) != saved['identity']:
            _fail('evidence_path_changed')
        for name, stamp in saved['files'].items():
            self._same(saved['directory'] / name, stamp)

    def _rejected(self, exc, wrote, *, rollback=False):
        if wrote:
            self._fused = True
        return {'status': 'uncertain' if wrote else 'rejected',
                'code': 'evidence_uncertain' if wrote else (exc.code if isinstance(exc, EvidenceError) else 'evidence_io_failed'),
                'restored' if rollback else 'applied': None if wrote else False,
                'wrote_nothing': not wrote}

    def _save_audit(self, plan):
        return _json({'version': 1, 'operation': 'save_evidence', 'hit_count': len(plan['hits']),
                      'payload_sha256': _hash(plan['payload'])})

    def confirm(self, plan_id, session, confirmed):
        with self._lock:
            wrote = False
            try:
                plan = self._consume(plan_id, session, confirmed, 'save')
                if len(self._saved) >= MAX_PENDING:
                    _fail('evidence_capacity')
                self._fresh(plan['hits'])
                directory = self.root / (self.directory_prefix + plan_id)
                wrote = True
                directory.mkdir(mode=0o700)
                identity = _dir(directory)
                files = {'backup': self._exclusive(directory / 'backup', b'')}
                audit = self._save_audit(plan)
                files['audit-intent.json'] = self._exclusive(directory / 'audit-intent.json', audit)
                stage = self._exclusive(directory / 'stage', plan['payload'])
                self._same(directory / 'stage', stage)
                self._fresh(plan['hits']); self._guard()
                if _dir(directory) != identity or os.path.lexists(directory / self.artifact_name):
                    _fail('evidence_path_changed')
                os.replace(directory / 'stage', directory / self.artifact_name)
                data, stamp = _read(directory / self.artifact_name)
                if data != plan['payload'] or stamp[0][:2] != stage[0][:2]:
                    _fail('evidence_verification_failed')
                files[self.artifact_name] = stamp
                files['audit-verified.json'] = self._exclusive(directory / 'audit-verified.json', audit)
                saved = {'session': session, 'directory': directory, 'identity': identity,
                         'files': files, 'hits': plan['hits']}
                self._saved_check(saved); self._fresh(saved['hits'])
                self._saved[plan_id] = saved
                return {'status': 'saved', 'applied': True, 'verified': True, 'wrote_nothing': False,
                        'save_id': plan_id, 'evidence': str((directory / self.artifact_name).relative_to(self.root)),
                        'payload_sha256': _hash(data), 'hits': [{key: value for key, value in hit.items() if key != 'text'} for hit in plan['hits']]}
            except Exception as exc:
                return self._rejected(exc, wrote)

    def _saved_for(self, save_id, session):
        self._session(session)
        saved = self._saved.get(save_id) if type(save_id) is str else None
        if saved is None:
            _fail('evidence_plan_invalid')
        if saved['session'] != session:
            _fail('evidence_session_invalid')
        self._saved_check(saved)
        return saved

    def read(self, save_id, session, *, offset=0, limit=2000):
        with self._lock:
            try:
                if type(offset) is not int or not 0 <= offset <= MAX_BYTES or type(limit) is not int or not 1 <= limit <= 2000:
                    _fail()
                saved = self._saved_for(save_id, session); self._fresh(saved['hits'])
                data, _ = _read(saved['directory'] / self.artifact_name)
                text = data.decode('utf-8')
                return {'text': text[offset:offset + limit], 'offset': offset, 'has_more': offset + limit < len(text)}
            except EvidenceError:
                raise
            except Exception:
                raise EvidenceError('evidence_io_failed') from None

    def read_context(self, save_id, hit_id, session):
        with self._lock:
            try:
                if type(hit_id) is not int or hit_id < 1:
                    _fail()
                saved = self._saved_for(save_id, session)
                self._fresh(saved['hits'])
                data, _ = _read(saved['directory'] / self.artifact_name)
                document = json.loads(data)
                contexts = self._contexts(document.get('contexts'), saved['hits'])
                result = next((item for item in contexts if item['hit_id'] == hit_id), None)
                if result is None:
                    _fail()
                self._saved_check(saved)
                self._fresh(saved['hits'])
                return result
            except EvidenceError:
                raise
            except Exception:
                raise EvidenceError('evidence_io_failed') from None

    def plan_rollback(self, save_id, session):
        with self._lock:
            try:
                self._saved_for(save_id, session)
                identifier = self._id()
                self._plans[identifier] = {'session': session, 'time': self._clock(), 'operation': 'rollback', 'save_id': save_id}
                return {'rollback_id': identifier, 'operation': 'remove_evidence', 'confirmation_required': True,
                        'wrote_nothing': True, 'expires_in_seconds': TTL}
            except EvidenceError:
                raise
            except Exception:
                raise EvidenceError('evidence_io_failed') from None

    def confirm_rollback(self, rollback_id, session, confirmed):
        with self._lock:
            wrote = False
            try:
                plan = self._consume(rollback_id, session, confirmed, 'rollback')
                saved = self._saved_for(plan['save_id'], session)
                directory = saved['directory']
                wrote = True
                self._exclusive(directory / 'audit-remove-intent.json', _json({'version': 1, 'operation': 'remove_evidence'}))
                self._saved_check(saved)
                (directory / self.artifact_name).unlink()
                if os.path.lexists(directory / self.artifact_name) or _dir(directory) != saved['identity']:
                    _fail('evidence_verification_failed')
                self._exclusive(directory / 'audit-removed.json', _json({'version': 1, 'operation': 'remove_evidence', 'verified': True}))
                self._saved.pop(plan['save_id'])
                return {'status': 'rolled_back', 'restored': True, 'verified': True, 'wrote_nothing': False,
                        'backups_audits_retained': True}
            except Exception as exc:
                return self._rejected(exc, wrote, rollback=True)

    def invalidate(self):
        with self._lock:
            self._plans.clear(); self._saved.clear()
            # No file deletion, automatic rollback, or reset of sticky uncertainty.
