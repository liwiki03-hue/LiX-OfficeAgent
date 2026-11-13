from __future__ import annotations

from copy import deepcopy
import hashlib
import json
import math
import os
from pathlib import Path
import re
import threading
import uuid

from pydantic import Field

from .auth import LocalAuth
from .memory import MemoryError, MemoryService
from .preference_memory import PreferenceMemory
from .writes import Strict


def local_embedding(text):
    features = re.findall(r'[a-z0-9]+|[\u4e00-\u9fff]', text.lower())
    features += [a+b for a, b in zip(features, features[1:])]
    vector = [0.0]*64
    for feature in features:
        index = int.from_bytes(hashlib.sha256(feature.encode('utf-8')).digest()[:2], 'big') % 64
        vector[index] += 1.0
    if not any(vector):
        vector[0] = 1.0
    norm = math.sqrt(sum(value*value for value in vector))
    return [value/norm for value in vector]


class LocalMemoryBackend:
    _guard = threading.Lock()
    _locks = {}

    def __init__(self, root):
        absolute = os.path.abspath(root)
        LocalAuth._check_path(Path(absolute))
        if os.name == 'nt' and not absolute.startswith('\\\\?\\'):
            absolute = '\\\\?\\UNC\\'+absolute[2:] if absolute.startswith('\\\\') else '\\\\?\\'+absolute
        self.root = Path(absolute)/'general-memory'
        with self._guard:
            self.lock = self._locks.setdefault(os.path.normcase(str(self.root)), threading.RLock())
        self.plans = {}
        self.uncertain = set()
        self.seen = {}

    def _path(self, account, kind):
        if kind != 'memory' or type(account) is not str or not re.fullmatch(r'[a-f0-9]{64}', account):
            raise MemoryError('memory_auth_required')
        path = self.root/account/'state.json'
        LocalAuth._check_path(path)
        return path

    def _read(self, account, kind, *, verify_audit=True):
        path = self._path(account, kind)
        if path.exists():
            payload = PreferenceMemory._read_bytes(path)
            def unique(pairs):
                result = {}
                for key, value in pairs:
                    if key in result:
                        raise MemoryError('memory_backend_failed')
                    result[key] = value
                return result
            state = json.loads(payload, object_pairs_hook=unique)
            if (type(state) is not dict or set(state) != {'revision', 'records', 'transaction_id'} or type(state['transaction_id']) is not str or not re.fullmatch('[a-f0-9]{32}',state['transaction_id']) or type(state['revision']) is not int
                    or state['revision'] < 1 or type(state['records']) is not list or len(state['records']) > 20):
                raise MemoryError('memory_backend_failed')
            identifiers = set()
            for record in state['records']:
                if type(record) is not dict or 'embedding' not in record:
                    raise MemoryError('memory_backend_failed')
                MemoryService._record({k:v for k,v in record.items() if k != 'embedding'}, content=True)
                MemoryService._vector(record['embedding'])
                if len(record['embedding']) != 64 or record['memory_id'] in identifiers:
                    raise MemoryError('memory_backend_failed')
                identifiers.add(record['memory_id'])
            if verify_audit:
                try:
                    audit = json.loads(PreferenceMemory._read_bytes(path.parent/state['transaction_id']/'audit.json'))
                    if audit.get('status') != 'verified' or audit.get('revision') != state['revision'] or audit.get('after_sha256') != hashlib.sha256(payload).hexdigest():
                        self.uncertain.add(account)
                except Exception:
                    self.uncertain.add(account)
        else:
            payload, state = b'', {'revision': 0, 'records': [], 'transaction_id':None}
        fingerprint = hashlib.sha256(payload).hexdigest()
        previous = self.seen.get(account)
        if previous and ((previous[0] == state['revision'] and previous[1] != fingerprint) or previous[0] > state['revision']):
            self.uncertain.add(account)
        self.seen[account] = (state['revision'], fingerprint)
        return payload, state

    def snapshot(self, *, account, kind):
        with self.lock:
            _, state = self._read(account, kind)
            return {'revision': state['revision'], 'uncertain': account in self.uncertain}

    def plan(self, *, account, kind, operation, expected, records, memory_ids):
        with self.lock:
            payload, state = self._read(account, kind)
            if self.snapshot(account=account, kind=kind) != expected or expected['uncertain'] or operation not in {'remember', 'delete', 'clear'}:
                raise MemoryError('memory_stale')
            if operation == 'remember':
                if len(records) != 1 or memory_ids or len(state['records']) >= 20:
                    raise MemoryError('memory_capacity')
                record = records[0]
                MemoryService._record({k:v for k,v in record.items() if k != 'embedding'}, content=True)
                MemoryService._vector(record['embedding'])
                if len(record['embedding']) != 64 or any(r['memory_id'] == record['memory_id'] for r in state['records']):
                    raise MemoryError('memory_invalid')
                state['records'].append(deepcopy(record))
            elif operation == 'delete':
                if records or len(memory_ids) != 1 or not any(r['memory_id'] == memory_ids[0] for r in state['records']):
                    raise MemoryError('memory_invalid')
                state['records'] = [r for r in state['records'] if r['memory_id'] not in memory_ids]
            else:
                if records or memory_ids:
                    raise MemoryError('memory_invalid')
                state['records'] = []
            state['revision'] += 1
            identifier = uuid.uuid4().hex
            if len(self.plans) >= 32:
                raise MemoryError('memory_capacity')
            self.plans[identifier] = (account, operation, payload, state)
            return {'plan_id': identifier}

    def _apply(self, account, operation, old, state):
        path = self._path(account, 'memory')
        current, _ = self._read(account, 'memory')
        if current != old or account in self.uncertain:
            return {'status': 'rejected', 'applied': False, 'uncertain': False}
        save_id = uuid.uuid4().hex
        state['transaction_id'] = save_id
        if type(state) is not dict or set(state) != {'revision','records','transaction_id'} or type(state['revision']) is not int or state['revision'] < 1 or type(state['records']) is not list or len(state['records']) > 20:
            raise MemoryError('memory_backend_failed')
        identifiers = set()
        for record in state['records']:
            MemoryService._record({key:value for key,value in record.items() if key != 'embedding'},content=True)
            MemoryService._vector(record.get('embedding'))
            if len(record['embedding']) != 64 or record['memory_id'] in identifiers:
                raise MemoryError('memory_backend_failed')
            identifiers.add(record['memory_id'])
        payload = json.dumps(state, ensure_ascii=False, separators=(',', ':'), allow_nan=False).encode('utf-8')
        if len(payload) > 65536:
            raise MemoryError('memory_capacity')
        job = path.parent/save_id
        touched = False
        try:
            job.mkdir(parents=True, exist_ok=False)
            LocalAuth._check_path(job)
            PreferenceMemory._exclusive(job/'backup.json', old or b'{}')
            if PreferenceMemory._read_bytes(job/'backup.json') != (old or b'{}'):
                raise MemoryError('memory_backend_failed')
            working = job/'copy.json'
            PreferenceMemory._exclusive(working, payload)
            if PreferenceMemory._read_bytes(working) != payload or json.loads(payload) != state:
                raise MemoryError('memory_backend_failed')
            audit = {'operation': operation, 'revision': state['revision'], 'source_present': bool(old),
                     'before_sha256': hashlib.sha256(old).hexdigest(), 'after_sha256': hashlib.sha256(payload).hexdigest(),
                     'independent_human_confirmed': True, 'status': 'prepared'}
            PreferenceMemory._exclusive(job/'prepared.json', json.dumps(audit).encode('utf-8'))
            current, _ = self._read(account, 'memory')
            if current != old or account in self.uncertain:
                raise MemoryError('memory_stale')
            LocalAuth._check_path(path)
            touched = True
            os.replace(working, path)
            self.seen[account] = (state['revision'], hashlib.sha256(payload).hexdigest())
            reopened, verified = self._read(account, 'memory', verify_audit=False)
            if reopened != payload or verified != state:
                raise MemoryError('memory_backend_failed')
            audit['status'] = 'verified'
            PreferenceMemory._exclusive(job/'audit.json', json.dumps(audit).encode('utf-8'))
            if json.loads(PreferenceMemory._read_bytes(job/'audit.json')) != audit:
                raise MemoryError('memory_backend_failed')
            return {'status': {'remember':'saved','delete':'deleted','clear':'cleared','rollback':'rolled_back'}[operation],
                    'applied': True, 'uncertain': False, 'save_id': save_id}
        except Exception:
            if touched:
                self.uncertain.add(account)
                return {'status':'uncertain','applied':None,'uncertain':True}
            raise MemoryError('memory_backend_failed') from None

    def confirm(self, plan_id, *, account, kind, confirmed):
        with self.lock:
            self._path(account, kind)
            saved = self.plans.get(plan_id)
            if confirmed is not True or not saved or saved[0] != account:
                return {'status':'rejected','applied':False,'uncertain':False}
            self.plans.pop(plan_id)
            _, operation, old, state = saved
            return self._apply(account, operation, old, state)

    def rollback(self, save_id, *, account, kind, expected, confirmed):
        with self.lock:
            path = self._path(account, kind)
            if confirmed is not True or type(save_id) is not str or not re.fullmatch('[a-f0-9]{32}', save_id):
                return {'status':'rejected','applied':False,'uncertain':False}
            if self.snapshot(account=account, kind=kind) != expected or expected['uncertain']:
                return {'status':'rejected','applied':False,'uncertain':False}
            job = path.parent/save_id
            audit = json.loads(PreferenceMemory._read_bytes(job/'audit.json'))
            backup = PreferenceMemory._read_bytes(job/'backup.json')
            old, state = self._read(account, kind)
            if (audit.get('status') != 'verified' or audit.get('revision') != state['revision']
                    or hashlib.sha256(old).hexdigest() != audit.get('after_sha256')
                    or hashlib.sha256(backup if audit['source_present'] else b'').hexdigest() != audit.get('before_sha256')):
                return {'status':'rejected','applied':False,'uncertain':False}
            if not audit['source_present'] and backup != b'{}':
                return {'status':'rejected','applied':False,'uncertain':False}
            restored = json.loads(backup) if audit['source_present'] else {'records': []}
            restored['revision'] = state['revision']+1
            return self._apply(account, 'rollback', old, restored)

    def catalog(self, *, account, kind, limit, offset):
        with self.lock:
            _, state = self._read(account, kind)
            return [{k:v for k,v in record.items() if k not in {'content','embedding'}} for record in state['records'][offset:offset+limit]]

    def read(self, memory_id, *, account, kind):
        with self.lock:
            _, state = self._read(account, kind)
            record = next(record for record in state['records'] if record['memory_id'] == memory_id)
            return {k:v for k,v in record.items() if k != 'embedding'}

    def retrieve(self, vector, *, account, kind, categories, top):
        with self.lock:
            _, state = self._read(account, kind)
            ranked = sorted((record for record in state['records'] if record['category'] in categories),
                            key=lambda record: (-sum(a*b for a,b in zip(vector, record['embedding'])), record['memory_id']))
            return [{k:v for k,v in record.items() if k != 'embedding'} for record in ranked[:top]]


class RememberArgs(Strict):
    content: str = Field(min_length=1, max_length=1000)
    category: str = Field(pattern='^(task|preference|summary)$')


class GeneralMemory:
    def __init__(self, root):
        self.service = MemoryService(LocalMemoryBackend(root), local_embedding)
        self.session = None
        self.account = None
        self.operations = {}

    def invalidate(self):
        if self.session is not None:
            self.service.revoke_session(self.session)
        self.session = self.account = None
        self.operations.clear()

    def bind(self, account):
        if not isinstance(account, str) or not re.fullmatch('[a-f0-9]{64}', account):
            raise MemoryError('memory_auth_required')
        if self.account != account or self.session is None:
            self.invalidate()
            self.session = self.service.open_session(account, authenticated=True)
            self.account = account
        return self.session

    def snapshot(self, account):
        session = self.bind(account)
        catalog = self.service.list_memories(session)
        return {'items':[self.service.read_memory(record['memory_id'], session) for record in catalog],
                'limit':20, 'recall_limit':3, 'embedding':'local_hash_64_v1', 'backup_deletion':'logical_delete_retains_transaction_backups'}

    def plan(self, account, operation, value=None):
        session = self.bind(account)
        if operation == 'remember':
            args = RememberArgs.model_validate(value)
            result = self.service.plan_remember(args.content, args.category, session, explicit=True)
        elif operation == 'delete':
            result = self.service.plan_delete(value, session)
        elif operation == 'clear' and value is None:
            result = self.service.plan_clear(session)
        else:
            raise MemoryError('memory_invalid')
        if 'plan_id' in result:
            self.operations[result['plan_id']] = operation
        return result

    def confirm(self, account, identifier, confirmed):
        session = self.bind(account)
        operation = self.operations.pop(identifier, None)
        if operation is None:
            raise MemoryError('memory_plan_invalid')
        return getattr(self.service, 'confirm_'+operation)(identifier, session, confirmed=confirmed)

    def recall(self, account, question):
        return self.service.model_context(self.service.recall(question[:200], self.bind(account)))
