"""Account-bound, opt-in host-only memory with confirmed proposals and manual notes.

Only direct-human text can become a volatile, repeated proposal; manual notes
require their own confirmation. No document text, model output, embedding, remote
API or shared document index. The store is not encrypted for the OS account.
"""
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

from .auth import LocalAuth
from .memory import classify_candidate

_ALLOWED = re.compile(r'^[\u4e00-\u9fffA-Za-z0-9 .:\-]{2,120}$')
_PHRASE = re.compile(r'^(?:请|以后|始终|总是|默认|我希望|我偏好|我习惯)[ ]*(?:把文档|将文档|修改文档|输出|回答|使用|按|你担任|你作为|你是).{1,90}$')
_FORMAT = re.compile(r'(?i)格式|排版|字体|字号|标题|段落|表格|列表|markdown|word|excel|ppt|docx|xlsx|pptx')
_ROLE = re.compile(r'(?i)角色|担任|作为|你是|编辑|审校|翻译|律师|教师|顾问')
_ID = re.compile(r'^[0-9a-f]{32}$')
_KEY = re.compile(r'^[0-9a-f]{64}$')
MAX_BYTES = 65536
_MANUAL_KINDS = ('persona', 'work', 'personal', 'writing', 'skill')
_AUTO_KINDS = ('format', 'role', 'work', 'skill')
_WORK = re.compile(r'^(?:我负责|我正在负责|我的长期项目是).{2,100}$')
_WORKFLOW = re.compile(r'^(?:以后|今后|每次|总是).{2,100}(?:先.{1,50}再.{1,50}|按.{1,60}步骤.{1,30})$')
_SKIP_GUARDS = re.compile(r'(?i)跳过.{0,8}(确认|校验|审计|检查)|绕过|无需.{0,8}确认|'
                          r'禁用.{0,8}(安全|检查)|bypass|disable.{0,8}(safety|validation)')


class PreferenceError(ValueError):
    def __init__(self, code='memory_unavailable'):
        self.code = code
        super().__init__(code)


def weight_at(timestamp, now):
    """Full weight for 10 days then halves every 10 more days."""
    if not all(type(t) in (int, float) and math.isfinite(t) and t >= 0 for t in (timestamp, now)):
        raise PreferenceError('memory_invalid')
    age_days = max(0.0, now - timestamp) / 86400
    return 1.0 if age_days <= 10 else 0.5 ** ((age_days - 10) / 10)


def candidate(question):
    """Conservative exact one-sentence user preference; reject mixed task content."""
    if (type(question) is not str or '\n' in question or not _ALLOWED.fullmatch(question.strip())
            or not _PHRASE.fullmatch(question.strip()) or not classify_candidate(question)['allowed']):
        return None
    text = question.strip()
    kind = 'format' if _FORMAT.search(text) else 'role' if _ROLE.search(text) else None
    if kind == 'format' and _ROLE.search(text):
        return None
    return (kind, text) if kind else None


def derived_candidate(question):
    """Only short direct-human declarations; no document or tool output enters here."""
    match = candidate(question)
    if match is not None:
        return match
    if type(question) is not str or '\n' in question or len(question) > 120:
        return None
    text = question.strip()
    if not classify_candidate(text)['allowed'] or _SKIP_GUARDS.search(text):
        return None
    if _WORK.fullmatch(text):
        return 'work', text
    if _WORKFLOW.fullmatch(text):
        return 'skill', text
    return None


def manual_text(text):
    """Conservative local-note limit; filtering is not a complete DLP system."""
    return (type(text) is str and text == text.strip() and 1 <= len(text) <= 300
            and classify_candidate(text)['allowed'] and not _SKIP_GUARDS.search(text))


class PreferenceMemory:
    def __init__(self, root, *, clock=time.time):
        self.root = Path(os.path.abspath(os.fspath(root))) / 'memory'
        self._clock = clock
        self._lock = threading.RLock()
        self._plans = {}
        self._observed = {}
        self._suggestion = {}
        self._disabled = False

    def invalidate(self):
        with self._lock:
            self._plans.clear()
            self._observed.clear()
            self._suggestion.clear()

    def _path(self, key):
        if type(key) is not str or not _KEY.fullmatch(key):
            raise PreferenceError('memory_auth_required')
        path = self.root / (key + '.json')
        LocalAuth._check_path(path)
        return path

    @staticmethod
    def _read(path):
        LocalAuth._check_path(path)
        try:
            before = path.lstat()
        except FileNotFoundError:
            return None, {'version': 1, 'enabled': False, 'revision': 0, 'items': []}
        if not stat.S_ISREG(before.st_mode) or before.st_size > MAX_BYTES or before.st_nlink != 1:
            raise PreferenceError()
        fd = os.open(path, os.O_RDONLY | getattr(os, 'O_NOFOLLOW', 0) | getattr(os, 'O_BINARY', 0))
        with os.fdopen(fd, 'rb') as stream:
            opened = os.fstat(stream.fileno())
            if (opened.st_dev, opened.st_ino) != (before.st_dev, before.st_ino):
                raise PreferenceError()
            data = stream.read(MAX_BYTES + 1)
        LocalAuth._check_path(path)
        after = path.lstat()
        if (after.st_dev, after.st_ino) != (opened.st_dev, opened.st_ino) or len(data) > MAX_BYTES:
            raise PreferenceError()
        def unique(pairs):
            result = {}
            for k, v in pairs:
                if k in result:
                    raise PreferenceError()
                result[k] = v
            return result
        record = json.loads(data.decode('utf-8'), object_pairs_hook=unique)
        if (type(record) is not dict or set(record) != {'version', 'enabled', 'revision', 'items'}
                or type(record['version']) is not int or record['version'] not in (1, 2)
                or type(record['enabled']) is not bool
                or type(record['revision']) is not int or record['revision'] < 1
                or type(record['items']) is not list or len(record['items']) > 30):
            raise PreferenceError()
        ids, manual_kinds = set(), set()
        for item in record['items']:
            fields = {'id', 'kind', 'text', 'confirmed_at'}
            if record['version'] == 2:
                fields.add('source')
            if (type(item) is not dict or set(item) != fields
                    or type(item['id']) is not str or not _ID.fullmatch(item['id'])
                    or item['id'] in ids or type(item['kind']) is not str
                    or type(item['confirmed_at']) not in (int, float)
                    or not math.isfinite(item['confirmed_at']) or item['confirmed_at'] < 0):
                raise PreferenceError()
            if record['version'] == 1:
                valid = candidate(item['text']) == (item['kind'], item['text'])
            elif item['source'] == 'manual':
                valid = (item['kind'] in _MANUAL_KINDS and manual_text(item['text'])
                         and item['kind'] not in manual_kinds)
                manual_kinds.add(item['kind'])
            else:
                valid = (item['source'] == 'proposal' and item['kind'] in _AUTO_KINDS
                         and derived_candidate(item['text']) == (item['kind'], item['text']))
            if not valid:
                raise PreferenceError()
            ids.add(item['id'])
        return data, record

    def snapshot(self, key):
        with self._lock:
            if self._disabled:
                raise PreferenceError('memory_uncertain')
            path = self._path(key)
            _, state = self._read(path)
            now = self._clock()
            items = [{**item, 'source': item.get('source', 'proposal'),
                      'weight': (1.0 if item.get('source') == 'manual' else
                                 round(weight_at(item['confirmed_at'], now), 6))}
                     for item in state['items']]
            suggestion = self._suggestion.get(key)
            if suggestion is not None and suggestion['expires'] <= time.monotonic():
                self._suggestion.pop(key, None)
                suggestion = None
            return {'enabled': state['enabled'], 'revision': state['revision'],
                    'items': items, 'suggestion': ({k: v for k, v in suggestion.items() if k != 'expires'}
                                                 if suggestion else None)}

    def observe(self, key, question):
        """Two distinct direct human turns; never infer consent from task/model text."""
        with self._lock:
            if self._disabled:
                return None
            path = self._path(key)
            _, state = self._read(path)
            if not state['enabled']:
                return None
            match = derived_candidate(question)
            if (match is None or _SKIP_GUARDS.search(match[1])
                    or any(i['text'] == match[1] for i in state['items'])):
                return None
            seen = self._observed.setdefault(key, {})
            if match not in seen and len(seen) >= 32:
                seen.pop(next(iter(seen)))  # Volatile bounded evidence, never stored history.
            seen[match] = min(2, seen.get(match, 0) + 1)
            if seen[match] >= 2:
                self._suggestion[key] = {'kind': match[0], 'text': match[1], 'repeated': True,
                                         'expires': time.monotonic() + 86400}
                return {'kind': match[0], 'text': match[1], 'repeated': True}
            return None

    def plan(self, key, action, value=None):
        with self._lock:
            if self._disabled:
                raise PreferenceError('memory_uncertain')
            path = self._path(key)
            old, state = self._read(path)
            if action == 'enable' and type(value) is bool:
                if state['enabled'] == value:
                    raise PreferenceError('memory_no_change')
            elif action == 'save':
                suggestion = self._suggestion.get(key)
                if (not state['enabled'] or suggestion is None or suggestion['expires'] <= time.monotonic()
                        or value != suggestion['text'] or len(state['items']) >= 30):
                    raise PreferenceError('memory_invalid')
            elif action == 'set_manual':
                if (not state['enabled'] or type(value) is not dict or set(value) != {'kind', 'text'}
                        or type(value['kind']) is not str or value['kind'] not in _MANUAL_KINDS
                        or not manual_text(value['text'])):
                    raise PreferenceError('memory_invalid')
                existing = [i for i in state['items'] if i['kind'] == value['kind']
                            and i.get('source') == 'manual']
                if existing and existing[0]['text'] == value['text']:
                    raise PreferenceError('memory_no_change')
                if not existing and len(state['items']) >= 30:
                    raise PreferenceError('memory_invalid')
                value = dict(value)  # Never retain a mutable bridge-owned request.
            elif action == 'delete':
                if type(value) is not str or not _ID.fullmatch(value) or not any(i['id'] == value for i in state['items']):
                    raise PreferenceError('memory_invalid')
            elif action == 'clear':
                if not state['items']:
                    raise PreferenceError('memory_no_change')
            else:
                raise PreferenceError('memory_invalid')
            now = time.monotonic()
            self._plans = {token: entry for token, entry in self._plans.items() if entry[4] > now}
            if len(self._plans) >= 16:
                raise PreferenceError('memory_capacity')
            token = secrets.token_hex(16)
            self._plans[token] = (key, hashlib.sha256(old or b'').hexdigest(), action, value, now + 300)
            return {'plan_id': token, 'action': action, 'value': value, 'confirmation_required': True}

    def confirm(self, key, plan_id, confirmed=False):
        with self._lock:
            if type(plan_id) is not str:
                raise PreferenceError('memory_plan_invalid')
            saved = self._plans.pop(plan_id, None)
            if saved is None:
                raise PreferenceError('memory_plan_invalid')
            if confirmed is not True:
                raise PreferenceError('memory_confirmation_required')
            owner, digest, action, value, deadline = saved
            if key != owner or time.monotonic() > deadline:
                raise PreferenceError('memory_plan_expired')
            path = self._path(key)
            old, record = self._read(path)
            if hashlib.sha256(old or b'').hexdigest() != digest:
                raise PreferenceError('memory_changed')
            if record['version'] == 1:
                record['version'] = 2
                for item in record['items']:
                    item['source'] = 'proposal'
            if action == 'enable':
                record['enabled'] = value
                if value is False:
                    self._suggestion.pop(key, None)
                    self._observed.pop(key, None)
            elif action == 'save':
                suggestion = self._suggestion.get(key)
                if (not record['enabled'] or suggestion is None or suggestion['expires'] <= time.monotonic()
                        or suggestion['text'] != value):
                    raise PreferenceError('memory_changed')
                record['items'].append({'id': secrets.token_hex(16), 'kind': suggestion['kind'],
                                        'text': value, 'confirmed_at': self._clock(), 'source': 'proposal'})
            elif action == 'set_manual':
                if not record['enabled'] or value['kind'] not in _MANUAL_KINDS or not manual_text(value['text']):
                    raise PreferenceError('memory_changed')
                record['items'] = [i for i in record['items'] if not (
                    i['kind'] == value['kind'] and i['source'] == 'manual')]
                record['items'].append({'id': secrets.token_hex(16), 'kind': value['kind'],
                                        'text': value['text'], 'confirmed_at': self._clock(),
                                        'source': 'manual'})
            elif action == 'delete':
                record['items'] = [i for i in record['items'] if i['id'] != value]
            elif action == 'clear':
                record['items'] = []
            record['revision'] += 1
            payload = json.dumps(record, ensure_ascii=False, separators=(',', ':'), allow_nan=False).encode('utf-8')
            if len(payload) > MAX_BYTES:
                raise PreferenceError('memory_invalid')
            try:
                self.root.mkdir(parents=True, exist_ok=True)
                LocalAuth._check_path(path)
                if old is not None:
                    backup = self.root / (key + '-' + secrets.token_hex(16) + '.bak')
                    self._exclusive(backup, old)
                    if self._read_bytes(backup) != old:
                        raise PreferenceError()
                temp = self.root / (key + '-' + secrets.token_hex(16) + '.tmp')
                self._exclusive(temp, payload)
                if self._read_bytes(temp) != payload:
                    raise PreferenceError()
                LocalAuth._check_path(path)
                current, _ = self._read(path)
                if current != old:
                    raise PreferenceError('memory_changed')
                os.replace(temp, path)
                actual, new = self._read(path)
                if actual != payload or new != record:
                    raise PreferenceError()
                audit = self.root / (key + '-' + secrets.token_hex(16) + '.audit.json')
                self._exclusive(audit, json.dumps({'action': action, 'revision': record['revision'],
                    'before_sha256': digest, 'after_sha256': hashlib.sha256(payload).hexdigest(),
                    'backup': 'created' if old is not None else 'not_applicable_new_artifact',
                    'independent_human_confirmed': True}, separators=(',', ':')).encode('utf-8'))
                if not self._read_bytes(audit):
                    raise PreferenceError()
                if action in {'delete', 'clear', 'set_manual'}:
                    # Account-local privacy erasure (including replaced manual text)
                    # supersedes keeping historical
                    # plaintext backups; audits retain hashes but no preference text.
                    backups = list(self.root.glob(key + '-*.bak'))
                    if len(backups) > 1024:
                        raise PreferenceError()
                    for previous in backups:
                        LocalAuth._check_path(previous)
                        if not stat.S_ISREG(previous.lstat().st_mode) or previous.lstat().st_nlink != 1:
                            raise PreferenceError()
                        previous.unlink()
            except Exception:
                self._disabled = True  # After attempted write, never claim no effect.
                raise PreferenceError('memory_uncertain') from None
            if action == 'save':
                self._suggestion.pop(key, None)
            return {'saved': True, 'action': action, 'revision': record['revision']}

    @staticmethod
    def _exclusive(path, payload):
        LocalAuth._check_path(path)
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, 'O_NOFOLLOW', 0)
                     | getattr(os, 'O_BINARY', 0), 0o600)
        with os.fdopen(fd, 'wb') as stream:
            if stream.write(payload) != len(payload):
                raise PreferenceError()
            stream.flush()
            os.fsync(stream.fileno())

    @staticmethod
    def _read_bytes(path):
        LocalAuth._check_path(path)
        info = path.lstat()
        if not stat.S_ISREG(info.st_mode) or info.st_nlink != 1 or info.st_size > MAX_BYTES:
            raise PreferenceError()
        with path.open('rb') as stream:
            data = stream.read(MAX_BYTES + 1)
        LocalAuth._check_path(path)
        if len(data) != info.st_size or path.lstat().st_ino != info.st_ino:
            raise PreferenceError()
        return data

    def recall(self, key):
        """At most five manual notes plus three decaying proposals, account-local."""
        state = self.snapshot(key)
        if not state['enabled']:
            return []
        notes = {item['kind']: item for item in state['items'] if item['source'] == 'manual'}
        proposals = (item for item in state['items'] if item['source'] == 'proposal')
        ranked = sorted(proposals, key=lambda item: (-item['weight'], -item['confirmed_at'], item['id']))
        return [notes[kind] for kind in _MANUAL_KINDS if kind in notes] + ranked[:3]
