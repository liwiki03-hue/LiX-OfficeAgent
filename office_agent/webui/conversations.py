import copy
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import time
import unicodedata

from ..auth import LocalAuth


class ConversationError(ValueError):
    def __init__(self, code):
        self.code = code
        super().__init__(code)


class Conversations:
    @staticmethod
    def _body(rows):
        extended = any(set(turn) - {'id','question','answer','route','status','created','context_turns','omitted_turns','finish_reason'} or turn['finish_reason'] == 'local_limit' for row in rows for turn in row['turns'])
        return json.dumps({'version': 2 if extended else 1, 'sessions': rows}, ensure_ascii=False).encode('utf8')

    def __init__(self, root, key):
        self.root = Path(root) / 'conversations'
        self.key = key
        self.account = None
        self.rows = []
        self.current = None
        self.warning = ''
        self.observed = None
        self.loaded = False

    def invalidate(self):
        self.account = None
        self.rows = []
        self.current = None
        self.loaded = False
        self.observed = None
        self.warning = ''

    def _path(self):
        key = self.key()
        if type(key) is not str or not re.fullmatch('[a-f0-9]{64}', key):
            raise ConversationError('chat_account_required')
        if self.account != key:
            self.invalidate()
            self.account = key
        path = self.root / key / 'events.json'
        LocalAuth._check_path(path)
        return path

    def _read(self, path):
        from .bridge import _read_bounded_json
        LocalAuth._check_path(path)
        if not path.exists():
            return None, []
        info = path.lstat()
        if info.st_nlink != 1 or info.st_size > 4 * 1024 * 1024:
            raise ConversationError('chat_history_corrupt')
        body = _read_bounded_json(path, 4 * 1024 * 1024)
        if type(body) is not dict or set(body) != {'version', 'sessions'} or type(body['version']) is not int or body['version'] not in (1, 2) or type(body['sessions']) is not list or len(body['sessions']) > 32:
            raise ConversationError('chat_history_corrupt')
        identifiers = set()
        for row in body['sessions']:
            if type(row) is not dict or set(row) != {'id', 'created', 'title', 'turns'} or not re.fullmatch('[a-f0-9]{32}', row['id']) or row['id'] in identifiers or type(row['created']) not in (int, float) or type(row['title']) is not str or len(row['title']) > 80 or type(row['turns']) is not list or len(row['turns']) > 100:
                raise ConversationError('chat_history_corrupt')
            identifiers.add(row['id'])
            for turn in row['turns']:
                required = {'id', 'question', 'answer', 'route', 'status', 'created', 'context_turns', 'omitted_turns', 'finish_reason'}
                allowed = required | ({'budget','usage','continuation_of'} if body['version'] == 2 else set())
                if type(turn) is not dict or not required <= set(turn) <= allowed or not re.fullmatch('[a-f0-9]{32}', turn['id']) or type(turn['question']) is not str or len(turn['question']) > 8000 or type(turn['answer']) is not str or len(turn['answer']) > 32000 or turn['status'] not in {'complete', 'failed', 'stopped', 'interrupted'} or type(turn['created']) not in (int, float):
                    raise ConversationError('chat_history_corrupt')
                if any(type(turn[name]) is not int or not 0 <= turn[name] <= 100 for name in ('context_turns', 'omitted_turns')):
                    raise ConversationError('chat_history_corrupt')
                if turn['finish_reason'] not in (('unknown', 'stop', 'length', 'content_filter', 'local_limit') if body['version'] == 2 else ('unknown', 'stop', 'length', 'content_filter')):
                    raise ConversationError('chat_history_corrupt')
                if 'budget' in turn:
                    budget = turn['budget']
                    if type(budget) is not dict or set(budget) != {'reply_token_reserve','context_window','estimated_input_tokens','output_parameter'} or any(type(budget[key]) is not int or not 0 <= budget[key] <= 10000000 for key in ('reply_token_reserve','context_window','estimated_input_tokens')) or budget['output_parameter'] not in ('max_tokens','max_completion_tokens','num_predict'):
                        raise ConversationError('chat_history_corrupt')
                if 'usage' in turn and (type(turn['usage']) is not dict or not set(turn['usage']) <= {'input_tokens','output_tokens','total_tokens'} or any(type(value) is not int or not 0 <= value <= 1000000000 for value in turn['usage'].values())):
                    raise ConversationError('chat_history_corrupt')
                if 'continuation_of' in turn:
                    link = turn['continuation_of']
                    if type(link) is not dict or set(link) != {'session_id','turn_id','source_digest'} or link['session_id'] != row['id'] or type(link['turn_id']) is not str or not re.fullmatch('[a-f0-9]{32}',link['turn_id']) or type(link['source_digest']) is not str or not re.fullmatch('[a-f0-9]{64}',link['source_digest']):
                        raise ConversationError('chat_history_corrupt')
                route = turn['route']
                if type(route) is not dict or set(route) != {'connection_id', 'connection_name', 'model_id', 'capability', 'revision'} or any(type(v) is not str or len(v) > 200 for v in route.values()) or route['capability'] != 'chat':
                    raise ConversationError('chat_history_corrupt')
        digest = hashlib.sha256(json.dumps(body, ensure_ascii=False).encode('utf8')).hexdigest()
        return digest, body['sessions']

    def ensure(self, enabled):
        path = self._path()
        if not self.loaded:
            self.loaded = True
            try:
                self.observed, rows = self._read(path)
                self.rows = rows
            except Exception:
                self.warning = 'chat_history_corrupt'
        if self.current is None:
            if self.rows:
                self.current = self.rows[-1]['id']
            else:
                self.new()

    def new(self):
        if len(self.rows) >= 32:
            raise ConversationError('chat_history_capacity')
        row = {'id': secrets.token_hex(16), 'created': time.time(), 'title': '新会话', 'turns': []}
        self.rows.append(row)
        self.current = row['id']
        return row

    def row(self, identifier=None):
        return next((row for row in self.rows if row['id'] == (identifier or self.current)), None)

    def public(self, enabled):
        self.ensure(enabled)
        body = self._body(self.rows)
        current = self.row()
        save_state = ('uncertain' if self.warning == 'chat_history_uncertain' else 'failed' if self.warning
                      else 'disabled' if not enabled else 'saved' if hashlib.sha256(body).hexdigest() == self.observed else 'unsaved')
        return {'current': self.current, 'sessions': copy.deepcopy(self.rows), 'saving_enabled': enabled,
                'location': str(self._path()), 'warning': self.warning,
                'save_state': save_state,
                'capacity': {'sessions': len(self.rows), 'max_sessions': 32,
                             'current_turns': len(current['turns']) if current else 0, 'max_turns': 100,
                             'bytes': len(body), 'max_bytes': 4 * 1024 * 1024},
                'scope': '仅普通问答文字、时间和实际模型来源；不恢复文件、批准或接口凭据。'}

    def persist(self, enabled, *, rows=None):
        if not enabled:
            return 'disabled'
        temporary = None
        attempted = False
        try:
            path = self._path()
            if self.warning in {'chat_history_corrupt', 'chat_history_uncertain'}:
                raise ConversationError(self.warning)
            observed, _ = self._read(path)
            if observed != self.observed:
                raise ConversationError('chat_history_conflict')
            body = self._body(self.rows if rows is None else rows)
            if len(body) > 4 * 1024 * 1024:
                raise ConversationError('chat_history_capacity')
            LocalAuth._check_path(path)
            path.parent.mkdir(parents=True, exist_ok=True)
            LocalAuth._check_path(path)
            temporary = path.with_name('.pending-' + secrets.token_hex(16))
            with temporary.open('xb') as stream:
                stream.write(body)
                stream.flush()
                os.fsync(stream.fileno())
            current, _ = self._read(path)
            if current != observed:
                raise ConversationError('chat_history_conflict')
            LocalAuth._check_path(path)
            attempted = True
            os.replace(temporary, path)
            temporary = None
            self.observed, _ = self._read(path)
            if self.observed != hashlib.sha256(body).hexdigest():
                raise ConversationError('chat_history_write_failed')
            self.warning = ''
            return 'saved'
        except Exception as error:
            if attempted and isinstance(error, PermissionError) and temporary is not None:
                try:
                    LocalAuth._check_path(temporary)
                    info = temporary.lstat()
                    with temporary.open('rb') as stream:
                        unchanged = info.st_nlink == 1 and info.st_size == len(body) and stream.read(len(body) + 1) == body
                    if unchanged and self._read(path)[0] == observed:
                        attempted = False
                except Exception:
                    pass
            self.warning = ('chat_history_uncertain' if attempted else
                            error.code if isinstance(error, ConversationError) else 'chat_history_write_failed')
            return 'uncertain' if self.warning == 'chat_history_uncertain' else 'not_saved'
        finally:
            if temporary is not None:
                try:
                    LocalAuth._check_path(temporary)
                    temporary.unlink()
                except Exception:
                    pass

    def remove(self, identifier, enabled):
        row = self.row(identifier)
        if row is None:
            raise ConversationError('chat_session_missing')
        original, selected = self.rows[:], self.current
        saved_rows = None
        if not enabled:
            observed, saved = self._read(self._path())
            if observed is not None and any(item['id'] == identifier for item in saved):
                if observed != self.observed:
                    self.warning = 'chat_history_conflict'
                    raise ConversationError(self.warning)
                saved_rows = [item for item in saved if item['id'] != identifier]
        self.rows = [item for item in self.rows if item['id'] != identifier]
        if self.current == identifier:
            if self.rows:
                self.current = self.rows[-1]['id']
            else:
                self.current = None
                self.new()
        outcome = self.persist(True, rows=saved_rows) if enabled or saved_rows is not None else 'disabled'
        if outcome == 'not_saved':
            self.rows, self.current = original, selected
        if outcome in {'not_saved', 'uncertain'}:
            raise ConversationError(self.warning)

    def rename(self, identifier, title, enabled):
        if type(title) is not str or not 1 <= len(title.strip()) <= 80 or any(unicodedata.category(char) in {'Cc', 'Cs'} for char in title):
            raise ConversationError('chat_title_invalid')
        row = self.row(identifier)
        if row is None:
            raise ConversationError('chat_session_missing')
        previous = row['title']
        row['title'] = title.strip()
        outcome = self.persist(enabled)
        if outcome == 'not_saved':
            row['title'] = previous
        if outcome in {'not_saved', 'uncertain'}:
            raise ConversationError(self.warning)

    def clear_saved(self, confirmed):
        if confirmed is not True:
            raise ConversationError('confirmation_required')
        path = self._path()
        LocalAuth._check_path(path)
        if path.exists():
            info = path.lstat()
            if info.st_nlink != 1 or not path.is_file():
                raise ConversationError('chat_history_corrupt')
            path.unlink()
        self.invalidate()
