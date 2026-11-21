import json
import os
import re
import sqlite3
import threading
import time
import uuid
from contextlib import closing, contextmanager
from pathlib import Path

from dotenv import dotenv_values

from . import plugin_catalog


LEARN_URL = 'https://learn.microsoft.com/api/mcp'
SECRET_NAME = re.compile(r'LIX_PLUGIN_[A-Z][A-Z0-9_]{0,79}\Z')
CONFIG_KEYS = {'name', 'transport', 'url', 'auth', 'credential_ref', 'private', 'source', 'program', 'args', 'env_refs', 'api_header', 'oauth_issuer'}
LEARN_REQUEST = re.compile(r'(?:查(?:询|找)?|(?:请)?搜索|检索|参考|使用|用)\s*(?:微软(?:官方)?文档|Microsoft\s*Learn)', re.I)


def config_fields(value):
    if not isinstance(value, dict) or value.keys() - CONFIG_KEYS:
        raise ValueError('插件配置字段无效。')
    name = value.get('name', '')
    if not isinstance(name, str) or not name.strip() or len(name) > 100 or not name.isprintable():
        raise ValueError('请填写 1 至 100 字的插件名称。')
    transport, auth = value.get('transport', 'http'), value.get('auth', 'none')
    if transport not in ('http', 'stdio') or auth not in ('none', 'bearer', 'oauth'):
        raise ValueError('传输或认证方式不可用。')
    source = value.get('source', '')
    if source and (not isinstance(source, str) or len(source) > 200 or not plugin_catalog.IDENTITY.fullmatch(source)):
        raise ValueError('目录身份无效。')
    result = dict(name=name.strip(), transport=transport, auth=auth, source=source, scope='user')
    if transport == 'http':
        ref = value.get('credential_ref', '') if auth == 'bearer' else ''
        if auth == 'bearer' and (not isinstance(ref, str) or not SECRET_NAME.fullmatch(ref)):
            raise ValueError('凭据引用应为 .env 中以 LIX_PLUGIN_ 开头的变量名。')
        header = value.get('api_header') or 'Authorization'
        if header not in ('Authorization', 'X-Api-Key') or type(value.get('private', False)) is not bool:
            raise ValueError('凭据标头或网络范围无效。')
        result.update(url=plugin_catalog.clean_url(value.get('url')), private=value.get('private', False),
                      credential_ref=ref, api_header=header)
        if auth == 'oauth' and value.get('oauth_issuer'):
            result['oauth_issuer'] = plugin_catalog.clean_url(value['oauth_issuer'])
        if source and result['private']:
            raise ValueError('目录连接只允许公共端点。私有服务请手动添加。')
    else:
        program, args, refs = value.get('program'), value.get('args', []), value.get('env_refs', {})
        if auth != 'none' or source or not isinstance(program, str) or not Path(program).is_absolute() or not Path(program).is_file():
            raise ValueError('请手动选择已安装的本地程序。')
        if not isinstance(args, list) or len(args) > 32 or any(not isinstance(s, str) or len(s) > 1024 or '\x00' in s for s in args):
            raise ValueError('启动参数应逐项填写，最多 32 项。')
        if not isinstance(refs, dict) or len(refs) > 16 or any(not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]{0,63}', k)
                or not isinstance(v, str) or not SECRET_NAME.fullmatch(v) for k, v in refs.items()):
            raise ValueError('环境变量须引用 .env 中以 LIX_PLUGIN_ 开头的变量。')
        result.update(program=str(Path(program).resolve()), args=args, env_refs=refs)
    return result


def same_target(row):
    return tuple(row.get(k) for k in ('transport', 'url', 'program', 'args', 'env_refs', 'auth', 'credential_ref', 'api_header', 'private', 'oauth_issuer'))


class PluginConnections:
    def __init__(self, profile, settings_root):
        self.profile, self.settings_root = Path(profile), Path(settings_root)
        self.db = self.profile / 'connections.sqlite3'
        self.lock = threading.RLock()
        self.operations = {}

    @contextmanager
    def database(self):
        from .bridge import _reject_links
        _reject_links(self.db)
        self.profile.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(self.db, timeout=5)) as db, db:
            db.execute('CREATE TABLE IF NOT EXISTS plugins(id TEXT PRIMARY KEY, data TEXT NOT NULL)')
            db.execute('CREATE TABLE IF NOT EXISTS catalog(term TEXT, cursor TEXT, at REAL, data TEXT, PRIMARY KEY(term,cursor))')
            db.execute('BEGIN IMMEDIATE')
            yield db

    @staticmethod
    def get(db, key, revision=None):
        if not isinstance(key, str):
            raise ValueError('连接标识无效。')
        result = db.execute('SELECT data FROM plugins WHERE id=?', (key,)).fetchone()
        if not result:
            raise ValueError('连接已移除，请刷新插件列表。')
        row = json.loads(result[0])
        if revision is not None and (type(revision) is not int or revision != row['revision']):
            raise ValueError('连接已更新，请刷新后重试。')
        return row

    @staticmethod
    def put(db, row):
        db.execute('INSERT OR REPLACE INTO plugins VALUES (?,?)', (row['id'], json.dumps(row, ensure_ascii=False)))

    def list(self):
        if not self.db.exists():
            return []
        with self.database() as db:
            return [json.loads(data) for data, in db.execute('SELECT data FROM plugins ORDER BY rowid')]

    def search(self, term, cursor=''):
        if not isinstance(term, str) or not term.strip() or len(term) > 100 or not term.isprintable():
            raise ValueError('请输入 1 至 100 字的搜索词。')
        if not isinstance(cursor, str) or len(cursor) > 1500:
            raise ValueError('目录翻页标识无效。')
        term = term.strip()
        with self.database() as db:
            cached = db.execute('SELECT at,data FROM catalog WHERE term=? AND cursor=?', (term, cursor)).fetchone()
            first = db.execute('SELECT data FROM catalog WHERE term=? AND cursor=?', (term, '')).fetchone() if cursor else None
        if cached and time.time() - cached[0] < 300:
            return dict(json.loads(cached[1]), cached_at=cached[0], offline=False)
        try:
            lookup = json.loads(first[0]).get('search_term', term) if first else term
            page = plugin_catalog.fetch_page(lookup, cursor)
            if not cursor and not page['items'] and len(term.split()) > 1:
                lookup = '-'.join(term.split()).casefold()
                page = plugin_catalog.fetch_page(lookup, '')
            page['search_term'] = lookup
        except Exception:
            if cached:
                return dict(json.loads(cached[1]), cached_at=cached[0], offline=True)
            raise ValueError('公开目录暂时无法访问。可重试或手动添加服务。') from None
        now = time.time()
        with self.database() as db:
            db.execute('INSERT OR REPLACE INTO catalog VALUES (?,?,?,?)', (term, cursor, now, json.dumps(page)))
            db.execute('DELETE FROM catalog WHERE rowid IN (SELECT rowid FROM catalog ORDER BY at DESC LIMIT -1 OFFSET 100)')
        return dict(page, cached_at=now, offline=False)

    def save(self, config, key=None, revision=None, separate=False):
        fields = config_fields(config)
        with self.database() as db:
            old = self.get(db, key, revision) if key else None
            if old and type(revision) is not int:
                raise ValueError('请刷新连接版本。')
            if fields['source'] and not (old and old['source'] == fields['source'] and old.get('url') == fields.get('url')):
                if not any(entry['source'] == fields['source'] and entry['url'] == fields['url']
                           for data, in db.execute('SELECT data FROM catalog') for entry in json.loads(data)['items']):
                    raise ValueError('目录地址已改变，请重新搜索或手动添加。')
            if not old and not separate:
                for data, in db.execute('SELECT data FROM plugins'):
                    row = json.loads(data)
                    if same_target(row) == same_target(fields):
                        return row, True
            if not old and db.execute('SELECT COUNT(*) FROM plugins').fetchone()[0] >= 48:
                raise ValueError('已达到 48 个本机连接上限。')
            changed = not old or same_target(old) != same_target(fields) or old['source'] != fields['source']
            row = dict(fields, id=key or uuid.uuid4().hex, revision=old['revision'] + 1 if old else 1,
                       tools=[] if changed else old['tools'], allowed=[] if changed else old['allowed'],
                       enabled=False if changed else old['enabled'], state='已添加' if changed else old['state'])
            self.put(db, row)
        if old:
            self.cancel(key)
            if changed and old.get('auth') == 'oauth':
                self._disconnect(old)
        return row, False

    def secret(self, name):
        if not isinstance(name, str) or not SECRET_NAME.fullmatch(name):
            return ''
        from .bridge import _reject_links
        target = self.settings_root / '.env'
        _reject_links(target)
        if target.is_file() and target.stat().st_size <= 1024 * 1024:
            return dotenv_values(target, encoding='utf-8-sig', interpolate=False).get(name) or ''
        return ''

    def runtime(self, row):
        return dict(row, _profile_dir=self.profile, _resolve_secret=self.secret)

    def _disconnect(self, row):
        from . import plugin_transport
        plugin_transport.disconnect(self.runtime(row))

    def cancel(self, key):
        with self.lock:
            for event in self.operations.get(key, ()):
                event.set()

    def cancel_all(self):
        with self.lock:
            for key in self.operations:
                self.cancel(key)

    @contextmanager
    def operation(self, row, cancelled=lambda: False):
        key = row['id']
        event = threading.Event()
        with self.lock:
            self.operations.setdefault(key, set()).add(event)
        def stale():
            return event.is_set() or cancelled()
        try:
            with self.database() as db:
                self.get(db, key, row['revision'])
            if stale():
                raise ValueError('本次插件操作已停止。')
            yield stale
        finally:
            with self.lock:
                self.operations[key].discard(event)
                if not self.operations[key]:
                    del self.operations[key]

    def probe(self, key, revision, login=False):
        from . import plugin_transport
        with self.database() as db:
            row = self.get(db, key, revision)
        error = ''
        try:
            with self.operation(row) as cancelled:
                action = plugin_transport.login if login else plugin_transport.probe
                tools = action(self.runtime(row), cancelled=cancelled)
                if cancelled():
                    raise ValueError('连接测试已取消。')
        except ValueError as exc:
            tools = []
            error = str(exc)[:220]
        with self.database() as db:
            row = self.get(db, key, revision)
            old = {t['name']: t['signature'] for t in row['tools'] if t['usable']}
            row['allowed'] = [t['name'] for t in tools if t['usable'] and t['name'] in row['allowed'] and old.get(t['name']) == t['signature']]
            row.update(tools=tools, enabled=bool(row['allowed'] and row['enabled']), revision=revision + 1)
            row['state'] = ('待配置' if error else '可用' if row['enabled'] else '已连接' if any(t['usable'] for t in tools) else '已连接但无可用工具')
            row['error'] = error
            self.put(db, row)
        self.cancel(key)
        return row

    def change(self, key, revision, action, allowed=()):
        if type(revision) is not int:
            raise ValueError('请刷新连接版本。')
        with self.database() as db:
            row = self.get(db, key, revision)
            if action == 'remove':
                db.execute('DELETE FROM plugins WHERE id=?', (key,))
            else:
                if action == 'enable':
                    names = {t['name'] for t in row['tools'] if t['usable']}
                    if not isinstance(allowed, list) or not 1 <= len(allowed) <= 8 or not all(isinstance(n, str) and n in names for n in allowed):
                        raise ValueError('请先测试连接，再选择已核对的只读工具。')
                    row.update(allowed=list(dict.fromkeys(allowed)), enabled=True, state='可用')
                elif action in ('disable', 'disconnect'):
                    row.update(allowed=[], enabled=False, state='已停用')
                    if action == 'disconnect':
                        row.update(tools=[], state='待配置')
                else:
                    raise ValueError('不支持的连接操作。')
                row['revision'] += 1
                self.put(db, row)
        self.cancel(key)
        if action in ('remove', 'disconnect') and row['auth'] == 'oauth':
            self._disconnect(row)
        return row

    def resolve(self, text, choice):
        public_learn = bool(LEARN_REQUEST.match(text.strip().removeprefix('请').removeprefix('帮我')))
        if not choice and not public_learn:
            return None
        if choice:
            if not isinstance(choice, dict) or set(choice) != {'id', 'tool', 'revision'}:
                raise ValueError('请重新选择本次插件。')
            with self.database() as db:
                row = self.get(db, choice['id'], choice['revision'])
            tool = choice['tool']
        else:
            candidates = [r for r in self.list() if r.get('url') == LEARN_URL and r['enabled'] and 'microsoft_docs_search' in r['allowed']]
            if len(candidates) != 1:
                raise ValueError('请在「设置 → 插件」启用 Microsoft Learn 文档搜索，或在输入区选择本次连接。')
            row, tool = candidates[0], 'microsoft_docs_search'
        if not row['enabled'] or tool not in row['allowed']:
            raise ValueError('插件或工具已停用，请在设置中核对。')
        return dict(id=row['id'], revision=row['revision'], name=row['name'], tool=tool, destination=row.get('url') or row.get('program'))

    def checked(self, ticket):
        with self.database() as db:
            row = self.get(db, ticket['id'], ticket['revision'])
        if not row['enabled'] or ticket['tool'] not in row['allowed']:
            raise ValueError('插件已停用。')
        return row

    def query(self, ticket, query_text, cancelled, before_dispatch=None):
        from . import plugin_transport
        row = self.checked(ticket)
        tool = next(t for t in row['tools'] if t['name'] == ticket['tool'] and t['usable'])
        with self.operation(row, cancelled) as stale:
            def authorize():
                self.checked(ticket)
                if stale():
                    raise ValueError('本次插件查询已停止。')
                if before_dispatch:
                    before_dispatch()
            config = self.runtime(row)
            config['_before_dispatch'] = authorize
            result = plugin_transport.query(config, tool, query_text, cancelled=stale)
            self.checked(ticket)
            if stale():
                raise ValueError('本次插件查询已停止。')
            return result

