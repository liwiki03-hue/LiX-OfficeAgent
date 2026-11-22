import hashlib
import json
import re
import time
import uuid
from copy import deepcopy
from dataclasses import dataclass
from urllib.parse import urlsplit

from ..rag import SECRET
from ..diagnostics import record


FIELDS = {
    '邮箱': re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_\x60{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+"),
    '证件号': re.compile(r'(?<![A-Za-z0-9])[1-9][0-9]{5}(?:19|20)[0-9]{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])[0-9]{3}[0-9Xx](?![A-Za-z0-9])'),
    '电话': re.compile(r'(?<![A-Za-z0-9])(?:(?:\+?86[- ]?)?1[3-9]\d[- ]?\d{4}[- ]?\d{4}|0\d{2,3}[- ]\d{7,8})(?![A-Za-z0-9])'),
}
LABELED = {
    '证件号': re.compile(r'((?:身份证(?:号码?|号)?|居民身份号码)\s*[:：]?\s*)([1-9]\d{16}[0-9Xx])'),
    '电话': re.compile(r'((?:联系电话|电话|手机|mobile|phone|tel)(?:号码?)?\s*[:：]\s*)(\+?[\d(][\d() -]{5,20}\d)', re.I),
}
CREDENTIAL = re.compile(r'''(?ix)(?:api[_ -]?key|access[_ -]?token|password|secret|密码|密钥)["']?\s*[:=：]\s*(?:"[^"\n]*"|'[^'\n]*'|[^\s"',;，；。}\]]+)|\bsk-[a-z0-9_-]{10,}|\bbearer\s+[a-z0-9._-]{10,}|-----BEGIN\ [A-Z\ ]*PRIVATE\ KEY-----[\s\S]*?(?:-----END\ [A-Z\ ]*PRIVATE\ KEY-----|\Z)''')
EXTRA_FILE = re.compile(r'(?:参考|参照|对照|比较|结合|读取|打开|看看?|使用|附上|选择|添加).{0,12}(?:另(?:一|外)|第二|额外|其他).{0,8}(?:文件|文档|简历|表格|报告|附件)|(?:另一份|额外的?)(?:文件|文档).{0,12}(?:参考|对照|比较|一起)')
TTL = 300


def scrub_secrets(text):
    return SECRET.sub('[凭据已隐藏]', CREDENTIAL.sub('[凭据已隐藏]', text))


def mask_text(text, keep=()):
    text = scrub_secrets(text)
    for name, pattern in FIELDS.items():
        if name not in keep:
            text = pattern.sub('[' + name + '已隐藏]', text)
            if name in LABELED:
                text = LABELED[name].sub(lambda match: match[1] + '[' + name + '已隐藏]', text)
    return text


def field_names(text):
    return [name for name, pattern in FIELDS.items() if pattern.search(text) or name in LABELED and LABELED[name].search(text)]


def wants_raw(question):
    return (bool(re.search(r'原样|原文|未脱敏|不(?:要)?(?:脱敏|隐藏|遮盖)', question))
            and not re.search(r'(?:不要|别|不必|不用|不能)(?:把.{0,24})?(?:原样|原文)', question))


def requested_fields(question):
    names = []
    if re.search(r'电话|手机|联系方式|个人信息', question):
        names.append('电话')
    if re.search(r'邮箱|邮件|联系方式|个人信息', question):
        names.append('邮箱')
    if re.search(r'证件|身份证|个人信息', question):
        names.append('证件号')
    return names or list(FIELDS)


def unsupported_reason(text):
    if re.search(r'(?:导出|读取|列出|发给|显示|打印).{0,16}(?:配置密钥|配置.*(?:密码|令牌)|环境变量|\.env|API\s*密钥)', text, re.I):
        return '文档问答不能读取或导出本机配置凭据。'
    if re.search(r'(?:扫描|遍历|读取|枚举).{0,8}(?:整个目录|所有目录|整[个台]电脑|整个磁盘|全部子目录)', text):
        return '一次可以选择一份参考文件。文档问答不支持读取整个目录。'
    if re.search(r'(?:执行|运行).{0,6}(?:shell|powershell|cmd(?:\.exe)?|系统命令|终端命令)', text, re.I):
        return '文档问答不提供系统命令执行。'
    return ''


def safe_tree(value, transform=mask_text):
    if isinstance(value, str):
        return transform(value)
    if isinstance(value, list):
        return [safe_tree(item, transform) for item in value]
    if isinstance(value, dict):
        result = {}
        for key, item in value.items():
            name = transform(key) if isinstance(key, str) else key
            if isinstance(key, str) and re.fullmatch(r'api[ _-]?key|access[ _-]?token|password|client[ _-]?secret|secret|密码|密钥', key, re.I):
                result[name] = '[凭据已隐藏]'
            else:
                result[name] = safe_tree(item, transform)
        return result
    return value


def clean_material(value, transform):
    if isinstance(value, str) and value.startswith(('{', '[')):
        try:
            return json.dumps(safe_tree(json.loads(value), transform), ensure_ascii=False)
        except ValueError:
            pass
    return safe_tree(value, transform)


def route_key(route):
    return (route.connection_id, route.model_id, route.revision, route.base_url, route.epoch)


def recipient(route):
    address = urlsplit(route.base_url)
    return f'{route.connection_name} · {route.model_id}（{address.netloc}）'


def fingerprint(payload):
    return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()).digest()


def record_decision(identifier, action, decision):
    record('scope', action, decision, correlation=hashlib.sha256(identifier.encode()).hexdigest()[:24])


@dataclass
class Decision:
    token: str
    turn: str
    action: str
    binding: tuple
    target: tuple
    reference: dict | None
    full: dict
    masked: dict
    label: str
    fields: list
    created: float
    raw: bool

    def public(self):
        return {'token': self.token, 'action': self.action, 'destination': self.label,
                'fields': self.fields, 'raw': self.raw}


@dataclass(frozen=True)
class PreparedFile:
    token: str
    owner: tuple
    created: float
    selected: object


class TurnScopes:
    def __init__(self, conversation):
        self.chat = conversation
        self.primary = None
        self.pending = {}
        self.approved = {}
        self.dispatch = {}
        self.attachments = {}
        self.reconfirm = set()
        self.local_only = set()
        self.prepared_file = None
        self.file_choice = 0
        self.selected_files = {}
        self.selected_skills = {}

    def clear(self):
        self.pending.clear()
        self.approved.clear()
        self.dispatch.clear()
        self.attachments.clear()
        self.reconfirm.clear()
        self.local_only.clear()
        self.prepared_file = None
        self.file_choice += 1
        self.selected_files.clear()
        self.selected_skills.clear()

    def forget(self, identifier):
        self.reconfirm.discard(identifier)
        self.local_only.discard(identifier)
        for mapping in (self.pending, self.approved, self.dispatch, self.attachments, self.selected_files,
                        self.selected_skills, self.chat.plugin_results):
            mapping.pop(identifier, None)

    def owner(self):
        api = self.chat.floating.api
        return (getattr(api._auth, 'account_key', lambda: '')() or '', self.chat.binding, self.chat.generation)

    def _file_owner(self, document_id, version):
        floating = self.chat.floating
        floating.authenticated()
        doc = floating.document
        if not doc or not doc.get('ready') or (doc['id'], doc['version']) != (document_id, version):
            raise ValueError('当前文档已切换，请重新选择参考文件。')
        floating.check_source()
        account, _, generation = self.owner()
        return account, generation, *self.chat._read_binding()[:3]

    def choose_file(self, document_id, version):
        from .turn_attachment import choose
        return self._prepare_file(document_id, version, lambda: choose(self.chat.floating))

    def drop_file(self, epoch, document_id, version, path):
        from .turn_attachment import select_dropped
        return self._prepare_file(document_id, version, lambda: select_dropped(path), epoch)

    def _prepare_file(self, document_id, version, select, expected_epoch=None):
        from .turn_attachment import recheck
        floating = self.chat.floating
        with floating.lock, self.chat.lock:
            if expected_epoch is not None and floating.epoch != expected_epoch:
                raise ValueError('当前文档已切换，请重新拖入参考文件。')
            owner = self._file_owner(document_id, version)
            self.file_choice += 1
            request = self.file_choice
        selected = select()
        with floating.lock, self.chat.lock:
            if request != self.file_choice or owner != self._file_owner(document_id, version):
                raise ValueError('选择已过期，请在当前对话重新选择。')
            if selected is None:
                return {'ok': True, 'cancelled': True}
            if selected.path == floating.app.path:
                raise ValueError('当前文档已在本轮材料中，请选择另一份文件。')
            prepared = PreparedFile(uuid.uuid4().hex, owner, time.monotonic(), recheck(selected))
            self.prepared_file = prepared
            return {'ok': True, 'file': {'token': prepared.token, 'name': mask_text(selected.name)}}

    def remove_file(self, token):
        with self.chat.floating.lock, self.chat.lock:
            self.chat.floating.authenticated()
            if self.prepared_file and self.prepared_file.token == token:
                self.prepared_file = None
                self.file_choice += 1
            return {'ok': True}

    def checked_file(self, token, document_id, version):
        from .turn_attachment import recheck
        item = self.prepared_file
        if (not item or type(token) is not str or item.token != token or time.monotonic() - item.created > TTL
                or item.owner != self._file_owner(document_id, version)):
            raise ValueError('参考文件选择已失效，请重新选择。')
        return recheck(item.selected)

    def recheck_attachment(self, identifier):
        from .turn_attachment import recheck
        selected = self.selected_files.get(identifier)
        if selected:
            recheck(selected)

    def bind_recipient(self, displayed):
        if self.primary is None:
            try:
                route = self.chat.floating.api._connections.capture('agent')
            except (AttributeError, OSError, ValueError):
                return
            if displayed is None or displayed == {key: getattr(route, key) for key in ('connection_id','model_id','revision')}:
                self.primary = (self.owner(), route_key(route))

    def model_allowed(self, route):
        return self.primary == (self.owner(), route_key(route))

    def target(self, route):
        return (*route_key(route), getattr(self.chat.floating.api._connections, 'document_selection_serial', 0))

    def selection_valid(self, ref):
        if not ref:
            return True
        floating = self.chat.floating
        if floating.paged.active():
            return floating.paged.references.get(ref['id']) == ref
        return ref in (floating.canvas.ai.read_reference, floating.canvas.ai.reference)

    def _binding(self, binding, identifier):
        return (*self.owner(), *binding, self.attachments.get(identifier, {}).get('sha256'))

    def ask(self, row, action, binding, target, reference, full, masked, label, fields, raw):
        identifier = row['id']
        self.pending[identifier] = Decision(uuid.uuid4().hex, identifier, action, self._binding(binding, identifier),
                                           target, deepcopy(reference), full, masked, label, fields, time.monotonic(), raw)
        if action == 'file':
            message = '要参考另一份文件，请选择它。仅在本轮读取所选文件。'
        elif action == 'citation':
            field, value = next(iter(full.items()))
            message = f'本次仅向 {label} 查询{field}「{value}」。允许发送吗？'
        else:
            scope = '选中的' + '、'.join(fields) + '原文' if raw else '本次脱敏材料'
            message = '要把' + scope + '发送给 ' + label + '，用于回答本条问题吗？'
            if identifier in self.attachments:
                message += ' 本次包含参考文件「' + mask_text(self.attachments[identifier]['name']) + '」。'
        self.chat.worker_owners.pop(identifier, None)
        self.chat.readers.discard(identifier)
        self.chat.update(identifier, 'scope_consent', message, status='waiting')
        record_decision(identifier, action, 'waiting')

    def file_ready(self, row, reference, binding):
        identifier = row['id']
        with self.chat.lock:
            if identifier in self.attachments:
                return True
            self.ask(row, 'file', binding, (), reference, {}, {}, '', [], False)
            return False

    def authorize_model(self, row, route, payload, reference, binding, *, raw=False, stage='model'):
        identifier = row['id']
        local = self.chat.floating.api._connections
        confirmed_local = getattr(local, 'confirmed_local_inference', lambda _: False)(route)
        secret = getattr(route, 'api_key', '')
        def protect(text, keep=()):
            if secret:
                text = text.replace(secret, '[凭据已隐藏]')
            return mask_text(text, keep)
        masked = safe_tree(payload, protect)
        for name in ('current_material', 'conversation'):
            if name in payload:
                masked[name] = clean_material(payload[name], protect)
        full = dict(masked)
        if 'current_material' in full:
            keep = list(FIELDS) if confirmed_local else requested_fields(row['question']) if raw and reference else []
            full['current_material'] = clean_material(payload['current_material'], lambda text: protect(text, keep))
        names = field_names(str(full.get('current_material', '')))
        raw = bool(raw and reference and names and not confirmed_local)
        if raw:
            full['conversation'] = '[]'
            masked['conversation'] = '[]'
        target = self.target(route)
        with self.chat.floating.lock, self.chat.lock:
            self.chat.floating.authenticated()
            if binding != self.chat._read_binding() or row['status'] in {'stopped','interrupted'}:
                raise ValueError('当前文档或回合已改变，本次材料未发送。')
            decision = self.approved.pop(identifier, None)
            if decision:
                prior, choice = decision
                if (prior.action != stage or prior.binding != self._binding(binding, identifier) or prior.target != target
                        or time.monotonic() - prior.created > TTL or not self.selection_valid(prior.reference)):
                    raise ValueError('原授权范围已改变或过期，请重新核对。')
                if fingerprint(prior.full) == fingerprint(full) and fingerprint(prior.masked) == fingerprint(masked):
                    self.dispatch[identifier] = prior
                    return prior.full if choice == 'allow' else prior.masked
            if confirmed_local and names:
                self.local_only.add(identifier)
            needs_destination = not confirmed_local and (not self.model_allowed(route) or identifier in self.attachments)
            if decision or raw or needs_destination or identifier in self.reconfirm:
                self.reconfirm.discard(identifier)
                if reference and not self.selection_valid(reference):
                    raise ValueError('原选区已失效，请重新选择后提问。')
                self.ask(row, stage, binding, target, reference, full, masked, recipient(route), names, raw)
                return None
            return full if confirmed_local else masked

    def dispatch_guard(self, row, route, binding, current):
        def guard(phase, state):
            if phase != 'model_start':
                return
            floating = self.chat.floating
            with floating.lock, self.chat.lock:
                floating.authenticated()
                floating.check_source()
                self.recheck_attachment(row['id'])
                if (not current() or row['status'] != 'running' or binding != self.chat._read_binding()
                        or route_key(floating.api._connections.capture('agent')) != route_key(route)):
                    raise ValueError('接收方、文档或回合已改变，本次未派发。')
                if row['id'] in self.local_only and not floating.api._connections.confirmed_local_inference(route):
                    raise ValueError('无法继续确认本机推理，本次原文未发送。')
                decision = self.dispatch.pop(row['id'], None)
                if decision:
                    if (decision.binding != self._binding(binding, row['id']) or decision.target != self.target(route)
                            or time.monotonic() - decision.created > TTL
                            or not self.selection_valid(decision.reference)):
                        raise ValueError('本次授权已过期或选区已改变，请重新核对。')
                    row['scope_dispatched'] = True
                    self.chat.persist()
                    if self.chat.storage_error:
                        row.pop('scope_dispatched', None)
                        raise ValueError('无法记录单次派发决定，本次未发送。')
        return guard

    def _lookup(self, token):
        if not isinstance(token, str) or len(token) != 32:
            raise ValueError('本次决定已失效。')
        item = next((p for p in self.pending.values() if p.token == token), None)
        if not item:
            raise ValueError('本次决定已处理或失效。')
        return item

    def decide(self, token, choice):
        floating = self.chat.floating
        with floating.lock, self.chat.lock:
            floating.authenticated()
            floating.check_source()
            item = self._lookup(token)
            row = self.chat.turns[item.turn]
            if row['stage'] != 'scope_consent' or row['status'] != 'waiting':
                raise ValueError('本次决定已失效。')
            expired = time.monotonic() - item.created > TTL
            stale = item.binding != self._binding(self.chat._read_binding(), item.turn) or not self.selection_valid(item.reference)
            if expired or stale:
                self.forget(item.turn)
                if (item.reference and item.action == 'model' and item.binding[:3] == self.owner()
                        and row['document_id'] == floating.document['id']):
                    self.chat.pending_scope = item.turn
                    self.chat.read_requests.pop(item.turn, None)
                    self.chat.update(item.turn, 'reference', '原授权已失效。请重新引用文字，再点「用新选区回答」。',
                                     status='failed', failure_code='scope_reselect')
                else:
                    self.chat.update(item.turn, 'failed', '授权已过期或内容已改变，请重新核对后继续。', status='failed')
                record_decision(item.turn, item.action, 'expired' if expired else 'stale')
                raise ValueError('本次决定已失效。')
            if choice == 'cancel':
                self.forget(item.turn)
                record_decision(item.turn, item.action, 'cancel')
                messages = {'file': '已取消，文件未读取。', 'plugin_model': '已取消向模型发送资料。插件查询已完成。'}
                self.chat.update(item.turn, 'stopped', messages.get(item.action, '已取消，内容未发送给该接收方。'), status='stopped')
                return {'ok': True}
            if item.action not in ('file', 'citation') and item.target != self.target(floating.api._connections.capture('agent')):
                self.pending.pop(item.turn)
                self.reconfirm.add(item.turn)
                self.chat._start_read(row, self.chat._read_binding()[3], item.reference)
                raise ValueError('接收方已改变，旧许可已失效。请核对新的发送范围。')
            if choice == 'preview' and item.action != 'file':
                return {'ok': True, 'preview': json.dumps(item.full, ensure_ascii=False, indent=2)}
            if choice not in (('choose',) if item.action == 'file' else ('allow','masked') if item.raw else ('allow',)):
                raise ValueError('决定无效。')
            if len(self.chat.readers) >= 2:
                raise ValueError('已有两个问题正在处理，请稍后再试。')
            self.pending.pop(item.turn)
            record_decision(item.turn, item.action, choice)
            if item.action != 'file':
                self.approved[item.turn] = (item, choice)
                self.chat._start_read(row, self.chat._read_binding()[3], item.reference)
                return {'ok': True}
            self.chat.update(item.turn, 'choosing_file', '等待选择参考文件。', status='waiting')
        from .turn_attachment import choose, read_selected
        try:
            selected = choose(floating)
            with floating.lock, self.chat.lock:
                floating.authenticated()
                floating.check_source()
                if (row['stage'] != 'choosing_file' or row['status'] != 'waiting' or
                        self.chat.turns.get(item.turn) is not row or item.binding != self._binding(self.chat._read_binding(), item.turn)
                        or time.monotonic() - item.created > TTL or not self.selection_valid(item.reference)):
                    raise ValueError('文档或回合已改变，未读取所选文件。')
                if selected is None:
                    self.chat.update(item.turn, 'stopped', '已取消，文件未读取。', status='stopped')
                else:
                    self.selected_files[item.turn] = selected
                    self.attachments[item.turn] = read_selected(selected)
                    self.chat._start_read(row, self.chat._read_binding()[3], item.reference)
        except Exception:
            with self.chat.lock:
                if self.chat.turns.get(item.turn) is row and row['status'] == 'waiting':
                    self.forget(item.turn)
                    self.chat.update(item.turn, 'failed', '参考文件未获准读取，请重新选择。', status='failed')
            raise
        return {'ok': True}
