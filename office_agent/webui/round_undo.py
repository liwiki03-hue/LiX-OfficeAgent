import base64
from copy import deepcopy
import hashlib
import io
import json
import os
from pathlib import Path
import re
import threading
import time
import uuid

from docx import Document
from docx.text.paragraph import Paragraph

from ..paged_documents import GuardError, W, changes as document_changes, eligibility, normalize, projection, sha
from ..writes import digest
from .bridge import _read_bounded_json, _reject_links


class RoundUndo:
    def __init__(self, floating):
        self.floating = floating
        self.lock = threading.RLock()
        self.records = {}
        self.loaded = set()
        self.active = None
        self.scope = None

    @property
    def busy(self):
        return self.active is not None

    def _actor(self):
        return hashlib.sha256((getattr(self.floating.api._auth, 'account_key', lambda: '')() or '').encode()).hexdigest()

    def _key(self, turn_id):
        return sha((self._actor() + ':' + str(self.floating.paged.source) + ':' + turn_id).encode())[:32]

    def _paths(self, key):
        root = self.floating.app.editor.audit / 'round-undo'
        _reject_links(root)
        return root / (key + '.json'), root / (key + '.docx')

    def _write(self, path, data):
        _reject_links(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        _reject_links(path)
        temporary = path.with_name(uuid.uuid4().hex + '.tmp')
        try:
            with temporary.open('xb') as stream:
                stream.write(data)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)

    def _save(self, row):
        path, _ = self._paths(row['key'])
        data = {key:value for key,value in row.items() if key not in {'before_data', 'available'}}
        self._write(path, json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def _ensure_scope(self):
        page = self.floating.paged
        scope = (self._actor(),str(page.source),page.source_version,self.floating.conversation.binding)
        if self.scope != scope and not self.busy:
            self.records.clear()
            self.loaded.clear()
            self.scope = scope

    def capture(self, turn_id, revision):
        page = self.floating.paged
        turn = self.floating.conversation.turns.get(turn_id)
        if not turn or turn.get('intent') != 'edit' or turn.get('parent_id') or not page.data or page.snapshot_revision != revision:
            return
        with self.lock:
            self._ensure_scope()
            row = self.records.get(turn_id)
            if row:
                if row.get('after_identity') and row['after_identity'] != page._snapshot_identity(page.data):
                    row.update(phase='blocked', available=False, message='本轮执行期间草稿另有编辑，为保留新内容，不能直接撤销整轮。')
                return
            if len(page.data) > 8*1024*1024:
                return
            while len(self.records) >= 4 or sum(len(item.get('before_data', b'')) for item in self.records.values()) + len(page.data) > 32*1024*1024:
                oldest = min(self.records, key=lambda key:self.records[key]['created'])
                retired = self.records.pop(oldest)
                retired.update(phase='expired', message='此轮恢复点已超出保留范围，不提供正文撤销。')
                self._save(retired)
                _, blob = self._paths(retired['key'])
                _reject_links(blob)
                blob.unlink(missing_ok=True)
            key = self._key(turn_id)
            row = {'schema':1,'key':key,'turn_id':turn_id,'actor':self._actor(),'source':str(page.source),
                'source_version':page.source_version,'source_key':self.floating.source_key,'context_id':self.floating.conversation.binding,
                'created':time.time(),'before_data':bytes(page.data),'before_sha256':sha(page.data),
                'before_identity':page._snapshot_identity(page.data),'before_revision':revision,
                'before_published_revision':page.published_revision,'before_published_id':page.published_id,
                'before_clean':revision==page.published_revision,'phase':'capturing','outputs':[],
                'after_revision':None,'after_identity':None,'available':False,'deleted':[],'message':''}
            _, blob = self._paths(key)
            self._write(blob, row['before_data'])
            self.records[turn_id] = row
            self._save(row)

    def applied(self, turn_id, revision):
        with self.lock:
            row = self.records.get(turn_id)
            if row and row['phase'] not in {'blocked','expired','complete'}:
                row.update(phase='applied', after_revision=revision, available=False)
                self._save(row)

    def changed(self):
        with self.lock:
            for row in self.records.values():
                if row.get('available'):
                    row.update(available=False, message='草稿有后续编辑，正在核对；新内容会保留。')

    def allows_change(self, token):
        with self.lock:
            row = self.records.get(self.active)
            return bool(row and row['phase']=='restoring' and row.get('token')==token)

    def refuse(self, turn_id, message):
        with self.lock:
            row = self.records.get(turn_id)
            if row and row['source']==str(self.floating.paged.source) and row['phase'] in {'ready','blocked'} and not self.busy:
                row.update(available=False,message=message[:500])
                self._save(row)

    def _load(self):
        page = self.floating.paged
        for turn in reversed(list(self.floating.conversation.turns.values())[-80:]):
            if len(self.records) >= 4:
                break
            identifier = turn['id']
            marker = (str(page.source), identifier)
            if identifier in self.records or marker in self.loaded or turn.get('intent') != 'edit' or turn.get('parent_id'):
                continue
            self.loaded.add(marker)
            path, blob = self._paths(self._key(identifier))
            if not path.is_file():
                continue
            try:
                row = _read_bounded_json(path, 128*1024)
                if (row.get('schema') != 1 or row.get('key') != self._key(identifier) or row.get('turn_id') != identifier
                        or row.get('actor') != self._actor() or row.get('source') != str(page.source)
                        or row.get('source_version') != page.source_version or row.get('source_key') != self.floating.source_key
                        or row.get('context_id') != self.floating.conversation.binding or len(row.get('outputs', [])) > 16):
                    continue
                if row.get('phase') not in {'complete','expired'}:
                    _reject_links(blob)
                    if blob.stat().st_size > 8*1024*1024:
                        continue
                    row['before_data'] = blob.read_bytes()
                    if sha(row['before_data']) != row['before_sha256']:
                        continue
                row['available'] = False
                if row.get('phase') in {'restoring','deleting','files_done','uncertain'}:
                    row.update(phase='uncertain', message='上次撤销尚未完成核对；请先核对已有结果，不会重复删除文件。')
                self.records[identifier] = row
            except (OSError, ValueError, KeyError, TypeError):
                continue

    def synced(self):
        page = self.floating.paged
        with self.lock:
            self._ensure_scope()
            self._load()
            identity = page._snapshot_identity(page.data)
            for row in self.records.values():
                if row['source'] != str(page.source) or row['source_version'] != page.source_version:
                    continue
                if row['phase'] == 'applied' and row['after_revision'] == page.revision:
                    row['after_identity'] = identity
                    row['phase'] = 'ready' if row['outputs'] else 'capturing'
                    self._save(row)
                if row['phase'] == 'ready':
                    row['available'] = identity == row['after_identity'] and bool(row['outputs']) and not self.busy
                    row['message'] = '' if row['available'] else '草稿已有后续编辑，为保留新内容，不能直接撤销这次修改。'

    def published(self, turn_id, result, candidate=None):
        with self.lock:
            row = self.records.get(turn_id)
            if not row or row['phase'] in {'blocked','expired','complete'} or not row.get('after_identity'):
                return
            if result['id'] not in [item['id'] for item in row['outputs']]:
                row['outputs'].append({'id':result['id'],'output':result['output'],'sha256':result['output_sha256']})
            page = self.floating.paged
            available = page.snapshot_revision==page.revision and page._snapshot_identity(page.data)==row['after_identity']
            row.update(phase='ready', available=available, message='' if available else '草稿已有后续编辑，为保留新内容，不能直接撤销这次修改。')
            self._save(row)
            if candidate and sha(candidate)==result.get('candidate_sha256'):
                return document_changes(row['before_data'],candidate)

    def public(self, turn_id):
        with self.lock:
            row = self.records.get(turn_id)
            if not row or row['source'] != str(self.floating.paged.source) or row['context_id'] != self.floating.conversation.binding or row['actor']!=self._actor():
                return None
            complete=self.floating.conversation.turns.get(turn_id,{}).get('status')=='complete'
            return {'status':row['phase'],'available':bool(row.get('available')) and complete,'busy':self.active==turn_id and row['phase'] not in {'files_done','uncertain'},
                'recoverable':row['phase'] in {'uncertain','files_done'},
                'message':row.get('message',''),'detail':row.get('error_detail',''),'outputs_count':len(row['outputs'])}

    def _audit(self, item):
        editor = self.floating.app.editor
        if not isinstance(item, dict) or not re.fullmatch(r'[a-f0-9]{32}', item.get('id','')):
            raise GuardError('撤销产物标识无效')
        path = editor.audit / item['id'] / 'audit.json'
        _reject_links(path)
        log = _read_bounded_json(path, 4*1024*1024)
        source = editor.space.file(log['source'])
        expected = source.with_name(source.stem + '.agent-' + item['id'][:8] + source.suffix)
        if (log.get('id') != item['id'] or log.get('kind') != 'paged_candidate' or log.get('status') not in {'verified','rolling_back','rolled_back'}
                or log.get('output') != str(expected) or item.get('output') != str(expected) or item.get('sha256') != log.get('output_sha256')):
            raise GuardError('产物与本轮审计不一致，未撤销')
        backup = editor.audit / item['id'] / ('backup' + source.suffix)
        if log.get('backup') != str(backup) or digest(editor.space.file(str(backup))) != log.get('source_sha256'):
            raise GuardError('备份已改变，未撤销')
        _reject_links(expected)
        if log['status'] == 'rolled_back' or log['status'] == 'rolling_back' and not expected.exists():
            if expected.exists():
                raise GuardError('已撤回位置又出现文件，保留现场')
            return log
        if digest(editor.space.file(str(expected))) != item['sha256']:
            raise GuardError('结果文件已被修改，保留新内容，未撤销')
        return log

    def _guard(self, turn_id, document_id, revision=None):
        page = self.floating.paged
        page.guard(document_id, revision)
        row = self.records.get(turn_id)
        turn = self.floating.conversation.turns.get(turn_id)
        if (not row or not turn or turn.get('intent') != 'edit' or row['source'] != str(page.source) or row['actor']!=self._actor()
                or row['source_version'] != page.source_version or row['source_key'] != self.floating.source_key
                or row['context_id'] != self.floating.conversation.binding):
            raise GuardError('此轮不属于当前编辑草稿，未撤销')
        return page, row

    def begin(self, turn_id, document_id, revision, approved):
        page = self.floating.paged
        with page.lock, self.lock:
            page, row = self._guard(turn_id, document_id, revision)
            if type(approved) is not bool:
                raise GuardError('撤销需要本轮独立确认')
            if not approved:
                return {'ok':True,'cancelled':True}
            if row['phase'] == 'complete':
                return {'ok':True,'complete':True}
            if self.floating.conversation.turns[turn_id]['status']!='complete':
                raise GuardError('本轮修改尚未结束，暂不撤销')
            if self.active == turn_id:
                return {'ok':True,**page.state(),'accepted':True,'token':row['token'],'phase':row['phase'],
                        'data':base64.b64encode(row['before_data']).decode() if row['phase']=='restoring' else None,
                        'draft_revision':row.get('restore_revision',revision),'editable':True,'capabilities':eligibility(row['before_data'])}
            if self.busy or self.floating.app.busy or page.publishing or page.preparing or page.uncertain:
                raise GuardError('当前步骤尚未完成，未撤销')
            if page.snapshot_revision != page.revision:
                raise GuardError('请先同步当前人工编辑，再核对撤销')
            identity = page._snapshot_identity(page.data)
            resuming = row['phase'] in {'uncertain','files_done'}
            if row['phase'] not in {'ready','uncertain','files_done'} or not row.get('before_data'):
                raise GuardError('缺少可靠的编辑前恢复点，未撤销')
            if identity != row['after_identity'] and not (resuming and identity == row['before_identity']):
                row.update(available=False,message='草稿已有后续编辑，为保留新内容，不能直接撤销这次修改。')
                raise GuardError(row['message'])
            group = {item['id'] for item in row['outputs']}
            hashes = {item['sha256'] for item in row['outputs']}
            if any(not value.get('undone') and key not in group and value.get('source_version') in hashes for key,value in self.floating.outputs.items()):
                raise GuardError('后续版本依赖本轮结果，请先撤销后续修改')
            if any(other['created'] > row['created'] and other['source'] == row['source'] and other['phase'] not in {'complete','expired'} and other['outputs'] for other in self.records.values()):
                raise GuardError('已有后续回合依赖当前草稿，请先撤销后续回合')
            for item in reversed(row['outputs']):
                self._audit(item)
            row.update(token=uuid.uuid4().hex, phase='restoring', available=False, restore_revision=revision+1,
                       document_id=document_id, message='正在撤销本次修改，恢复编辑前正文。')
            self.active = turn_id
            self.floating.app.busy = True
            try:
                self._save(row)
            except (OSError,ValueError):
                self.active = None
                self.floating.app.busy = False
                row.update(phase='ready',available=False,message='撤销状态未保存，正文与结果文件保持。')
                raise
            return {'ok':True,**page.state(),'token':row['token'],'data':base64.b64encode(row['before_data']).decode(),
                    'draft_revision':revision+1,'editable':True,'capabilities':eligibility(row['before_data']),'phase':'restoring'}

    def restored(self, turn_id, token, document_id, revision, encoded, manifest):
        page = self.floating.paged
        with page.lock, self.lock:
            self._guard(turn_id,document_id)
            row = self.records.get(turn_id)
            if not row or self.active != turn_id or row.get('token') != token or row['phase'] != 'restoring' or row['document_id'] != document_id or revision != row['restore_revision']:
                raise GuardError('撤销恢复回执已过期，未删除产物')
            page.guard(document_id)
            raw = base64.b64decode(encoded, validate=True) if isinstance(encoded,str) and len(encoded) <= 12*1024*1024 else b''
            data = normalize(raw, page.source.read_bytes())
            if projection(data) != projection(row['before_data']):
                raise GuardError('可见草稿与编辑前恢复点不一致，未删除产物')
            document = Document(io.BytesIO(data))
            paragraphs = [Paragraph(node,document) for node in document.element.body.iter(W+'p')]
            if (not isinstance(manifest,list) or len(manifest)!=len(paragraphs) or len(manifest)>3000
                    or any(not isinstance(item,dict) or set(item)!={'id','text'} or not isinstance(item['id'],str)
                           or not 1<=len(item['id'])<=100 or item['text']!=paragraph.text for item,paragraph in zip(manifest,paragraphs))
                    or len({item['id'] for item in manifest})!=len(manifest)):
                raise GuardError('恢复结构身份不匹配，未删除产物')
            page.changed(document_id, revision, token)
            page.sync(document_id, revision, encoded, manifest)
            row.update(phase='deleting',message='正文已恢复，正在核对并撤回本轮结果文件。')
            row['restored_identity'] = page._snapshot_identity(page.data)
            try:
                self._save(row)
            except (OSError,ValueError):
                row.update(phase='uncertain',message='正文已恢复，撤销状态未保存；结果文件尚未撤回，请核对后恢复。')
                self.active=None
                self.floating.app.busy=False
                raise
            threading.Thread(target=self._delete,args=(turn_id,token),daemon=True).start()
            return {'ok':True,**page.state()}

    def status(self, turn_id, token, document_id):
        with self.floating.paged.lock, self.lock:
            self._guard(turn_id,document_id)
            row = self.records.get(turn_id)
            if not row or row.get('token') != token or row.get('document_id') != document_id:
                raise GuardError('撤销状态归属已改变')
            return {'ok':True,'phase':row['phase'],'message':row.get('message','')}

    def abort(self, turn_id, token, document_id):
        with self.floating.paged.lock, self.lock:
            self._guard(turn_id,document_id)
            row = self.records.get(turn_id)
            if not row or row.get('token') != token or row.get('document_id') != document_id:
                raise GuardError('撤销状态归属已改变')
            if row['phase'] in {'restoring','files_done'}:
                row.update(phase='uncertain',available=False,message='撤销尚未完成全部核对，当前草稿和已知文件状态保留；请核对后恢复。')
                self.active = None
                self.floating.app.busy = False
                self._save(row)
            return {'ok':True,'phase':row['phase']}

    def _delete(self, turn_id, token):
        row = None
        try:
            row = self.records[turn_id]
            for item in reversed(row['outputs']):
                self._audit(item)
            for item in reversed(row['outputs']):
                with self.lock:
                    if self.active != turn_id or row.get('token') != token:
                        raise GuardError('撤销归属已改变，停止后续处理')
                log = self._audit(item)
                if log['status'] != 'rolled_back':
                    log = self.floating.app.editor.rollback(item['id'], confirmed=True)
                with self.lock:
                    if item['id'] not in row['deleted']:
                        row['deleted'].append(item['id'])
                    self._save(row)
                self.floating.app.emit('undone',log)
            with self.floating.paged.lock, self.lock:
                row.update(phase='files_done',message='结果文件已撤回，正在核对编辑器正文。')
                self._save(row)
        except Exception as error:
            with self.lock:
                self.active = None
                self.floating.app.busy = False
                if row is not None:
                    row.update(phase='uncertain',available=False,message='正文或部分文件已处理，撤销尚未全部完成。请核对已有结果后恢复。',error_detail=str(error)[:2000])
                    try:
                        self._save(row)
                    except (OSError,ValueError):
                        row['error_detail'] += '\n撤销状态未能保存；重启后需从既有审计重新核对。'

    def finish(self, turn_id, token, document_id, revision):
        page = self.floating.paged
        with page.lock, self.lock:
            page, row = self._guard(turn_id,document_id,revision)
            if row['phase'] == 'complete':
                return {'ok':True,'complete':True,**page.state()}
            if self.active != turn_id or row.get('token') != token or row['phase'] != 'files_done':
                raise GuardError('撤销尚未完成文件核对')
            if page.snapshot_revision != revision or page._snapshot_identity(page.data) != row['restored_identity']:
                raise GuardError('恢复后正文又有编辑，保留现场，未宣称撤销完成')
            if any(self._audit(item)['status'] != 'rolled_back' for item in row['outputs']):
                raise GuardError('文件撤回结果尚未全部确认')
            page.published_id = row['before_published_id']
            page.published_revision = revision if row['before_clean'] else row['before_published_revision']
            page.publication_rolled_back = True
            page.round_undo_restored = {'turn_id':turn_id,'revision':revision,'before_dirty':not row['before_clean']}
            page.publisher.plans.clear()
            self.floating.app.last_verified_id = page.published_id
            row.update(phase='complete',available=False,message='已撤销这次修改，正文已恢复到本轮修改前。'+('此前人工编辑仍保留在草稿中。' if not row['before_clean'] else ''))
            self.active = None
            self.floating.app.busy = False
            self._save(row)
            message = row['message']
        self.floating.conversation.undo_completed(turn_id,message)
        return {'ok':True,'complete':True,**page.state()}
