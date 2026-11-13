import hashlib
import json
from pathlib import Path
import secrets
import time

from ..rag_evidence import RagEvidence, EvidenceError, _json, _hash, _dir, MAX_BYTES, TTL
from ..writes import Workspace
from ..settings_store import SettingsError


class ChatArchive(RagEvidence):
    directory_prefix = '.office-chat-export-'

    def _save_audit(self, plan):
        return _json({'version':1,'operation':'export_chat','payload_sha256':_hash(plan['payload']),
                      'selection':plan['selection'],'format':plan['format'],'metadata':plan['metadata'],'partial':plan['partial']})


class ChatExports:
    def __init__(self, connections):
        self.connections = connections
        self.entries = {}

    def _binding(self):
        owner = self.connections
        if not owner.authenticated():
            raise SettingsError('not_authenticated')
        book = owner.conversations
        if book is None or book.row() is None:
            raise SettingsError('chat_session_missing')
        return owner.state.epoch, book.current

    def plan(self, payload, choose):
        owner = self.connections
        with owner.lock:
            binding = self._binding()
            partial = type(payload) is dict and payload.get('allow_partial') is True
            if owner.app.busy and not (partial and owner.chat_running):
                return {'ok':False,'reason':'chat_export_in_progress'}
            if type(payload) is not dict or not {'session_id','turn_ids','format','metadata'} <= set(payload) <= {'session_id','turn_ids','format','metadata','content','allow_partial'} or type(payload.get('allow_partial',False)) is not bool or payload.get('content','turns') not in ('answer','turns') or payload['format'] not in ('md','txt') or type(payload['metadata']) is not bool or type(payload['turn_ids']) is not list or not 1 <= len(payload['turn_ids']) <= 100 or any(type(item) is not str for item in payload['turn_ids']) or len(set(payload['turn_ids'])) != len(payload['turn_ids']):
                raise SettingsError('chat_export_invalid')
            book = owner.conversations
            row = book.row()
            if row['id'] != payload['session_id']:
                raise SettingsError('chat_context_changed')
            ids = set(payload['turn_ids'])
            turns = [turn for turn in row['turns'] if turn['id'] in ids]
            if len(turns) != len(ids):
                raise SettingsError('chat_export_invalid')
            if partial and (owner.chat_running != (turns[0]['id'],row['id']) or len(turns) != 1 or not turns[0]['answer']):
                raise SettingsError('chat_export_partial_unavailable')
            identity = hashlib.sha256(json.dumps([(turn['id'],turn['question'],turn['route'],turn['created']) for turn in turns],ensure_ascii=False,sort_keys=True).encode('utf8')).hexdigest()
            blocks = ['未完成快照：仅含用户选择导出时已收到的部分内容，不代表最终回答。'] if partial else []
            for turn in turns:
                if payload.get('content','turns') == 'turns':
                    blocks += ['## 用户' if payload['format'] == 'md' else '用户', turn['question']]
                blocks += ['## 回答' if payload['format'] == 'md' else '回答', turn['answer']]
                if turn['status'] != 'complete' or turn['finish_reason'] != 'stop':
                    blocks.append('状态说明：' + {'failed':'失败，以下内容可能不完整','stopped':'已停止本地等待，远端状态未知','interrupted':'未完成'}.get(turn['status'], {'length':'达到输出预算，未完整','local_limit':'达到本地限制，未完整','content_filter':'服务过滤','unknown':'完整性未确认'}.get(turn['finish_reason'],'需要核对')))
                if 'continuation_of' in turn:
                    blocks.append('关联续写：' + turn['continuation_of']['turn_id'])
                if payload['metadata']:
                    blocks.append('时间：' + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(turn['created'])) + '；接口：' + turn['route']['connection_name'] + '；模型：' + turn['route']['model_id'])
            raw = ('\n\n'.join(blocks) + '\n').encode('utf8')
            if len(raw) > MAX_BYTES:
                raise SettingsError('chat_export_capacity')
            snapshot = hashlib.sha256(json.dumps(row, ensure_ascii=False, sort_keys=True).encode('utf8')).hexdigest()
        folder = choose()
        with owner.lock:
            if self._binding() != binding:
                raise SettingsError('chat_context_changed')
            if not folder:
                return {'ok':False,'reason':'cancelled'}
            if not partial and hashlib.sha256(json.dumps(book.row(),ensure_ascii=False,sort_keys=True).encode('utf8')).hexdigest() != snapshot:
                raise SettingsError('chat_context_changed')
            folder = Path(folder)
            if not folder.is_absolute():
                raise EvidenceError('evidence_path_changed')
            folder_identity = _dir(folder)
            space = Workspace(folder)
            if space.root != folder or _dir(folder) != folder_identity:
                raise EvidenceError('evidence_path_changed')
            archive = ChatArchive(space, lambda key:None)
            archive.artifact_name = 'conversation.' + payload['format']
            session = secrets.token_hex(16)
            archive.ensure_available(session)
            token = archive._id()
            archive._plans[token] = {'session':session,'time':archive._clock(),'operation':'save','payload':raw,'hits':[], 'selection':[turn['id'] for turn in turns],'format':payload['format'],'metadata':payload['metadata'],'partial':partial}
            self.entries = {key:value for key,value in self.entries.items() if time.monotonic()-value['created'] < TTL or value.get('saved')}
            if len(self.entries) >= 8:
                raise SettingsError('chat_export_capacity')
            self.entries[token] = {'archive':archive,'session':session,'binding':binding,'snapshot':snapshot,'created':time.monotonic(),'partial':partial,'identity':identity,'ids':ids}
            path = archive.root / (archive.directory_prefix + token) / archive.artifact_name
            return {'ok':True,'plan':{'plan_id':token,'path':str(path),'turns':len(turns),'format':payload['format'],'metadata':payload['metadata'],'partial':partial,'bytes':len(raw),'sha256':_hash(raw),'confirmation_required':True,'expires_in_seconds':TTL,'wrote_nothing':True}}

    def action(self, identifier, operation, approved=False):
        owner = self.connections
        with owner.lock:
            entry = self.entries.get(identifier) if type(identifier) is str else None
            if not entry or entry['binding'] != self._binding():
                raise SettingsError('chat_export_expired')
            archive, session = entry['archive'], entry['session']
            if operation == 'confirm':
                digest = hashlib.sha256(json.dumps(owner.conversations.row(),ensure_ascii=False,sort_keys=True).encode('utf8')).hexdigest()
                selected = [turn for turn in owner.conversations.row()['turns'] if turn['id'] in entry['ids']]
                identity = hashlib.sha256(json.dumps([(turn['id'],turn['question'],turn['route'],turn['created']) for turn in selected],ensure_ascii=False,sort_keys=True).encode('utf8')).hexdigest()
                if (identity != entry['identity'] if entry['partial'] else digest != entry['snapshot']) or (owner.app.busy and not (entry['partial'] and owner.chat_running)):
                    raise SettingsError('chat_context_changed')
                if (archive.root / (archive.directory_prefix + identifier)).exists():
                    archive._plans.pop(identifier,None)
                    return {'ok':True,'result':{'status':'rejected','code':'chat_export_path_conflict','applied':False,'wrote_nothing':True}}
                outcome = archive.confirm(identifier,session,approved)
                if outcome['status'] == 'saved':
                    entry['saved'] = True
                    outcome['path'] = str(archive.root / outcome['evidence'])
                return {'ok':outcome['status'] in ('saved','rejected'), 'result':outcome}
            if operation == 'open':
                saved = archive._saved_for(identifier,session)
                import os
                os.startfile(str(saved['directory']))
                return {'ok':True}
            if operation == 'rollback_plan':
                result = archive.plan_rollback(identifier,session)
                entry['rollback'] = result['rollback_id']
                return {'ok':True,'plan':result}
            if operation == 'rollback':
                return {'ok':True,'result':archive.confirm_rollback(entry.get('rollback'),session,approved)}
            raise SettingsError('chat_export_invalid')
