from copy import deepcopy
import threading
import uuid


class DocumentAccess:
    levels = {'edit', 'save_copy'}

    def __init__(self, floating):
        self.floating = floating
        self.lock = threading.RLock()
        self.serial = 0
        self.grant = None
        self.level = 'save_copy'

    def reset(self, *, default=False):
        with self.lock:
            self.serial += 1
            self.grant = None
            if default:
                self.level = 'save_copy'

    def state(self):
        with self.lock:
            document = self.floating.document
            grant = self.grant
            if grant and (not document or grant['document_id'] != document['id'] or
                          grant['source_version'] != document['version'] or grant['epoch'] != self.floating.epoch):
                self.reset()
                grant = None
            return {**deepcopy(grant or {'id': '', 'level': self.level}), 'serial': self.serial, 'active': bool(grant)}

    def set(self, document_id, version, epoch, level, confirmed, *, origin='user_selection'):
        self.floating.authenticated()
        self.floating.check_source()
        with self.lock:
            document = self.floating.document
            if (type(epoch) is not int or epoch != self.floating.epoch or not document or
                    document_id != document['id'] or version != document['version'] or not document['ready']):
                raise ValueError('文档已改变，请重新核对编辑方式')
            if not isinstance(level, str) or level not in self.levels or type(confirmed) is not bool:
                raise ValueError('编辑方式无效')
            if not confirmed:
                raise ValueError('未应用本次文档的编辑方式')
            if self.floating.paged.uncertain or self.floating.canvas.uncertain:
                raise ValueError('结果待核实，请先核对副本与审计')
            self.reset()
            self.level = level
            self.grant = {'id': uuid.uuid4().hex, 'level': level, 'document_id': document_id,
                          'source_version': version, 'epoch': epoch, 'origin': origin,
                          'file_identity': list(document['file_identity'])}
            state = self.state()
            self.floating.notice('document_access_changed', state)
            return {'ok': True, 'access': state}

    def for_instruction(self):
        self.floating.authenticated()
        self.floating.check_source()
        with self.lock:
            state = self.state()
            if state['id']:
                return self.require(state['id'], 'edit')
            document = self.floating.document
            return self.set(document['id'], document['version'], self.floating.epoch,
                            self.level, True, origin='user_instruction')['access']

    def require(self, identifier, action, document_id=None):
        self.floating.authenticated()
        self.floating.check_source()
        with self.lock:
            if self.floating.paged.uncertain or self.floating.canvas.uncertain:
                self.reset()
                raise ValueError('结果待核实，已收回自动修改权限；请先核对副本')
            state = self.state()
            document = self.floating.document
            if (not isinstance(identifier, str) or not identifier or state['id'] != identifier or
                    not document or not document['ready'] or state.get('file_identity') != document['file_identity'] or
                    document_id is not None and document_id != document['id']):
                raise ValueError('本次文档的修改授权已失效，操作未自动执行')
            if action not in {'edit', 'publish', 'open_copy'} or (action == 'publish' and state['level'] != 'save_copy'):
                raise ValueError('当前仅修改草稿，请主动保存或切换为自动存副本')
            return state
