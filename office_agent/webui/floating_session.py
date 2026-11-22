from copy import deepcopy
import base64
import hashlib
import json
import os
from pathlib import Path
import queue
import re
import threading
import time
import uuid

from ..instruction_runtime import InstructionRuntime
from ..documents import classify_document, read_word_text, inspect_xlsx
from ..writes import digest
from ..edit_metrics import for_edit
from .bridge import Api, _reject_links, _read_bounded_json
from .document_open import OPEN_FORMATS, read_local_view
from .paged_editor import PagedEditor
from .office_canvas import OfficeCanvas, OfficeCanvasApi
from .document_access import DocumentAccess
from .document_conversation import (
    DocumentConversation, STALE_SELECTION, TURN_LIMIT, edit_history_question,
    intent as conversation_intent, requested_read_scope, scope_choice,
)
from .round_undo import RoundUndo


class FloatingSession:
    def __init__(self, api):
        self.api, self.app = api, api._app
        self.lock = threading.RLock()
        self.admission_lock = threading.Lock()
        self.notifications = queue.Queue(maxsize=1024)
        self.runtime = None
        self.native = None
        self.epoch = 0
        self.seq = 0
        self.events = []
        self.document = None
        self.read_token = None
        self.drop = None
        self.pending_draft_drop = False
        self.outputs = {}
        self.protected = set()
        self.locators = {}
        self.cell_locators = {}
        self.references = []
        self.protected_cells = set()
        self.app.floating_plan_guard = self.check_plan
        self.app.floating_constraints = lambda: ['宿主明确保护的正文段落索引（从0计）：' + str(sorted(self.protected)),
                                                '用户引用的宿主位置（只表示目标，不授予写入）：' + json.dumps(self.references, ensure_ascii=False)]
        self.stage = 'idle'
        self.view = None
        self.queued_for_next_version = []
        self.copy_continuation = None
        self._editor_metrics_wait = None
        self.open_request = None
        self.open_trace = []
        self.local_read_grant = None
        self._open_queue = queue.Queue(maxsize=1)
        self._open_worker = None
        self._open_cancel = None
        self._trace_dirty = False
        self.preferred_paged = False
        self.access = DocumentAccess(self)
        self.paged = PagedEditor(self)
        self.canvas = OfficeCanvas(self)
        self.conversation = DocumentConversation(self)
        self.round_undo = RoundUndo(self)

    def authenticated(self):
        if not self.api._auth.status().get('authenticated', False):
            raise ValueError('请先登录本机账户')

    def notice(self, kind, payload):
        if hasattr(self, 'conversation'):
            self.conversation.event(kind, payload)
        if kind in {'document_ready', 'admission_failed', 'draft_model_error', 'draft_publishing', 'draft_published', 'draft_publish_failed', 'stop_requested', 'stale_action_blocked'}:
            from ..diagnostics import record
            record('document', kind, 'failed' if kind.endswith(('failed', 'error')) else 'recorded',
                   correlation=self.document['id'] if self.document else 'none')
        try:
            self.notifications.put_nowait((kind, payload))
        except queue.Full:
            if self.runtime is not None:
                self.runtime.stop()
            self.app._stop_task.set()

    def reset(self, *, preserve_open=None):
        with self.lock:
            if self._editor_metrics_wait is not None:
                trace, token, _, _ = self._editor_metrics_wait
                trace.stop(token, outcome='cancelled')
                self._editor_metrics_wait = None
            self.conversation.reset()
            self.access.reset(default=True)
            self.paged.close()
            self.canvas.close()
            if self.runtime is not None:
                self.runtime.stop()
            self.runtime = self.app.instruction_runtime = None
            self.epoch += 1
            self.events.clear()
            while True:
                try:
                    self.notifications.get_nowait()
                except queue.Empty:
                    break
            self.document = None
            self.read_token = None
            self.drop = None
            self.pending_draft_drop = False
            self.outputs.clear()
            self.protected.clear()
            self.locators.clear()
            self.cell_locators.clear()
            self.references.clear()
            self.protected_cells.clear()
            self.stage = 'idle'
            self.copy_continuation = None
            self.view = None
            if self._open_cancel and self.open_request is not preserve_open:
                self._open_cancel.set()
            if preserve_open is None:
                while True:
                    try:
                        self._open_queue.get_nowait()['cancel'].set()
                    except queue.Empty:
                        break
            self.local_read_grant = None
            self.open_request = preserve_open

    def dropped(self, paths):
        if len(paths) != 1:
            self.notice('admission_failed', {'text': '一次只接收一份文档；未读取任何正文'})
            return
        with self.lock:
            if self.app.busy:
                self.notice('admission_failed', {'text': '当前任务进行中；停止并等待当前步骤结束后再拖入，当前文档未切换'})
                return
            if not self.paged.switch_allowed() or not self.canvas.switch_allowed():
                self.drop = (uuid.uuid4().hex, str(paths[0]))
                self.pending_draft_drop = True
                self.notice('draft_switch_requested', {'text':'新拖入文件已保留。请先保存副本、保留恢复草稿或放弃草稿，再继续打开。'})
                return
            self.pending_draft_drop = False
            request = {'id': uuid.uuid4().hex, 'drop_id': uuid.uuid4().hex,
                       'path': str(paths[0]), 'phase': 'receiving', 'origin': 'native_drop',
                       'started': time.monotonic(), 'cancel': threading.Event()}
            if self._open_cancel:
                self._open_cancel.set()
            try:
                self._open_queue.get_nowait()
            except queue.Empty:
                pass
            self._open_queue.put_nowait(request)
            self._trace('native_drop_received', request)
            if self._open_worker is None or not self._open_worker.is_alive():
                self._open_worker = threading.Thread(target=self._open_loop, daemon=True)
                self._open_worker.start()
            return {'drop_id': request['drop_id'], 'open_request_id': request['id']}

    def _trace(self, phase, request=None, code='ok'):
        request = request or self.open_request
        if not request:
            return
        record = {'open_request_id': request['id'], 'drop_id': request['drop_id'],
                  'epoch': self.epoch, 'document_id': self.document['id'] if self.document and self.document.get('open_request_id') == request['id'] else None,
                  'stage': phase, 'code': code, 'elapsed_ms': round((time.monotonic() - request['started']) * 1000, 2),
                  'instruction_revision': self.runtime.revision if self.runtime else 0}
        self.open_trace.append(record)
        self.open_trace = self.open_trace[-128:]
        self._trace_dirty = True
        self.notice('open_phase', record)

    def _store_trace(self):
        root = getattr(self.api._auth, 'root', None)
        if not self._trace_dirty or root is None:
            return
        temporary = None
        try:
            import os
            directory = Path(root) / 'diagnostics'
            target = directory / 'open-pipeline.json'
            _reject_links(directory)
            directory.mkdir(parents=True, exist_ok=True)
            _reject_links(target)
            temporary = directory / (uuid.uuid4().hex + '.tmp')
            with temporary.open('x', encoding='utf-8') as stream:
                json.dump({'version': 1, 'events': self.open_trace}, stream, ensure_ascii=False)
            os.replace(temporary, target)
            self._trace_dirty = False
        except (OSError, ValueError):
            self._trace_dirty = False
            self.notice('native_warning', {'text': '打开过程的本机诊断记录未能保存；当前正文状态以界面为准'})
        finally:
            if temporary is not None:
                try:
                    temporary.unlink(missing_ok=True)
                except OSError:
                    pass

    def _open_loop(self):
        while True:
            with self.lock:
                try:
                    request = self._open_queue.get_nowait()
                except queue.Empty:
                    self._open_worker = None
                    return
                self.open_request = request
                self.drop = (request['drop_id'], request['path'])
                self._open_cancel = request['cancel']
            try:
                self._open_native(request)
            except Exception as error:
                with self.lock:
                    if self.open_request is not request or request['cancel'].is_set():
                        continue
                    request['phase'] = 'failed'
                    request['error'] = str(error)[:200] if isinstance(error, ValueError) else '本地打开未完成，请检查格式、文件占用或目录权限后重试'
                    self.stage = 'open_failed'
                    self._trace('open_failed', request, 'open_rejected' if isinstance(error, ValueError) else 'open_failed')
                    self.notice('admission_failed', {'text': request['error'], 'open_request_id': request['id']})

    def _open_native(self, request):
        self.authenticated()
        if request['cancel'].is_set():
            return
        if Path(request['path']).suffix.lower() not in OPEN_FORMATS:
            raise ValueError('悬浮球正文打开支持DOCX/XLSX/PPTX；DOC/XLS/PPT旧格式写入仍暂停')
        reuse = False
        if self.document and self.app.path == Path(request['path']):
            try:
                self.check_source()
                reuse = True
            except ValueError:
                pass
        if reuse and self.document.get('ready'):
            self.document.update(open_request_id=request['id'], drop_id=request['drop_id'])
            request['phase'] = 'ready'
            self._trace('document_session_bound', request, 'existing_document')
            return
        request['phase'] = 'admission_pending'
        self._trace('admission_started', request)
        result = {'ok': True} if reuse else self.admit(request['path'], _opening_request=request)
        with self.lock:
            if request['cancel'].is_set():
                return
            self.open_request = request
            self._open_cancel = request['cancel']
            self.drop = (request['drop_id'], request['path'])
            if not result.get('ok'):
                raise ValueError('文件准入被拒绝；原生候选已保留，可重试或选择其他文件')
            self.document.update(open_request_id=request['id'], drop_id=request['drop_id'],
                                 source_origin='native_drop', capabilities=deepcopy(OPEN_FORMATS[Path(request['path']).suffix.lower()]))
            grant = {'id': uuid.uuid4().hex, 'document_id': self.document['id'], 'version': self.document['version'],
                     'file_identity': list(self.document['file_identity']), 'drop_id': request['drop_id'],
                     'actor': self.api._memory_key() if hasattr(self.api, '_memory_key') else None,
                     'deadline': time.monotonic() + 60}
            self.local_read_grant = grant
            self.document['local_read_grant_reference'] = grant['id']
            self._trace('admission_succeeded', request)
            self._trace('local_read_authorized', request)
            request['phase'] = 'parsing'
            self.stage = 'parsing'
            self._trace('parse_started', request)
            path, version, token = self.app.path, self.document['version'], self.read_token
        parsed = read_local_view(path, request['cancel'].is_set, expected_version=version)
        with self.lock:
            if request['cancel'].is_set() or self.open_request is not request:
                return
            if (self.local_read_grant is not grant or time.monotonic() > grant['deadline'] or
                    self.document['id'] != grant['document_id'] or self.document['version'] != grant['version'] or
                    self.document['file_identity'] != grant['file_identity'] or
                    grant['actor'] != (self.api._memory_key() if hasattr(self.api, '_memory_key') else None)):
                raise ValueError('本地读取授权已过期或归属改变，请重新拖入或明确重试打开')
            self.read(token, True, _payload=parsed)
            request['phase'] = 'partial_ready' if self.view.get('has_more') or any(s.get('cells_truncated') for s in self.view.get('sheets', [])) else 'ready'
            request['process'] = parsed.get('process')
            self.stage = 'idle'
            self.drop = None
            self._trace('document_content_ready', request)
            self._trace('document_session_bound', request)
            if self.runtime and self.runtime.messages and not self.runtime.closed:
                try:
                    self._start_ready_task()
                except Exception:
                    self.notice('instruction_waiting', {'text': '正文已打开，要求仍保留；请选择有效Agent模型后点击继续处理已排队要求'})

    def admit_drop(self):
        self.authenticated()
        with self.lock:
            candidate = self.drop
        if not candidate:
            raise ValueError('没有待准入的原生文件候选')
        if not self.dropped((candidate[1],)):
            raise ValueError('尚未开始打开；请处理当前任务或未保存草稿，原候选仍保留')
        return {'ok': True, 'opening': True}

    def admit(self, candidate, *, _opening_request=None):
        if not self.admission_lock.acquire(blocking=False):
            raise ValueError('已有文件正在准入，请等待当前选择完成')
        try:
            return self._admit(candidate, _opening_request=_opening_request)
        finally:
            self.admission_lock.release()

    def _admit(self, candidate, *, _opening_request=None):
        self.authenticated()
        if self.round_undo.busy:
            raise ValueError('正在核对本轮撤销，请完成后再切换文档')
        self.paged.before_admit()
        self.canvas.before_admit()
        admission_epoch = self.epoch
        if _opening_request and _opening_request['cancel'].is_set():
            return {'ok': False, 'reason': 'open_cancelled'}
        if self.app.busy:
            raise ValueError('请先停止并等待当前任务结束，再切换文档')
        path = Path(candidate)
        if (not path.is_absolute() or str(path).startswith('\\\\') or path.is_symlink() or
                not path.is_file() or path.stat().st_size > 64 * 1024 * 1024):
            raise ValueError('只接收64 MiB以内的真实本地文档，拒绝目录、网络路径及链接')
        classification = classify_document(path)
        expected = {suffix: capability['kind'] for suffix, capability in OPEN_FORMATS.items()}
        if (path.suffix.lower() not in expected or classification.kind != expected[path.suffix.lower()] or
                not classification.readable or classification.protected):
            raise ValueError('格式不匹配、损坏、受保护或暂不支持；未读取正文')
        result = self.app.admit_file(str(path))
        if not result.get('ok'):
            return result
        return self._bind_admitted(path, classification, admission_epoch, _opening_request)

    def _bind_admitted(self, path, classification, admission_epoch, _opening_request):
        with self.lock:
            if self.epoch != admission_epoch or _opening_request and _opening_request['cancel'].is_set():
                if self.epoch == admission_epoch:
                    self.reset(preserve_open=_opening_request)
                self.app._clear_private_state()
                return {'ok': False, 'reason': 'open_cancelled'}
            self.app._memory_context = None
            self.app._multi_scope = None
            self.api._multi_selection.invalidate() if hasattr(self.api, '_multi_selection') else None
            recent_file = [event for event in self.app.poll() if event.get('type') == 'file'][-1:]
            for event in recent_file:
                self.app._push(event)
            self.reset(preserve_open=_opening_request)
            self.document = {'id': uuid.uuid4().hex, 'name': self.app.path.name,
                             'version': digest(self.app.editor.space.file(str(self.app.path))),
                             'kind': classification.kind, 'ready': False,
                             'editable': path.suffix.lower() in {'.docx', '.xlsx', '.pptx'}}
            self.document['capabilities'] = deepcopy(OPEN_FORMATS[path.suffix.lower()])
            self.document['source_origin'] = 'file_picker'
            info = self.app.path.lstat()
            self.document['file_identity'] = [info.st_dev, info.st_ino]
            key = json.dumps([str(self.app.path).casefold(), self.document['version'], self.document['file_identity']])
            self.source_key = hashlib.sha256(key.encode()).hexdigest()
            self.read_token = uuid.uuid4().hex
            self.notice('document_admitted', deepcopy(self.document))
            return {'ok': True}

    def choose(self):
        self.authenticated()
        if not self.admission_lock.acquire(blocking=False):
            raise ValueError('已有文件正在选择或准入，请完成当前选择')
        try:
            if self.app.busy:
                raise ValueError('任务运行中，不能切换文件')
            from webview import FileDialog
            selected = self.app.window.create_file_dialog(FileDialog.OPEN,
                file_types=('文档 (*.docx;*.xlsx;*.pptx)',))
            if not selected:
                return {'ok': False, 'reason': 'cancelled'}
            result = self._admit(selected[0])
            if result.get('ok') and self.document:
                result = {**result, 'document_id': self.document['id'], 'epoch': self.epoch}
            return result
        finally:
            self.admission_lock.release()

    def check_source(self):
        if not self.document or not self.app.editor or not self.app.path:
            raise ValueError('请先选择文档')
        path = self.app.editor.space.file(str(self.app.path))
        info = path.lstat()
        if digest(path) != self.document['version'] or [info.st_dev, info.st_ino] != self.document['file_identity']:
            self.access.reset()
            if self.runtime:
                self.runtime.stop()
            self.app.pending_id = None
            self.document['ready'] = False
            self.stage = 'changed'
            raise ValueError('文件已在外部改变，请重新选择并授权读取；旧计划不可执行')
        return path

    def read(self, token, allowed, offset=0, *, _payload=None):
        self.authenticated()
        if type(allowed) is not bool or token != self.read_token or type(offset) is not int or not 0 <= offset <= 100000:
            raise ValueError('读取确认已失效')
        if not allowed:
            self.notice('read_denied', {'text': '未读取正文；可稍后重新授权'})
            return {'ok': True, 'denied': True}
        path = self.check_source()
        document, source_epoch, editor = self.document, self.epoch, self.app.editor
        locators, cell_locators = {}, {}
        if path.suffix.lower() == '.docx':
            view = _payload['view'] if _payload is not None else read_word_text(path, offset=offset, limit=20, max_chars=10000)
            from docx import Document
            from docx.text.paragraph import Paragraph
            positions = {int(key): value for key, value in _payload['positions'].items()} if _payload is not None else {}
            paragraph_index = 0
            for index, block in enumerate(Document(path).iter_inner_content() if _payload is None else []):
                if isinstance(block, Paragraph):
                    positions[index] = paragraph_index
                    paragraph_index += 1
            for block in view['blocks']:
                if block['kind'] == 'paragraph' and not block.get('text_truncated'):
                    identifier = hashlib.sha256((document['id'] + ':' + document['version'] + ':' + str(block['index'])).encode()).hexdigest()
                    locators[identifier] = positions[block['index']]
                    block['locator_id'] = identifier
                    block['paragraph_index'] = positions[block['index']]
        elif path.suffix.lower() == '.xlsx':
            view = _payload['view'] if _payload is not None else inspect_xlsx(path)
            remaining = 400
            for sheet in view['sheets']:
                cells = sheet['cells'][:min(100, remaining)]
                sheet['cells_truncated'] = sheet.get('cells_truncated', False) or len(cells) < len(sheet['cells'])
                sheet['cells'] = cells
                remaining -= len(cells)
                for cell in cells:
                    key = json.dumps([document['id'], document['version'], sheet['title'], cell['coordinate']])
                    identifier = hashlib.sha256(key.encode()).hexdigest()
                    cell_locators[identifier] = (sheet['title'], cell['coordinate'])
                    cell['locator_id'] = identifier
        elif path.suffix.lower() == '.pptx':
            from .document_open import parse_payload
            view = _payload['view'] if _payload is not None else parse_payload(path, offset)['view']
        else:
            raise ValueError('本主视图支持DOCX/XLSX/PPTX')
        with self.lock:
            if (self.document is not document or self.epoch != source_epoch or self.app.editor is not editor or
                    self.app.path != path or self.read_token != token):
                raise ValueError('文档归属已改变；旧读取结果已作废')
            self.authenticated()
            self.check_source()
            self.locators.update(locators)
            self.cell_locators.update(cell_locators)
            self.document['ready'] = True
            self.document['coverage'] = {'loaded_blocks': len(view.get('blocks', [])),
                                         'loaded_cells': sum(len(sheet['cells']) for sheet in view.get('sheets', [])),
                                         'offset': offset,
                                         'model_context': 'Agent按需读取获准范围，未声明全文已发送或理解'}
            self.view = deepcopy(view)
            self.app._read_authorized_path = path
            self.conversation.restore()
            self.conversation.resume_opening()
            self.notice('document_ready', {'view': view, 'version': document['version']})
        return {'ok': True, 'view': view}

    def protect(self, identifier, protected):
        self.authenticated()
        self.check_source()
        if type(protected) is not bool or identifier not in self.locators and identifier not in self.cell_locators:
            raise ValueError('保护位置无效，请重新读取')
        if self.app.busy and self.runtime and self.runtime.executing:
            raise ValueError('当前副本事务已开始；请在完成后设置下一轮保护')
        targets = self.protected if identifier in self.locators else self.protected_cells
        location = self.locators.get(identifier) if identifier in self.locators else self.cell_locators[identifier]
        if protected:
            targets.add(location)
        else:
            targets.discard(location)
        if self.runtime:
            self.runtime.submit(uuid.uuid4().hex, '宿主保护位置：段落（从0计）' + str(sorted(self.protected)) + '；单元格' + str(sorted(self.protected_cells)))
            with self.runtime.lock:
                self.runtime.protected_targets = self.protection_targets()
                if protected:
                    self.runtime.needs_clarification = False
                    self.stage = 'idle'
            if protected and not self.app.busy:
                with self.app._lock:
                    self.app._start_task(self.runtime.messages[0]['text'], clarified=True)
        return {'ok': True, 'protected': sorted(self.protected)}

    def reference(self, document_id, version, identifier):
        self.authenticated()
        self.check_source()
        if not self.document['ready'] or document_id != self.document['id'] or version != self.document['version']:
            raise ValueError('引用位置已失效，请核对当前文档')
        if identifier in self.locators:
            index = self.locators[identifier]
            block = next(item for item in self.view['blocks'] if item.get('locator_id') == identifier)
            record = {'paragraph_index': index, 'text_sha256': hashlib.sha256(block['text'].encode()).hexdigest()}
            label = '正文第' + str(index + 1) + '段'
        elif identifier in self.cell_locators:
            sheet, cell = self.cell_locators[identifier]
            record = {'sheet': sheet, 'cell': cell}
            label = sheet + '!' + cell
        else:
            raise ValueError('此位置未完整读取或不能唯一定位，请先读取该范围')
        record.update(document_id=document_id, document_version=version, locator_id=identifier)
        self.references = (self.references + [record])[-8:]
        return {'ok': True, 'label': label, 'reference': record}

    def protection_targets(self):
        return [{'document_version': self.document['version'], 'paragraph_index': index}
                for index in sorted(self.protected)] + [
            {'document_version': self.document['version'], 'sheet': sheet, 'cell': cell}
            for sheet, cell in sorted(self.protected_cells)]

    def resolve_protection(self):
        self.authenticated()
        for identifier, index in self.locators.items():
            if index in self.protected:
                return self.protect(identifier, True)
        for identifier, cell in self.cell_locators.items():
            if cell in self.protected_cells:
                return self.protect(identifier, True)
        raise ValueError('请先在文档视图勾选明确保护的位置；不猜测段落或单元格')

    def check_plan(self, payload):
        if self.canvas.blocks_original():
            raise ValueError('交互草稿已改变，磁盘原稿的旧Agent计划失效；请保存并选择副本后重新规划')
        request = self.app.editor.plans[payload['id']].request
        if request.kind in {'word_replace', 'word_format'}:
            identifier = next((key for key, index in self.locators.items() if index == request.paragraph_index), None)
        elif request.kind == 'excel_set':
            identifier = next((key for key, cell in self.cell_locators.items() if cell == (request.sheet, request.cell.upper())), None)
        else:
            identifier = None
        if identifier:
            payload['location'] = {'locator_id': identifier, 'document_version': self.document['version']}
        if self.protected_cells:
            if request.kind == 'excel_set':
                affected = {(request.sheet, request.cell.upper())}
            elif request.kind == 'excel_range_set':
                from openpyxl.utils.cell import range_boundaries, get_column_letter
                left, top, right, bottom = range_boundaries(request.first_cell + ':' + request.last_cell)
                affected = {(request.sheet, get_column_letter(column) + str(row))
                            for column in range(left, right + 1) for row in range(top, bottom + 1)}
            else:
                raise ValueError('当前操作无法证明不影响受保护单元格，请改用有界单元格修改')
            if affected & self.protected_cells:
                raise ValueError('计划触及明确保护的单元格，已拒绝')
        if request.kind in {'word_replace', 'word_batch_replace'}:
            from ..professional_edits import _protect
            from types import SimpleNamespace
            _protect(payload['before'], payload['after'], SimpleNamespace(
                preserve_numbers=True, preserve_citations=True, preserve_formulas=True, protected_literals=[]))
        if not self.protected:
            return
        if request.kind in {'word_replace', 'word_format'} and request.paragraph_index not in self.protected:
            return
        if request.kind == 'word_batch_replace':
            from docx import Document
            paragraphs = Document(self.app.path).paragraphs
            if all(request.old not in paragraphs[index].text for index in self.protected):
                return
        raise ValueError('计划可能触及宿主明确保护的段落；已拒绝，请改用不跨保护段落的局部修改')

    def persist(self, record):
        from ..rag import SECRET
        if SECRET.search(record['text']):
            raise ValueError('要求中疑似包含凭据；未保存、未发送，请移除后重试')
        actor = hashlib.sha256((getattr(self.api._auth, 'account_key', lambda: '')() or '').encode()).hexdigest()
        scope = base64.urlsafe_b64encode(hashlib.sha256((actor + ':' + self.source_key).encode()).digest()).decode().rstrip('=')
        directory = self.app.editor.audit / 'inbox' / scope / self.runtime.session_id
        _reject_links(directory)
        directory.mkdir(parents=True, exist_ok=True)
        _reject_links(directory)
        target = directory / (str(record['revision']) + '.json')
        payload = dict(record)
        payload['text_sha256'] = hashlib.sha256(record['text'].encode()).hexdigest()
        payload['actor'] = actor
        with target.open('x', encoding='utf-8') as stream:
            json.dump(payload, stream, ensure_ascii=False)
            stream.flush()
            import os
            os.fsync(stream.fileno())

    def restore_drafts(self):
        import os
        self.authenticated()
        self.check_source()
        actor = hashlib.sha256((getattr(self.api._auth, 'account_key', lambda: '')() or '').encode()).hexdigest()
        scope = base64.urlsafe_b64encode(hashlib.sha256((actor + ':' + self.source_key).encode()).digest()).decode().rstrip('=')
        root = self.app.editor.audit / 'inbox' / scope
        _reject_links(root)
        if not root.exists():
            return {'ok': True, 'tasks': [], 'limited': False}
        directories, limited = [], False
        examined = 0
        with os.scandir(root) as entries:
            for entry in entries:
                examined += 1
                if examined > 64:
                    limited = True
                    break
                if re.fullmatch('[a-f0-9]{32}', entry.name) and not entry.is_symlink() and entry.is_dir(follow_symlinks=False):
                    directories.append(Path(entry.path))
        tasks, remaining = [], 1024 * 1024
        for directory in sorted(directories, key=lambda item: item.lstat().st_mtime_ns, reverse=True)[:5]:
            _reject_links(directory)
            records = []
            complete = True
            for number in range(1, 65):
                path = directory / (str(number) + '.json')
                _reject_links(path)
                if not path.exists():
                    for later in range(number+1,65):
                        later_path=directory/(str(later)+'.json')
                        _reject_links(later_path)
                        if later_path.exists():raise ValueError('历史要求记录存在缺口，拒绝恢复，避免遗漏修改或取消')
                    break
                size = path.stat().st_size
                if size > 20000 or size > remaining:
                    limited = True
                    complete = False
                    break
                remaining -= size
                record = _read_bounded_json(path, 20000)
                text = record.get('text')
                if (record.get('actor') != actor or record.get('document_version') != self.document['version'] or
                        record.get('session_id') != directory.name or type(record.get('revision')) is not int or record.get('revision') != number or
                        type(text) is not str or len(text) > 4000 or
                        hashlib.sha256(text.encode()).hexdigest() != record.get('text_sha256')):
                    raise ValueError('历史要求记录已改变或不匹配，拒绝恢复')
                records.append(record)
            extra=directory/'65.json'
            _reject_links(extra)
            if extra.exists():
                limited = True
                complete = False
            if records and complete:
                from ..requirement_history import project_requirements
                requirements=project_requirements(records)
                messages=[row['text'] for row in requirements if not row['cancelled']]
                tasks.append({'id': directory.name, 'messages': messages,
                              'cancelled_count':sum(row['cancelled'] for row in requirements),
                              'changed_count':sum(bool(row['history']) for row in requirements),
                              'status': '历史要求；执行状态须核对，不恢复运行、读取或写入批准'})
        return {'ok': True, 'tasks': tasks, 'limited': limited or len(directories) > 5}

    def _start_ready_task(self):
        if self.canvas.blocks_original():
            raise ValueError('请保存并选择表格/演示副本后重新生成Agent方案')
        if self.preferred_paged and self.app.path and self.app.path.suffix.lower()=='.docx':
            self.notice('instruction_waiting',{'text':'要求已保存。分页草稿载入后，请引用目标文字并处理已保存要求；尚未调用模型。'})
            return
        if self.paged.active():
            raise ValueError('分页编辑已启用，请引用当前草稿后发送；不会从磁盘原稿重启模型任务')
        self.check_source()
        self.access.for_instruction()
        self.app._pending_clarification = None
        self.app.pending_id = None
        self.api._connections.state.mode = 'agent'
        with self.app._lock:
            self.app._start_task(self.runtime.messages[0]['text'], clarified=True)
        if self.app.busy:
            self._trace('agent_started')

    def submit(self, session_id, client_id, text, open_request_id=None, document_id=None, version=None):
        if self.canvas.blocks_original():
            raise ValueError('表格/演示交互草稿尚未作为Agent输入；请保存副本并选择该副本继续，避免修改旧原稿')
        if self.paged.active():
            raise ValueError('请在分页编辑器引用最新草稿后发送，不沿用原磁盘内容的旧入口')
        self.authenticated()
        with self.lock:
            waiting = bool(self.document and not self.document['ready'] and self.open_request and self.open_request['phase'] in {'parsing', 'failed'})
            if not self.document or not self.document['ready'] and not waiting:
                raise ValueError('请先授权并查看文档')
            if self.document.get('open_request_id') and (open_request_id != self.document['open_request_id'] or
                    document_id != self.document['id'] or version != self.document['version']):
                raise ValueError('文件打开归属已改变，旧要求未提交；请核对当前文件')
            self.check_source()
            if self.runtime is not None and session_id != self.runtime.session_id:
                if client_id not in self.runtime.ids or session_id:
                    raise ValueError('当前任务归属已改变，请重新核对要求')
            if self.runtime is None:
                if session_id:
                    raise ValueError('任务已失效')
                from ..agent import configured_decision_limit
                self.runtime = InstructionRuntime(self.document['version'], persist=self.persist, notify=self.notice,
                                                  model_limit=configured_decision_limit())
                bound_runtime = self.runtime
                bound_runtime.notify = lambda kind, payload: self.notice(kind, payload) if self.runtime is bound_runtime else None
                bound_runtime.protected_targets = self.protection_targets()
                self.app.instruction_runtime = self.runtime
            duplicate = client_id in self.runtime.ids
            with self.runtime.lock:
                record = self.runtime.submit(client_id, text)
                if (not duplicate and re.search(r'段|表格|单元格|[A-Z]{1,3}[1-9][0-9]*', text) and
                        re.search(r'(?:不[要得]?|别|勿).{0,12}(?:改|动)|(?:保持|保留).{0,12}(?:不变|原样)', text)):
                    self.runtime.needs_clarification = True
                    self.stage = 'needs_clarification'
                    self.notice('needs_clarification', {'text': '要保护哪个位置？请在文档视图勾选明确段落或单元格。此要求已收到；确认位置前暂停新方案，不沿用旧计划。'})
            if duplicate:
                return {'ok': True, 'record': record, 'session_id': self.runtime.session_id, 'duplicate': True}
            self._trace('user_instruction_accepted')
            if waiting:
                self.notice('instruction_waiting', {'text': '要求已保存，等待文档解析；尚未发送模型请求'})
            elif not self.app.busy and not self.runtime.needs_clarification:
                self._start_ready_task()
            return {'ok': True, 'record': record, 'session_id': self.runtime.session_id}

    def confirm(self, session_id, plan_id, revision, version, approved, access_id=None):
        self.authenticated()
        authorization = self.access.require(access_id, 'publish') if access_id is not None else None
        if authorization:
            approved = True
        if type(approved) is not bool or not self.runtime or session_id != self.runtime.session_id:
            raise ValueError('确认归属已失效')
        if not approved:
            with self.runtime.lock:
                self.runtime.plans.pop(plan_id, None)
            if self.app.pending_id == plan_id:
                self.app.pending_id = None
            self.notice('plan_cancelled', {'id': plan_id})
            return {'ok': True, 'cancelled': True}
        self.check_source()
        route = getattr(self.runtime, 'model_route', None)
        if route is not None and self.api._connections.capture('agent').revision != route['revision']:
            self.runtime.plans.clear()
            raise ValueError('模型连接配置已改变，请重新生成方案并确认')
        with self.runtime.transaction_start(plan_id, revision, version):
            if self.app.busy:
                self.runtime.executing = False
                raise ValueError('当前任务仍在运行，请等方案完成后确认')
            if plan_id not in self.app.editor.plans:
                self.runtime.executing = False
                raise ValueError('计划不存在')
            self.check_plan(self.app.editor.plans[plan_id].public())
            self.app.pending_id = plan_id
            self.stage = 'executing'
            self.executing_revision = revision
            self._trace('access_authorized' if authorization else 'human_confirmed')
            if authorization:
                self.notice('document_access_used', {'plan_id': plan_id, 'access': authorization})
            result = self.app.execute()
            if not result.get('ok'):
                self.runtime.executing = False
        return result

    def stop(self, *, announce=True):
        self.authenticated()
        self.conversation.stop()
        self.copy_continuation = None
        had_access = bool(self.access.state()['id'])
        publication_stop = self.paged.stop()
        self.access.reset()
        with self.canvas.lock:
            self.canvas.ai.invalidate()
            self.canvas.plans.clear()
        was_opening = bool(self.open_request and self.open_request['phase'] in {'parsing', 'receiving', 'admission_pending'})
        if self._open_cancel:
            self._open_cancel.set()
        with self.lock:
            while True:
                try:
                    self._open_queue.get_nowait()['cancel'].set()
                except queue.Empty:
                    break
        if self.open_request and self.open_request['phase'] in {'parsing', 'receiving', 'admission_pending'}:
            self.open_request['phase'] = 'cancelled'
            self._trace('open_cancelled', code='cancelled')
        if self.runtime:
            self.runtime.stop()
        self.app.pending_id = None
        self.stage = 'stopped'
        result = self.app.cancel_task()
        if publication_stop:
            text = ('已停止后续派发，正在停止当前副本校验。' if publication_stop == 'cancelling'
                    else '已停止后续派发，当前副本已进入写入，正在核验结果。')
            self.notice('stop_requested', {'text': text, 'publication_pending': True})
            return {'ok': True, 'stopped': True, 'publication_stop': publication_stop}
        if announce:
            self.notice('stop_requested', {'text': '已停止后续派发。已发送的远端请求可能继续处理或计费。已开始的副本事务将安全收尾。'})
        if result.get('reason') == 'no_active_model_task' and (was_opening or had_access):
            return {'ok': True, 'local_open_cancelled': was_opening, 'stopped': True}
        return result

    def start_again(self, fresh=False):
        if self.paged.active():
            return self.paged.start_again(fresh)
        self.authenticated()
        self.check_source()
        if self.app.busy:
            raise ValueError('请等待当前步骤结束')
        if self.runtime and self.runtime.uncertain:
            raise ValueError('结果不确定，不能在本会话重试。请先核对审计与产物，再明确重新选择文档')
        if fresh is True:
            if self.runtime:
                self.runtime.stop()
            self.runtime = self.app.instruction_runtime = None
            self.app.pending_id = None
            self.notice('new_round', {'text': '新一轮已准备；旧计划和批准失效，当前文档读取许可保留。'})
        elif self.runtime and not self.runtime.closed and not self.runtime.uncertain:
            with self.app._lock:
                self.app._start_task(self.runtime.messages[0]['text'], clarified=True)
        else:
            raise ValueError('请先开始新一轮并填写要求')
        return {'ok': True}

    def rollback(self, output_id, approved):
        self.authenticated()
        if type(approved) is not bool or output_id not in self.outputs or self.outputs[output_id].get('undone'):
            raise ValueError('撤销入口已失效')
        output = self.outputs[output_id]
        if any(not row.get('undone') and row.get('source_version') == output.get('output_sha256') for row in self.outputs.values()):
            raise ValueError('后续版本依赖此副本，请先按倒序撤销后续输出')
        if not approved:
            return {'ok': True, 'cancelled': True}
        if self.app.busy:
            raise ValueError('任务运行中不能撤销')
        self.app.last_verified_id = output_id
        return self.app.undo()

    def begin_round_undo(self, turn_id, document_id, revision, approved):
        try:
            return self.round_undo.begin(turn_id, document_id, revision, approved)
        except (ValueError, OSError) as error:
            self.round_undo.refuse(turn_id, '核对撤销所需文件时遇到问题，正文与结果文件保留。' if isinstance(error,OSError) else str(error))
            raise

    def select_output(self, output_id, read_open=False, access_id=None):
        self.authenticated()
        if type(read_open) is not bool:
            raise ValueError('打开副本的读取选项无效')
        if not self.admission_lock.acquire(blocking=False):
            raise ValueError('已有文件正在选择或准入，请完成当前选择')
        metric_handoff = None
        handoff_trace = None
        handoff_span = None
        try:
            row = self.outputs.get(output_id)
            if not row or row.get('undone') or digest(Path(row['output'])) != row.get('output_sha256'):
                raise ValueError('副本不存在、已撤销或内容已改变，请重新选择文件')
            access = self.access.require(access_id, 'open_copy') if access_id is not None else None
            consent = row.get('continue_editing')
            serial = self.access.serial
            level = access['level'] if access and row.get('source_version') == self.document['version'] else None
            if consent and consent['serial'] == serial and row.get('source_version') == self.document['version']:
                level = consent['level']
            if self.protected or self.protected_cells:
                level = None
            continuation = []
            model_limit = 12
            if row.get('automatic_requirements'):
                from ..requirement_history import active_requirements
                runtime = self.runtime
                if not access or not level or not runtime or runtime.closed or self.stage == 'stopped':
                    raise ValueError('当前要求已停止或授权已改变，未自动打开或继续执行副本')
                clients = {record['id']: client for client, record in runtime.ids.items() if record.get('kind') != 'requirement_change'}
                continuation = [{**item, 'client_id': clients.get(item['id'], item['id'])} for item in active_requirements(runtime) if item['version'] > runtime.applied]
                if not continuation:
                    raise ValueError('当前没有待执行要求，未自动继续旧请求')
                model_limit = None if runtime.model_limit is None else max(0, runtime.model_limit - runtime.model_calls)
            saved = deepcopy(self.outputs)
            conversation = self.conversation.state()
            queued = [m['text'] for m in self.runtime.messages if m['revision'] > self.runtime.applied] if self.runtime else []
            handoff_trace = for_edit(self, row.get('conversation_id')) if read_open else None
            if handoff_trace is not None:
                handoff_span = handoff_trace.start('reopen')
                metric_handoff = self.conversation.take_edit_metrics()
            result = self._admit(row['output'])
            if result.get('ok'):
                self.outputs = saved
                self.queued_for_next_version = queued
                if self.document['version'] != row.get('output_sha256'):
                    raise ValueError('副本在打开期间改变，未沿用修改权限，请重新选择文件')
                if read_open:
                    self.read(self.read_token, True)
                    self.conversation.rebind(conversation)
                    if metric_handoff is not None:
                        self.conversation.restore_edit_metrics(metric_handoff)
                        metric_handoff = None
                    if handoff_trace is not None:
                        handoff_trace.stop(handoff_span)
                        handoff_span = None
                    if level and self.access.serial == serial + 1:
                        self.access.set(self.document['id'], self.document['version'], self.epoch, level, True)
                        if continuation:
                            self.runtime = InstructionRuntime(self.document['version'], persist=self.persist,
                                                              notify=self.notice, model_limit=model_limit)
                            self.app.instruction_runtime = self.runtime
                            for item in continuation:
                                self.runtime.submit(item.get('client_id') or item['id'], item['text'])
                            self.copy_continuation = self.document['id']
                            if handoff_trace is not None and handoff_trace.current():
                                self._editor_metrics_wait = (handoff_trace, handoff_trace.start('editor'),
                                                             self.document['id'], self.document['version'])
                            self.queued_for_next_version = []
                            self.conversation.update(self.conversation.active_task, 'loading_editor', '副本已验证并准入，正在载入编辑器。')
            return result
        finally:
            if metric_handoff is not None:
                self.conversation.restore_edit_metrics(metric_handoff)
            if handoff_trace is not None:
                handoff_trace.stop(handoff_span, outcome='error')
            self.admission_lock.release()

    def select_history_output(self, identifier):
        from .manual_history import checked_output, resolve
        self.authenticated()
        if self.app.editor is None:
            raise ValueError('请先选择当前目录中的文档')
        if not self.admission_lock.acquire(blocking=False):
            raise ValueError('已有文件正在准入，请完成当前选择')
        try:
            epoch, editor = self.epoch, self.app.editor
            expected = resolve(editor, identifier)['output_sha256']
            path = checked_output(editor, identifier)
            if self.epoch != epoch or self.app.editor is not editor:
                raise ValueError('当前文档已切换，旧副本打开请求已失效')
            result = self._admit(str(path))
            if result.get('ok') and self.document.get('version') != expected:
                self.reset()
                raise ValueError('重新准入期间副本已改变；请重新选择当前文件')
            return result
        finally:
            self.admission_lock.release()

    def preview_output(self, output_id, allowed, offset=0):
        self.authenticated()
        row = self.outputs.get(output_id)
        if allowed is not True or type(offset) is not int or not 0 <= offset <= 100000 or not row or row.get('undone'):
            raise ValueError('副本读取未获确认或记录失效')
        path = self.app.editor.space.file(row['output'])
        source_epoch = self.epoch
        if digest(path) != row.get('output_sha256'):
            raise ValueError('副本内容已改变，请重新选择并授权读取')
        from .document_open import parse_payload
        view = parse_payload(path, offset)['view']
        if path.suffix.lower() == '.xlsx':
            remaining = 400
            for sheet in view['sheets']:
                cells = sheet['cells'][:min(100, remaining)]
                sheet['cells_truncated'] = sheet.get('cells_truncated', False) or len(cells) < len(sheet['cells'])
                sheet['cells'] = cells
                remaining -= len(cells)
        if self.epoch != source_epoch or self.outputs.get(output_id) is not row:
            raise ValueError('当前文档已切换，旧副本读取结果作废')
        if digest(self.app.editor.space.file(row['output'])) != row.get('output_sha256'):
            raise ValueError('副本在读取期间改变，拒绝沿用此次内容')
        return {'ok': True, 'view': view, 'output_id': output_id, 'document_version': row['output_sha256']}

    def collect(self, cursor=0):
        self.authenticated()
        if type(cursor) is not int or cursor < 0:
            raise ValueError('事件游标无效')
        with self.lock:
            incoming = self.app.poll()
            while True:
                try:
                    kind, payload = self.notifications.get_nowait()
                    incoming.append({'type': kind, 'payload': payload})
                except queue.Empty:
                    break
            for event in incoming:
                kind = event['type']
                if (kind == 'worker_finished' and not self.paged.active() and self.runtime and not self.runtime.closed and
                        self.runtime.applied < self.runtime.revision and not self.app.busy and
                        not self.runtime.executing and not self.runtime.needs_clarification):
                    self.check_source()
                    with self.app._lock:
                        self.app._start_task(self.runtime.messages[0]['text'], clarified=True)
                if kind == 'status' and event.get('level') == 'danger':
                    self.stage = 'uncertain' if self.runtime and self.runtime.executing else 'failed'
                if kind == 'plan' and (not self.runtime or event['plan']['id'] not in self.runtime.plans):
                    continue
                if kind == 'plan':
                    self._trace('plan_ready')
                if kind == 'output':
                    payload = event['payload']
                    if event.get('undone'):
                        if payload['id'] in self.outputs:
                            self.outputs[payload['id']]['undone'] = True
                            self.paged.output_rolled_back(payload['id'])
                        self.stage = 'rolled_back'
                    elif payload.get('status') == 'verified':
                        row = deepcopy(payload)
                        row['instruction_revision'] = getattr(self, 'executing_revision', 0)
                        row['source_version'] = self.document['version']
                        if row.get('output'):
                            row['output_sha256'] = digest(Path(row['output']))
                        self.outputs[payload['id']] = row
                        self.stage = 'verified'
                        self._trace('output_verified')
                self.seq += 1
                self.events.append({'schema_version': 1, 'event_seq': self.seq, 'event_id': uuid.uuid4().hex,
                                    'epoch': self.epoch, 'session_id': self.runtime.session_id if self.runtime else '',
                                    'run_id': self.runtime.run_id if self.runtime else '',
                                    'step_id': self.runtime.step_id if self.runtime else '',
                                    'instruction_revision': self.runtime.revision if self.runtime else 0,
                                    'document_version': self.document['version'] if self.document else '', 'event': event})
            self.events = self.events[-512:]
            if self.runtime and self.runtime.executing and not self.app.busy:
                self.runtime.transaction_finished(uncertain=self.stage != 'verified')
            plans = list(self.runtime.plans.values()) if self.runtime else []
            self._store_trace()
            if self.native is not None:
                from .orb_visuals import display_status
                self.native.set_status(display_status(self.stage, self.app.busy,
                    bool(self.open_request and self.open_request.get('phase') in {'receiving','admission_pending','parsing'}),
                    bool(plans or any(p['status']=='pending' for p in self.paged.proposals.values()))))
            return {'ok': True, 'epoch': self.epoch, 'cursor': self.seq,
                    'conversation': self.conversation.state(),
                    'access': self.access.state(),
                    'gap': bool(self.events and cursor < self.events[0]['event_seq'] - 1),
                    'events': deepcopy([e for e in self.events if e['event_seq'] > cursor]),
                    'document': deepcopy(self.document), 'read_token': self.read_token,
                    'open_request': {key: deepcopy(value) for key, value in self.open_request.items() if key in {'id', 'drop_id', 'phase', 'error', 'process'}} if self.open_request else None,
                    'open_trace': deepcopy(self.open_trace),
                    'view': deepcopy(self.view), 'queued_for_next_version': list(self.queued_for_next_version),
                    'drop_pending': self.drop is not None,
                    'session_id': self.runtime.session_id if self.runtime else '',
                    'revision': self.runtime.revision if self.runtime else 0,
                    'applied_revision': self.runtime.applied if self.runtime else 0,
                    'model_calls': self.runtime.model_calls if self.runtime else 0,
                    'actual_model': getattr(self.runtime, 'model_route', None),
                    'needs_clarification': bool(self.runtime and self.runtime.needs_clarification),
                    'presentation': self.native.mode if self.native is not None else 'review',
                    'messages': deepcopy(self.runtime.messages) if self.runtime else [],
                    'protected': sorted(self.protected),
                    'protected_cells': [list(cell) for cell in sorted(self.protected_cells)],
                    'plans': deepcopy(plans), 'outputs': deepcopy(list(self.outputs.values())),
                    'busy': self.app.busy, 'stage': self.stage, 'draft_drop_pending':self.pending_draft_drop, 'paged':self.paged.state() if self.paged.active() else None,
                    'canvas':self.canvas.state() if self.canvas.active() else None}


from .presentation_workspace import PresentationApi


class FloatingApi(OfficeCanvasApi,PresentationApi,Api):
    def __init__(self, app, **options):
        super().__init__(app, **options)
        self._floating = FloatingSession(self)
        self._window_chrome = None
        from ..download_cache import DownloadCache
        self._download_cache = DownloadCache(self._settings._root)
        from .window_preferences import WindowPreferences
        self._ui_preferences = WindowPreferences(self._settings._root)

    def floating_intro(self, remember=False):
        def action():
            self._floating.authenticated()
            if type(remember) is not bool:
                raise ValueError('引导偏好无效')
            if remember:
                self._ui_preferences.set_intro_seen()
            return {'ok': True, 'intro_seen': self._ui_preferences.read().get('intro_seen', False)}
        return self._floating_call(action)

    def floating_theme(self, theme=None):
        def action():
            self._floating.authenticated()
            if theme is not None:
                self._ui_preferences.set_theme(theme)
            return {'ok': True, 'theme': self._ui_preferences.read().get('theme')}
        return self._floating_call(action)

    def floating_download_cache(self, kind):
        def action():
            self._floating.authenticated()
            if self._app.busy:
                raise ValueError('请等待当前任务结束再检查缓存')
            return {'ok': True, 'plan': self._download_cache.preview(kind)}
        return self._floating_call(action)

    def floating_quarantine_download(self, identifier, confirmed=False):
        def action():
            self._floating.authenticated()
            if self._app.busy:
                raise ValueError('请等待当前任务结束再处理缓存')
            return {'ok': True, 'result': self._download_cache.quarantine(identifier, confirmed)}
        return self._floating_call(action)

    def floating_diagnostics(self):
        def action():
            self._floating.authenticated()
            from ..diagnostics import get
            preview = get().preview()
            identifier = uuid.uuid4().hex
            self._diagnostic_preview = (identifier, time.monotonic(), json.dumps({key: value for key, value in preview.items() if key != 'path'}, ensure_ascii=False).encode('utf-8'))
            return {'ok': True, 'diagnostics': preview, 'preview_id': identifier}
        return self._floating_call(action)

    def floating_export_diagnostics(self, preview_id):
        def action():
            self._floating.authenticated()
            preview = getattr(self, '_diagnostic_preview', None)
            if not preview or preview[0] != preview_id or time.monotonic() - preview[1] > 300:
                raise ValueError('诊断预览已失效，请重新查看')
            import webview
            from ..downloads import safe
            selected = self._app.window.create_file_dialog(webview.SAVE_DIALOG, save_filename='LiX-diagnostics.json', file_types=('JSON (*.json)',))
            if not selected:
                return {'ok': True, 'cancelled': True}
            target = safe(Path(selected if isinstance(selected, str) else selected[0]))
            if target.suffix.lower() != '.json':
                raise ValueError('请选择新的 JSON 文件')
            with target.open('xb') as stream:
                stream.write(preview[2])
                stream.flush()
                os.fsync(stream.fileno())
            self._diagnostic_preview = None
            return {'ok': True, 'saved': True}
        return self._floating_call(action)

    def floating_window_state(self):
        def action():
            self._floating.authenticated()
            return self._window_chrome.state() if self._window_chrome else {'ok': True, 'available': False}
        return self._floating_call(action)

    def floating_window_trace(self, value):
        def action():
            self._floating.authenticated()
            return self._window_chrome.trace(value) if self._window_chrome else {'ok': True, 'available': False}
        return self._floating_call(action)

    def floating_window_pointer(self, value):
        def action():
            self._floating.authenticated()
            if not self._window_chrome:
                raise ValueError('窗口尚未就绪')
            return self._window_chrome.pointer(value)
        return self._floating_call(action)

    def test_model_connection_draft(self, draft, refresh=False):
        return self._connections.probe_draft(draft, refresh)

    def floating_window_action(self, action):
        def apply():
            from .window_chrome import WINDOW_ACTIONS
            self._floating.authenticated()
            if not self._window_chrome:
                raise ValueError('当前界面不在原生桌面窗口中')
            if type(action) is not str or action not in WINDOW_ACTIONS:
                raise ValueError('不支持的窗口操作')
            return self._window_chrome.action(action)
        return self._floating_call(apply)

    def floating_window_regions(self, value):
        def apply():
            self._floating.authenticated()
            if not self._window_chrome:
                return {'ok': True, 'available': False}
            return self._window_chrome.set_drag_regions(value)
        return self._floating_call(apply)

    def _clear_private_state(self):
        self._diagnostic_preview = None
        if hasattr(self,'_presentations'):
            self._presentations.invalidate()
        if hasattr(self, '_floating'):
            self._floating.reset()
        super()._clear_private_state()

    def _floating_call(self, action, *args):
        try:
            return action(*args)
        except Exception as exc:
            if isinstance(exc, ValueError):
                return {'ok': False, 'reason': 'floating_rejected', 'error': str(exc)[:240]}
            if isinstance(exc, OSError):
                return {'ok': False, 'reason': 'floating_io_failed', 'error': '本机文件操作未完成，请检查路径长度、占用或目录权限；核对已有产物后再操作'}
            return {'ok': False, 'reason': 'floating_failed', 'error': '此步骤未完成，请核对当前文档和记录；没有自动重试'}

    def floating_snapshot(self, cursor=0):
        return self._floating_call(self._floating.collect, cursor)

    def paged_editor_ready(self, document_id, revision):
        def action():
            self._floating.paged.guard(document_id, revision)
            waiting = getattr(self._floating, '_editor_metrics_wait', None)
            if waiting is not None and waiting[2:] == (document_id, self._floating.document['version']):
                waiting[0].stop(waiting[1])
                self._floating._editor_metrics_wait = None
            runtime = self._floating.runtime
            if self._floating.paged.editable and runtime and not runtime.closed and runtime.applied < runtime.revision and not self._floating.paged.deferred:
                from ..requirement_history import active_requirements
                text = '\n'.join(row['text'] for row in active_requirements(runtime))
                if re.search(r'最后|文末|末尾|结尾', text):
                    self._floating.copy_continuation = document_id
            if self._floating.paged.editable and self._floating.copy_continuation == document_id:
                if any(row.get('kind') == 'paged_editable_copy' and row.get('output_sha256') == self._floating.document['version'] for row in self._floating.outputs.values()):
                    self._floating.notice('draft_copy_continued', {'text': '可编辑副本已打开，正在继续处理已发送要求。'})
                else:
                    self._floating.conversation.update(self._floating.conversation.active_task, 'editor_ready', '编辑器已就绪，正在继续处理已发送要求。')
            return {'ok': True}
        return self._floating_call(action)

    def conversation_observe(self, client_id, phase, elapsed_ms):
        def action():
            floating = self._floating
            floating.authenticated()
            if (client_id not in floating.conversation.turns or not isinstance(phase, str)
                    or not re.fullmatch(r'(?:local_echo|accepted_visible|first_effective_visible|final_answer_visible|stage_visible:[a-z_]{1,32})', phase)
                    or type(elapsed_ms) not in (int, float) or not 0 <= elapsed_ms <= 86400000):
                raise ValueError('界面计时记录无效')
            value = {'id': client_id, 'phase': phase, 'elapsed_ms': round(elapsed_ms, 2)}
            floating.conversation.measurements.append(value)
            floating.conversation.measurements = floating.conversation.measurements[-256:]
            from ..diagnostics import record
            record('conversation', 'client_phase', 'observed', correlation=client_id, code=phase, elapsed_ms=round(elapsed_ms))
            return {'ok': True}
        return self._floating_call(action)

    def conversation_undo_begin(self, turn_id, document_id, revision, approved=False):
        return self._floating_call(self._floating.begin_round_undo, turn_id, document_id, revision, approved)

    def conversation_undo_restored(self, turn_id, token, document_id, revision, data, manifest):
        return self._floating_call(self._floating.round_undo.restored, turn_id, token, document_id, revision, data, manifest)

    def conversation_undo_status(self, turn_id, token, document_id):
        return self._floating_call(self._floating.round_undo.status, turn_id, token, document_id)

    def conversation_undo_finish(self, turn_id, token, document_id, revision):
        return self._floating_call(self._floating.round_undo.finish, turn_id, token, document_id, revision)

    def conversation_undo_abort(self, turn_id, token, document_id):
        return self._floating_call(self._floating.round_undo.abort, turn_id, token, document_id)

    def conversation_begin(self, client_id, text, document_id, version, revision=0, reference=None, context_id=None, local_only=False, reply_to=None, plugin_choice=None, displayed_recipient=None, resources=None):
        def action():
            floating = self._floating
            floating.authenticated()
            if not floating.document or floating.document['id'] != document_id or floating.document['version'] != version:
                raise ValueError('文档或来源版本已改变；消息没有进入其他文件')
            if context_id is not None and context_id != floating.conversation.binding:
                raise ValueError('对话已切换，旧消息未发送')
            base_intent = conversation_intent(text)
            local = base_intent in {'status', 'stop'}
            opening = bool(floating.open_request and floating.open_request.get('phase') == 'parsing')
            at_limit = len(floating.conversation.turns) >= TURN_LIMIT
            known = floating.conversation.turns.get(client_id) if isinstance(client_id, str) else None
            known_read = known is not None and known['intent'] == 'read'
            selection_message = base_intent == 'read' and requested_read_scope(text) in {'selection', 'mixed'}
            pending_scope = floating.conversation.turns.get(floating.conversation.pending_scope)
            repeats_question = pending_scope and pending_scope['stage'] == 'reference' and pending_scope['question'] == text.strip()
            resolving = bool(scope_choice(text) or repeats_question)
            allow_reselect = selection_message or resolving
            selection_prompt = selection_message and not (floating.paged.active() or floating.canvas.active() or resolving)
            local_history = base_intent == 'read' and edit_history_question(text)
            skip_reference = local or known_read or at_limit and not resolving
            if local_only and not (skip_reference or selection_prompt or local_history):
                return {'ok': False, 'reason': 'model_required', 'error': '请先在输入区选择有效的文档 Agent 模型；要求已保留。'}
            if skip_reference or opening and not floating.document.get('ready'):
                reference_value = None
            elif floating.paged.active():
                floating.paged.guard(document_id, None if allow_reselect else revision)
                reference_value = floating.paged.references.get(reference) if isinstance(reference, str) and revision == floating.paged.revision else None
            elif floating.canvas.active():
                floating.canvas.guard(document_id, floating.canvas.revision if allow_reselect else revision)
                references = (floating.canvas.ai.read_reference, floating.canvas.ai.reference)
                reference_value = next((item for item in references if item and revision == floating.canvas.revision and
                                        (item['id'] == reference or reference is None and not allow_reselect)), None)
            else:
                floating.check_source()
                reference_value = None
            if reference_value and reference_value.get('draft_revision') != revision:
                reference_value = None
            stale = reference is not None and reference_value is None
            if stale and not (skip_reference or allow_reselect):
                raise ValueError(STALE_SELECTION)
            with floating.lock:
                turn, duplicate = floating.conversation.receive(client_id, text, document_id, revision, reference_value, context_id,
                                                                 reply_to, stale, plugin_choice, displayed_recipient, resources)
            if turn['intent'] == 'stop' and not duplicate:
                result = floating.stop()
                stop_text = {'cancelling': '正在停止副本校验，草稿保留。', 'committing': '已停止后续派发，当前副本正在完成核验。'}.get(
                    result.get('publication_stop'), '已停止后续派发，已有要求与结果保留。')
                if turn.get('transient'):
                    turn.update(stage='stopped', text=stop_text, answer=stop_text, sequence=1)
                    floating.conversation.control_turn = turn
                else:
                    floating.conversation.update(client_id, 'stopped', stop_text, status='complete')
            return {'ok': True, 'handled': turn['intent'] != 'edit', 'duplicate': duplicate,
                    'turn': deepcopy(turn if turn.get('transient') else floating.conversation.turns[turn['id']]),
                    'edit_text': turn.get('effective_question', text), 'conversation': floating.conversation.state()}
        return self._floating_call(action)

    def conversation_choose_file(self, document_id, version):
        return self._floating_call(self._floating.conversation.scopes.choose_file, document_id, version)

    def conversation_remove_file(self, token):
        return self._floating_call(self._floating.conversation.scopes.remove_file, token)

    def conversation_skill_choices(self, document_id, version):
        def action():
            from ..skills_backend import SkillsBackend, SkillsError
            floating = self._floating
            with floating.lock:
                floating.conversation.scopes._file_owner(document_id, version)
                try:
                    return {'ok': True, 'items': SkillsBackend(floating.app.path.parent).turn_choices()}
                except SkillsError:
                    raise ValueError('无法读取当前工作区的技能，请检查技能设置。') from None
        return self._floating_call(action)

    def conversation_plugin_consent(self, client_id, document_id, version, context_id, token, approved):
        def action():
            from .turn_scope import record_decision
            floating = self._floating
            conversation = floating.conversation
            with floating.lock, conversation.lock:
                floating.authenticated()
                floating.check_source()
                row = conversation.turns.get(client_id)
                preview = conversation.plugin_previews.get(client_id)
                if (not (type(approved) is bool or approved in ('masked', 'show')) or not row or not preview or preview.token != token
                        or row['stage'] != 'plugin_consent' or row['status'] != 'waiting'
                        or context_id != conversation.binding or not floating.document
                        or (document_id, version) != (floating.document['id'], floating.document['version'])):
                    raise ValueError('这次插件许可已结束。')
                if approved is False:
                    conversation.plugin_previews.pop(client_id)
                    record_decision(client_id, 'plugin', 'cancel')
                    conversation.update(client_id, 'stopped', '已取消，内容未发送。', status='stopped')
                    return {'ok': True}
                try:
                    preview.require_current(conversation, row)
                except ValueError:
                    conversation.plugin_previews.pop(client_id)
                    conversation.update(client_id, 'failed', '选区、许可或接收方已改变，请重新发起查询。', status='failed')
                    raise
                if approved == 'show':
                    return {'ok': True, 'query': preview.query}
                if client_id in conversation.readers or len(conversation.readers) >= 2:
                    raise ValueError('已有查询正在处理，请稍后重试。')
                conversation.plugin_approvals[client_id] = (conversation.plugin_previews.pop(client_id), approved == 'masked')
                record_decision(client_id, 'plugin', 'masked' if approved == 'masked' else 'allow')
                saved = conversation.read_requests.get(client_id)
                conversation._start_read(row, row['draft_revision'], saved[0] if saved else None)
                return {'ok': True}
        return self._floating_call(action)

    def conversation_plugin_source(self, client_id, index):
        def action():
            import webbrowser
            from .plugin_turns import usable_sources
            self._floating.authenticated()
            row = self._floating.conversation.turns.get(client_id) or {}
            links = usable_sources(row.get('plugin_sources', []))
            if type(index) is not int or not 0 <= index < len(links):
                raise ValueError('找不到这次查询的来源。')
            return {'ok': bool(webbrowser.open(links[index]['url']))}
        return self._floating_call(action)

    def conversation_scope_decide(self, token, decision):
        return self._floating_call(self._floating.conversation.scopes.decide, token, decision)

    def conversation_retry(self, client_id, document_id, version, displayed_recipient=None):
        def action():
            floating = self._floating
            floating.authenticated()
            floating.check_source()
            if floating.document['id'] != document_id or floating.document['version'] != version:
                raise ValueError('文档已改变，未恢复旧任务')
            row = floating.conversation.turns.get(client_id)
            if not row or row['document_id'] != document_id or row['status'] not in {'failed', 'interrupted', 'uncertain'}:
                raise ValueError('此回合当前不能重试；没有重复派发')
            if row['intent'] == 'clarify':
                raise ValueError('请连同范围重新发送原问题。')
            if row.get('failure_code') == 'number_conflict':
                raise ValueError('请先调整与数字顺序冲突的要求，再继续处理')
            if (row.get('publication') or {}).get('reason') == 'draft_advanced':
                raise ValueError('请预览或保存当前草稿，旧发布请求已失效')
            if row['intent'] == 'read':
                floating.conversation.scopes.bind_recipient(displayed_recipient)
                floating.conversation.retry_answer(row)
                return {'ok': True}
            if floating.app.busy or floating.paged.preparing or floating.paged.publishing or floating.canvas.publishing:
                raise ValueError('已有步骤仍在安全收尾，未重复创建文件')
            if floating.paged.uncertain or floating.canvas.uncertain or row['status'] == 'uncertain':
                result = floating.paged.publication_result
                if not result:
                    raise ValueError('结果仍待核实，请先核对审计和文件；没有重新发布')
                target = floating.app.editor.audit / result['id'] / 'audit.json'
                _reject_links(target)
                receipt = _read_bounded_json(target, 2 * 1024 * 1024)
                output = floating.app.editor.space.file(receipt['output'])
                if receipt.get('status') != 'verified' or receipt.get('source_sha256') != version or digest(output) != receipt.get('output_sha256') or receipt.get('output_sha256') != receipt.get('candidate_sha256'):
                    raise ValueError('尚无可核验的成功回执；保留结果待核实，未重复执行')
                receipt = {key: value for key, value in receipt.items() if key not in {'authorization', 'continue_editing', 'automatic_requirements'}}
                receipt['source_version'] = receipt['source_sha256']
                floating.paged.uncertain = False
                floating.paged.publication_result = {**result, 'status': 'verified', 'text': '已核对成功回执，未重复发布。'}
                floating.stage = 'verified'
                floating.outputs[receipt['id']] = receipt
                copy_only = receipt.get('kind') == 'paged_editable_copy'
                if not copy_only:
                    floating.paged.published_revision = result['draft_revision']
                floating.conversation.update(client_id, 'copy_verified' if copy_only else 'complete',
                    '可编辑副本的成功回执已核对；可以从这个已有文件继续。' if copy_only else '已找回并核对改稿成功回执；没有再次发布。', status='waiting' if copy_only else 'complete',
                    results=[{'id': receipt['id'], 'name': output.name, 'kind': receipt.get('kind'), 'sha256': receipt['output_sha256'], 'verified': True}])
                if copy_only:
                    floating.access.reset()
                    grant = floating.access.for_instruction()
                    if not floating.runtime or floating.runtime.closed:
                        floating.runtime = InstructionRuntime(version, persist=floating.persist, notify=floating.notice, model_limit=12)
                        floating.app.instruction_runtime = floating.runtime
                        floating.runtime.submit(client_id, row.get('effective_question', row['question']))
                    receipt.update(automatic_requirements=True, authorization=grant,
                                   continue_editing={'serial': floating.access.serial, 'level': grant['level']})
                    return {'ok': True, 'reconciled': True, 'output_id': receipt['id'], 'access_id': grant['id']}
                return {'ok': True, 'reconciled': True, 'output_id': receipt['id']}
            for known in row['results']:
                if known.get('verified') and known['id'] not in floating.outputs:
                    restored = self.conversation_output(known['id'], document_id, version)
                    if not restored.get('ok'):
                        raise ValueError('已有结果无法核对；未重复生成文件，请先查看审计与产物')
            final = next((item for item in reversed(list(floating.outputs.values())) if item.get('kind') != 'paged_editable_copy' and not item.get('undone') and any(known['id'] == item['id'] for known in row['results'])), None)
            if final:
                floating.conversation.update(client_id, 'existing_result', '找到已有验证改稿，将打开最新结果；原要求保留，没有重复执行。', status='waiting')
                return {'ok': True, 'output_id': final['id']}
            floating.access.reset()
            floating.access.for_instruction()
            floating.conversation.active_task = client_id
            floating.conversation.update(client_id, 'resuming', '正在从最后已知阶段恢复；复用已验证成果，不重复创建工作副本。')
            copy = next((item for item in floating.outputs.values() if item.get('kind') == 'paged_editable_copy' and item.get('source_version') == version and not item.get('undone')), None)
            if copy:
                if not floating.runtime or floating.runtime.closed:
                    floating.runtime = InstructionRuntime(version, persist=floating.persist, notify=floating.notice, model_limit=12)
                    floating.app.instruction_runtime = floating.runtime
                    floating.runtime.submit(client_id, row.get('effective_question', row['question']))
                copy.update(automatic_requirements=True, authorization=floating.access.state(),
                            continue_editing={'serial': floating.access.serial, 'level': floating.access.state()['level']})
                return {'ok': True, 'output_id': copy['id'], 'access_id': floating.access.state()['id']}
            if floating.paged.active():
                paged = floating.paged
                paged.copy_error = None
                paged.start_blocked = None
                if not floating.runtime or floating.runtime.closed:
                    paged.submit(uuid.uuid4().hex, row['question'], None, document_id, version, floating.runtime.session_id if floating.runtime else '', paged.inbox_id)
                elif paged.editable:
                    paged.automatic_reference(document_id, paged.revision, floating.runtime.session_id, floating.runtime.revision)
                    paged.start()
            return {'ok': True}
        return self._floating_call(action)

    def conversation_new(self, context_id, document_id, version):
        def action():
            floating = self._floating
            floating.authenticated()
            with floating.lock:
                floating.check_source()
                if not floating.document.get('ready') or (document_id, version) != (floating.document['id'], floating.document['version']):
                    raise ValueError('文档或版本已改变，请重新读取')
                if (floating.app.busy or floating.conversation.readers or floating.paged.preparing or floating.paged.publishing
                        or floating.paged.applying or floating.canvas.publishing or floating.canvas.ai.running is not None
                        or floating.round_undo.busy or floating.paged.uncertain or floating.canvas.uncertain):
                    raise ValueError('上一对话仍在处理或核对结果；请等待结束后开始新对话。')
                state = floating.conversation.new_conversation(context_id)
                floating.stop(announce=False)
                floating.runtime = floating.app.instruction_runtime = None
                floating.stage = 'idle'
                return {'ok': True, 'conversation': state}
        return self._floating_call(action)

    def conversation_archive(self, document_id, version, archive_id=None):
        def action():
            floating = self._floating
            floating.authenticated()
            with floating.lock, floating.conversation.lock:
                floating.check_source()
                if not floating.document.get('ready') or (document_id, version) != (floating.document['id'], floating.document['version']):
                    raise ValueError('文档或版本已改变，请重新读取')
                if archive_id is None:
                    return {'ok': True, 'archives': floating.conversation.archives()}
                return {'ok': True, 'turns': floating.conversation.archived(archive_id)}
        return self._floating_call(action)

    def conversation_output(self, identifier, document_id, version, archive_id=None):
        def action():
            floating = self._floating
            floating.authenticated()
            floating.check_source()
            if floating.document['id'] != document_id or floating.document['version'] != version or not isinstance(identifier, str) or not re.fullmatch('[a-f0-9]{32}', identifier):
                raise ValueError('结果归属已改变')
            turns = floating.conversation.archived(archive_id) if archive_id is not None else floating.conversation.turns.values()
            known = next((result for turn in turns for result in turn['results'] if result.get('id') == identifier and result.get('verified')), None)
            if not known:
                raise ValueError('此结果不属于当前对话')
            target = floating.app.editor.audit / identifier / 'audit.json'
            _reject_links(target)
            receipt = _read_bounded_json(target, 2 * 1024 * 1024)
            path = floating.app.editor.space.file(receipt['output'])
            if receipt.get('id') != identifier or receipt.get('status') != 'verified' or digest(path) != known['sha256'] or receipt.get('output_sha256') != known['sha256']:
                raise ValueError('结果文件或审计已改变，未打开旧记录')
            receipt = {key: value for key, value in receipt.items() if key not in {'authorization', 'continue_editing', 'automatic_requirements'}}
            receipt['source_version'] = receipt.get('source_sha256')
            if archive_id is None:
                floating.outputs[identifier] = receipt
            return {'ok': True, 'output': receipt}
        return self._floating_call(action)

    def floating_set_access(self, document_id, version, epoch, level, confirmed=False):
        return self._floating_call(self._floating.access.set, document_id, version, epoch, level, confirmed)

    def floating_capabilities(self):
        def action():
            self._floating.authenticated()
            return {'ok':True,'event_owner':'floating_snapshot'}
        return self._floating_call(action)

    def paged_open(self, document_id, version):
        return self._floating_call(self._floating.paged.open,document_id,version)

    def paged_enable_workspace(self):
        def action():
            self._floating.authenticated()
            self._floating.preferred_paged=True
            return {'ok':True}
        return self._floating_call(action)

    def paged_start_pending(self, document_id, revision, reference_id, instruction_revision, session_id):
        return self._floating_call(self._floating.paged.start_pending,document_id,revision,reference_id,instruction_revision,session_id)

    def paged_automatic_reference(self, document_id, revision, session_id, instruction_revision):
        return self._floating_call(self._floating.paged.automatic_reference, document_id, revision, session_id, instruction_revision)

    def paged_automatic_editable_copy(self, document_id, revision, inbox_id, session_id, instruction_revision, access_id):
        return self._floating_call(self._floating.paged.automatic_editable_copy, document_id, revision,
                                  inbox_id, session_id, instruction_revision, access_id)

    def paged_requirements(self, document_id, source_version, inbox_id):
        return self._floating_call(self._floating.paged.requirements,document_id,source_version,inbox_id)

    def paged_change_requirement(self, document_id, source_version, inbox_id, session_id, client_id, target_id, version, instruction_revision, action, text):
        return self._floating_call(self._floating.paged.change_requirement,document_id,source_version,inbox_id,session_id,client_id,target_id,version,instruction_revision,action,text)

    def paged_changed(self, document_id, revision, apply_token=None):
        return self._floating_call(self._floating.paged.changed,document_id,revision,apply_token)

    def paged_sync(self, document_id, revision, data, manifest, bootstrap=False):
        def action():
            try:
                return self._floating.paged.sync(document_id,revision,data,manifest,bootstrap)
            except (ValueError,OSError):
                if bootstrap and self._floating.paged.document_id==document_id and self._floating.paged.revision==revision:
                    self._floating.paged.editable=False
                    self._floating.paged.reasons=['无修改往返或结构身份校验未通过；当前仅供查看，原稿不变']
                raise
        return self._floating_call(action)

    def paged_reference(self, document_id, revision, target):
        return self._floating_call(self._floating.paged.reference,document_id,revision,target)

    def paged_preview_latest_draft(self, publication_id, document_id, source_version, revision):
        return self._floating_call(self._floating.paged.preview_latest_draft,publication_id,document_id,source_version,revision)

    def paged_reference_locator(self, document_id, revision, locator_id, start, end):
        return self._floating_call(self._floating.paged.reference_locator,document_id,revision,locator_id,start,end)

    def paged_submit(self, client_id, text, reference_id, document_id, source_version, session_id, inbox_id):
        return self._floating_call(self._floating.paged.submit,client_id,text,reference_id,document_id,source_version,session_id,inbox_id)

    def paged_prepare_apply(self, proposal_id, document_id, revision, access_id=None):
        return self._floating_call(self._floating.paged.prepare_apply,proposal_id,document_id,revision,access_id)

    def paged_applied(self, token, document_id, prior_revision, revision):
        return self._floating_call(self._floating.paged.applied,token,document_id,prior_revision,revision)

    def paged_prepare_operations(self, document_id, revision, instruction_revision, operations):
        return self._floating_call(self._floating.paged.prepare_operations,document_id,revision,instruction_revision,operations)

    def paged_applied_operations(self, token, document_id, prior_revision, revision):
        return self._floating_call(self._floating.paged.applied_operations,token,document_id,prior_revision,revision)

    def paged_id_map(self, document_id):
        return self._floating_call(self._floating.paged.id_map,document_id)

    def paged_reject(self, proposal_id):
        return self._floating_call(self._floating.paged.reject,proposal_id)

    def paged_plan(self, document_id, revision):
        return self._floating_call(self._floating.paged.plan,document_id,revision)

    def paged_publish(self, plan_id, document_id, revision, confirmed=False, access_id=None):
        return self._floating_call(self._floating.paged.publish,plan_id,document_id,revision,confirmed,access_id)

    def paged_plan_editable_copy(self, document_id, revision):
        return self._floating_call(self._floating.paged.plan_editable_copy, document_id, revision)

    def paged_publish_editable_copy(self, plan_id, document_id, revision, confirmed=False, continue_editing=False):
        return self._floating_call(self._floating.paged.publish_editable_copy, plan_id, document_id, revision, confirmed, continue_editing)

    def paged_restore(self, document_id):
        return self._floating_call(self._floating.paged.restore,document_id)

    def paged_use_original(self, document_id):
        return self._floating_call(self._floating.paged.use_original,document_id)

    def paged_permit_switch(self, document_id, revision, choice):
        return self._floating_call(self._floating.paged.permit_switch,document_id,revision,choice)

    def select_file(self):
        return self.floating_choose()

    def execute(self):
        if self._floating.canvas.blocks_original():
            return {'ok':False, 'reason':'draft_changed', 'error':'交互草稿已改变，旧执行入口失效；请保存并选择副本后重新规划'}
        if self._floating.runtime is not None:
            return {'ok': False, 'reason': 'floating_confirmation_required', 'error': '请回到文档助手审阅当前版本的方案并独立确认'}
        return super().execute()

    def ask(self, question, options=None):
        if self._connections.state.mode == 'agent' and (self._floating.runtime is not None or self._floating.paged.active() or self._floating.canvas.active()):
            return {'ok': False, 'reason': 'floating_task_active', 'error': '请返回悬浮文档助手追加要求；不会从其他入口重启当前任务'}
        return super().ask(question, options)

    def _rag_scope(self):
        authenticated, busy, token = super()._rag_scope()
        active = hasattr(self, '_floating') and self._floating.runtime is not None and not self._floating.runtime.closed
        return authenticated, busy or active, token

    def logout(self):
        with self._floating.lock:
            runtime = self._floating.runtime
            if runtime is not None:
                with runtime.lock:
                    if runtime.executing or self._app.busy:
                        return {'ok': False, 'reason': 'busy'}
                    return super().logout()
            return super().logout()

    def login(self, username, password):
        with self._floating.lock:
            pending = self._floating.drop if self._floating.document is None and self._floating.runtime is None else None
            result = super().login(username, password)
            if result.get('ok') and result.get('authenticated') and pending is not None:
                self._floating.drop = pending
                try:
                    self._floating.admit_drop()
                except (OSError, ValueError) as exc:
                    self._floating.notice('admission_failed', {'text': str(exc)[:200]})
            return result

    def register(self, username, password):
        with self._floating.lock:
            return super().register(username, password)

    def reset_accounts(self, confirmed=False):
        with self._floating.lock:
            if self._floating.runtime and self._floating.runtime.executing:
                return {'ok': False, 'reason': 'busy'}
            return super().reset_accounts(confirmed)

    def floating_choose(self):
        return self._floating_call(self._floating.choose)

    def floating_close_document(self):
        def action():
            self._floating.authenticated()
            if self._app.busy:raise ValueError('当前步骤尚未结束，不能关闭文档')
            self._floating.paged.before_admit()
            self._floating.canvas.before_admit()
            self._floating.reset()
            self._app._clear_private_state()
            return {'ok':True}
        return self._floating_call(action)

    def floating_admit_drop(self):
        return self._floating_call(self._floating.admit_drop)

    def floating_read(self, token, allowed, offset=0):
        return self._floating_call(self._floating.read, token, allowed, offset)

    def floating_protect(self, identifier, protected):
        return self._floating_call(self._floating.protect, identifier, protected)

    def floating_reference(self, document_id, version, identifier):
        return self._floating_call(self._floating.reference, document_id, version, identifier)

    def floating_resolve_protection(self):
        return self._floating_call(self._floating.resolve_protection)

    def floating_submit(self, session_id, client_id, text, open_request_id=None, document_id=None, version=None):
        return self._floating_call(self._floating.submit, session_id, client_id, text, open_request_id, document_id, version)

    def floating_confirm(self, session_id, plan_id, revision, version, approved=False, access_id=None):
        return self._floating_call(self._floating.confirm, session_id, plan_id, revision, version, approved, access_id)

    def floating_stop(self):
        return self._floating_call(self._floating.stop)

    def floating_start_again(self, fresh=False):
        if type(fresh) is not bool:
            return {'ok': False, 'error': '参数无效'}
        return self._floating_call(self._floating.start_again, fresh)

    def floating_restore_drafts(self):
        return self._floating_call(self._floating.restore_drafts)

    def floating_rollback(self, output_id, approved=False):
        return self._floating_call(self._floating.rollback, output_id, approved)

    def floating_select_output(self, output_id, read_open=False, access_id=None):
        return self._floating_call(self._floating.select_output, output_id, read_open, access_id)

    def floating_select_history_output(self, identifier):
        return self._floating_call(self._floating.select_history_output, identifier)

    def floating_preview_output(self, output_id, allowed=False, offset=0):
        return self._floating_call(self._floating.preview_output, output_id, allowed, offset)

    def floating_present(self, mode):
        def action():
            self._floating.authenticated()
            if self._floating.native:
                self._floating.native.paged_document = self._floating.paged.active()
                self._floating.native.present(mode)
            return {'ok': True}
        return self._floating_call(action)

    def floating_restore_position(self, position):
        def action():
            self._floating.authenticated()
            return {'ok': bool(self._floating.native and self._floating.native.restore_position(position))}
        return self._floating_call(action)
