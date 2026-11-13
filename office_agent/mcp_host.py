import json
from pathlib import Path
import threading
import time

from .auth import LocalAuth
from .documents import inspect_file, search_document
from .rag import SECRET
from .rag_sources import SourceExtractor
from .writes import Editor, WordReplace, ExcelSet, digest


def scalar(kind, **limits):
    return {'type': kind, **limits}


TOOLS = {
    'inspect_document': ('查看本机已授权文件的结构与安全计数，不返回正文、文件名或路径。', {}),
    'word_structure': ('查看当前Word正文段落及表格数量，不返回内容。', {}),
    'search_current_document': ('在当前文件搜索，返回有界命中数量，不返回摘录。', {'query': scalar('string', minLength=1, maxLength=200)}),
    'plan_word_replace': ('只生成当前Word的替换计划；不会写入，必须在本机控制窗口另行确认。', {
        'paragraph_index': scalar('integer', minimum=0, maximum=10000), 'old': scalar('string', minLength=1, maxLength=500), 'new': scalar('string', maxLength=500)}),
    'plan_excel_set': ('只生成当前Excel字面值修改计划；不会写入，必须在本机控制窗口另行确认。工作表从0编号。', {
        'sheet_index': scalar('integer', minimum=0, maximum=100), 'cell': scalar('string', minLength=2, maxLength=10),
        'value': scalar(['string','integer','number','boolean','null'], maxLength=500)}),
}


class HostError(ValueError):
    pass


def arguments(name, values):
    if name not in TOOLS or type(values) is not dict or set(values) != set(TOOLS[name][1]):
        raise HostError('invalid_arguments')
    if len(json.dumps(values, ensure_ascii=True, allow_nan=False)) > 7000:
        raise HostError('invalid_arguments')
    for key, rule in TOOLS[name][1].items():
        value = values[key]
        if isinstance(value, (dict, list)) or type(value) not in (str, int, float, bool, type(None)):
            raise HostError('invalid_arguments')
        kind = rule['type']
        if kind == 'integer' and (type(value) is not int or not rule['minimum'] <= value <= rule['maximum']):
            raise HostError('invalid_arguments')
        if kind == 'string' and type(value) is not str:
            raise HostError('invalid_arguments')
        if type(value) is str and (not rule.get('minLength', 0) <= len(value) <= rule['maxLength'] or
                                   any(ord(char) < 32 or 0xd800 <= ord(char) <= 0xdfff for char in value) or SECRET.search(value)):
            raise HostError('invalid_arguments')
    return values


class McpHost:
    def __init__(self, auth=None, *, clock=time.monotonic, editor_factory=Editor):
        self.auth = auth if auth is not None else LocalAuth()
        self.clock, self.editor_factory = clock, editor_factory
        self.lock = threading.RLock()
        self.path = self.editor = self.fingerprint = None
        self.extractor = None
        self.expires = 0
        self.pending = {}
        self.history = {}
        self.disabled = False

    def revoke(self):
        with self.lock:
            if self.editor is not None:
                self.editor.plans.clear()
            self.pending.clear()
            self.history.clear()
            self.path = self.editor = self.fingerprint = None
            self.extractor = None
            self.expires = 0

    def logout(self):
        with self.lock:
            self.revoke()
            self.auth.logout()

    def authorize(self, selected):
        with self.lock:
            self.revoke()
            if self.auth.status()['authenticated'] is not True:
                raise HostError('not_authenticated')
            path = Path(selected).absolute()
            if path.suffix.lower() not in {'.docx','.xlsx'}:
                raise HostError('unsupported_document')
            editor = self.editor_factory(path.parent, require_confirmation=True, render_validation=True)
            checked = editor.space.file(str(path))
            extractor = SourceExtractor(editor.space)
            extractor.extract(checked.name)
            self.path, self.editor, self.fingerprint = checked, editor, digest(checked)
            self.extractor = extractor
            self.expires = self.clock() + 900

    def _guard(self):
        if self.disabled:
            raise HostError('outcome_uncertain')
        if self.auth.status()['authenticated'] is not True:
            self.revoke()
            raise HostError('not_authenticated')
        if self.path is None:
            raise HostError('file_not_authorized')
        if self.clock() >= self.expires:
            self.revoke()
            raise HostError('authorization_expired')
        try:
            current = self.editor.space.file(str(self.path)) == self.path and self.extractor._snapshot(self.path.name)[2] == self.fingerprint
        except Exception:
            current = False
        if not current:
            self.revoke()
            raise HostError('source_changed')

    def _prune(self):
        for key, created in list(self.pending.items()):
            if self.clock() - created >= 300:
                self.pending.pop(key)
                self.editor.plans.pop(key, None)

    def call(self, name, values):
        with self.lock:
            try:
                self._guard()
                arguments(name, values)
                self._prune()
                if name.startswith('plan_') and len(self.pending) + len(self.history) >= 8:
                    raise HostError('plan_capacity')
                if name == 'inspect_document':
                    report = inspect_file(self.path)
                    result = {'kind': report['kind'], 'readable': report['readable'],
                              'protected': report['protected'], 'macros': report['macros'],
                              'external_links': report['external_links'], 'risk_count': len(report['risks'])}
                elif name == 'word_structure':
                    if self.path.suffix.lower() != '.docx':
                        raise HostError('unsupported_document')
                    from docx import Document
                    doc = Document(self.path)
                    result = {'paragraph_count': len(doc.paragraphs), 'table_count': len(doc.tables)}
                elif name == 'search_current_document':
                    hits = search_document(self.path, values['query'], max_results=10)
                    result = {'hit_count': len(hits), 'limit': 10, 'limit_reached': len(hits) == 10,
                              'content_returned': False, 'support_assessed': False}
                else:
                    if name == 'plan_word_replace':
                        request = WordReplace(path=str(self.path), **values)
                    else:
                        if self.path.suffix.lower() != '.xlsx':
                            raise HostError('unsupported_document')
                        from openpyxl import load_workbook
                        book = load_workbook(self.path, read_only=True, keep_links=False)
                        try:
                            if values['sheet_index'] >= len(book.sheetnames):
                                raise HostError('invalid_arguments')
                            request = ExcelSet(path=str(self.path), sheet=book.sheetnames[values['sheet_index']], cell=values['cell'], value=values['value'])
                        finally:
                            book.close()
                    self._guard()
                    plan = self.editor.plan(request)
                    self.pending[plan.id] = self.clock()
                    result = {'plan_id': plan.id, 'status': 'awaiting_local_review', 'wrote_nothing': True,
                              'confirmation_required': True, 'expires_in_seconds': 300}
                self._guard()
                reply = {'ok': True, 'result': result}
                if len(json.dumps(reply, ensure_ascii=True, allow_nan=False)) > 2048:
                    raise HostError('response_capacity')
                return reply
            except HostError as error:
                return {'ok': False, 'reason': str(error)}
            except Exception:
                return {'ok': False, 'reason': 'operation_refused'}

    def preview(self, plan_id):
        with self.lock:
            self._guard()
            self._prune()
            if plan_id not in self.pending:
                raise HostError('plan_expired')
            return self.editor.plans[plan_id].public()

    def confirm_local(self, plan_id, approved=False):
        with self.lock:
            self.preview(plan_id)
            if approved is not True:
                self.pending.pop(plan_id)
                self.editor.plans.pop(plan_id, None)
                return {'status': 'cancelled', 'wrote_nothing': True}
            self.pending.pop(plan_id)
            try:
                result = self.editor.execute(plan_id, confirmed=True)
                if type(result) is not dict or result.get('status') != 'verified' or result.get('id') != plan_id:
                    raise HostError('outcome_uncertain')
                self._guard()
                self.history[plan_id] = result
                return result
            except Exception:
                self.editor.plans.pop(plan_id, None)
                self.disabled = True
                raise HostError('outcome_uncertain') from None

    def rollback_local(self, plan_id, approved=False):
        with self.lock:
            self._guard()
            if approved is not True or plan_id not in self.history:
                raise HostError('confirmation_required')
            try:
                result = self.editor.rollback(plan_id, confirmed=True)
                if type(result) is not dict or result.get('status') != 'rolled_back' or result.get('id') != plan_id:
                    raise HostError('outcome_uncertain')
                self.history.pop(plan_id)
                return result
            except Exception:
                self.disabled = True
                raise HostError('outcome_uncertain') from None
