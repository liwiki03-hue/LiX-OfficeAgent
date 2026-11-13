from copy import deepcopy
from collections import Counter
import json
import os
import re
import threading
import time
import uuid

from pydantic import BaseModel, ConfigDict, Field

from .. import office_canvas as engine
from ..agent import OfficeAgent, ToolSpec, configured_decision_limit
from ..instruction_runtime import InstructionRuntime, InstructionPause, StaleInstruction
from ..paged_documents import sha
from ..professional_edits import _protect
from ..requirement_history import active_requirements, change_requirement
from ..writes import GuardError


class CanvasSuggestion(BaseModel):
    model_config = ConfigDict(extra='forbid', strict=True)
    reference_id: str = Field(min_length=1, max_length=100)
    old: str = Field(max_length=4000)
    new: str = Field(max_length=4000)
    reason: str = Field(min_length=1, max_length=500)


class CanvasAI:
    def __init__(self, canvas):
        self.canvas = canvas
        self.reference = None
        self.runtime = None
        self.proposals = {}
        self.running = None
        self.error = ''
        self.sequence = 0

    @property
    def floating(self):
        return self.canvas.floating

    def invalidate(self):
        self.sequence += 1
        if self.runtime:
            self.runtime.stop()
        self.reference = None
        for proposal in self.proposals.values():
            if proposal['status'] == 'pending':
                proposal['status'] = 'stale'

    def state(self):
        runtime = self.runtime
        return {'protocol': 1, 'sequence': self.sequence, 'document_id': self.canvas.document_id, 'draft_revision': self.canvas.revision,
                'running': self.running is not None, 'error': self.error,
                'session_id': runtime.session_id if runtime else '',
                'instruction_revision': runtime.revision if runtime else 0,
                'requirements': active_requirements(runtime) if runtime else [],
                'reference': deepcopy(self.reference), 'proposals': deepcopy(list(self.proposals.values())),
                'scope': '一个普通单元格或已有文字片段；建议采用不等于发布副本'}

    def _guard(self, document_id, revision):
        self.canvas.guard(document_id, revision)
        self.canvas._editing()
        if self.canvas.publishing or self.canvas.previewing:
            raise GuardError('当前正在发布或预览；要求和输入请保留，稍后重试')

    def select(self, document_id, revision, item_id, target, allow_numeric_change=False):
        with self.canvas.lock:
            self._guard(document_id, revision)
            if type(allow_numeric_change) is not bool or not isinstance(target, str) or not isinstance(item_id, str):
                raise GuardError('引用参数无效')
            suffix = self.canvas.source.suffix.lower()
            item = engine.item_for(engine.load(self.canvas.data, suffix), suffix, item_id)
            if suffix == '.xlsx':
                row, col = engine.address(target)
                view = engine.view(self.canvas.data, suffix, item_id, row, col)
                value = next((cell for cell in view['cells'] if cell['id'] == target), None)
                if not value or not value['editable'] or value.get('formula'):
                    raise GuardError('仅能引用已有普通可编辑单元格；公式和富文本不在AI首版范围')
                if (item['name'], target) in self.floating.protected_cells:
                    raise GuardError('目标已明确保护，不能提出修改建议')
                old = value['value']
                action = 'number' if value['type'] == 'n' else 'text'
                if action != 'number' and allow_numeric_change:
                    raise GuardError('数值许可仅适用于数值格；普通文本仍保留数字和单位')
                if action == 'number' and not allow_numeric_change:
                    raise GuardError('数值修改需要明确允许；普通润色不会自动改变数值')
            else:
                if allow_numeric_change:
                    raise GuardError('PPT文字建议首版保留数字，不支持解除数值保护')
                view = engine.view(self.canvas.data, suffix, item_id)
                value = next((run for run in view['runs'] if run['id'] == target), None)
                if not value:
                    raise GuardError('请选择当前支持的普通文字片段')
                old, action = value['text'], 'text'
            if len(old) > 4000:
                raise GuardError('目标文字超过4000字符，请缩小范围')
            from ..rag import SECRET
            if SECRET.search(old):
                raise GuardError('目标疑似包含凭据，请移除敏感内容后再生成建议')
            for proposal in self.proposals.values():
                if proposal['status'] == 'pending':
                    proposal['status'] = 'stale'
            while len(self.proposals) >= 16:
                self.proposals.pop(next(iter(self.proposals)))
            self.reference = {'id': uuid.uuid4().hex, 'document_id': document_id, 'draft_revision': revision,
                              'source_version': sha(self.canvas.original), 'draft_sha256': sha(self.canvas.data),
                              'item_id': item_id, 'target': target, 'old': old, 'action': action,
                              'allow_numeric_change': allow_numeric_change, 'location': value.get('location', target),
                              'item_name': item['name']}
            self.sequence += 1
            return {'ok': True, 'reference': deepcopy(self.reference), 'ai': self.state()}

    def _current(self, reference):
        return (self.canvas.active() and self.reference is not None and self.reference['id'] == reference['id'] and
                self.canvas.revision == reference['draft_revision'] and sha(self.canvas.data) == reference['draft_sha256'])

    def submit(self, document_id, revision, reference_id, session_id, client_id, text):
        with self.canvas.lock:
            self._guard(document_id, revision)
            reference = self.reference
            if not reference or reference['id'] != reference_id or not self._current(reference):
                raise GuardError('引用已失效；请重新选择当前草稿目标，输入保留')
            if self.runtime is None or self.runtime.closed:
                if session_id and (self.runtime is None or session_id != self.runtime.session_id):
                    raise GuardError('任务会话已改变')
                runtime = InstructionRuntime(sha(self.canvas.original), persist=self.floating.persist,
                                             model_limit=configured_decision_limit())
                runtime.notify = lambda kind, payload: self.floating.notice(kind, payload) if self.runtime is runtime and self.canvas.active() else None
                self.runtime = self.floating.runtime = self.floating.app.instruction_runtime = runtime
            elif session_id != self.runtime.session_id:
                raise GuardError('任务会话已改变')
            duplicate = client_id in self.runtime.ids
            record = self.runtime.submit(client_id, text)
            if not duplicate:
                self.sequence += 1
                self.canvas.plans.clear()
                for proposal in self.proposals.values():
                    if proposal['status'] == 'pending': proposal['status'] = 'stale'
            started = False if duplicate else self.start()
            return {'ok': True, 'record': record, 'duplicate': duplicate, 'started': started, 'ai': self.state()}

    def change(self, document_id, revision, session_id, client_id, target_id, version, instruction_revision, action, text):
        with self.canvas.lock:
            self._guard(document_id, revision)
            if not self.runtime or self.runtime.session_id != session_id:
                raise GuardError('任务会话已改变')
            record, duplicate = change_requirement(self.runtime, client_id, target_id, version, instruction_revision, action, text)
            if not duplicate:
                self.sequence += 1
                self.canvas.plans.clear()
                for proposal in self.proposals.values():
                    if proposal['status'] == 'pending': proposal['status'] = 'stale'
            return {'ok': True, 'record': record, 'duplicate': duplicate, 'ai': self.state()}

    def start(self):
        if self.running is not None or self.floating.app.busy or not self.reference or not self.runtime or self.runtime.closed:
            return False
        if not active_requirements(self.runtime) or not self._current(self.reference):
            return False
        try:
            route = self.floating.api._connections.capture('agent')
        except Exception:
            self.sequence += 1
            self.error = '模型配置不可用；要求已保留，请修复后主动处理'
            return False
        runtime, reference = self.runtime, deepcopy(self.reference)
        owner = uuid.uuid4().hex
        self.running = owner
        self.sequence += 1
        self.floating.app._canvas_ai_owner = owner
        self.floating.app.busy = self.floating.app._model_task_active = True
        runtime.model_route = route.public()
        self.error = ''
        self.floating.notice('canvas_ai_running', {'document_id': reference['document_id'], 'model': route.public()})
        threading.Thread(target=self._run, args=(runtime, reference, route, owner), daemon=True).start()
        return True

    def _run(self, runtime, reference, route, owner):
        app = self.floating.app
        def current():
            return self.running == owner and self.runtime is runtime and self.floating.runtime is runtime and not runtime.closed and self._current(reference)
        def boundary(phase, state):
            runtime.boundary(phase, state)
            if not current(): raise StaleInstruction('当前草稿或任务已变化')
            if phase == 'before_model':
                self.canvas.guard(reference['document_id'], reference['draft_revision'])
                if self.floating.api._connections.capture('agent').public() != route.public():
                    raise InstructionPause('接口或模型配置已改变，请核对后重新处理')
                state.required.goal = json.dumps({'requirements': [row['text'] for row in active_requirements(runtime)],
                                                 'reference': reference}, ensure_ascii=False)
                if not active_requirements(runtime): raise InstructionPause('没有有效要求')
        def dispatch(name, args):
            with self.canvas.lock:
                if not current() or runtime.request_revision != runtime.revision:
                    raise StaleInstruction('旧提案已失效')
                parsed = CanvasSuggestion.model_validate(args)
                from ..rag import SECRET
                if SECRET.search(parsed.reason): raise GuardError('修改原因疑似包含凭据')
                if name != 'propose_canvas_edit' or parsed.reference_id != reference['id'] or parsed.old != reference['old']:
                    raise GuardError('提案引用或原文不符')
                if self.floating.api._connections.capture('agent').public() != route.public():
                    raise GuardError('模型配置已变化，请重新核对')
                operation = self._operation(reference, parsed.new)
                self.canvas.guard(reference['document_id'], reference['draft_revision'])
                engine.transform(self.canvas.data, self.canvas.source.suffix.lower(), operation)
                if len(self.proposals) >= 16: raise GuardError('提案达到16项，请结束本轮后重新开始')
                proposal = {'id': uuid.uuid4().hex, 'reference_id': reference['id'], 'old': parsed.old,
                            'new': parsed.new, 'reason': parsed.reason, 'status': 'pending',
                            'draft_revision': reference['draft_revision'], 'instruction_revision': runtime.revision,
                            'session_id': runtime.session_id, 'model': route.public(), 'operation': operation,
                            'expires': time.monotonic() + 300}
                self.proposals[proposal['id']] = proposal
                self.sequence += 1
                self.floating.notice('canvas_ai_proposal', deepcopy(proposal))
                return {'id': proposal['id'], 'status': 'pending', 'wrote_nothing': True}
        try:
            model = self.floating.api._connections.build_model(route)
            limit = configured_decision_limit()
            tool = ToolSpec(name='propose_canvas_edit', description='为宿主当前选定单元格或文字片段提出修改；不采用、不写文件，old必须与引用一致。', schema_model=CanvasSuggestion)
            from ..skills_backend import workspace_skill_context
            skill_context = workspace_skill_context(self.canvas.source.parent, on_load=lambda receipt:self.floating.notice('skill_loaded',receipt))
            agent = OfficeAgent(model, dispatch, [tool], self.canvas.source.parent, persist=False,
                                decision_limit=min(8, limit) if limit is not None else 8, context_tokens=route.context_window,
                                reply_tokens=int(os.getenv('OFFICE_REPLY_BUDGET', '512')), runtime_boundary=boundary,
                                stop_requested=lambda: not current(), skill_context=skill_context)
            state = agent.run(agent.new_state(json.dumps(reference, ensure_ascii=False), constraints=['仅提出当前目标的修改建议', '源文件不变，采用和发布由用户分别决定']))
            if current() and not any(row['status'] == 'pending' and row['instruction_revision'] == runtime.revision for row in self.proposals.values()):
                self.error = '未形成可采用的建议；请核对保护要求、范围和模型回复后主动处理'
                self.sequence += 1
                self.floating.notice('canvas_ai_failed', {'text': self.error})
            elif current() and state.answer:
                self.floating.notice('canvas_ai_answer', {'text': state.answer, 'model': route.public()})
        except Exception:
            if current():
                self.error = '建议未完成；草稿和要求保留，请核对引用及模型后主动处理'
                self.sequence += 1
                self.floating.notice('canvas_ai_failed', {'text': self.error})
        finally:
            with self.canvas.lock:
                valid = current()
                if self.running == owner:
                    self.running = None
                    self.sequence += 1
                if getattr(app, '_canvas_ai_owner', None) == owner:
                    app._canvas_ai_owner = None
                    app.busy = app._model_task_active = False
                if valid and runtime.applied < runtime.revision:
                    self.start()

    def _operation(self, reference, value):
        if value == reference['old']: raise GuardError('没有实际变化')
        from ..rag import SECRET
        if SECRET.search(value): raise GuardError('建议疑似包含凭据，不能采用')
        if not reference['allow_numeric_change']:
            from types import SimpleNamespace
            from ..technical_profile import observations
            _protect(reference['old'], value, SimpleNamespace(preserve_numbers=True, preserve_citations=True, preserve_formulas=True, protected_literals=[]))
            quantities = lambda text: Counter(item for kind, item in observations(text) if kind == 'quantity')
            if (quantities(reference['old']) != quantities(value) or
                    re.findall(r'℃|℉|°[CF]', reference['old']) != re.findall(r'℃|℉|°[CF]', value)):
                raise GuardError('建议改变了可识别的数值单位，请保留原单位')
        return {'item_id': reference['item_id'], 'target': reference['target'], 'action': reference['action'],
                'value': value, 'expected': reference['old']}

    def adopt(self, document_id, revision, proposal_id):
        with self.canvas.lock:
            self._guard(document_id, revision)
            proposal = self.proposals.get(proposal_id)
            runtime, reference = self.runtime, self.reference
            if self.running is not None or not proposal or not runtime or runtime.closed or not reference or not self._current(reference):
                raise GuardError('提案不可采用；请等待请求结束并重新核对引用')
            if (proposal['status'] != 'pending' or proposal['session_id'] != runtime.session_id or
                    proposal['instruction_revision'] != runtime.revision or proposal['draft_revision'] != revision or
                    proposal['reference_id'] != reference['id'] or time.monotonic() > proposal['expires']):
                raise GuardError('提案已过期或要求已改变')
            if self.floating.api._connections.capture('agent').public() != proposal['model']:
                raise GuardError('模型配置已变化，请重新生成建议')
            result = self.canvas.apply(document_id, revision, proposal['operation'])
            proposal['status'] = 'adopted'
            self.sequence += 1
            self.canvas.summaries[-1]['ai'] = {key: deepcopy(proposal[key]) for key in
                                             ('id', 'reason', 'model', 'session_id', 'instruction_revision')}
            return {**result, 'proposal_id': proposal_id, 'ai': self.state()}

    def reject(self, document_id, revision, proposal_id):
        with self.canvas.lock:
            self._guard(document_id, revision)
            proposal = self.proposals.get(proposal_id)
            if not proposal or proposal['status'] != 'pending': raise GuardError('提案已失效')
            proposal['status'] = 'rejected'
            self.sequence += 1
            return {'ok': True, 'ai': self.state()}

    def resume(self, document_id, revision, reference_id, session_id, instruction_revision):
        with self.canvas.lock:
            self._guard(document_id, revision)
            if (not self.runtime or self.runtime.closed or self.runtime.session_id != session_id or
                    type(instruction_revision) is not int or instruction_revision != self.runtime.revision or
                    not self.reference or self.reference['id'] != reference_id or not self._current(self.reference)):
                raise GuardError('引用或要求已失效，请核对后主动处理')
            return {'ok': True, 'started': self.start(), 'ai': self.state()}

    def stop(self, document_id, revision):
        with self.canvas.lock:
            self._guard(document_id, revision)
            self.invalidate()
            self.canvas.plans.clear()
            return {'ok': True, 'remote_cancelled': False, 'ai': self.state()}
