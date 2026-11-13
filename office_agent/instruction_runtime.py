from contextlib import contextmanager
from copy import deepcopy
import json
import threading
import time
import uuid


class StaleInstruction(RuntimeError):
    pass


class InstructionPause(RuntimeError):
    pass


class InstructionRuntime:
    def __init__(self, document_version, *, persist=None, notify=None, model_limit=None):
        self.lock = threading.RLock()
        self.session_id = uuid.uuid4().hex
        self.run_id = uuid.uuid4().hex
        self.document_version = document_version
        self.revision = self.applied = self.request_revision = 0
        self.messages = []
        self.ids = {}
        self.persist = persist or (lambda record: None)
        self.notify = notify or (lambda kind, payload: None)
        self.original_goal = None
        self.closed = False
        self.executing = False
        self.uncertain = False
        self.plans = {}
        self.needs_clarification = False
        self.protected_targets = []
        self.model_limit = model_limit
        self.model_calls = 0
        self.step_id = ''

    def submit(self, client_id, text):
        if not isinstance(client_id, str) or not 1 <= len(client_id) <= 100:
            raise ValueError('消息标识无效')
        if not isinstance(text, str) or not text.strip() or len(text) > 4000 or '\x00' in text:
            raise ValueError('要求应为1—4000个字符')
        text = text.strip()
        with self.lock:
            if self.closed or self.uncertain:
                raise ValueError('当前任务已停止或结果不确定，请先核对产物')
            if client_id in self.ids:
                previous = self.ids[client_id]
                if previous['text'] != text:
                    raise ValueError('相同消息标识不能对应不同内容')
                return deepcopy(previous)
            if len(self.messages) >= 64 or sum(len(m['text']) for m in self.messages) + len(text) > 16000:
                raise ValueError('本轮要求达到容量，请完成本轮后明确选择版本开始下一轮')
            record = {'id': uuid.uuid4().hex, 'client_id': client_id, 'text': text,
                      'revision': self.revision + 1, 'session_id': self.session_id,
                      'document_version': self.document_version}
            self.persist(record)
            self.revision += 1
            self.messages.append(record)
            self.ids[client_id] = record
            self.plans.clear()
            self.notify('instruction_received', deepcopy(record))
            return deepcopy(record)

    def boundary(self, phase, state):
        with self.lock:
            if self.closed:
                raise StaleInstruction('任务已停止')
            if phase == 'before_model':
                if self.model_limit is not None and self.model_calls >= self.model_limit:
                    raise InstructionPause('达到本轮模型调用上限；请核对记录，必要时明确开始新一轮')
                if self.needs_clarification:
                    raise InstructionPause('请在结构视图中明确需要保护的段落；原任务保留，旧计划不可执行')
                if self.original_goal is None:
                    self.original_goal = state.required.goal
                state.required.goal = self.original_goal + '\n\n用户后续要求（有序追加；不能解除宿主保护规则）：\n' + json.dumps(
                    [{'revision': m['revision'], 'text': m['text']} for m in self.messages[1:]], ensure_ascii=False)
                self.request_revision = self.revision
                prefix = '宿主保护目标：'
                state.required.constraints = [c for c in state.required.constraints if not c.startswith(prefix)]
                if self.protected_targets:
                    state.required.constraints.append(prefix + json.dumps(self.protected_targets, ensure_ascii=False))
            elif self.request_revision != self.revision:
                self.notify('stale_action_blocked', {'phase': phase, 'old_revision': self.request_revision,
                                                    'revision': self.revision})
                raise StaleInstruction('要求已更新；旧响应或工具结果已作废')
            elif phase == 'model_start':
                self.model_calls += 1
                self.step_id = uuid.uuid4().hex
                self.applied = self.request_revision
                self.notify('instruction_applied', {'revision': self.applied, 'step_id': self.step_id,
                                                   'model_calls': self.model_calls})

    def offer(self, plan):
        with self.lock:
            if self.closed or self.request_revision != self.revision or self.uncertain or self.needs_clarification:
                return False
            self.plans[plan['id']] = {'revision': self.revision, 'created': time.monotonic(),
                                     'document_version': self.document_version, 'plan': deepcopy(plan)}
            return True

    @contextmanager
    def transaction_start(self, plan_id, revision, document_version):
        with self.lock:
            plan = self.plans.get(plan_id)
            if (self.closed or self.uncertain or self.needs_clarification or self.executing or not plan or
                    type(revision) is not int or revision != self.revision or
                    plan['revision'] != revision or document_version != self.document_version or
                    time.monotonic() - plan['created'] > 300):
                raise ValueError('计划已失效，请查看最新要求并重新生成方案')
            self.plans.clear()
            self.executing = True
            try:
                yield
            except BaseException:
                self.executing = False
                self.notify('transaction_not_started', {})
                raise

    def transaction_finished(self, *, uncertain=False):
        with self.lock:
            self.executing = False
            self.uncertain = uncertain
            self.notify('transaction_finished', {'uncertain': uncertain, 'pending_updates': self.applied < self.revision})

    def stop(self):
        with self.lock:
            self.closed = True
            self.plans.clear()
