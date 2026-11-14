from copy import deepcopy
import base64
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import threading
import time
import uuid

from .. import office_canvas as engine
from ..fidelity import _convert, _run, _temporary_root, compare_documents, package_projection, runtime_path
from ..paged_documents import CandidatePublisher, sha
from ..render_identity import environment_identity
from ..render_handoff import implementation_identity
from ..writes import GuardError, _audit_path, digest
from ..canvas_drafts import CanvasDraftStore, package_identity
from ..canvas_ranges import paste_operations, read_range
from .. import canvas_display
from .. import canvas_geometry
from .. import canvas_panes
from ..canvas_drafts import implementation as canvas_implementation


class OfficeCanvas:
    def __init__(self, floating):
        self.floating = floating
        self.lock = threading.RLock()
        self.close()

    def close(self):
        with self.lock:
            if hasattr(self, 'ai'):
                self.ai.invalidate()
            from .canvas_ai import CanvasAI
            self.ai = CanvasAI(self)
            self.document_id = None
            self.source = None
            self.data = self.original = None
            self.revision = self.published_revision = 0
            self.operations = []
            self.summaries = []
            self.undo_groups = []
            self.redo_groups = []
            self.range_plan = None
            self.history_origin = 'session'
            self.plans = {}
            self.preview_cache = None
            self.preview_token = None
            self.calculation_cache = {}
            self.calculation_revision = None
            self.calculation_edges = {}
            self.window_cache = None
            self.publishing = self.previewing = self.uncertain = False
            self.permit = None
            self.last_error = ''
            self.recovery = self.checkpoint = None
            self.recovery_state = 'none'
            self.recovery_status = ''
            self.recovery_dismissed = False
            self.checkpoint_revision = None

    def active(self):
        return self.document_id is not None and self.floating.document and self.floating.document['id'] == self.document_id

    def blocks_original(self):
        return self.active() and bool(self.operations)

    def state(self):
        with self.lock:
            return self._state()

    def _state(self):
        outputs = [row for row in list(self.floating.outputs.values()) if not row.get('undone') and row.get('document_id') == self.document_id]
        last = outputs[-1] if outputs else None
        return {'document_id': self.document_id, 'draft_revision': self.revision, 'published_revision': self.published_revision,
                'dirty': bool(self.operations) and self.revision != self.published_revision, 'publishing': self.publishing,
                'previewing': self.previewing, 'uncertain': self.uncertain, 'error': self.last_error, 'operation_count': len(self.operations),
                'undo_count':len(self.undo_groups), 'redo_count':len(self.redo_groups), 'history_origin':self.history_origin,
                'needs_copy_for_agent': bool(self.operations),
                'ai': self.ai.state(),
                'recovery_available': self._recovery_available(), 'recovery_state': self.recovery_state,
                'recovery_status': ('本机草稿有更新待保存' if self.checkpoint_revision is not None and self.checkpoint_revision != self.revision else self.recovery_status), 'checkpoint_revision': self.checkpoint_revision,
                'recovery_saved_at': self.recovery[0]['saved_at'] if self.recovery else None,
                'recovery_location': str(self.drafts.target) if self.active() and hasattr(self, 'drafts') else None,
                'last_output': {'id':last['id'], 'name':Path(last['output']).name, 'path':last['output']} if last else None,
                'storage': '编辑在内存中；主动保存本机草稿后，可重新准入原文件恢复内容，不恢复批准'}

    def _recovery_available(self):
        return bool(self.recovery and not self.recovery_dismissed and self.revision == 0)

    def _editing(self):
        if self._recovery_available():
            raise GuardError('请先恢复本机草稿，或明确本次使用原稿')

    def guard(self, document_id, revision=None):
        self.floating.authenticated()
        self.floating.check_source()
        if not self.active() or document_id != self.document_id or revision is not None and (type(revision) is not int or revision != self.revision):
            raise GuardError('文档或草稿版本已改变，请重新核对')
        if self.uncertain:
            raise GuardError('已有结果不确定，请先核对副本及审计，不能重复写入')

    def open(self, document_id, version):
        with self.lock:
            self.floating.authenticated()
            source = self.floating.check_source()
            doc = self.floating.document
            if not doc['ready'] or document_id != doc['id'] or version != doc['version'] or source.suffix.lower() not in {'.xlsx', '.pptx'}:
                raise GuardError('请先准入并授权当前XLSX/PPTX')
            if not self.active():
                if self.floating.app.busy:
                    raise GuardError('请等待当前文档任务结束')
                if source.stat().st_size > engine.MAX_BYTES:
                    raise GuardError('当前文件超过8 MiB交互编辑预算；保留已有有界读取')
                data = source.read_bytes()
                if sha(data) != version:
                    raise GuardError('读取期间源文件改变')
                engine.load(data, source.suffix.lower())
                initial_view = engine.view(data, source.suffix.lower())
                self.close()
                self.document_id, self.source = document_id, source
                self.original = self.data = data
                self.source_guard = CandidatePublisher(self.floating.app.editor, source, version)
                self.drafts = CanvasDraftStore(self.floating.app.editor.audit, self.floating.api._memory_key(),
                    self.floating.source_key, version, doc['file_identity'], source.suffix.lower())
                try:
                    self.recovery = self.drafts.read()
                    if self.recovery:
                        self.recovery_state = 'available'
                        self.recovery_status = '找到本机草稿；核对后恢复，或本次使用原稿'
                except (ValueError, OSError):
                    self.recovery_state = 'invalid'
                    self.recovery_status = '本机草稿损坏、身份不符或不兼容；未恢复，原稿仍可使用'
                return {'ok': True, **self.state(), 'view': initial_view}
            return {'ok': True, **self.state(), 'view': engine.view(self.data, self.source.suffix.lower())}

    def view(self, document_id, revision, item_id, row=1, col=1):
        with self.lock:
            self.guard(document_id, revision)
            if type(row) is not int or type(col) is not int or not 1 <= row <= 10000 or not 1 <= col <= 256:
                raise GuardError('工作表视窗位置无效')
            key = (document_id, revision, item_id, row, col)
            cached = self.window_cache
            if cached and cached[0] == key and cached[1] is self.data:
                self.window_cache = None
                return {'ok': True, **self.state(), 'view': deepcopy(cached[2])}
            return {'ok': True, **self.state(), 'view': self._view(item_id, row, col)}

    def _view(self, item_id, row, col):
        if self.calculation_revision != self.revision:
            self.calculation_cache = {}
            self.calculation_edges = {}
            self.calculation_revision = self.revision
        result = engine.view(self.data, self.source.suffix.lower(), item_id, row, col, self.calculation_cache)
        graph = result.get('formula_graph') or {}
        self.calculation_edges[item_id] = graph.get('edges')
        return result

    def prefetch(self, document_id, revision, item_id, row=1, col=1):
        with self.lock:
            self.guard(document_id, revision)
            if type(row) is not int or type(col) is not int or not 1 <= row <= 10000 or not 1 <= col <= 256:
                raise GuardError('工作表视窗位置无效')
            if self.source.suffix.lower() != '.xlsx' or self.publishing or self.previewing or self.floating.app.busy:
                raise GuardError('当前步骤不预取工作表窗口')
            key = (document_id, revision, item_id, row, col)
            cached = self.window_cache
            if not cached or cached[0] != key or cached[1] is not self.data:
                self.window_cache = (key, self.data, self._view(item_id, row, col))
            return {'ok': True, 'document_id': document_id, 'draft_revision': revision, 'prefetched': True}

    def apply(self, document_id, revision, operation):
        return self.apply_many(document_id, revision, [operation])

    def _check_editing(self):
        self._editing()
        if self.publishing or self.previewing and self.source.suffix.lower() != '.pptx' or self.floating.app.busy and self.ai.running is None:
            raise GuardError('正在验证或执行任务，请等待此步骤结束')

    def _check_protection(self, operation, data):
        if self.source.suffix.lower() == '.xlsx' and isinstance(operation, dict):
            item = engine.item_for(engine.load(data, '.xlsx'), '.xlsx', operation.get('item_id'))
            if operation.get('action') != canvas_panes.ACTION and any(sheet == item['name'] and canvas_geometry.affects(operation, coordinate)
                   for sheet, coordinate in self.floating.protected_cells):
                raise GuardError('此单元格已由用户明确保护，请先取消保护')

    def _build_group(self, operations):
        if not isinstance(operations, list) or not 1 <= len(operations) <= 100 or len(self.operations)+len(operations)>100:
            raise GuardError('本草稿最多100项操作，请缩小区域或保存副本后继续')
        operations = deepcopy(operations)
        targets = set()
        candidate, summaries = self.data, []
        for operation in operations:
            if not isinstance(operation, dict):
                raise GuardError('编辑组操作无效')
            target = (operation.get('item_id'), operation.get('target'))
            if not all(isinstance(value, str) for value in target) or target in targets:
                raise GuardError('编辑组目标无效或重复，请拆分有依赖的操作')
            targets.add(target)
            self._check_protection(operation, candidate)
            candidate, summary = engine.transform(candidate, self.source.suffix.lower(), operation)
            summaries.append(summary)
        self.source_guard.check_source()
        return candidate, summaries, operations

    def _invalidate_edit(self, changed=None):
        self.ai.invalidate()
        if self.floating.runtime is not None:
            self.floating.runtime.stop()
        self.floating.app.pending_id = None
        self.floating.app.editor.plans.clear()
        self.plans.clear(); self.preview_cache = None
        self.window_cache = None
        if changed is None:
            self.calculation_cache = {}
            self.calculation_revision = None
            self.calculation_edges = {}
        else:
            affected = set(changed)
            unknown = {item for item, _ in changed if not isinstance(self.calculation_edges.get(item), dict)}
            for item in {item for item, _ in changed} - unknown:
                edges = self.calculation_edges[item]
                pending = {ref for sheet, ref in affected if sheet == item}
                while pending:
                    downstream = {ref for ref, dependencies in edges.items() if (item, ref) not in affected and pending.intersection(dependencies)}
                    affected.update((item, ref) for ref in downstream)
                    pending = downstream
            self.calculation_cache = {key: value for key, value in self.calculation_cache.items() if key not in affected and key[0] not in unknown}
            self.calculation_revision = self.revision + 1
        self.permit = None; self.last_error = ''
        self.checkpoint = self.range_plan = None

    def apply_many(self, document_id, revision, operations, *, redo=False):
        with self.lock:
            self.guard(document_id, revision)
            self._check_editing()
            candidate, summaries, operations = self._build_group(operations)
            if redo:
                for summary, previous in zip(summaries, self.redo_groups[-1]['summaries']):
                    if 'ai' in previous:
                        summary['ai'] = deepcopy(previous['ai'])
            self._invalidate_edit(None if redo else {(operation['item_id'], operation['target']) for operation in operations})
            self.data = candidate
            self.operations.extend(operations); self.summaries.extend(summaries)
            self.undo_groups.append(len(operations))
            if redo:
                self.redo_groups.pop()
            else:
                self.redo_groups.clear()
            self.revision += 1
            return {'ok': True, **self.state(), 'change': summaries[-1], 'changes':deepcopy(summaries), 'group_operations':len(operations)}

    def undo(self, document_id, revision):
        with self.lock:
            self.guard(document_id, revision)
            if self.publishing or self.previewing or self.floating.app.busy or not self.operations:
                raise GuardError('当前不能撤回草稿操作')
            count = self.undo_groups[-1] if self.undo_groups else 1
            removed = self.operations[-count:]
            removed_summaries = deepcopy(self.summaries[-count:])
            for operation in removed:
                self._check_protection(operation, self.data)
            remaining = self.operations[:-count]
            data = self.original
            for operation in remaining:
                data, _ = engine.transform(data, self.source.suffix.lower(), operation)
            self.source_guard.check_source()
            self._invalidate_edit()
            self.data = data; self.operations = remaining; self.summaries = self.summaries[:-count]
            if self.undo_groups:
                self.undo_groups.pop()
            self.redo_groups.append({'operations':deepcopy(removed), 'summaries':removed_summaries})
            self.revision += 1
            return {'ok': True, **self.state()}

    def redo(self, document_id, revision):
        with self.lock:
            self.guard(document_id, revision)
            if self.publishing or self.previewing or self.floating.app.busy or not self.redo_groups:
                raise GuardError('当前没有可重做的草稿编辑组')
            return self.apply_many(document_id, revision, self.redo_groups[-1]['operations'], redo=True)

    def copy_range(self, document_id, revision, item_id, start, end):
        with self.lock:
            self.guard(document_id, revision)
            if self.source.suffix.lower() != '.xlsx':
                raise GuardError('区域操作仅适用于当前工作表')
            result = read_range(self.data, item_id, start, end)
            return {'ok':True, 'draft_revision':revision, **result}

    def selection_scope(self, document_id, revision, item_id, axis, index):
        from ..canvas_ranges import used_selection
        with self.lock:
            self.guard(document_id, revision)
            if self.source.suffix.lower() != '.xlsx':
                raise GuardError('行列选择仅适用于工作表')
            return {'ok': True, **used_selection(self.data, item_id, axis, index)}

    def prepare_range(self, document_id, revision, item_id, start, text, mode):
        with self.lock:
            self.guard(document_id, revision)
            self._check_editing()
            if self.source.suffix.lower() != '.xlsx':
                raise GuardError('区域粘贴仅适用于当前工作表')
            operations, selected = paste_operations(self.data, item_id, start, text, mode)
            plan = self._stage_group(revision,operations,mode,**selected)
            return {'ok':True, 'plan':{key:deepcopy(plan[key]) for key in ('id','revision','changes','mode','start','end','rows','cols')},
                    'draft_only':True, 'changed_cells':len(operations), 'unchanged_cells':selected['rows']*selected['cols']-len(operations)}

    def _stage_group(self, revision, operations, mode, **details):
        _, changes, operations = self._build_group(operations)
        plan = {'id':uuid.uuid4().hex,'revision':revision,'snapshot':sha(self.data),
                'protection':sorted(self.floating.protected_cells),'deadline':time.monotonic()+300,
                'implementation':canvas_implementation(),'operations':operations,'changes':changes,'mode':mode,**details}
        self.range_plan = plan
        return plan

    def prepare_geometry(self, document_id, revision, item_id, target, action, value):
        with self.lock:
            self.guard(document_id, revision)
            self._check_editing()
            if self.source.suffix.lower() != '.xlsx' or action not in canvas_geometry.ACTIONS:
                raise GuardError('行列尺寸仅适用于当前工作表')
            raw = engine.load(self.data, '.xlsx')
            item = engine.item_for(raw, '.xlsx', item_id)
            row, col = engine.address(target)
            operation = {'item_id':item_id, 'target':target, 'action':action, 'value':value,
                         'expected':canvas_geometry.snapshot(engine.xml(raw[item['part']]), action, row, col)}
            plan = self._stage_group(revision,[operation],'geometry')
            return {'ok':True, 'draft_only':True, 'plan':{key:deepcopy(plan[key]) for key in ('id','revision','changes','mode')}}

    def prepare_panes(self, document_id, revision, item_id, rows, cols):
        with self.lock:
            self.guard(document_id, revision)
            self._check_editing()
            if self.source.suffix.lower() != '.xlsx':
                raise GuardError('冻结窗格仅适用于工作表')
            raw = engine.load(self.data,'.xlsx')
            item = engine.item_for(raw,'.xlsx',item_id)
            operation = {'item_id':item_id,'target':'A1','action':canvas_panes.ACTION,'value':{'rows':rows,'cols':cols},
                         'expected':canvas_panes.identity(engine.xml(raw[item['part']]))}
            plan = self._stage_group(revision,[operation],'panes')
            return {'ok':True,'draft_only':True,'plan':{key:deepcopy(plan[key]) for key in ('id','revision','changes','mode')}}

    def commit_range(self, document_id, revision, plan_id, confirmed):
        with self.lock:
            self.guard(document_id, revision)
            plan = self.range_plan
            if type(confirmed) is not bool:
                raise GuardError('采用区域需要明确确认')
            if (not plan or plan['id'] != plan_id or plan['revision'] != revision or time.monotonic()>plan['deadline']
                    or plan['implementation'] != canvas_implementation()
                    or plan['snapshot'] != sha(self.data) or plan['protection'] != sorted(self.floating.protected_cells)):
                raise GuardError('区域预览已过期，请重新核对；未采用到草稿')
            self.range_plan = None
            if not confirmed:
                return {'ok':True, 'cancelled':True, **self.state()}
            return self.apply_many(document_id, revision, plan['operations'])

    def draft_plan(self, document_id, revision, action):
        with self.lock:
            self.guard(document_id, revision)
            if self.publishing or self.previewing or self.floating.app.busy:
                raise GuardError('请等待当前步骤结束，再核对本机草稿')
            if action == 'save':
                self._editing()
                if not self.operations or self.checkpoint_revision == revision:
                    raise GuardError('没有需要更新的本机草稿')
                record, encoded = self.drafts.prepare(self.original, self.data, self.operations)
                changes = deepcopy(self.summaries)
                candidate = self.data
                previous = self.drafts.read()
                record_hash = previous[1] if previous else None
            elif action == 'restore':
                if self.revision or self.operations or self.recovery_dismissed:
                    raise GuardError('当前已有编辑或已选择原稿，不能覆盖恢复')
                found = self.drafts.read()
                if not found:
                    raise GuardError('未找到当前准确源版本的本机草稿')
                record, record_hash = found
                candidate, changes = self._replay_draft(record)
                encoded = None
            else:
                raise GuardError('本机草稿操作无效')
            self.source_guard.check_source()
            identifier = uuid.uuid4().hex
            self.checkpoint = {'id': identifier, 'action': action, 'revision': revision, 'deadline': time.monotonic()+300,
                               'record': record, 'encoded': encoded, 'record_hash': record_hash,
                               'candidate': candidate, 'changes': changes}
            return {'ok': True, 'plan': {'id': identifier, 'action': action, 'source_name': self.source.name,
                'count': len(record['operations']), 'saved_at': record['saved_at'], 'location': str(self.drafts.target),
                'changes': changes, 'notice': '仅保存或恢复编辑内容；不恢复模型请求、接口凭据、写入或撤销批准；原文件不变'}}

    def _replay_draft(self, record):
        data, changes = self.original, []
        try:
            for operation in record['operations']:
                self._check_protection(operation, data)
                data, summary = engine.transform(data, self.source.suffix.lower(), operation)
                changes.append(summary)
            if package_identity(data, self.source.suffix.lower()) != record['candidate_package']:
                raise GuardError('本机草稿操作回放与保存身份不符')
        except (KeyError, TypeError, OverflowError):
            raise GuardError('本机草稿操作无效，未载入') from None
        return data, changes

    def draft_confirm(self, document_id, revision, plan_id, confirmed):
        with self.lock:
            self.guard(document_id, revision)
            plan = self.checkpoint
            if type(confirmed) is not bool:
                raise GuardError('本机草稿需要独立明确确认')
            if (not plan or plan_id != plan['id'] or revision != plan['revision'] or time.monotonic() > plan['deadline']
                    or self.publishing or self.previewing or self.floating.app.busy):
                raise GuardError('本机草稿计划已失效，请重新核对')
            self.checkpoint = None
            if not confirmed:
                return {'ok': True, 'cancelled': True, **self.state()}
            self.source_guard.check_source()
            if plan['action'] == 'save':
                try:
                    self.drafts.write(plan['record'], plan['encoded'], plan['record_hash'])
                    self.source_guard.check_source()
                    self.recovery = self.drafts.read()
                except (ValueError, OSError):
                    self.recovery_state = 'failed'
                    self.recovery_status = '本机草稿保存未确认；内存编辑保留，请核对本机位置或生成副本'
                    raise GuardError(self.recovery_status) from None
                self.checkpoint_revision = revision
                self.recovery_state = 'saved'
                self.recovery_status = '本机草稿已保存；不是已交付副本'
            else:
                found = self.drafts.read()
                if not found or found[1] != plan['record_hash']:
                    raise GuardError('本机草稿记录在确认前改变，请重新核对')
                self._replay_draft(found[0])
                self.source_guard.check_source()
                self.ai.invalidate()
                from .canvas_ai import CanvasAI
                self.ai = CanvasAI(self)
                if self.floating.runtime:
                    self.floating.runtime.stop()
                self.floating.runtime = None
                self.floating.app.pending_id = None
                self.floating.app.editor.plans.clear()
                self.data = plan['candidate']
                self.operations = deepcopy(plan['record']['operations'])
                self.summaries = deepcopy(plan['changes'])
                self.undo_groups = [1]*len(self.operations)
                self.redo_groups = []
                self.range_plan = None
                self.history_origin = 'restored_operations'
                self.recovery = found
                self.revision += 1
                self.plans.clear(); self.preview_cache = None; self.permit = None
                self.recovery_dismissed = True
                self.checkpoint_revision = self.revision
                self.recovery_state = 'restored'
                self.recovery_status = '草稿已恢复；模型要求和批准未恢复，保存副本仍需独立确认'
            return {'ok': True, **self.state()}

    def use_original(self, document_id, revision):
        with self.lock:
            self.guard(document_id, revision)
            if self.revision or self.operations or self.floating.app.busy or self.previewing or self.publishing:
                raise GuardError('当前不能切换为原稿')
            self.checkpoint = None
            self.recovery_dismissed = True
            self.recovery_status = '本次使用原稿；已保存的本机草稿保留'
            return {'ok': True, **self.state()}

    def before_admit(self):
        if not self.switch_allowed():
            raise GuardError('有未保存的表格/演示草稿；请先保存副本或明确放弃')
        self.permit = None

    def switch_allowed(self):
        return not self.active() or not self.publishing and not self.previewing and not self.uncertain and (not self.state()['dirty'] or self.permit == self.revision)

    def discard(self, document_id, revision, confirmed):
        with self.lock:
            self.guard(document_id, revision)
            if confirmed is not True or self.publishing or self.previewing:
                raise GuardError('放弃草稿需要明确确认，且须等待当前步骤结束')
            self.permit = self.revision
            self.ai.invalidate()
            return {'ok': True}

    def preview(self, document_id, revision, item_id, page=1):
        with self.lock:
            self.guard(document_id, revision)
            if self.previewing or self.publishing or self.floating.app.busy:
                raise GuardError('已有预览或发布正在进行')
            if type(page) is not int or not 1 <= page <= 100:
                raise GuardError('预览页码无效，最多预览100页')
            suffix = self.source.suffix.lower()
            items = engine.inventory(engine.load(self.data, suffix), suffix)
            item = engine.item_for(engine.load(self.data, suffix), suffix, item_id)
            index = next(index for index, entry in enumerate(items) if entry['id'] == item['id']) if suffix == '.pptx' else page-1
            data = self.data; self.previewing = True
            token = self.preview_token = object()
        try:
            runtime = runtime_path()
            identity = environment_identity(runtime)
            worker = Path(__file__).with_name('office_canvas_worker.py')
            key = (sha(data), index, identity, implementation_identity(), sha(worker.read_bytes()), 'png-1400-v1')
            with self.lock:
                self.guard(document_id, revision)
                if self.preview_token is not token:
                    raise GuardError('预览所属文档已改变')
                if identity.get('complete') and self.preview_cache and self.preview_cache[0] == key:
                    return {**deepcopy(self.preview_cache[1]), 'document_id': document_id, 'draft_revision': revision, 'item_id': item_id}
            with tempfile.TemporaryDirectory(prefix='ofc-', dir=_temporary_root()) as raw:
                scratch = Path(raw); source = scratch/('document'+suffix); source.write_bytes(data); runs = []
                package_projection(source, for_render=True)
                pdf = _convert(source, scratch/'pdf', runtime, 30000, runs)
                _run((sys.executable, str(worker), str(pdf), str(index), str(scratch)), 15000, runs)
                report = json.loads((scratch/'page.json').read_text('utf8'))
                if suffix == '.pptx' and report.get('pages') != len(items):
                    raise GuardError('渲染页数与幻灯片清单不符，不能按页序猜测定位；请使用文字清单并核对原文件')
                image = scratch/'page.png'
                if image.stat().st_size > 3*1024*1024:
                    raise GuardError('页面图像超过展示预算')
                result = {'ok': True, 'document_id': document_id, 'draft_revision': revision, 'item_id': item_id,
                          'image': 'data:image/png;base64,'+base64.b64encode(image.read_bytes()).decode(), **report,
                          'renderer': 'LibreOffice/PDFium · 静态页图，不是PowerPoint/Excel实时画布', 'processes': runs}
            with self.lock:
                self.guard(document_id, revision)
                if implementation_identity() != key[3] or sha(worker.read_bytes()) != key[4]:
                    raise GuardError('转换或页图实现已变化，请重新核对预览')
                if identity.get('complete') and environment_identity(runtime) != identity:
                    raise GuardError('渲染环境已变化，请重新生成页图')
                if identity.get('complete'): self.preview_cache = (key, deepcopy(result))
            return result
        finally:
            with self.lock:
                if self.preview_token is token:
                    self.previewing = False
                    self.preview_token = None

    def plan(self, document_id, revision):
        with self.lock:
            self.guard(document_id, revision)
            self._editing()
            if self.publishing or self.previewing or self.floating.app.busy or not self.state()['dirty']:
                raise GuardError('请先完成实际编辑并等待当前步骤结束')
            proof = engine.verify(self.original, self.data, self.source.suffix.lower(), self.operations)
            for operation in self.operations:
                self._check_protection(operation, self.data)
            changes = deepcopy(self.summaries)
            if self.source.suffix.lower() == '.xlsx':
                raw = engine.load(self.data, '.xlsx')
                calculators = {}
                for operation, change in zip(self.operations, changes):
                    if operation['action'] != 'formula':
                        continue
                    item = engine.item_for(raw, '.xlsx', operation['item_id'])
                    if item['id'] not in calculators:
                        calculators[item['id']] = engine.calculator_for(raw, item)
                    change['calculation_at_adoption'] = change['calculation']
                    change['calculation'] = calculators[item['id']].calculate(operation['target'])
                    change['final_formula'] = '='+calculators[item['id']].cells[operation['target']].findtext(engine.S+'f')
                    change['calculation_scope'] = '最终草稿；28位十进制预览，保存后由Office重算'
            identifier = uuid.uuid4().hex
            self.plans.clear()
            plan = {'id': identifier, 'document_id': document_id, 'draft_revision': revision, 'source': str(self.source),
                    'source_sha256': sha(self.original), 'candidate_sha256': sha(self.data), 'changes': changes,
                    'output': str(self.source.with_name(f'{self.source.stem}.agent-{identifier[:8]}{self.source.suffix}')),
                    'proof': proof, 'confirmation_required': True, 'mode': '净稿副本；保留原包非目标部件，不支持原生修订',
                    'instruction_binding': self._instruction_binding(),
                    'protection_binding': sorted(self.floating.protected_cells),
                    'implementation': canvas_implementation(),
                    'deadline': time.monotonic()+300, 'candidate': self.data, 'operations': deepcopy(self.operations)}
            self.plans[identifier] = plan
            return {'ok': True, 'plan': {key: value for key, value in plan.items() if key not in {'deadline', 'candidate', 'operations'}}}

    def publish(self, document_id, revision, plan_id, confirmed, access_id=None):
        with self.lock:
            self.guard(document_id, revision)
            authorization = self.floating.access.require(access_id, 'publish', document_id) if access_id is not None else None
            if authorization:
                if revision <= self.published_revision or self.ai.running is not None:
                    raise GuardError('当前草稿已保存或仍在处理，未重复自动保存')
                confirmed = True
            plan = self.plans.pop(plan_id, None)
            if confirmed is not True:
                return {'ok': True, 'cancelled': True}
            if not plan or plan['draft_revision'] != revision or time.monotonic() > plan['deadline'] or self.publishing or self.floating.app.busy:
                raise GuardError('计划过期、草稿改变或已有任务；请重新预览确认')
            if plan['instruction_binding'] != self._instruction_binding():
                raise GuardError('要求或模型配置已改变，请重新预览确认')
            if plan['protection_binding'] != sorted(self.floating.protected_cells):
                raise GuardError('保护范围已改变，请重新核对计划')
            if plan['implementation'] != canvas_implementation():
                raise GuardError('编辑实现已改变，请重新核对计划')
            self.source_guard.check_source()
            if authorization:
                plan['authorization'] = authorization
            self.publishing = self.floating.app.busy = True
            threading.Thread(target=self._publish, args=(plan, self.source_guard, self.floating.app.editor, self.floating.epoch), daemon=True).start()
            return {'ok': True, 'publishing': True}

    def _instruction_binding(self):
        runtime = self.ai.runtime
        if runtime is None:
            return None
        route = getattr(runtime, 'model_route', None)
        current = None
        if route:
            from ..settings_store import SettingsError
            try:
                current = self.floating.api._connections.capture('agent').public()
            except SettingsError as error:
                if str(error) not in {'model_selection_required', 'model_selection_invalid', 'model_configuration_changed'}:
                    raise
                current = {'unavailable': True, 'reason': str(error)}
        return {'session_id': runtime.session_id, 'revision': runtime.revision,
                'model': current}

    def _publish(self, plan, source_guard, editor, epoch):
        source = source_guard.source
        job = editor.audit/plan['id']; candidate = job/('working'+source.suffix)
        log = {key: value for key, value in plan.items() if key not in {'deadline', 'candidate', 'operations'}}
        log.update(kind='office_canvas', target='已批准的表格/演示操作', before='原稿', after='批准草稿', risks=['净稿；Office人工排版未验证'], status='failed')
        output = Path(plan['output']); linked = False
        try:
            source_guard.check_source()
            _audit_path(editor.audit, directory=True, missing=True); editor.audit.mkdir(exist_ok=True)
            _audit_path(editor.audit, directory=True); job.mkdir(); _audit_path(job, directory=True)
            backup = job/('backup'+source.suffix); shutil.copy2(source, backup); log['backup'] = str(backup)
            if digest(backup) != plan['source_sha256']: raise GuardError('备份摘要不一致')
            with candidate.open('xb') as stream:
                stream.write(plan['candidate']); stream.flush(); os.fsync(stream.fileno())
            log['verification'] = engine.verify(backup.read_bytes(), candidate.read_bytes(), source.suffix.lower(), plan['operations'])
            render = compare_documents(editor.space.root, backup, candidate)
            log['render_diff'] = render
            if render.get('inputs_unchanged') is not True or render.get('render', {}).get('status') not in {'equivalent', 'different'}:
                raise GuardError('渲染或重开验证未通过，未发布副本')
            source_guard.check_source()
            if digest(candidate) != plan['candidate_sha256'] or digest(backup) != plan['source_sha256']:
                raise GuardError('验证期间候选或备份发生变化')
            target = editor.space.file(str(output), existing=False)
            os.link(candidate, target); linked = True
            log.update(status='verified', output_sha256=digest(target))
            if log['output_sha256'] != plan['candidate_sha256']: raise GuardError('输出身份不一致')
            (job/'audit.json').write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf8')
            with self.lock:
                if self.floating.epoch != epoch or self.document_id != plan['document_id']:
                    return
                self.published_revision = plan['draft_revision']
                self.floating.outputs[plan['id']] = {**log, 'source_version': plan['source_sha256'],
                    'instruction_revision': (plan.get('instruction_binding') or {}).get('revision', 0)}
                self.floating.stage = 'verified'
                self.floating.notice('draft_published', {'receipt': log, 'draft_revision': plan['draft_revision']})
        except Exception:
            with self.lock:
                current = self.floating.epoch == epoch and self.document_id == plan['document_id']
                uncertain = linked or output.exists()
                log['status'] = 'uncertain' if uncertain else 'failed'
                if current:
                    self.uncertain = uncertain
                    self.last_error = '副本结果待核实，请勿重复发布' if uncertain else '生成或渲染验证失败；未发布副本，草稿保留'
                    self.floating.notice('draft_publish_failed', {'text': self.last_error, 'uncertain': self.uncertain})
                try:
                    if job.is_dir(): (job/'audit.json').write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf8')
                except OSError:
                    self.uncertain = True
        finally:
            try:
                if candidate.is_file(): candidate.unlink()
            finally:
                if self.floating.epoch == epoch and self.document_id == plan['document_id']: self.publishing = False
                self.floating.app.busy = False


class OfficeCanvasApi:
    def canvas_prepare_panes(self, document_id, revision, item_id, rows, cols):
        return self._floating_call(self._floating.canvas.prepare_panes, document_id, revision, item_id, rows, cols)

    def canvas_commit_panes(self, document_id, revision, plan_id, confirmed=False):
        return self._floating_call(self._floating.canvas.commit_range, document_id, revision, plan_id, confirmed)

    def canvas_prepare_geometry(self, document_id, revision, item_id, target, action, value):
        return self._floating_call(self._floating.canvas.prepare_geometry, document_id, revision, item_id, target, action, value)

    def canvas_commit_geometry(self, document_id, revision, plan_id, confirmed=False):
        return self._floating_call(self._floating.canvas.commit_range, document_id, revision, plan_id, confirmed)

    def canvas_copy_range(self, document_id, revision, item_id, start, end):
        return self._floating_call(self._floating.canvas.copy_range, document_id, revision, item_id, start, end)

    def canvas_prepare_range(self, document_id, revision, item_id, start, text, mode='text'):
        return self._floating_call(self._floating.canvas.prepare_range, document_id, revision, item_id, start, text, mode)

    def canvas_commit_range(self, document_id, revision, plan_id, confirmed=False):
        return self._floating_call(self._floating.canvas.commit_range, document_id, revision, plan_id, confirmed)

    def canvas_redo(self, document_id, revision):
        return self._floating_call(self._floating.canvas.redo, document_id, revision)

    def canvas_draft_plan(self, document_id, revision, action):
        return self._floating_call(self._floating.canvas.draft_plan, document_id, revision, action)

    def canvas_draft_confirm(self, document_id, revision, plan_id, confirmed=False):
        return self._floating_call(self._floating.canvas.draft_confirm, document_id, revision, plan_id, confirmed)

    def canvas_use_original(self, document_id, revision):
        return self._floating_call(self._floating.canvas.use_original, document_id, revision)

    def canvas_ai_reference(self, document_id, revision, item_id, target, allow_numeric_change=False):
        return self._floating_call(self._floating.canvas.ai.select, document_id, revision, item_id, target, allow_numeric_change)

    def canvas_selection_scope(self, document_id, revision, item_id, axis, index):
        return self._floating_call(self._floating.canvas.selection_scope, document_id, revision, item_id, axis, index)

    def canvas_ai_reference_range(self, document_id, revision, item_id, start, end, numeric_targets=None):
        return self._floating_call(self._floating.canvas.ai.select_range, document_id, revision, item_id, start, end, numeric_targets)

    def canvas_ai_submit(self, document_id, revision, reference_id, session_id, client_id, text):
        return self._floating_call(self._floating.canvas.ai.submit, document_id, revision, reference_id, session_id, client_id, text)

    def canvas_ai_change(self, document_id, revision, session_id, client_id, target_id, version, instruction_revision, action, text):
        return self._floating_call(self._floating.canvas.ai.change, document_id, revision, session_id, client_id, target_id, version, instruction_revision, action, text)

    def canvas_ai_resume(self, document_id, revision, reference_id, session_id, instruction_revision):
        return self._floating_call(self._floating.canvas.ai.resume, document_id, revision, reference_id, session_id, instruction_revision)

    def canvas_ai_adopt(self, document_id, revision, proposal_id, access_id=None):
        return self._floating_call(self._floating.canvas.ai.adopt, document_id, revision, proposal_id, access_id)

    def canvas_ai_reject(self, document_id, revision, proposal_id):
        return self._floating_call(self._floating.canvas.ai.reject, document_id, revision, proposal_id)

    def canvas_ai_stop(self, document_id, revision):
        return self._floating_call(self._floating.canvas.ai.stop, document_id, revision)

    def canvas_open(self, document_id, version):
        return self._floating_call(self._floating.canvas.open, document_id, version)

    def canvas_view(self, document_id, revision, item_id, row=1, col=1):
        return self._floating_call(self._floating.canvas.view, document_id, revision, item_id, row, col)

    def canvas_prefetch(self, document_id, revision, item_id, row=1, col=1):
        return self._floating_call(self._floating.canvas.prefetch, document_id, revision, item_id, row, col)

    def canvas_apply(self, document_id, revision, operation):
        return self._floating_call(self._floating.canvas.apply, document_id, revision, operation)

    def canvas_undo(self, document_id, revision):
        return self._floating_call(self._floating.canvas.undo, document_id, revision)

    def canvas_preview(self, document_id, revision, item_id, page=1):
        return self._floating_call(self._floating.canvas.preview, document_id, revision, item_id, page)

    def canvas_plan(self, document_id, revision):
        return self._floating_call(self._floating.canvas.plan, document_id, revision)

    def canvas_publish(self, document_id, revision, plan_id, confirmed=False, access_id=None):
        return self._floating_call(self._floating.canvas.publish, document_id, revision, plan_id, confirmed, access_id)

    def canvas_discard(self, document_id, revision, confirmed=False):
        return self._floating_call(self._floating.canvas.discard, document_id, revision, confirmed)
