import json
import hashlib
from pathlib import Path
import tempfile
import threading
import zipfile

from .output_review import OutputReview
from ..output_dependencies import file_identity
from ..range_review_archive import RangeReviewArchive, FORMAT, validate
from ..writes import GuardError, digest, _audit_record, _task_literal_audit
from ..review_diagnostics import initialize_review_diagnostics, review_phase, review_request, review_stage


class RangeReview(OutputReview):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.cache = None
        initialize_review_diagnostics(self)

    def _close_cache(self):
        cache, self.cache = self.cache, None
        if cache:
            cache['timer'].cancel()
            cache['directory'].cleanup()

    def invalidate(self, retain_conversions=False):
        with self._lock:
            with self.diagnostic_lock:
                self.diagnostics.clear()
            from ..render_handoff import clear
            if retain_conversions is not True:
                clear()
            self._close_cache()
            super().invalidate()

    def start(self, target, receipts):
        with self._lock:
            self._close_cache()
            return super().start(target, receipts)

    def _archive(self, editor):
        if self.archive_backend is None or self.archive_backend.root != editor.space.root:
            self.archive_backend = RangeReviewArchive(editor.space)
        return self.archive_backend

    def _verification(self, audit):
        if _task_literal_audit(audit):
            return {'reference_impact': None}
        return json.loads(audit['after'])['items'][0].get('verification', {})

    @review_stage('source_output_and_environment_identity')
    def _pair(self, editor, path, entry, refresh=False):
        receipt = self.dependencies._id(entry['target'])
        record = self.dependencies._record(editor, receipt)
        if not record.get('readable') or record['status'] != 'verified':
            raise GuardError('professional_review_changed')
        audit, _ = _audit_record(editor.space, editor.audit / receipt)
        source, output = Path(record['source']), Path(record['output'])
        routed = _task_literal_audit(audit)
        if path not in (source, output) or (audit['kind'] not in {'range_docx', 'block_docx'} and not routed):
            raise GuardError('professional_review_scope')
        identities = [file_identity(editor, item) for item in (source, output)]
        if any(item is None for item in identities):
            raise GuardError('professional_review_changed')
        if digest(source) != audit['source_sha256'] or digest(output) != audit['output_sha256']:
            raise GuardError('professional_review_changed')
        from ..fidelity import runtime_path, compare_documents
        from ..render_identity import environment_identity
        with review_phase('environment_identity'):
            environment = environment_identity(runtime_path())
        report = entry.get('_render_report', audit['render_diff'])
        if not environment['complete']:
            self._close_cache()
        legacy_environment = 'render_environment' not in report
        if environment['complete'] and not legacy_environment and report.get('render_environment') != environment:
            self._close_cache()
            if not refresh or not report.get('anchor_contexts'):
                raise GuardError('professional_review_environment_changed')
            report = compare_documents(editor.space.root, source, output, static_docx=True, range_anchors=report['anchor_contexts'])
            if report.get('inputs_unchanged') is not True or report.get('render', {}).get('status') not in {'equivalent', 'different'} or report.get('render_environment') != environment:
                raise GuardError('professional_review_revalidation_failed')
            entry['_render_report'] = report
        alignment = report['render'].get('anchor_alignment', [])
        if not 1 <= len(alignment) <= 12 or report['inputs_unchanged'] is not True:
            raise GuardError('professional_review_alignment_missing')
        pair = {'source': audit['source_sha256'], 'output': audit['output_sha256'],
                'analysis': self._hash({'alignment': alignment, 'projection': report.get('static_projection_sha256'), 'runtime': report.get('runtime_sha256'), 'environment': environment,
                                       'reference_impact': self._verification(audit).get('reference_impact')})}
        if identities != [file_identity(editor, item) for item in (source, output)]:
            raise GuardError('professional_review_changed')
        return pair, identities, (source, output), {**report, '_page_environment_verified': environment['complete'] and report.get('render_environment') == environment,
                                                   '_observed_environment': environment}

    @review_stage('read_consent_and_pair')
    def _ensure(self, editor, path, entry):
        if not entry.get('consent'):
            raise GuardError('professional_review_read_consent')
        pair, identities, paths, report = self._pair(editor, path, entry)
        if entry.get('consent') != [pair, identities]:
            self._close_cache()
            raise GuardError('professional_review_read_consent')
        return pair, paths, report

    @review_request
    def open(self, identifier, approved=False):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            if type(approved) is not bool:
                raise GuardError('professional_invalid_request')
            if not approved:
                return {'status': 'cancelled', 'wrote_nothing': True}
            pair, identities, paths, report = self._pair(editor, path, entry, refresh=True)
            self._close_cache()
            old = entry.get('pair')
            items = entry.get('review_items')
            changed = old is not None and old != pair
            alignment = report['render']['anchor_alignment']
            if not items or changed or len(items) != len(alignment):
                items = [{'index': i, 'state': 'needs_review' if changed else 'unreviewed', 'reason': ''} for i in range(len(alignment))]
            entry.update(pair=pair, consent=[pair, identities], review_items=items)
            verification = self._verification(_audit_record(editor.space, editor.audit / entry['target'])[0])
            return {'review_id': identifier, 'revision': entry['revision'], 'alignment': alignment,
                    'reference_impact': verification.get('reference_impact'), 'items': items, 'source_changed_since_archive': changed,
                    'pair': pair, 'source_name': paths[0].name, 'output_name': paths[1].name,
                    'write_approval_granted': False, 'restores_rollback_authority': False,
                    'page_environment_verified': report['_page_environment_verified'],
                    'notice': '静态域呈现；审阅状态是用户记录，不代表事实正确、Word验证或写入批准。' +
                              ('字体/配置身份未获当前验证，仅恢复历史定位和用户记录；当前页图复用被拒绝。' if not report['_page_environment_verified'] else '')}
        return self._call(operation)

    @review_stage('validated_pdf_cache')
    def _pdfs(self, identifier, pair, paths, report, observed_environment=None, root=None, defer_environment_check=False):
        from ..fidelity import _temporary_root, runtime_path, _convert, bounded_bytes
        from ..range_edits import static_render_parts
        from ..render_identity import environment_identity
        environment = observed_environment if observed_environment is not None else environment_identity(runtime_path())
        if not environment['complete'] or report.get('render_environment') != environment:
            self._close_cache()
            raise GuardError('professional_review_environment_changed')
        if self.cache and self.cache['key'] == [identifier, pair]:
            return self.cache
        self._close_cache()
        runtime = runtime_path()
        if digest(runtime) != report.get('runtime_sha256'):
            raise GuardError('professional_review_renderer_changed')
        temporary = tempfile.TemporaryDirectory(prefix='ofd-review-', dir=_temporary_root())
        directory = Path(temporary.name)
        processes, pdfs, projections = [], [], []
        try:
            from ..render_handoff import take
            with review_phase('verified_conversion_handoff'):
                retained = take(root if root is not None else paths[0].parent, report)
            for i, path in enumerate(paths):
                data = bounded_bytes(path)
                copy = directory / f'input-{i}.docx'
                copy.write_bytes(data)
                with zipfile.ZipFile(copy) as archive:
                    parts = static_render_parts({name: archive.read(name) for name in archive.namelist()})
                projection = hashlib.sha256(json.dumps({name: hashlib.sha256(payload).hexdigest() for name, payload in parts.items()}, sort_keys=True).encode('utf8')).hexdigest()
                if report.get('static_projection_sha256', [None, None])[i] != projection:
                    raise GuardError('professional_review_projection_changed')
                with zipfile.ZipFile(copy, 'w', zipfile.ZIP_DEFLATED) as archive:
                    for name, payload in parts.items():
                        archive.writestr(name, payload)
                projections.append(projection)
                if retained is None:
                    with review_phase('convert_before' if i == 0 else 'convert_after'):
                        pdfs.append(_convert(copy, directory / f'pdf-{i}', runtime, 30000, processes))
                else:
                    pdf = directory / f'retained-{i}.pdf'
                    pdf.write_bytes(retained[i])
                    pdfs.append(pdf)
            if digest(runtime) != report.get('runtime_sha256'):
                raise GuardError('professional_review_renderer_changed')
            if not defer_environment_check:
                with review_phase('environment_identity_after_conversion'):
                    if environment_identity(runtime) != environment:
                        raise GuardError('professional_review_environment_changed')
            timer = threading.Timer(300, self._expire_cache, args=(identifier,))
            timer.daemon = True
            cache = {'key': [identifier, pair], 'directory': temporary, 'pdfs': pdfs, 'hashes': [digest(p) for p in pdfs],
                     'projections': projections, 'processes': processes, 'timer': timer, 'conversion_reused': retained is not None,
                     'pending_environment': defer_environment_check}
            if defer_environment_check:
                return cache
            self.cache = cache
            timer.start()
            return self.cache
        except Exception:
            temporary.cleanup()
            raise

    def _expire_cache(self, identifier):
        with self._lock:
            if self.cache and self.cache['key'][0] == identifier:
                self._close_cache()

    @review_request
    def page(self, identifier, index, side, offset=0):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            pair, paths, report = self._ensure(editor, path, entry)
            alignment = report['render']['anchor_alignment']
            if type(index) is not int or not 0 <= index < len(alignment) or side not in ('before', 'after') or type(offset) is not int:
                raise GuardError('professional_invalid_request')
            row = alignment[index]
            pages = row[side + '_pages']
            if row[side + '_status'] != 'unique' or not 0 <= offset < len(pages) or len(pages) > 2 or any(type(p) is not int or not 1 <= p <= 40 for p in pages):
                raise GuardError('professional_review_alignment_missing')
            cache = None
            try:
                cache = self._pdfs(identifier, pair, paths, report, report['_observed_environment'], editor.space.root, defer_environment_check=True)
                from ..research_page import render_page
                position = 0 if side == 'before' else 1
                with review_phase('raster_before' if side == 'before' else 'raster_after'):
                    result = render_page(cache['pdfs'][position], cache['hashes'][position], pages[offset] - 1, [0, 0, 0, 0])
                self._ensure(editor, path, entry)
                if cache.get('pending_environment'):
                    cache['pending_environment'] = False
                    self.cache = cache
                    cache['timer'].start()
                return {**result, 'review_id': identifier, 'item': index, 'side': side, 'projection_sha256': cache['projections'][position],
                        'conversion_processes': cache['processes'], 'conversion_reused': cache.get('conversion_reused', False),
                        'notice': ('复用本进程五分钟内已验证转换；当前文档、投影和环境已重验。' if cache.get('conversion_reused') else '本次审阅重新转换。') +
                                  '指定渲染器的静态域页面；不是 Microsoft Word 或逐词像素验证。'}
            except Exception:
                if cache and cache.get('pending_environment') and cache is not self.cache:
                    cache['timer'].cancel()
                    cache['directory'].cleanup()
                self._close_cache()
                raise
        return self._call(operation)

    def mark(self, identifier, index, state, reason=''):
        def operation(editor, path, binding):
            entry = self._entry(identifier, binding)
            self._ensure(editor, path, entry)
            if type(index) is not int or not 0 <= index < len(entry['review_items']):
                raise GuardError('professional_invalid_request')
            candidate = [dict(item) for item in entry['review_items']]
            candidate[index] = {'index': index, 'state': state, 'reason': reason}
            document = {**self._document(editor, entry, entry['snapshot']), 'review_items': candidate}
            validate(document)
            entry['review_items'] = candidate
            entry['revision'] += 1
            return {'items': candidate, 'revision': entry['revision'], 'write_approval_granted': False}
        return self._archive_call(operation)

    def _document(self, editor, entry, snapshot):
        return {**super()._document(editor, entry, snapshot), 'format': FORMAT,
                'pair': entry.get('pair'), 'review_items': entry.get('review_items')}

    def plan_archive(self, identifier, revision):
        with self._lock:
            def operation(editor, path, binding):
                entry = self._entry(identifier, binding)
                self._ensure(editor, path, entry)
                result = super(RangeReview, self).plan_archive(identifier, revision)
                if result.get('ok'):
                    self.archive_plans[result['result']['plan_id']]['item_stamp'] = self._hash([entry['pair'], entry['review_items']])
                    return result['result']
                raise GuardError(result['reason'])
            return self._archive_call(operation)

    def confirm_archive(self, identifier, approved=False):
        def operation(editor, path, binding):
            plan = self.archive_plans.get(identifier) if type(identifier) is str else None
            if approved is True and plan:
                entry = self._entry(plan['review'], binding)
                self._ensure(editor, path, entry)
                if plan.get('item_stamp') != self._hash([entry['pair'], entry['review_items']]):
                    raise GuardError('professional_review_changed')
            reply = super(RangeReview, self).confirm_archive(identifier, approved)
            if not reply['ok']:
                raise GuardError(reply['reason'])
            return reply['result']
        return self._archive_call(operation)

    def restore_archive(self, save_id):
        def operation(editor, path, binding):
            self._close_cache()
            reply = super(RangeReview, self).restore_archive(save_id)
            if reply.get('ok'):
                document, _ = self.archive_backend.inspect(save_id)
                entry = self.reviews[reply['result']['review_id']]
                entry.update(pair=document['pair'], review_items=document['review_items'])
                reply['result'].update(items=document['review_items'], read_consent_required=True)
            if not reply['ok']:
                raise GuardError(reply['reason'])
            return reply['result']
        return self._archive_call(operation)
