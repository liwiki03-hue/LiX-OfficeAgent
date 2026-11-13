import json
from pathlib import Path
import re
import stat

from .rag_evidence import _read
from .writes import GuardError, _task_literal_audit


def file_identity(editor, path):
    editor.space.file(str(path), existing=path.exists())
    try:
        info = path.lstat()
    except FileNotFoundError:
        return None
    if (not stat.S_ISREG(info.st_mode) or info.st_nlink != 1 or path.is_symlink() or
            getattr(info, 'st_file_attributes', 0) & 0x400):
        raise GuardError('professional_dependency_stale')
    return [info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns]


class OutputDependencies:
    def __init__(self):
        self.root = None
        self.known = {}

    def invalidate(self):
        self.root = None
        self.known.clear()

    def _bind(self, editor):
        stat = editor.space.root.stat()
        root = (str(editor.space.root), stat.st_dev, stat.st_ino)
        if root != self.root:
            self.known.clear()
            self.root = root

    def _id(self, value):
        if type(value) is not str or not re.fullmatch('[a-f0-9]{32}', value):
            raise GuardError('professional_dependency_invalid')
        return value

    def admit(self, editor, receipt, source, parent=None):
        self._bind(editor)
        self._id(receipt)
        if parent is not None:
            self._id(parent)
        if receipt not in self.known and len(self.known) >= 32:
            raise GuardError('professional_dependency_capacity')
        self.known[receipt] = {'source': str(source), 'parent': parent}

    def _record(self, editor, receipt):
        try:
            path = editor.space.file(str(editor.audit / receipt / 'audit.json'))
            encoded, stamp = _read(path)
            audit = json.loads(encoded)
            source = Path(audit['source'])
            output = Path(audit['output'])
            expected = source.with_name(f'{source.stem}.agent-{receipt[:8]}{source.suffix}')
            if (audit['id'] != receipt or source.parent != editor.space.root or source.suffix.lower() != '.docx' or
                    output != expected or type(audit['status']) is not str):
                raise ValueError()
            editor.space.file(str(source), existing=source.exists())
            editor.space.file(str(output), existing=output.exists())
            if _task_literal_audit(audit):
                parents = []
            else:
                parents = [r['citation_manifest'].get('parent_receipt') for r in json.loads(audit['after'])['items'] if 'citation_manifest' in r]
            parent = next((p for p in parents if p is not None), None)
            if parent is not None:
                self._id(parent)
            if len(set(p for p in parents if p is not None)) > 1:
                raise ValueError()
            output_state = 'present_not_revalidated' if output.exists() else 'absent'
            output_identity = file_identity(editor, output)
            previous = self.known.get(receipt, {})
            if (any(key in previous and previous[key] != value for key, value in [('source', str(source)), ('output', str(output)), ('parent', parent)]) or
                    (audit['status'] == 'rolled_back' and output_state != 'absent')):
                raise ValueError()
            if _read(path)[1] != stamp:
                raise ValueError()
            record = {'source': str(source), 'output': str(output), 'parent': parent, 'status': audit['status'],
                      'receipt_id': receipt, 'output_name': output.name, 'output_state': output_state, 'readable': True,
                      '_audit_identity': json.loads(json.dumps(stamp)), '_output_identity': output_identity}
            self.known[receipt] = record
            return record
        except Exception:
            return {**self.known.get(receipt, {}), 'receipt_id': receipt, 'status': 'unavailable_or_refused', 'readable': False}

    def preview(self, editor, target, receipts):
        self._bind(editor)
        self._id(target)
        if type(receipts) is not list or len(receipts) > 8 or len(set(self._id(p) for p in receipts)) != len(receipts):
            raise GuardError('professional_dependency_invalid')
        ids = set(self.known) | set(receipts) | {target}
        if len(ids) > 32:
            raise GuardError('professional_dependency_capacity')
        for receipt in ids:
            self.known.setdefault(receipt, {})
        rows = {receipt: self._record(editor, receipt) for receipt in sorted(ids)}
        parent = rows[target]
        blockers = []
        for receipt, row in rows.items():
            if receipt == target or row['status'] == 'rolled_back':
                continue
            reason = ('unavailable_receipt' if not row['readable'] else 'parent_receipt' if row.get('parent') == target else
                      'source_output' if parent.get('output') and row.get('source') == parent['output'] else None)
            if reason:
                blockers.append({'receipt_id': receipt, 'reason': reason, 'status': row['status']})
        return {'target': target, 'records': [{k: v for k, v in row.items() if k not in {'source', 'output', 'parent'} and not k.startswith('_')} for row in rows.values()],
                'blockers': blockers, 'target_readable': parent['readable'], 'known_receipts_only': True,
                'complete_dependency_graph': False, 'write_approval_granted': False, 'restores_rollback_authority': False,
                'wrote_nothing': True}

    def check(self, editor, target):
        result = self.preview(editor, target, [])
        if not result['target_readable'] or result['blockers']:
            raise GuardError('professional_output_in_use')
        row = next(r for r in result['records'] if r['receipt_id'] == target)
        if row.get('output_state') != 'present_not_revalidated' or row['status'] != 'verified':
            raise GuardError('professional_dependency_stale')
