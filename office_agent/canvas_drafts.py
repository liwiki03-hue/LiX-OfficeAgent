from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import uuid

from . import office_canvas as engine
from .rag_evidence import _read, EvidenceError, MAX_BYTES
from .writes import GuardError, _audit_path, digest


def package_identity(data, suffix):
    return engine.sha(json.dumps({name: engine.sha(value) for name, value in engine.load(data, suffix).items()}, sort_keys=True).encode())


def implementation():
    return [digest(Path(engine.__file__)), digest(Path(engine.__file__).with_name('canvas_formulas.py')),
            digest(Path(engine.__file__).with_name('canvas_geometry.py')), digest(Path(engine.__file__).with_name('canvas_panes.py'))]


def _pairs(values):
    if len(dict(values)) != len(values):
        raise GuardError('本机草稿包含重复字段')
    return dict(values)


def _constant(value):
    raise GuardError('本机草稿包含非法数值')


class CanvasDraftStore:
    def __init__(self, audit, actor, source_key, source_version, file_identity, suffix):
        if not re.fullmatch('[a-f0-9]{64}', source_key) or suffix not in {'.xlsx', '.pptx'}:
            raise GuardError('本机草稿来源身份无效')
        self.audit = audit
        self.actor = hashlib.sha256(actor.encode()).hexdigest()
        self.key, self.version, self.file_identity, self.suffix = source_key, source_version, file_identity, suffix
        self.directory = audit / 'canvas-drafts' / source_key[:24]
        self.target = self.directory / ('draft-' + self.actor[:12] + '.json')

    def _directory(self, create=False):
        for path in [self.audit, self.audit / 'canvas-drafts', self.directory.parent, self.directory]:
            if not path.exists() and not create:
                _audit_path(path.parent, directory=True)
                _audit_path(path, directory=True, missing=True)
                return False
            _audit_path(path, directory=True, missing=create)
            if create:
                if path == self.directory and not path.exists():
                    from itertools import islice
                    if len(list(islice(path.parent.iterdir(), 33))) >= 32:
                        raise GuardError('本项目已保存32份来源草稿，请先核对本机草稿目录')
                path.mkdir(exist_ok=True)
            _audit_path(path, directory=True)
        return True

    def read(self):
        if not self._directory():
            return None
        _audit_path(self.target, directory=False, missing=True)
        if not self.target.exists():
            return None
        try:
            data, stamp = _read(self.target)
            record = json.loads(data, object_pairs_hook=_pairs, parse_constant=_constant)
            self.validate(record)
            return record, engine.sha(data)
        except (ValueError, KeyError, TypeError, OSError, EvidenceError, RecursionError):
            raise GuardError('本机草稿损坏、身份不符或不兼容；未恢复，原稿仍可使用') from None

    def validate(self, record):
        fields = {'schema_version', 'actor', 'source_key', 'source_version', 'file_identity', 'suffix', 'implementation', 'operations', 'candidate_package', 'saved_at'}
        if (type(record) is not dict or set(record) != fields or type(record['schema_version']) is not int or record['schema_version'] != 1
                or record['actor'] != self.actor or record['source_key'] != self.key or record['source_version'] != self.version
                or record['file_identity'] != self.file_identity or record['suffix'] != self.suffix or record['implementation'] != implementation()
                or type(record['operations']) is not list or not 1 <= len(record['operations']) <= 100
                or type(record['candidate_package']) is not str or not re.fullmatch('[a-f0-9]{64}', record['candidate_package'])
                or type(record['saved_at']) is not str or len(record['saved_at']) > 40):
            raise GuardError('本机草稿身份、版本或容量不符')
        datetime.fromisoformat(record['saved_at'])
        for operation in record['operations']:
            if (type(operation) is not dict or set(operation) != {'item_id', 'target', 'action', 'value', 'expected'}
                    or any(type(operation[key]) is not str for key in ('item_id', 'target', 'action'))
                    or operation['action'] not in {'text', 'number', 'clear', 'format', 'formula', engine.canvas_panes.ACTION, *engine.canvas_geometry.ACTIONS}):
                raise GuardError('本机草稿操作不兼容')
            if operation['action'] in engine.canvas_geometry.ACTIONS:
                expected = operation['expected']
                if self.suffix != '.xlsx' or expected is not None and (type(expected) is not dict or len(expected) > 32 or any(type(key) is not str or type(value) is not str for key, value in expected.items())):
                    raise GuardError('本机尺寸草稿前置条件无效')
            elif type(operation['expected']) is not str:
                raise GuardError('本机草稿文字前置条件无效')

    def prepare(self, original, candidate, operations):
        engine.verify(original, candidate, self.suffix, operations)
        record = {'schema_version': 1, 'actor': self.actor, 'source_key': self.key, 'source_version': self.version,
                  'file_identity': self.file_identity, 'suffix': self.suffix, 'implementation': implementation(),
                  'operations': deepcopy(operations), 'candidate_package': package_identity(candidate, self.suffix),
                  'saved_at': datetime.now(timezone.utc).isoformat()}
        self.validate(record)
        encoded = json.dumps(record, ensure_ascii=False, allow_nan=False).encode('utf8')
        if len(encoded) > MAX_BYTES:
            raise GuardError('本机草稿记录超过512 KiB，请先生成副本；内存编辑仍保留')
        return record, encoded

    def write(self, record, encoded, expected_hash):
        self.validate(record)
        if len(encoded) > MAX_BYTES or json.loads(encoded) != record:
            raise GuardError('本机草稿快照不一致')
        self._directory(True)
        from itertools import islice
        entries = list(islice(self.directory.iterdir(), 9))
        if len(entries) >= 8 or sum(path.lstat().st_size for path in entries) > 4*1024*1024:
            raise GuardError('本机草稿目录达到容量，请核对遗留记录；内存编辑保留')
        previous = self.read()
        if (previous[1] if previous else None) != expected_hash:
            raise GuardError('本机草稿在计划后已变化，请重新核对')
        parent = self.directory.stat()
        temporary = self.directory / (uuid.uuid4().hex + '.tmp')
        try:
            with temporary.open('xb') as stream:
                stream.write(encoded)
                stream.flush()
                os.fsync(stream.fileno())
            self._directory()
            if self.read() != previous or (self.directory.stat().st_dev, self.directory.stat().st_ino) != (parent.st_dev, parent.st_ino):
                raise GuardError('本机草稿在保存期间发生变化，请重新核对')
            os.replace(temporary, self.target)
            if _read(self.target)[0] != encoded:
                raise GuardError('本机草稿保存结果待核对；请保留内存编辑')
        finally:
            temporary.unlink(missing_ok=True)
