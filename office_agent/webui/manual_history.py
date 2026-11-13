from datetime import datetime, timezone
from itertools import islice
import json
from pathlib import Path
import re

from ..writes import GuardError, _audit_path, digest


KINDS = {'paged_candidate': {'.docx'}, 'paged_editable_copy': {'.docx'}, 'office_canvas': {'.xlsx', '.pptx'}}
STATES = {'verified': '副本已验证', 'failed': '发布失败', 'uncertain': '结果待核实', 'rolled_back': '副本已撤销', 'rolling_back': '撤销结果待核实'}
LIMIT = 512 * 1024


def _root(editor):
    _audit_path(editor.space.root, directory=True)
    _audit_path(editor.audit, directory=True, missing=True)
    root = editor.space.root.resolve(strict=True)
    audit = editor.audit.resolve(strict=False)
    if audit.parent != root:
        raise GuardError('历史审计目录不在当前授权目录内')
    return audit


def _record(editor, directory):
    _audit_path(directory, directory=True)
    path = directory / 'audit.json'
    _audit_path(path, directory=False)
    if path.stat().st_size > LIMIT:
        raise GuardError('副本审计超过读取上限')
    with path.open('rb') as stream:
        raw = stream.read(LIMIT + 1)
    if len(raw) > LIMIT:
        raise GuardError('副本审计超过读取上限')
    row = json.loads(raw)
    if not isinstance(row, dict) or row.get('kind') not in KINDS:
        return None
    if row.get('id') != directory.name or not re.fullmatch('[a-f0-9]{32}', directory.name):
        raise GuardError('副本事务身份不符')
    source, output = Path(row['source']), Path(row['output'])
    for candidate in (source, output):
        if not candidate.is_absolute() or candidate.suffix.lower() not in KINDS[row['kind']]:
            raise GuardError('副本路径或格式不符')
        if candidate.parent != editor.space.root:
            raise GuardError('副本不属于当前授权目录')
        _audit_path(candidate, directory=False, missing=True)
        if editor.space.file(str(candidate), existing=candidate.exists()) != candidate:
            raise GuardError('副本路径身份不符')
    expected = source.with_name(f'{source.stem}.agent-{directory.name[:8]}{source.suffix}')
    if output != expected or row.get('status') not in STATES:
        raise GuardError('副本名称或事务状态不符')
    status = row['status']
    valid_hash = isinstance(row.get('output_sha256'), str) and re.fullmatch('[a-f0-9]{64}', row['output_sha256'])
    if status == 'verified' and not valid_hash:
        raise GuardError('已验证副本缺少身份摘要')
    text = STATES[status]
    return {'session_dir': 'manual:' + row['id'], 'record_kind': 'manual_copy', 'transaction_id': row['id'],
            'file': str(source), 'output': str(output), 'output_sha256': row.get('output_sha256'),
            'question': output.name, 'format': output.suffix[1:].upper(), 'status': status,
            'created_at': datetime.fromtimestamp(path.stat().st_mtime, timezone.utc).isoformat(),
            'time_basis': '审计文件更新时间', 'current': False, 'history_available': True,
            'history_truncated': False, 'event_count': 1,
            'result_summary': text + ('；简化文字副本，原样式与分页已重建' if row['kind']=='paged_editable_copy' else '') + '；打开时重新校验文件并请求读取同意，不恢复旧批准。',
            'history_status': '来自副本事务审计；刷新重新读取'}


def listing(editor):
    result = {'records': [], 'warning': ''}
    if editor is None:
        return result
    try:
        audit = _root(editor)
        if not audit.exists():
            return result
        candidates = list(islice(audit.iterdir(), 501))
        warnings = []
        if len(candidates) > 500:
            warnings.append('审计目录超过500项，本次仅读取首500项，历史可能不完整')
        remaining = 8 * 1024 * 1024
        for directory in candidates[:500]:
            if not re.fullmatch('[a-f0-9]{32}', directory.name):
                continue
            try:
                _audit_path(directory, directory=True)
                artifact = directory / 'audit.json'
                _audit_path(artifact, directory=False, missing=True)
                if not artifact.exists():
                    continue
                size = artifact.stat().st_size
                if size > remaining:
                    warnings.append('副本审计读取达到8 MiB上限，历史可能不完整')
                    break
                remaining -= size
                row = _record(editor, directory)
                if row:
                    result['records'].append(row)
            except (OSError, ValueError, TypeError, KeyError):
                warnings.append('部分副本审计无法安全读取；已有副本不受影响，可刷新重读')
        result['records'].sort(key=lambda row: (row['created_at'], row['session_dir']), reverse=True)
        if len(result['records']) > 100:
            warnings.append('仅显示扫描范围中最近100份副本')
        result['records'] = result['records'][:100]
        result['warning'] = '。'.join(dict.fromkeys(warnings))
    except (OSError, ValueError):
        result['warning'] = '副本历史暂不可读取；磁盘副本不受影响，请核对审计目录后刷新'
    return result


def resolve(editor, identifier):
    if not isinstance(identifier, str) or not re.fullmatch('[a-f0-9]{32}', identifier):
        raise GuardError('副本记录身份无效')
    row = _record(editor, _root(editor) / identifier)
    if row is None:
        raise GuardError('此记录不是手工副本事务')
    return row


def checked_output(editor, identifier):
    row = resolve(editor, identifier)
    if row['status'] != 'verified':
        raise GuardError(STATES[row['status']] + '，不能从此记录打开')
    if not Path(row['output']).exists():
        raise GuardError('副本文件不存在；请确认是否已移动或删除')
    try:
        path = editor.space.file(row['output'])
    except FileNotFoundError:
        raise GuardError('副本文件不存在；请确认是否已移动或删除') from None
    if path.stat().st_size > 50 * 1024 * 1024:
        raise GuardError('副本超过重新核对的大小上限')
    if digest(path) != row['output_sha256']:
        raise GuardError('副本已被外部修改；请重新选择当前文件，不沿用旧记录')
    return path


def detail(editor, identifier):
    row = resolve(editor, identifier)
    status = STATES[row['status']]
    if row['status'] == 'verified':
        try:
            checked_output(editor, identifier)
            status += ' · 当前摘要一致'
        except (OSError, ValueError) as error:
            status = str(error)
    return {'ok': True, 'readonly': True, 'session_dir': row['session_dir'], 'record_kind': 'manual_copy',
            'events': [{'type': 'output', 'text': row['question'] + '\n' + status}],
            'history_available': True, 'history_truncated': False, 'summary': status,
            'warning': '只读事务记录；打开需重新准入，不恢复写入或撤销批准'}
