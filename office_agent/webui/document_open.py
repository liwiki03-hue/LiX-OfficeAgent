from dataclasses import asdict
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time

from ..documents import read_word_text, inspect_xlsx


OPEN_FORMATS = {
    '.docx': {'kind': 'word', 'view': True, 'parse': True, 'edit': True, 'render': True},
    '.xlsx': {'kind': 'spreadsheet', 'view': True, 'parse': True, 'edit': True, 'render': True},
    '.pptx': {'kind': 'presentation', 'view': True, 'parse': True, 'edit': True, 'render': True},
}


def parse_payload(path, offset=0):
    path = Path(path)
    if path.suffix.lower() == '.pptx':
        from ..presentations import read_pptx_structure
        view = read_pptx_structure(path, offset=offset, limit=20)
        return {'view': {**view, 'has_more': view['truncated']}, 'positions': {}}
    if path.suffix.lower() == '.docx':
        from docx import Document
        from docx.text.paragraph import Paragraph
        view = read_word_text(path, offset=offset, limit=20, max_chars=10000)
        positions, index = {}, 0
        for number, block in enumerate(Document(path).iter_inner_content()):
            if isinstance(block, Paragraph):
                positions[str(number)] = index
                index += 1
        return {'view': view, 'positions': positions}
    if path.suffix.lower() != '.xlsx':
        raise ValueError('unsupported_format')
    view = inspect_xlsx(path)
    remaining = 400
    for sheet in view['sheets']:
        cells = sheet['cells'][:min(100, remaining)]
        sheet['cells_truncated'] = sheet.get('cells_truncated', False) or len(cells) < len(sheet['cells'])
        sheet['cells'] = cells
        remaining -= len(cells)
    return {'view': view, 'positions': {}}


def read_local_view(path, cancelled, *, expected_version):
    path = Path(path)
    if path.suffix.lower() not in OPEN_FORMATS or path.is_symlink() or path.stat().st_size > 64 * 1024 * 1024:
        raise ValueError('不支持的本地文档或文件超出64 MiB限制')
    with tempfile.TemporaryDirectory(prefix='office-open-') as raw:
        scratch = Path(raw)
        source, target = scratch / ('input' + path.suffix.lower()), scratch / 'view.json'
        with path.open('rb') as stream:
            data = stream.read(64 * 1024 * 1024 + 1)
        if len(data) > 64 * 1024 * 1024:
            raise ValueError('文件超出64 MiB限制')
        if hashlib.sha256(data).hexdigest() != expected_version:
            raise ValueError('文件已改变，旧读取授权失效；请重新拖入当前版本')
        source.write_bytes(data)
        command = (sys.executable, str(Path(__file__).with_name('document_open_worker.py')), str(source), str(target))
        if os.name == 'nt':
            from ..diagram_process import _run_owned
            run = _run_owned(command, timeout_ms=30000, cancelled=cancelled)
            process = asdict(run)
            if run.code != 'process_completed' or run.exit_code != 0:
                raise ValueError('解析已取消' if run.code == 'render_cancelled' else '解析超时或失败；可重试打开此文件')
        else:
            started = time.monotonic()
            result = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
            process = {'exit_code': result.returncode, 'elapsed_seconds': time.monotonic() - started}
        if cancelled():
            raise ValueError('解析已取消')
        if not target.is_file() or target.stat().st_size > 1024 * 1024:
            raise ValueError('解析结果超过有界视图限制，未启动模型')
        report = json.loads(target.read_text(encoding='utf-8'))
        if report.get('error'):
            raise ValueError('文档正文无法解析；请检查文件是否损坏、加密或包含不支持的结构')
        return {**report, 'process': process}
