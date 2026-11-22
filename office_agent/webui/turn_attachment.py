from dataclasses import dataclass, field
from pathlib import Path

from ..documents import classify_document
from ..presentations import _items, _load
from ..rag_sources import MAX_SOURCE_BYTES, SourceError, SourceExtractor, _identity, _ordinary
from ..writes import GuardError, Workspace
from .document_open import OPEN_FORMATS


MAX_MATERIAL = 8000


@dataclass(frozen=True)
class SelectedFile:
    path: Path = field(repr=False)
    stamp: tuple = field(repr=False)

    @property
    def name(self):
        return self.path.name


def _selected(path):
    if (not path.is_absolute() or path.as_posix().startswith('//') or '..' in path.parts or
            any(':' in part for part in path.parts[1:]) or path.name.lower().startswith('.env') or
            path.suffix.lower() not in OPEN_FORMATS):
        raise ValueError('请选择本机的 DOCX、XLSX 或 PPTX 文件。')
    try:
        for parent in reversed(path.parents):
            _ordinary(parent.lstat(), directory=True)
        info = path.lstat()
        _ordinary(info)
        if info.st_size > MAX_SOURCE_BYTES:
            raise ValueError('本次参考文件超过 32 MiB 的只读提取上限。')
        admitted = Workspace(path.parent).file(str(path))
    except (SourceError, OSError, GuardError):
        raise ValueError('请选择可读取的普通本地文件，无法使用目录、链接或已移走的文件。') from None
    return SelectedFile(admitted, _identity(info))


def choose(floating):
    from webview import FileDialog

    floating.authenticated()
    window = floating.app.window
    if window is None:
        raise ValueError('文件选择器不可用。')
    result = window.create_file_dialog(FileDialog.OPEN, allow_multiple=False,
                                       file_types=('参考文档 (*.docx;*.xlsx;*.pptx)',))
    if not result:
        return None
    if not isinstance(result, (list, tuple)) or len(result) != 1 or not isinstance(result[0], str):
        raise ValueError('本轮只能选择一个参考文件。')
    return _selected(Path(result[0]))


def recheck(selected):
    try:
        if isinstance(selected, SelectedFile) and _selected(selected.path) == selected:
            return selected
    except ValueError:
        pass
    raise ValueError('参考文件已改变或不可用，请重新选择。')


def _slide_blocks(path, expected):
    parts, slides, fingerprint = _load(path)
    if fingerprint != expected:
        raise SourceError('source_changed')
    for index, slide in enumerate(slides, 1):
        items, _, _ = _items(parts, slide)
        for item in items:
            if item['text'] and item['type'] != 'notes':
                yield {'location': f'幻灯片 {index}', 'text': item['text']}


def read_selected(selected):
    if not isinstance(selected, SelectedFile):
        raise ValueError('请通过文件选择器允许读取一个参考文件。')
    path = selected.path
    try:
        if _selected(path) != selected:
            raise SourceError('source_changed')
        report = classify_document(path)
        if (not report.readable or report.kind != OPEN_FORMATS[path.suffix.lower()]['kind'] or
                report.protected or report.macro_enabled or
                set(report.suspicious_features) & {'activex', 'embedded_object', 'dde_marker', 'script_marker',
                                                   'unparseable_protection_xml', 'large_protection_part_uninspected'}):
            raise GuardError('参考文件格式不匹配、受保护或含活动内容，未用于本轮回答。')
        reader = SourceExtractor(Workspace(path.parent))
        data, stamp, fingerprint = reader._snapshot(path.name)
        if stamp != selected.stamp:
            raise SourceError('source_changed')
        blocks = (_slide_blocks(path, fingerprint) if path.suffix.lower() == '.pptx'
                  else reader._parse(data, path.suffix.lower(), path.name))
        lines, sources, length, truncated = [], [], 0, False
        for block in blocks:
            line = block['location'] + ' · ' + block['text']
            remaining = MAX_MATERIAL - length - bool(lines)
            if remaining > 0:
                lines.append(line[:remaining])
                location = block['location']
                if remaining > len(location) + 3:
                    sources.append({'location': location, 'text': block['text'][:remaining - len(location) - 3]})
                length += len(lines[-1]) + (len(lines) > 1)
            if remaining < len(line):
                truncated = True
                break
        if _selected(path) != selected:
            raise SourceError('source_changed')
        if not lines:
            raise GuardError('所选文件没有可供本次参考的正文、表格值或幻灯片文字。')
        return {'name': selected.name, 'sha256': fingerprint, 'material': '\n'.join(lines), 'truncated': truncated, 'sources': sources}
    except SourceError as error:
        if error.code == 'source_changed':
            raise ValueError('参考文件已改变，请重新选择并允许读取。') from None
        raise ValueError('该参考文件含不支持的结构或超过只读提取上限。') from None
    except GuardError:
        raise
    except (OSError, ValueError, KeyError, IndexError, SyntaxError, TypeError, RuntimeError):
        raise ValueError('参考文件无法读取或已损坏，请重新选择。') from None
