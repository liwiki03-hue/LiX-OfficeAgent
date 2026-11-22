import hashlib
import os
from pathlib import Path
from types import SimpleNamespace

from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
import pytest

from office_agent.webui import turn_attachment as attachment


def _picker(*paths):
    def pick(_dialog, *, allow_multiple, file_types):
        assert not allow_multiple
        return tuple(map(str, paths))

    return SimpleNamespace(authenticated=lambda: True, app=SimpleNamespace(window=SimpleNamespace(create_file_dialog=pick)))


def _document(path, marker):
    if path.suffix == '.docx':
        book = Document()
        book.add_paragraph(marker)
    elif path.suffix == '.xlsx':
        book = Workbook()
        book.active['B3'] = marker
    else:
        book = Presentation()
        slide = book.slides.add_slide(book.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text = marker
    book.save(path)


@pytest.mark.parametrize('suffix, location', [('.docx', 'body:1/paragraph'), ('.xlsx', 'B3'), ('.pptx', '幻灯片 1')])
def test_only_read_the_picked_document_once_approved(tmp_path, monkeypatch, suffix, location):
    path = tmp_path / ('参照样本' + suffix)
    sibling = tmp_path / ('相邻文件' + suffix)
    marker = '独有指标 37 ms'
    _document(path, marker)
    _document(sibling, '相邻内容不可见')
    fingerprint = hashlib.sha256(path.read_bytes()).hexdigest()
    reads = []
    open_file = Path.open
    scan = os.scandir

    def open_selected(file, *args, **kwargs):
        if file.parent == tmp_path:
            assert file == path and 'w' not in str(args)
            reads.append(file.name)
        return open_file(file, *args, **kwargs)

    def scan_directory(directory):
        assert Path(directory) != tmp_path
        return scan(directory)

    monkeypatch.setattr(Path, 'open', open_selected)
    monkeypatch.setattr(os, 'scandir', scan_directory)
    assert attachment.choose(_picker()) is None
    with pytest.raises(ValueError, match='只能选择一个'):
        attachment.choose(_picker(path, sibling))
    selected = attachment.choose(_picker(path))
    assert not reads and selected.name == path.name
    receipt = attachment.read_selected(selected)
    assert reads and set(reads) == {path.name}
    assert receipt['name'] == path.name and receipt['sha256'] == fingerprint
    assert location in receipt['material'] and marker in receipt['material']
    assert '相邻内容不可见' not in str(receipt) and str(tmp_path) not in str(receipt)
    assert not receipt['truncated']
    assert hashlib.sha256(path.read_bytes()).hexdigest() == fingerprint


def test_refuse_stale_binding_protected_sheet_and_client_path(tmp_path):
    path = tmp_path / '样本.xlsx'
    _document(path, '初始版本')
    selected = attachment.choose(_picker(path))
    replacement = tmp_path / '替换.xlsx'
    _document(replacement, '已替换版本')
    os.replace(replacement, path)
    with pytest.raises(ValueError, match='已改变'):
        attachment.read_selected(selected)
    with pytest.raises(ValueError, match='文件选择器'):
        attachment.read_selected(path)
    workbook = Workbook()
    workbook.active['A1'] = '受保护内容'
    workbook.active.protection.enable()
    workbook.save(path)
    with pytest.raises(ValueError, match='受保护'):
        attachment.read_selected(attachment.choose(_picker(path)))
    alias = tmp_path / '硬链接.xlsx'
    os.link(path, alias)
    with pytest.raises(ValueError, match='普通本地文件'):
        attachment.choose(_picker(alias))


def test_attachment_bounds_the_excerpt_and_reports_coverage(tmp_path):
    path = tmp_path / '较长材料.docx'
    book = Document()
    book.add_paragraph('文字' * 2100)
    book.add_paragraph('补充' * 2100)
    book.add_paragraph('末尾标记')
    book.save(path)
    receipt = attachment.read_selected(attachment.choose(_picker(path)))
    assert receipt['truncated'] and len(receipt['material']) == attachment.MAX_MATERIAL
    assert '末尾标记' not in receipt['material']
