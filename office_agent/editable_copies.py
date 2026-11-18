from copy import deepcopy
import io
import time
import uuid

from docx import Document

from .paged_documents import CandidatePublisher, GuardError, W, eligibility, parts, sha, xml
from .edit_metrics import publication_metrics


def body_blocks(data):
    values = parts(data)
    settings = xml(values['word/settings.xml']) if 'word/settings.xml' in values else None
    if settings is not None and any(settings.find('.//' + W + name) is not None for name in ('documentProtection', 'writeProtection')):
        raise GuardError('文档含编辑保护；请先在原编辑软件中取得授权并解除保护，再重新打开')
    body = xml(values['word/document.xml']).find(W + 'body')
    blocked = {'drawing', 'pict', 'object', 'altChunk', 'sdt', 'ins', 'del', 'moveFrom', 'moveTo',
               'footnoteReference', 'endnoteReference', 'oMath', 'oMathPara', 'fldSimple', 'instrText',
               'sym', 'ruby', 'contentPart', 'subDoc', 'ptab', 'noBreakHyphen', 'softHyphen',
               'customXml', 'vanish', 'webHidden', 'fldChar', 'ffData'}
    if body is None or any(node.tag.split('}')[-1] in blocked for node in body.iter() if isinstance(node.tag, str)):
        raise GuardError('此文档含图片、公式、域、脚注、修订或其他复杂内容，暂不支持安全创建简化副本；请在原编辑软件中另存')
    paragraphs = list(body.iter(W + 'p'))
    if len(paragraphs) > 2000:
        raise GuardError('简化副本最多处理2000段，请先拆分材料')

    def text(node):
        return ''.join(item.text or '' if item.tag == W + 't' else '\t' if item.tag == W + 'tab' else '\n'
                       for run in node.iter(W + 'r') for item in run if item.tag in {W + 't', W + 'tab', W + 'br', W + 'cr'})

    blocks, cells = [], 0
    for node in body:
        if node.tag == W + 'p':
            blocks.append({'kind': 'paragraph', 'text': text(node)})
        elif node.tag == W + 'tbl':
            if (node.find('.//' + W + 'vMerge') is not None or node.find('.//' + W + 'gridSpan') is not None
                    or any(cell.find('.//' + W + 'tbl') is not None for cell in node.iter(W + 'tc'))):
                raise GuardError('合并或嵌套表格暂不支持简化副本，请在原编辑软件中另存')
            rows = []
            width = None
            for row in node.findall(W + 'tr'):
                entries = [[text(p) for p in cell.findall(W + 'p')] for cell in row.findall(W + 'tc')]
                width = len(entries) if width is None else width
                cells += len(entries)
                if not entries or len(entries) != width or any(not entry for entry in entries) or cells > 400:
                    raise GuardError('表格结构或数量超出简化副本范围（最多400格普通表格）')
                rows.append(entries)
            if not rows:
                raise GuardError('空表格结构暂不支持简化副本')
            blocks.append({'kind': 'table', 'rows': rows})
        elif node.tag != W + 'sectPr':
            raise GuardError('正文包含未支持结构，不能省略后声称已创建完整文字副本')
    if sum(len(text(p)) for p in paragraphs) > 400000:
        raise GuardError('正文超过简化副本的40万字符预算')
    return blocks


class EditableCopyPublisher(CandidatePublisher):
    @publication_metrics
    def plan_copy(self, *, document_id, draft_revision, operation_id=None):
        self.phase('preparing', '正在整理正文与普通表格，准备可编辑副本。')
        self.check_source()
        original = self.source.read_bytes()
        if sha(original) != self.source_sha256:
            raise GuardError('原文件在读取期间已改变，请重新读取后创建副本')
        blocks = body_blocks(original)
        book = Document()
        for block in blocks:
            if block['kind'] == 'paragraph':
                book.add_paragraph(block['text'])
            else:
                table = book.add_table(rows=len(block['rows']), cols=len(block['rows'][0]))
                for row, entries in zip(table.rows, block['rows']):
                    for cell, texts in zip(row.cells, entries):
                        if texts[0]:
                            cell.paragraphs[0].text = texts[0]
                        for text in texts[1:]:
                            cell.add_paragraph(text)
        if not blocks:
            book.add_paragraph('')
            blocks = [{'kind': 'paragraph', 'text': ''}]
        stream = io.BytesIO()
        self.phase('serializing', '正在生成工作副本，尚未发布文件。')
        book.save(stream)
        candidate = stream.getvalue()
        if not eligibility(candidate)['editable'] or body_blocks(candidate) != blocks:
            raise GuardError('简化副本的正文或结构重开校验未通过，未生成文件')
        identifier = operation_id or uuid.uuid4().hex
        output = self.editor.space.file(str(self.source.with_name(f'{self.source.stem}.agent-{identifier[:8]}.docx')), existing=False)
        self.plans.clear()
        self.plans[identifier] = {
            'id': identifier, 'kind': 'paged_editable_copy', 'source': str(self.source),
            'source_sha256': self.source_sha256, 'document_id': document_id,
            'draft_revision': draft_revision, 'candidate_sha256': sha(candidate),
            'output': str(output), 'mode': 'body_text', 'normalization_parts': [],
            'changes': [], 'retained': '正文和普通表格文字，包括超链接的显示文字；不改写数字、单位或引文文字',
            'limitations': ['原样式、自动编号、颜色和分页将重建，不保证原版式。',
                            '不带入页眉页脚、批注、书签、链接地址及自定义元数据。',
                            '副本打开后仍须通过编辑器的无修改往返检查；未通过则继续只读。'],
            'summary': {'paragraphs': len(list(book.element.body.iter(W + 'p'))), 'tables': len(book.tables),
                        'characters': sum(len(p.text or '') for p in book.element.body.iter(W + 't'))},
            'normalization': 'User-approved body-text copy. Body paragraph and ordinary table text preserved exactly; original styles, automatic numbering, page layout, headers, footers, comments, bookmarks, link destinations and custom metadata omitted.',
            'deadline': time.monotonic() + 300, 'candidate': candidate,
        }
        return deepcopy(self.public(identifier))
