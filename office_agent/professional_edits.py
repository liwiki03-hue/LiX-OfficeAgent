from difflib import SequenceMatcher
import hashlib
import json
import re
from typing import Literal
import uuid
from pathlib import Path

from lxml import etree as ET
from pydantic import Field

from .provenance import SourceDocument, EvidenceFragment, identifier, acquired_at
from .rag import SECRET
from .writes import Strict, Plan, GuardError, digest

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML = '{http://www.w3.org/XML/1998/namespace}'
PARSER_VERSION = 'docx-structural-edit-v1'
MAX_TARGETS = 4000
NUMBER = re.compile(r'(?<![\d.])[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?(?:\s?(?:%|ms|s|kg|mg|g|cm|mm|m|K|°C|MPa|Hz|kHz|MHz|GB|MB))?')
CITATION = re.compile(r'\[[^\[\]\n]{1,120}\]|\([^()\n]{0,100}(?:19|20)\d{2}[a-z]?[^()\n]{0,40}\)|10\.\d{4,9}/[^\s<>]+')
FORMULA = re.compile(r'\$[^$\n]+\$|\\\([^\n]*?\\\)')


def parser_identity():
    return {'version': PARSER_VERSION, 'source_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}


class LocalChange(Strict):
    locator_id: str = Field(pattern=r'^loc:[0-9a-f]{64}$')
    expected_text_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    new_text: str = Field(min_length=1, max_length=8000)
    reason: str = Field(min_length=1, max_length=400)
    evidence_ids: list[str] = Field(default_factory=list, max_length=12)


class ConstrainedDocx(Strict):
    kind: Literal['constrained_docx'] = 'constrained_docx'
    path: str
    source_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    changes: list[LocalChange] = Field(min_length=1, max_length=12)
    preserve_numbers: bool = True
    preserve_citations: bool = True
    preserve_formulas: bool = True
    protected_literals: list[str] = Field(default_factory=list, max_length=30)
    format_policy: Literal['preserve_runs'] = 'preserve_runs'


def _hash_text(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


def _canonical(node):
    return ET.tostring(node, method='c14n')


def _text(node):
    return ''.join(value.text or '' for value in node.iter(W + 't'))


def _support(node, continued=False):
    if continued:
        return 'merged_cell_continuation'
    if not list(node.iter(W + 't')):
        return 'empty_or_nontext_paragraph'
    if len(_text(node)) > 8000:
        return 'paragraph_capacity'
    if any(child.tag not in {W + 'pPr', W + 'r'} for child in node):
        return 'hyperlink_bookmark_field_or_revision'
    if any(child.tag not in {W + 'rPr', W + 't'} for run in node.findall(W + 'r') for child in run):
        return 'field_image_note_break_or_special_run'
    if any(n.tag in {W + 'rPrChange', W + 'pPrChange'} for n in node.iter()):
        return 'format_revision'
    return None


def _inventory(parts, source_sha256):
    root = ET.fromstring(parts['word/document.xml'])
    body = root.find(W + 'body')
    if body is None:
        raise GuardError('professional_invalid_document')
    entries = []
    headings = []
    def add(node, location, continued=False):
        text = _text(node)
        style = node.find('./' + W + 'pPr/' + W + 'pStyle')
        name = style.get(W + 'val') if style is not None else None
        if name and re.fullmatch(r'Heading[1-9]', name, re.I):
            level = int(name[-1])
            del headings[level - 1:]
            headings.append('[redacted]' if SECRET.search(text) else text[:300])
        loc = 'loc:' + identifier(source_sha256, location, _canonical(node).hex())
        entries.append({'locator_id': loc, 'location': location, 'text': text,
                        'text_sha256': _hash_text(text), 'style': name, 'heading_path': tuple(headings),
                        'writable': _support(node, continued) is None,
                        'refusal': _support(node, continued), 'node': node})
        if len(entries) > MAX_TARGETS:
            raise GuardError('professional_capacity')
    for position, node in enumerate(body):
        if node.tag == W + 'p':
            add(node, f'body:{position}/paragraph')
        elif node.tag == W + 'tbl':
            for row_index, row in enumerate(node.findall(W + 'tr')):
                for cell_index, cell in enumerate(row.findall(W + 'tc')):
                    merge = cell.find('./' + W + 'tcPr/' + W + 'vMerge')
                    continued = merge is not None and merge.get(W + 'val') != 'restart'
                    for paragraph_index, paragraph in enumerate(cell.findall(W + 'p')):
                        add(paragraph, f'body:{position}/table/row:{row_index}/physical-cell:{cell_index}/paragraph:{paragraph_index}', continued)
    return root, entries


def snapshot(editor, path, *, offset=0, limit=20, _range_edit=False):
    from .advanced_writes import _parts
    if type(offset) is not int or not 0 <= offset <= MAX_TARGETS or type(limit) is not int or not 1 <= limit <= 40:
        raise GuardError('professional_invalid_page')
    source = editor.space.file(str(path))
    if source.suffix.lower() != '.docx':
        raise GuardError('professional_docx_required')
    version = digest(source)
    parts = _parts(source, range_edit=_range_edit)
    _, entries = _inventory(parts, version)
    document = SourceDocument(id=identifier(source.name.casefold()), source_name=source.name,
                              source_sha256=version, parser_version=PARSER_VERSION,
                              scope_id=identifier(str(editor.space.root)), retrieved_at=acquired_at())
    selected = entries[offset:offset + limit]
    records = []
    budget = 0
    for entry in selected:
        if len(entry['text']) > 8000 or SECRET.search(entry['text']):
            records.append({key: value for key, value in entry.items() if key not in {'node', 'text', 'heading_path'}} |
                           {'text': '', 'writable': False, 'refusal': 'sensitive_or_oversize', 'evidence': None})
            continue
        budget += len(entry['text'])
        if budget > 24000:
            break
        fragment = EvidenceFragment(id='ev:' + identifier(version, entry['location'], entry['text']),
            document_id=document.id, source_sha256=version, location=entry['location'],
            heading_path=entry['heading_path'], original_text=entry['text'],
            writable_locator_id=entry['locator_id'] if entry['writable'] else None)
        records.append({key: value for key, value in entry.items() if key not in {'node', 'heading_path'}} |
                       {'evidence': fragment.model_dump(mode='json')})
    if editor.space.file(str(source)) != source or digest(source) != version:
        raise GuardError('professional_source_changed')
    next_offset = offset + len(records)
    return {'document': document.model_dump(mode='json'), 'items': records, 'total_targets': len(entries),
            'next_offset': next_offset if next_offset < len(entries) else None,
            'read_only': True, 'page_numbers_available': False,
            'limitations': ['Structural locations are not rendered page numbers', 'Evidence existence does not establish claim support',
                            'Unsupported paragraph structures are read-only; table coordinates are physical OOXML cells']}


def _protect(before, after, request):
    from .technical_profile import verify
    verify(before, after, getattr(request, 'editing_profile', 'general'), getattr(request, 'profile_rule_sha256', None))
    for enabled, pattern, code in [(request.preserve_numbers, NUMBER, 'numbers'),
                                   (request.preserve_citations, CITATION, 'citations'),
                                   (request.preserve_formulas, FORMULA, 'formulas')]:
        if enabled and pattern.findall(before) != pattern.findall(after):
            raise GuardError('professional_protected_' + code)
    for literal in request.protected_literals:
        if not literal or len(literal) > 200 or before.count(literal) != after.count(literal):
            raise GuardError('professional_protected_literal')
    if SECRET.search(before) or SECRET.search(after) or any(ord(c) < 32 for c in after):
        raise GuardError('professional_sensitive_or_control_text')


def _rewrite(node, before, after):
    texts = list(node.iter(W + 't'))
    positions, cursor = [], 0
    for text in texts:
        value = text.text or ''
        positions.append((cursor, cursor + len(value), text))
        cursor += len(value)
    for action, start, end, new_start, new_end in reversed(SequenceMatcher(None, before, after, autojunk=False).get_opcodes()):
        if action == 'equal':
            continue
        replacement = after[new_start:new_end]
        first = next((i for i, (left, right, _) in enumerate(positions) if left <= start < right), len(positions) - 1)
        if start == len(before):
            first = len(positions) - 1
        left, right, text = positions[first]
        if start == end:
            value = text.text or ''
            point = start - left
            text.text = value[:point] + replacement + value[point:]
            continue
        last = next(i for i in range(len(positions) - 1, -1, -1) if positions[i][0] < end and positions[i][1] >= end)
        if first == last:
            value = text.text or ''
            text.text = value[:start - left] + replacement + value[end - left:]
        else:
            text.text = (text.text or '')[:start - left] + replacement
            for index in range(first + 1, last):
                positions[index][2].text = ''
            tail_start, _, tail = positions[last]
            tail.text = (tail.text or '')[end - tail_start:]
    for text in texts:
        value = text.text or ''
        if value.startswith(' ') or value.endswith(' '):
            text.set(XML + 'space', 'preserve')
    if _text(node) != after:
        raise GuardError('professional_run_mapping_failed')


def transform(parts, request):
    if request.kind == 'block_docx':
        from .block_edits import transform as block_transform
        return block_transform(parts, request)
    if request.kind == 'range_docx':
        from .range_edits import transform as range_transform
        return range_transform(parts, request)
    if request.kind == 'structured_docx':
        from .structured_edits import transform as structured_transform
        return structured_transform(parts, request)
    if request.kind == 'tracked_assisted_docx':
        from .tracked_edits import transform as tracked_transform
        return tracked_transform(parts, request)
    if request.kind == 'assisted_docx':
        from .assisted_edits import transform as assisted_transform
        return assisted_transform(parts, request)
    if request.kind == 'citation_docx':
        from .citation_lifecycle import transform as citation_transform
        return citation_transform(parts, request)
    if request.kind == 'grounded_docx':
        from .grounded_edits import transform as grounded_transform
        return grounded_transform(parts, request)
    root, entries = _inventory(parts, request.source_sha256)
    targets = {entry['locator_id']: entry for entry in entries}
    evidence = {'ev:' + identifier(request.source_sha256, entry['location'], entry['text']) for entry in entries if not SECRET.search(entry['text'])}
    seen = set()
    text_budget = 0
    before_items, after_items = [], []
    for change in request.changes:
        if SECRET.search(change.reason) or not change.reason.strip():
            raise GuardError('professional_invalid_reason')
        if change.locator_id in seen:
            raise GuardError('professional_overlapping_changes')
        seen.add(change.locator_id)
        entry = targets.get(change.locator_id)
        if entry is None or entry['text_sha256'] != change.expected_text_sha256:
            raise GuardError('professional_stale_locator')
        if not entry['writable']:
            raise GuardError('professional_unsupported_target:' + entry['refusal'])
        text_budget += len(entry['text']) + len(change.new_text)
        if text_budget > 16000:
            raise GuardError('professional_capacity')
        if entry['text'] == change.new_text:
            raise GuardError('professional_no_changes')
        if any(value not in evidence for value in change.evidence_ids):
            raise GuardError('professional_evidence_missing')
        _protect(entry['text'], change.new_text, request)
        before_items.append({'location': entry['location'], 'text': entry['text']})
        _rewrite(entry['node'], entry['text'], change.new_text)
        after_items.append({'location': entry['location'], 'text': change.new_text, 'reason': change.reason,
                            'evidence_ids': change.evidence_ids, 'claim_support': 'not_evaluated',
                            'format_policy': request.format_policy})
    changed = dict(parts)
    changed['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return changed, before_items, after_items


def make_plan(editor, request):
    from .advanced_writes import _parts
    from .documents import inspect_file
    source = editor.space.file(request.path)
    if source.suffix.lower() != '.docx' or digest(source) != request.source_sha256:
        raise GuardError('professional_source_changed')
    inspection = inspect_file(source)
    external_refused = bool(inspection.get('external_links'))
    if external_refused and request.kind in {'range_docx', 'block_docx'}:
        from .range_edits import validate_external_relationships
        validate_external_relationships(_parts(source, range_edit=True))
        external_refused = False
    if inspection.get('protected') or inspection.get('macros') or external_refused:
        raise GuardError('professional_unsafe_document')
    _, before, after = transform(_parts(source, range_edit=request.kind in {'range_docx', 'block_docx'}), request)
    if digest(source) != request.source_sha256:
        raise GuardError('professional_source_changed')
    warning = '独立确认整组修改；保护检查不证明语义等价；新增文字沿用首个受影响 run 的格式；非原生修订'
    if request.kind == 'tracked_assisted_docx':
        warning = '原生 Word 插入/删除修订；全部接受对应已批准净稿，全部拒绝对应原稿；现有修订拒绝；Word 界面效果仍须实机核对'
    if request.kind == 'structured_docx':
        warning = '普通段落净稿插入与审阅批注；新增段落没有原生结构修订；整组引用影响须独立确认；分页可变化'
    if request.kind == 'range_docx':
        warning = '版本绑定局部修改；混合格式须明确继承策略；非目标部件逐字节保持；净稿和文字修订分别验证；Word人工未验证'
    if request.kind == 'block_docx':
        warning = '普通段落块净稿操作；明确边界/样式及完整章节；不提供结构原生修订；书签目标调整后域结果保留且可能过时，须人工核对'
    plan = Plan(uuid.uuid4().hex, request, request.source_sha256, 'version-bound structural changes',
                json.dumps({'items': before}, ensure_ascii=False), json.dumps({'items': after}, ensure_ascii=False),
                (warning,), True)
    editor.plans[plan.id] = plan
    return plan


def verify(before, after, request):
    if request.kind == 'block_docx':
        from .block_edits import verify as block_verify
        return block_verify(before, after, request)
    if request.kind == 'range_docx':
        from .range_edits import verify as range_verify
        return range_verify(before, after, request)
    if request.kind == 'structured_docx':
        from .structured_edits import verify as structured_verify
        return structured_verify(before, after, request)
    expected, _, _ = transform(before, request)
    if set(after) != set(before) or any(value != after[name] for name, value in before.items() if name != 'word/document.xml'):
        raise GuardError('professional_non_target_part_changed')
    if _canonical(ET.fromstring(expected['word/document.xml'])) != _canonical(ET.fromstring(after['word/document.xml'])):
        raise GuardError('professional_non_target_structure_changed')
