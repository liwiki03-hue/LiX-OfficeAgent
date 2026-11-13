"""Restricted diagram display primitives; no tools, authority, I/O or model calls.

This deliberately supports a small Mermaid subset, not arbitrary Mermaid. SVG
must be flattened by the isolated renderer first (styles become safe attributes).
Passing validation is NOT a rendering or business-semantics proof.
"""
from dataclasses import dataclass
import hashlib
import math
import re
from xml.etree import ElementTree as ET

KINDS = frozenset({'classDiagram', 'flowchart', 'sequenceDiagram', 'stateDiagram-v2', 'erDiagram'})
MAX_REPAIRS = 2
SVG_NS = 'http://www.w3.org/2000/svg'
_ID = r'[A-Za-z][A-Za-z0-9_]{0,31}'
_LABEL = r'[A-Za-z0-9 _-]{1,60}'
_NODE = '(' + _ID + r')(?:\[' + _LABEL + r'\]|\{' + _LABEL + r'\}|\(' + _LABEL + r'\))?'
_NUMBER = r'[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?'


class DiagramError(ValueError):
    def __init__(self, code: str, message: str = '图表超出受限安全范围，请简化后重试'):
        self.code = code
        super().__init__(message)


def _require(condition, code='unsupported_source'):
    if not condition:
        raise DiagramError(code)


@dataclass(frozen=True)
class DiagramSource:
    kind: str
    source: str
    nodes: tuple[str, ...]
    edges: int
    sha256: str

    def evidence(self) -> dict:
        # Only metadata belongs in model evidence. Source/SVG are display-only.
        return {'kind': self.kind, 'nodes': len(self.nodes), 'edges': self.edges,
                'source_sha256': self.sha256, 'rendered': False}


def validate_source(source: str, expected_kind: str | None = None) -> DiagramSource:
    _require(isinstance(source, str) and 0 < len(source) <= 8192)
    _require(all(c == '\n' or c == '\r' or c == '\t' or 32 <= ord(c) <= 126 for c in source))
    normalized = '\n'.join(line.rstrip() for line in source.splitlines()).strip()
    lines = [line.strip() for line in normalized.splitlines() if line.strip()]
    _require(1 < len(lines) <= 128)
    header = lines.pop(0)
    kind = 'flowchart' if re.fullmatch(r'flowchart (?:TD|TB|BT|LR|RL)', header) else header
    _require(kind in KINDS and (expected_kind is None or expected_kind == kind))
    nodes, participants, edges, block, members = set(), set(), 0, None, 0
    for line in lines:
        matched = False
        if block:
            if line == '}':
                block = None
            else:
                member = (r'[+~-]?(?:' + _ID + r'(?:\[\])?\s+' + _ID + '|' + _ID +
                          r'\([A-Za-z0-9_ ,]{0,80}\)(?:\s+' + _ID + r'(?:\[\])?)?)')
                if kind == 'erDiagram':
                    member = _ID + r'\s+' + _ID + r'(?:\s+(?:PK|FK|UK))?'
                _require(bool(re.fullmatch(member, line)))
                members += 1
                _require(members <= 12)
            continue
        if kind == 'flowchart':
            edge = re.fullmatch(_NODE + r'\s*(-->|---|-\.->|==>)(?:\|' + _LABEL + r'\|)?\s*' + _NODE, line)
            node = re.fullmatch(_NODE, line)
            if edge:
                nodes.update((edge[1], edge[3])); edges += 1; matched = True
            elif node:
                nodes.add(node[1]); matched = True
        elif kind == 'classDiagram':
            declared = re.fullmatch(r'class (' + _ID + r')(?:\s*(\{))?', line)
            edge = re.fullmatch('(' + _ID + r')(?:\s+"(?:\d+|\*|[01]\.\.[0-9*])")?\s+' +
                r'(<\|--|--\|>|\*--|--\*|o--|--o|-->|<--|--|\.\.>|<\.\.|\.\.|<\|\.\.|\.\.\|>)\s+' +
                r'(?:"(?:\d+|\*|[01]\.\.[0-9*])"\s+)?(' + _ID + r')(?:\s*:\s*' + _LABEL + ')?', line)
            if declared:
                nodes.add(declared[1]); block = declared[1] if declared[2] else None
                members = 0; matched = True
            elif edge:
                nodes.update((edge[1], edge[3])); edges += 1; matched = True
        elif kind == 'sequenceDiagram':
            declared = re.fullmatch(r'participant (' + _ID + r')(?: as ' + _LABEL + ')?', line)
            edge = re.fullmatch('(' + _ID + r')(->>|-->>|->|-->|-x|--x|-\)|--\))(' + _ID + r'):\s*' + _LABEL, line)
            if declared:
                _require(declared[1] not in participants)
                participants.add(declared[1]); nodes.add(declared[1]); matched = True
            elif edge:
                _require(edge[1] in participants and edge[3] in participants)
                edges += 1; matched = True
        elif kind == 'stateDiagram-v2':
            sid = r'(\[\*\]|' + _ID + ')'
            edge = re.fullmatch(sid + r'\s*-->\s*' + sid + r'(?:\s*:\s*' + _LABEL + ')?', line)
            declared = re.fullmatch(r'state "' + _LABEL + '" as (' + _ID + ')', line)
            if edge:
                nodes.update(n for n in (edge[1], edge[2]) if n != '[*]')
                edges += 1; matched = True
            elif declared:
                nodes.add(declared[1]); matched = True
        else:
            declared = re.fullmatch('(' + _ID + r')\s*\{', line)
            edge = re.fullmatch('(' + _ID + r')\s+(\|\||o\||\|o|\}o|\}\|)(--|\.\.)(\|\||o\||\|o|o\{|\|\{)\s+(' + _ID + r')\s*:\s*' + _LABEL, line)
            if declared:
                nodes.add(declared[1]); block = declared[1]; members = 0; matched = True
            elif edge:
                nodes.update((edge[1], edge[5])); edges += 1; matched = True
        _require(matched and len(nodes) <= 12 and edges <= 36)
    _require(block is None and bool(nodes))
    return DiagramSource(kind, normalized, tuple(sorted(nodes)), edges,
                         hashlib.sha256(normalized.encode('utf-8')).hexdigest())


_ELEMENTS = frozenset({'svg', 'g', 'path', 'rect', 'circle', 'ellipse', 'line', 'polyline',
                       'polygon', 'text', 'tspan', 'defs', 'marker'})
_NUMERIC = frozenset({'x', 'y', 'x1', 'y1', 'x2', 'y2', 'cx', 'cy', 'r', 'rx', 'ry', 'width',
                     'height', 'refX', 'refY', 'markerWidth', 'markerHeight', 'stroke-width',
                     'font-size', 'opacity', 'fill-opacity', 'stroke-opacity', 'dx', 'dy', 'textLength'})
_FIXED = {'font-family': {'Arial', 'sans-serif', 'serif', 'monospace'},
          'text-anchor': {'start', 'middle', 'end'}, 'dominant-baseline': {'auto', 'middle', 'central', 'hanging'},
          'alignment-baseline': {'auto', 'middle', 'central', 'baseline'},
          'font-weight': {'normal', 'bold', '400', '600', '700'}, 'font-style': {'normal', 'italic'},
          'stroke-linecap': {'butt', 'round', 'square'}, 'stroke-linejoin': {'miter', 'round', 'bevel'},
          'fill-rule': {'nonzero', 'evenodd'}, 'markerUnits': {'strokeWidth', 'userSpaceOnUse'},
          'orient': {'auto', 'auto-start-reverse'}, 'preserveAspectRatio': {'xMidYMid meet', 'xMinYMin meet', 'none'},
          'role': {'img', 'graphics-document', 'graphics-symbol'}, 'aria-hidden': {'true', 'false'}}


def _numbers(value: str):
    values = [float(n) for n in re.findall(_NUMBER, value)]
    _require(all(math.isfinite(n) and abs(n) <= 10000 for n in values), 'unsafe_svg')
    return values


def _color(value):
    if value in {'none', 'currentColor', 'transparent', 'black', 'white', 'red', 'blue', 'gray', 'grey'}:
        return True
    if re.fullmatch(r'#[0-9a-fA-F]{3}(?:[0-9a-fA-F]{3})?', value):
        return True
    rgb = re.fullmatch(r'rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)', value)
    return bool(rgb and all(int(n) <= 255 for n in rgb.groups()))


def validate_svg(svg: str) -> str:
    code = 'unsafe_svg'
    _require(isinstance(svg, str), code)
    try:
        size = len(svg.encode('utf-8'))
    except UnicodeError:
        raise DiagramError(code) from None
    _require(0 < size <= 1_048_576, code)
    _require(not re.search(r'<!|<\?|\x00', svg), code)
    try:
        root = ET.fromstring(svg)
    except (ET.ParseError, ValueError):
        raise DiagramError(code) from None
    _require(root.tag == '{' + SVG_NS + '}svg', code)
    view = root.get('viewBox', '')
    _require(bool(re.fullmatch(_NUMBER + r'\s+' + _NUMBER + r'\s+' + _NUMBER + r'\s+' + _NUMBER, view)), code)
    x, y, width, height = _numbers(view)
    _require(1 <= width <= 4096 and 1 <= height <= 4096 and width * height <= 8_388_608, code)
    stack, count, text_size, ids, refs = [(root, 0, False)], 0, 0, set(), []
    while stack:
        element, depth, in_defs = stack.pop()
        count += 1
        _require(count <= 4096 and depth <= 32, code)
        _require(element.tag.startswith('{' + SVG_NS + '}'), code)
        tag = element.tag[len(SVG_NS) + 2:]
        _require(tag in _ELEMENTS and (element is root or tag != 'svg'), code)
        text_size += len(element.text or '') + len(element.tail or '')
        _require(text_size <= 8192, code)
        for name, value in element.attrib.items():
            _require(len(value) <= 32768, code)
            if name in _NUMERIC:
                _require(bool(re.fullmatch(_NUMBER, value)), code)
                number = _numbers(value)[0]
                if name in {'opacity', 'fill-opacity', 'stroke-opacity'}:
                    _require(0 <= number <= 1, code)
                if name in {'width', 'height', 'r', 'rx', 'ry', 'markerWidth', 'markerHeight', 'stroke-width', 'font-size', 'textLength'}:
                    _require(number >= 0, code)
                if name == 'font-size':
                    _require(number <= 128, code)
                if name == 'stroke-width':
                    _require(number <= 20, code)
                if name in {'markerWidth', 'markerHeight'}:
                    _require(number <= 64, code)
            elif name in {'fill', 'stroke'}:
                _require(_color(value), code)
            elif name == 'id':
                _require(bool(re.fullmatch(r'[A-Za-z_][A-Za-z0-9_-]{0,127}', value)) and value not in ids, code)
                ids.add(value)
            elif name == 'class':
                _require(bool(re.fullmatch(r'[A-Za-z0-9_ -]{1,512}', value)), code)
            elif name in {'marker-start', 'marker-mid', 'marker-end'}:
                _require(not in_defs and tag not in {'defs', 'marker'}, code)
                ref = re.fullmatch(r'url\(#([A-Za-z_][A-Za-z0-9_-]{0,127})\)', value)
                _require(bool(ref), code)
                refs.append(ref[1])
            elif name == 'd':
                _require(bool(re.fullmatch(r'[MmLlHhVvCcSsQqTtAaZz0-9eE.,+\-\s]{1,32768}', value)), code)
                _numbers(value)
            elif name in {'points', 'stroke-dasharray'}:
                _require(bool(re.fullmatch(r'[0-9eE.,+\-\s]{1,32768}', value)), code)
                _numbers(value)
            elif name == 'transform':
                _require(bool(re.fullmatch(r'(?:(?:translate|rotate|scale|matrix|skewX|skewY)\([0-9eE.,+\-\s]+\)\s*){1,8}', value)), code)
                _numbers(value)
            elif name == 'viewBox':
                _require(bool(re.fullmatch(_NUMBER + r'\s+' + _NUMBER + r'\s+' + _NUMBER + r'\s+' + _NUMBER, value)), code)
                _numbers(value)
            elif name in _FIXED:
                _require(value in _FIXED[name], code)
            elif name == 'aria-label':
                _require(len(value) <= 500, code)
            else:
                raise DiagramError(code)
        stack.extend((child, depth + 1, in_defs or tag in {'defs', 'marker'}) for child in element)
    _require(all(ref in ids for ref in refs), code)
    root.set('width', format(width, 'g'))
    root.set('height', format(height, 'g'))
    ET.register_namespace('', SVG_NS)
    return ET.tostring(root, encoding='unicode')


def diagram_prompt(kind: str, goal: str) -> str:
    _require(isinstance(kind, str) and kind in KINDS)
    _require(isinstance(goal, str) and 0 < len(goal) <= 1000 and bool(goal.strip()) and '\x00' not in goal)
    return (f'为以下需求生成 {kind} 图表。只输出 Mermaid 源码，不要代码围栏或解释。\n'
            '先声明类型；flowchart 使用 TD 或 LR。节点最多 12、边最多 36；使用简短英文标识和标签。'
            '类图用独立 class 声明及简单属性/方法，关系带基数；时序图先用 participant 声明所有参与者。'
            '仅简单关系与声明，不使用复合状态、子图、注释、HTML、链接、指令配置、click、classDef、样式或自定义资源。'
            '不要缩写；复杂图拆分。渲染通过不代表业务正确，导出仍须用户独立批准。\n需求：' + goal.strip())


def repair_prompt(kind: str, goal: str, error_code: str, attempt: int) -> str:
    _require(type(attempt) is int and 1 <= attempt <= MAX_REPAIRS)
    _require(isinstance(error_code, str) and error_code in {'unsupported_source', 'renderer_failed', 'unsafe_svg', 'render_timeout'})
    # Never send raw renderer exceptions, SVG, or prior source into model evidence.
    return diagram_prompt(kind, goal) + f'\n第 {attempt}/{MAX_REPAIRS} 次受限修正：{error_code}。简化并重新生成完整源码。'
