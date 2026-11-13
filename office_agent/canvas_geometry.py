from copy import deepcopy
from decimal import Decimal, InvalidOperation
import re

from lxml import etree as ET

from .writes import GuardError


S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ACTIONS = {'row_height', 'column_width'}


def _number(value, low, high):
    if type(value) not in {int, float, str}:
        raise GuardError('尺寸必须是有限数值')
    try:
        number = Decimal(str(value))
    except InvalidOperation:
        raise GuardError('尺寸数值无效') from None
    if not number.is_finite() or not low <= number <= high:
        raise GuardError('尺寸超出允许范围')
    return number


def _columns(root):
    containers = root.findall(S+'cols')
    if len(containers) > 1:
        raise GuardError('多个列尺寸容器暂不支持调整')
    nodes = list(containers[0]) if containers else []
    previous = 0
    for node in nodes:
        try:
            low, high = int(node.get('min', '')), int(node.get('max', ''))
        except ValueError:
            raise GuardError('列尺寸范围无效') from None
        if node.tag != S+'col' or not previous < low <= high <= 16384:
            raise GuardError('列尺寸重复、交叠或顺序无效')
        previous = high
    return nodes


def _target(root, action, row, col):
    if action not in ACTIONS or type(row) is not int or type(col) is not int or not 1 <= row <= 10000 or not 1 <= col <= 256:
        raise GuardError('行列尺寸目标无效')
    if root.find(S+'sheetProtection') is not None:
        raise GuardError('受保护工作表不能调整尺寸')
    if len(root.findall(S+'sheetData')) != 1:
        raise GuardError('工作表数据容器身份无效')
    if action == 'row_height':
        nodes = root.findall('./'+S+'sheetData/'+S+'row')
        matches = [node for node in nodes if node.get('r') == str(row)]
        if len(matches) > 1:
            raise GuardError('行标识重复，不能调整尺寸')
        node = matches[0] if matches else None
    else:
        node = next((node for node in _columns(root) if int(node.get('min')) <= col <= int(node.get('max'))), None)
    if node is not None:
        if any(node.get(key, '0') not in {'0', 'false'} for key in ('hidden', 'collapsed', 'outlineLevel', 'bestFit', 'thickTop', 'thickBot')):
            raise GuardError('隐藏、分组、自动适应或特殊行列请保留原尺寸')
        if any(key.startswith('{') for key in node.attrib):
            raise GuardError('扩展行列属性暂不支持调整')
        key, limit = ('ht', 409) if action == 'row_height' else ('width', 255)
        if node.get(key) is not None:
            _number(node.get(key), 0, limit)
    return node


def snapshot(root, action, row, col):
    node = _target(root, action, row, col)
    return dict(node.attrib) if node is not None else None


def describe(root, action, row, col):
    try:
        attrs = snapshot(root, action, row, col)
        key = 'ht' if action == 'row_height' else 'width'
        value = attrs.get(key) if attrs else None
        defaults = root.find(S+'sheetFormatPr')
        fallback = defaults.get('defaultRowHeight' if key == 'ht' else 'defaultColWidth') if defaults is not None else None
        if fallback is not None:
            fallback = float(_number(fallback, 0, 409 if key == 'ht' else 255))
        return {'editable': True, 'value': float(value) if value is not None else None, 'default':fallback,
                'source': 'explicit' if value is not None else 'automatic', 'unit': 'pt' if key == 'ht' else 'character'}
    except GuardError as error:
        return {'editable': False, 'value': None, 'source': 'unsupported', 'reason': str(error)}


def dimensions(root, row, col, rows, cols, *, row_indices=None, col_indices=None):
    return {'rows': {str(r): describe(root, 'row_height', r, col) for r in (row_indices if row_indices is not None else range(row, row+rows))},
            'columns': {str(c): describe(root, 'column_width', row, c) for c in (col_indices if col_indices is not None else range(col, col+cols))},
            'display': '结构网格采用尺寸近似展示；自动行高和字体相关列宽不代表Excel实际排版'}


def affects(operation, coordinate):
    if operation.get('action') not in ACTIONS:
        return operation.get('target') == coordinate
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', coordinate or '')
    target = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,4})', operation.get('target', ''))
    if not match or not target:
        raise GuardError('保护范围身份无效，不能调整行列尺寸')
    return match[2] == target[2] if operation['action'] == 'row_height' else match[1] == target[1]


def apply(root, operation, row, col):
    action, value = operation['action'], operation['value']
    node = _target(root, action, row, col)
    current = dict(node.attrib) if node is not None else None
    if current != operation['expected']:
        raise GuardError('行列尺寸前置条件已变化，请重新预览')
    number = None if value is None else _number(value, 6 if action == 'row_height' else 2, 200 if action == 'row_height' else 80)
    key = 'ht' if action == 'row_height' else 'width'
    before = current.get(key) if current else None
    if (number is None and before is None) or (number is not None and before is not None and Decimal(before) == number and (current or {}).get('customHeight' if action == 'row_height' else 'customWidth') == '1'):
        raise GuardError('行列尺寸没有实际变化')
    if action == 'row_height':
        container = root.find(S+'sheetData')
        if container is None:
            raise GuardError('工作表缺少数据容器')
        rows = list(container)
        try:
            identities = [int(entry.get('r', '')) for entry in rows]
        except ValueError:
            raise GuardError('行位置无效') from None
        if identities != sorted(set(identities)) or any(entry.tag != S+'row' or str(identity) != entry.get('r') or not 1 <= identity <= 1048576 for entry, identity in zip(rows, identities)):
            raise GuardError('行顺序或标识无效')
        if node is None:
            node = ET.Element(S+'row', r=str(row))
            container.insert(sum(value < row for value in identities), node)
        custom = 'customHeight'
        label = '整行 '+str(row)
    else:
        container = root.find(S+'cols')
        if container is None:
            container = ET.Element(S+'cols')
            data = root.find(S+'sheetData')
            if data is None:
                raise GuardError('工作表缺少数据容器')
            root.insert(root.index(data), container)
        if node is None:
            node = ET.Element(S+'col', min=str(col), max=str(col))
            index = next((index for index, entry in enumerate(container) if int(entry.get('min')) > col), len(container))
            container.insert(index, node)
        else:
            index, low, high = container.index(node), int(node.get('min')), int(node.get('max'))
            if low < col:
                left = deepcopy(node);left.set('max', str(col-1));container.insert(index, left)
            if col < high:
                right = deepcopy(node);right.set('min', str(col+1));container.insert(container.index(node)+1, right)
            node.set('min', str(col));node.set('max', str(col))
        custom = 'customWidth'
        label = '整列（第'+str(col)+'列）'
    if number is None:
        node.attrib.pop(key, None);node.attrib.pop(custom, None)
    else:
        node.set(key, format(number.normalize(), 'f'));node.set(custom, '1')
    return {'target': operation['target'], 'action': action, 'before': before or '自动/默认',
            'after': str(number) if number is not None else '自动/默认', 'scope': label+'，包含未显示单元格',
            'unit': '磅' if action == 'row_height' else '字符宽度', 'derived': '仅修改尺寸；保留单元格、公式缓存和其他行列属性'}
