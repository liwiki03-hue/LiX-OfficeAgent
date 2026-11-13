import hashlib
import re

from lxml import etree as ET

from .writes import GuardError


S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ACTION = 'freeze_panes'


def _position(value):
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', value or '')
    if not match:
        raise GuardError('工作表位置无效')
    col = 0
    for char in match[1]:
        col = col*26+ord(char)-64
    row = int(match[2])
    if row > 1048576 or col > 16384:
        raise GuardError('工作表位置越界')
    return row, col


def _column(number):
    text = ''
    while number:
        number, rest = divmod(number-1, 26)
        text = chr(65+rest)+text
    return text


def _view(root):
    containers = root.findall(S+'sheetViews')
    if len(containers) > 1:
        raise GuardError('多个工作表视图容器暂不支持调整')
    container = containers[0] if containers else None
    if container is None:
        return None
    if len(container) != 1 or container[0].tag != S+'sheetView' or container[0].get('workbookViewId', '0') != '0':
        raise GuardError('多个窗口或扩展视图请保留原设置')
    view = container[0]
    if view.get('rightToLeft', '0') not in {'0','false'} or any(key.startswith('{') for key in view.attrib):
        raise GuardError('从右向左或扩展视图暂不支持冻结调整')
    if len(view.findall(S+'pane')) > 1:
        raise GuardError('窗格标识重复，不能猜测冻结范围')
    if any(child.tag not in {S+'pane',S+'selection'} for child in view):
        raise GuardError('含特殊窗口部件，请保留原视图')
    return view


def identity(root):
    _view(root)
    node = root.find(S+'sheetViews')
    return hashlib.sha256(ET.tostring(node)).hexdigest() if node is not None else 'absent'


def _counts(value):
    if type(value) is not dict or set(value) != {'rows','cols'} or any(type(value[key]) is not int for key in value):
        raise GuardError('冻结范围必须明确行数与列数')
    if not 0 <= value['rows'] <= 4 or not 0 <= value['cols'] <= 3:
        raise GuardError('首版最多冻结前4行和前3列，其余保留为滚动区域')
    return value['rows'], value['cols']


def _merge_boundaries(root, rows, cols):
    for node in root.findall('./'+S+'mergeCells/'+S+'mergeCell'):
        values = (node.get('ref') or '').split(':')
        if len(values) not in {1,2}:
            raise GuardError('合并位置无效，请保留原窗格')
        r1,c1 = _position(values[0]);r2,c2 = _position(values[-1])
        if r1 > r2 or c1 > c2 or rows and r1 <= rows < r2 or cols and c1 <= cols < c2:
            raise GuardError('合并单元格跨越冻结边界，请选择其他边界')


def inspect(root):
    try:
        view = _view(root)
        pane = view.find(S+'pane') if view is not None else None
        rows = cols = 0
        if pane is not None:
            if pane.get('state') != 'frozen' or any(key not in {'xSplit','ySplit','topLeftCell','activePane','state'} for key in pane.attrib):
                raise GuardError('拆分窗格或扩展窗格只保留，不在结构网格中模拟')
            try:
                rows,cols = int(pane.get('ySplit','0')),int(pane.get('xSplit','0'))
            except ValueError:
                raise GuardError('冻结位置不是明确整数') from None
            _counts({'rows':rows,'cols':cols})
            if rows == cols == 0:
                raise GuardError('冻结窗格缺少有效范围')
            if pane.get('topLeftCell'):
                r,c = _position(pane.get('topLeftCell'))
                if r <= rows or c <= cols:
                    raise GuardError('冻结窗格滚动位置与边界冲突')
            expected = 'bottomRight' if rows and cols else 'bottomLeft' if rows else 'topRight'
            if pane.get('activePane',expected) != expected:
                raise GuardError('活动窗格与冻结边界不匹配')
            _merge_boundaries(root,rows,cols)
        protected = root.find(S+'sheetProtection') is not None
        return {'supported':True,'editable':not protected,'rows':rows,'cols':cols,
                'reason':'受保护工作表不能调整窗格' if protected else '',
                'scope':'冻结文件开头的行列；冻结区计入30行×16列预算'}
    except GuardError as error:
        return {'supported':False,'editable':False,'rows':0,'cols':0,'reason':str(error)}


def window(root, row, col):
    info = inspect(root)
    rows,cols = info['rows'],info['cols']
    start_row,start_col = max(row,rows+1),max(col,cols+1)
    r_indices = list(range(1,rows+1))+list(range(start_row,min(start_row+30-rows,10001)))
    c_indices = list(range(1,cols+1))+list(range(start_col,min(start_col+16-cols,257)))
    return info,r_indices,c_indices


def apply(root, operation):
    if operation['target'] != 'A1':
        raise GuardError('冻结操作只接受当前工作表视图目标')
    info = inspect(root)
    if not info['editable']:
        raise GuardError(info['reason'])
    rows,cols = _counts(operation['value'])
    _merge_boundaries(root,rows,cols)
    if identity(root) != operation['expected']:
        raise GuardError('窗格前置条件已改变，请重新预览')
    if (rows,cols) == (info['rows'],info['cols']):
        raise GuardError('冻结范围没有实际变化')
    container = root.find(S+'sheetViews')
    if container is None:
        container = ET.Element(S+'sheetViews')
        following = next((node for node in root if node.tag not in {S+'sheetPr',S+'dimension'}), None)
        root.insert(root.index(following) if following is not None else len(root),container)
        view = ET.SubElement(container,S+'sheetView',workbookViewId='0')
    else:
        view = container[0]
    selections = view.findall(S+'selection')
    original = next((node for node in selections if node.get('pane') in {None,'bottomRight','bottomLeft','topRight'}), None)
    attributes = dict(original.attrib) if original is not None else {}
    if any(key not in {'pane','activeCell','activeCellId','sqref'} for key in attributes):
        raise GuardError('扩展选区属性暂不支持调整')
    for node in list(view):
        view.remove(node)
    if rows or cols:
        active = 'bottomRight' if rows and cols else 'bottomLeft' if rows else 'topRight'
        pane = ET.SubElement(view,S+'pane',state='frozen',topLeftCell=_column(cols+1)+str(rows+1),activePane=active)
        if cols:pane.set('xSplit',str(cols))
        if rows:pane.set('ySplit',str(rows))
        attributes = {'pane':active,'activeCell':_column(cols+1)+str(rows+1),'sqref':_column(cols+1)+str(rows+1)}
    else:
        attributes.pop('pane',None)
    ET.SubElement(view,S+'selection',**attributes)
    return {'target':'A1','action':ACTION,'before':f"前{info['rows']}行 / 前{info['cols']}列",
            'after':f'前{rows}行 / 前{cols}列' if rows or cols else '取消冻结', 'scope':'工作表视图',
            'derived':'仅修改窗格及活动选区；单元格、公式、格式和源文件不变'}
