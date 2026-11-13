import csv
import io
import re

from . import office_canvas as engine
from .writes import GuardError


MAX_CELLS = 100
MAX_TEXT = 65536


def rectangle(start, end):
    if not isinstance(start, str) or not isinstance(end, str):
        raise GuardError('区域位置必须为A1形式的单元格地址')
    r1, c1 = engine.address(start)
    r2, c2 = engine.address(end)
    top, bottom = sorted((r1, r2))
    left, right = sorted((c1, c2))
    if (bottom-top+1)*(right-left+1) > MAX_CELLS:
        raise GuardError('区域超过100格，请缩小选择后分组处理')
    return top, left, bottom, right


def parse_tsv(text):
    if not isinstance(text, str) or len(text) > MAX_TEXT or '\0' in text:
        raise GuardError('粘贴内容无效或超过64 Ki字符，请缩小区域')
    try:
        rows = []
        for row in csv.reader(io.StringIO(text, newline=''), delimiter='\t', strict=True):
            if not row or not 1 <= len(row) <= MAX_CELLS or any(len(value) > 4000 for value in row):
                raise GuardError('粘贴内容必须为矩形，且每格不超过4000字符')
            rows.append(row)
            if len(rows)*len(rows[0]) > MAX_CELLS:
                raise GuardError('粘贴区域超过100格，请分组处理')
            if len(row) != len(rows[0]):
                raise GuardError('粘贴各行列数不一致，请先整理为矩形')
    except csv.Error:
        raise GuardError('粘贴内容的引号或分隔格式无效') from None
    if not rows:
        raise GuardError('请输入要粘贴的制表符分隔内容')
    return rows


def read_range(data, item_id, start, end):
    top, left, bottom, right = rectangle(start, end)
    first = engine.view(data, '.xlsx', item_id, top, left)
    for ref in first['merges']:
        a, _, b = ref.partition(':')
        points = []
        for value in (a, b or a):
            match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', value)
            if not match or int(match[2])>1048576 or engine.col_number(match[1])>16384:
                raise GuardError('工作表合并范围无效，请保留原稿并核对来源')
            points.append((int(match[2]), engine.col_number(match[1])))
        (r1, c1), (r2, c2) = points
        if r1>r2 or c1>c2:
            raise GuardError('工作表合并范围顺序无效')
        if max(top, r1) <= min(bottom, r2) and max(left, c1) <= min(right, c2):
            raise GuardError('区域包含合并单元格；区域复制/粘贴暂不开放，请使用承载格的单格编辑')
    windows = {}
    rows = []
    for row in range(top, bottom+1):
        values = []
        for col in range(left, right+1):
            key = (top+(row-top)//30*30, left+(col-left)//16*16)
            if key not in windows:
                current = first if key == (top, left) else engine.view(data, '.xlsx', item_id, *key)
                windows[key] = {cell['id']:cell for cell in current['cells']}
            target = engine.column(col)+str(row)
            values.append(windows[key].get(target, {'id':target,'value':'','type':'inlineStr','formula':None,'editable':not first['protected']}))
        rows.append(values)
    stream = io.StringIO(newline='')
    csv.writer(stream, delimiter='\t', lineterminator='\n').writerows([
        ['='+cell['formula'] if cell['formula'] is not None else cell['value'] for cell in row] for row in rows])
    tsv = stream.getvalue()
    if len(tsv) > MAX_TEXT:
        raise GuardError('区域复制内容超过64 Ki字符，请缩小范围')
    return {'item_id':item_id, 'start':engine.column(left)+str(top), 'end':engine.column(right)+str(bottom),
            'rows':bottom-top+1, 'cols':right-left+1, 'cells':rows, 'tsv':tsv,
            'copy_scope':'原始文字、数值及公式表达式；不是显示格式或旧公式缓存'}


def paste_operations(data, item_id, start, text, mode):
    if not isinstance(mode, str) or mode not in {'text', 'numbers'}:
        raise GuardError('粘贴类型仅支持文本或明确的十进制识别')
    values = parse_tsv(text)
    row, col = engine.address(start)
    end = engine.column(col+len(values[0])-1)+str(row+len(values)-1)
    selected = read_range(data, item_id, start, end)
    operations = []
    for old_row, new_row in zip(selected['cells'], values):
        for before, value in zip(old_row, new_row):
            if not before['editable']:
                raise GuardError('区域包含公式、富文本或受保护单元格；未采用任何粘贴内容')
            numeric = mode == 'numbers' and re.fullmatch(r'[+-]?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?', value) is not None
            action = 'number' if numeric else 'text'
            same_kind = before['type'] == 'n' if numeric else before['type'] in {'s', 'inlineStr'}
            if value == before['value'] and same_kind or not value and not before['value']:
                continue
            if numeric and before['type'] == 'n' and before['value']:
                from decimal import Decimal
                if Decimal(value) == Decimal(before['value']):
                    continue
            operations.append({'item_id':item_id,'target':before['id'],'action':action,'value':value,'expected':before['value']})
    if not operations:
        raise GuardError('区域没有实际变化，未创建编辑组')
    return operations, {key: selected[key] for key in ('item_id', 'start', 'end', 'rows', 'cols')}
