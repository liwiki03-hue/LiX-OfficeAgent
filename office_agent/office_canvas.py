from copy import deepcopy
from decimal import Decimal, InvalidOperation
import io
import posixpath
import re
import zipfile

from lxml import etree as ET

from .paged_documents import xml, sha
from .rag_sources import _package, SourceError
from .writes import GuardError
from .canvas_formulas import Calculator
from . import canvas_display
from . import canvas_geometry
from . import canvas_panes


S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
MAX_BYTES = 8 * 1024 * 1024


def load(data, suffix):
    if suffix not in {'.xlsx', '.pptx'} or not isinstance(data, bytes) or not 0 < len(data) <= MAX_BYTES:
        raise GuardError('表格/演示编辑仅接收8 MiB以内普通XLSX/PPTX')
    try:
        checked = _package(data, suffix)
        if any(any(token in name.lower() for token in ('embeddings/', 'activex', 'connections', 'customui', 'signatures')) for name in checked):
            raise GuardError('嵌入、活动内容、数据连接或数字签名暂不支持编辑；原文件未改变')
        with zipfile.ZipFile(io.BytesIO(data)) as book:
            raw = {name: book.read(name) for name in checked}
        images = {name for name in raw if '/media/' in name.lower()}
        for name, value in raw.items():
            if not name.endswith('.rels'):
                continue
            base = posixpath.dirname(posixpath.dirname(name)) if name != '_rels/.rels' else ''
            for rel in xml(value):
                if rel.get('Type', '').endswith('/image'):
                    target = rel.get('Target', '')
                    images.add(posixpath.normpath(posixpath.join(base, target)) if not target.startswith('/') else target[1:])
        for name in images:
            value = raw.get(name, b'')
            if not (value.startswith(b'\x89PNG\r\n\x1a\n') or value.startswith(b'\xff\xd8\xff')):
                raise GuardError('当前文件内编辑仅支持PNG/JPEG媒体；其他图片/音视频保留在原有只读路径，不进入转换')
        for node in checked.values():
            if isinstance(node, ET._Element):
                for item in node.iter():
                    if item.tag == S+'f' and any(token in (item.text or '').upper() for token in ('[', ']', '!', '|', 'WEBSERVICE', 'RTD', 'DDE')):
                        raise GuardError('当前编辑不接收外部或跨表公式；请保留原稿另行处理')
        return raw
    except (SourceError, zipfile.BadZipFile, KeyError, ValueError) as error:
        if isinstance(error, GuardError):
            raise
        raise GuardError('文档包损坏、不安全或超出编辑预算') from None


def pack(raw):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as book:
        for name, value in raw.items():
            book.writestr(name, value)
    if stream.tell() > MAX_BYTES:
        raise GuardError('编辑结果超过8 MiB限制，请缩小修改范围')
    return stream.getvalue()


def encoded(node):
    return ET.tostring(node, encoding='UTF-8', xml_declaration=True, standalone=True)


def relations(raw, part):
    name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part)+'.rels')
    return {node.get('Id'): posixpath.normpath(posixpath.join(posixpath.dirname(part), node.get('Target', ''))) if not node.get('Target', '').startswith('/') else node.get('Target')[1:]
            for node in xml(raw[name])} if name in raw else {}


def col_number(value):
    number = 0
    for char in value:
        number = number*26+ord(char)-64
    return number


def address(value):
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,4})', value or '')
    if not match or col_number(match[1]) > 256 or int(match[2]) > 10000:
        raise GuardError('首版编辑范围为A1至IV10000，请选择有界区域')
    return int(match[2]), col_number(match[1])


def column(number):
    text = ''
    while number:
        number, remainder = divmod(number-1, 26)
        text = chr(65+remainder)+text
    return text


def inventory(raw, suffix):
    if suffix == '.xlsx':
        root = xml(raw['xl/workbook.xml'])
        protection = root.find(S+'workbookProtection')
        if protection is not None and (any(protection.get(key, '0') not in {'0', 'false'} for key in ('lockStructure', 'lockWindows', 'lockRevision')) or any('password' in key.lower() or 'hashvalue' in key.lower() for key in protection.attrib)):
            raise GuardError('受保护工作簿仅支持现有只读路径')
        rels = relations(raw, 'xl/workbook.xml')
        items = [{'id': 'sheet-'+str(index), 'name': item.get('name'), 'part': rels.get(item.get(R+'id')), 'hidden': item.get('state', 'visible') != 'visible'} for index, item in enumerate(root.findall('./'+S+'sheets/'+S+'sheet'))]
        if not 1 <= len(items) <= 32 or any(not item['part'] or item['part'] not in raw for item in items):
            raise GuardError('工作表数量超过32或关系无效')
        if sum(len(xml(raw[item['part']]).findall('.//'+S+'c')) for item in items) > 100000:
            raise GuardError('单元格超过十万格编辑预算，请拆分工作簿')
        return items
    root = xml(raw['ppt/presentation.xml'])
    rels = relations(raw, 'ppt/presentation.xml')
    items = [{'id': 'slide-'+str(index), 'name': '第 '+str(index+1)+' 页', 'part': rels.get(item.get(R+'id'))} for index, item in enumerate(root.findall('./'+P+'sldIdLst/'+P+'sldId'))]
    if not 1 <= len(items) <= 100 or any(not item['part'] or item['part'] not in raw for item in items):
        raise GuardError('幻灯片超过100页或页面关系无效')
    return items


def item_for(raw, suffix, item_id):
    for item in inventory(raw, suffix):
        if item['id'] == item_id:
            return item
    raise GuardError('当前工作表或幻灯片标识已失效')


def calculator_for(raw, item, root=None, cache=None):
    shared = xml(raw['xl/sharedStrings.xml']).findall(S+'si') if 'xl/sharedStrings.xml' in raw else []
    return Calculator(root if root is not None else xml(raw[item['part']]), cache, canvas_display.date1904(raw), item['id'], shared)


def _style(root, cell):
    if root is None:
        return {}
    xfs, fonts, fills = root.find(S+'cellXfs'), root.find(S+'fonts'), root.find(S+'fills')
    index = int(cell.get('s', '0'))
    if xfs is None or not 0 <= index < len(xfs):
        raise GuardError('单元格样式引用无效')
    xf = xfs[index]
    if fonts is None or not 0 <= int(xf.get('fontId', '0')) < len(fonts) or fills is not None and not 0 <= int(xf.get('fillId', '0')) < len(fills):
        raise GuardError('字体或底纹样式引用无效')
    font = fonts[int(xf.get('fontId', '0'))] if fonts is not None else None
    align = xf.find(S+'alignment')
    value = {'bold': False, 'italic': False}
    if font is not None:
        for key, tag in [('bold', 'b'), ('italic', 'i')]:
            node = font.find(S+tag)
            value[key] = node is not None and node.get('val', '1') not in {'0', 'false'}
        for key, tag in [('font', 'name'), ('size', 'sz')]:
            node = font.find(S+tag)
            if node is not None:
                value[key] = node.get('val')
        color = font.find(S+'color')
        if color is not None and re.fullmatch(r'[0-9a-fA-F]{8}', color.get('rgb', '')):
            value['color'] = '#'+color.get('rgb')[2:]
    if align is not None:
        value['align'] = align.get('horizontal', 'general')
    if fills is not None and int(xf.get('fillId', '0')) < len(fills):
        color = fills[int(xf.get('fillId', '0'))].find('./'+S+'patternFill/'+S+'fgColor')
        if color is not None and re.fullmatch(r'[0-9a-fA-F]{8}', color.get('rgb', '')):
            value['background'] = '#'+color.get('rgb')[2:]
    return value


def format_allowed(sheet, cell, styles):
    if styles is None:
        return False
    xfs = styles.find(S+'cellXfs')
    index = int(cell.get('s', '0'))
    if xfs is None or not 0 <= index < len(xfs):
        return False
    xf = xfs[index]
    if xf.get('xfId', '0') != '0':
        return False
    base = styles.find(S+'cellStyleXfs')
    if base is not None and len(base) and xf.get('applyFont') in {'0', 'false'} and base[0].get('fontId', '0') != xf.get('fontId', '0'):
        return False
    if cell.get('s') is None:
        row = cell.getparent()
        if row is not None and row.get('s', '0') != '0':
            return False
        cols = sheet.find(S+'cols')
        if cols is not None:
            if len(cols) > 16384:
                raise GuardError('列格式定义超过工作簿预算')
            number = address(cell.get('r'))[1]
            if any(int(node.get('min')) <= number <= int(node.get('max')) and node.get('style', '0') != '0' for node in cols):
                return False
    return True


def table_cells(root, width, height):
    cells = {}
    limitations = []
    for index, table in enumerate(root.findall('.//'+A+'tbl')):
        rows = table.findall(A+'tr')
        columns = table.findall('./'+A+'tblGrid/'+A+'gridCol')
        grid = [row.findall(A+'tc') for row in rows]
        if not rows or not columns or len(rows)*len(columns) > 100 or any(len(row) != len(columns) for row in grid):
            limitations.append('表格'+str(index+1)+'：网格不规则或超过100格，只读')
            continue
        if any(any(cell.get(key, '1' if key in {'gridSpan', 'rowSpan'} else '0') not in ({'1'} if key in {'gridSpan', 'rowSpan'} else {'0', 'false'}) for key in ('gridSpan', 'rowSpan', 'hMerge', 'vMerge')) for row in grid for cell in row):
            limitations.append('表格'+str(index+1)+'：含合并结构，只读')
            continue
        frame = next((parent for parent in table.iterancestors() if parent.tag == P+'graphicFrame'), None)
        if frame is None:
            limitations.append('表格'+str(index+1)+'：宿主框架未知，只读')
            continue
        box = None
        transform = frame.find(P+'xfrm')
        widths = [int(node.get('w', '0')) for node in columns]
        heights = [int(node.get('h', '0')) for node in rows]
        if transform is not None and transform.find(A+'off') is not None and transform.find(A+'ext') is not None and all(value > 0 for value in widths+heights):
            off, ext = transform.find(A+'off'), transform.find(A+'ext')
            if sum(widths) == int(ext.get('cx')) and sum(heights) == int(ext.get('cy')) and not any(parent.tag == P+'grpSp' for parent in frame.iterancestors()) and not any(transform.get(key, '0') not in {'0', 'false'} for key in ('rot', 'flipH', 'flipV')):
                box = [int(off.get('x')), int(off.get('y'))]
        for row_index, row in enumerate(grid):
            for col_index, cell in enumerate(row):
                body = cell.find(A+'txBody')
                if body is None or any(node.tag not in {A+'bodyPr', A+'lstStyle', A+'p'} for node in body) or any(node.tag not in {A+'pPr', A+'r', A+'endParaRPr'} for para in body.findall(A+'p') for node in para) or any(node.tag not in {A+'rPr', A+'t'} for run in body.findall('.//'+A+'r') for node in run):
                    limitations.append('表格'+str(index+1)+'：第'+str(row_index+1)+'行第'+str(col_index+1)+'列含域/特殊结构，只读')
                    continue
                cell_box = None if box is None else [(box[0]+sum(widths[:col_index]))/width, (box[1]+sum(heights[:row_index]))/height, widths[col_index]/width, heights[row_index]/height]
                cells[cell] = {'location': '表格'+str(index+1)+' · 第'+str(row_index+1)+'行第'+str(col_index+1)+'列', 'box': cell_box}
    return cells, limitations


def view(data, suffix, item_id=None, row=1, col=1, calculation_cache=None):
    raw = load(data, suffix)
    items = inventory(raw, suffix)
    item = item_for(raw, suffix, item_id or next((item['id'] for item in items if not item.get('hidden')), items[0]['id']))
    root = xml(raw[item['part']])
    common = {'items': [{key: value for key, value in entry.items() if key != 'part'} for entry in items], 'item_id': item['id'], 'kind': suffix[1:], 'engine': 'source-preserving-ooxml-v1'}
    if suffix == '.pptx':
        size = xml(raw['ppt/presentation.xml']).find(P+'sldSz')
        width, height = int(size.get('cx')), int(size.get('cy'))
        tables, limitations = table_cells(root, width, height)
        runs = []
        for index, node in enumerate(root.findall('.//'+A+'r')):
            text = node.find(A+'t')
            if text is None or len(text.text or '') > 4000:
                continue
            shape = next((parent for parent in node.iterancestors() if parent.tag == P+'sp'), None)
            cell = next((parent for parent in node.iterancestors() if parent.tag == A+'tc'), None)
            if shape is None and cell not in tables:
                continue
            prop = node.find(A+'rPr')
            transform = shape.find('./'+P+'spPr/'+A+'xfrm') if shape is not None else None
            box = tables[cell]['box'] if cell in tables else None
            if transform is not None and transform.find(A+'off') is not None and transform.find(A+'ext') is not None:
                off, ext = transform.find(A+'off'), transform.find(A+'ext')
                if not any(parent.tag == P+'grpSp' for parent in shape.iterancestors()) and not any(transform.get(key, '0') not in {'0', 'false'} for key in ('rot', 'flipH', 'flipV')):
                    box = [int(off.get('x'))/width, int(off.get('y'))/height, int(ext.get('cx'))/width, int(ext.get('cy'))/height]
            runs.append({'id': 'run-'+str(index), 'text': text.text or '', 'bold': prop is not None and prop.get('b') == '1', 'italic': prop is not None and prop.get('i') == '1', 'size': prop.get('sz') if prop is not None else None, 'box': box, 'location': tables[cell]['location'] if cell in tables else '普通文字'})
        if len(runs) > 1000:
            raise GuardError('本页文字片段超过一千项，请使用有界读取')
        return {**common, 'width': width, 'height': height, 'runs': runs, 'limitations': limitations, 'readonly_objects': '支持普通文字及100格以内未合并表格的已有文字片段；不改变表格几何。图片、图表、母版、域、合并表格和空白格保留。'+('；'.join(limitations))}
    if type(row) is not int or type(col) is not int or not 1 <= row <= 10000 or not 1 <= col <= 256:
        raise GuardError('工作表视窗位置无效')
    if item.get('hidden'):
        raise GuardError('隐藏工作表不在本轮交互编辑范围')
    panes, row_indices, col_indices = canvas_panes.window(root,row,col)
    visible_rows, visible_cols = set(row_indices), set(col_indices)
    protected = root.find(S+'sheetProtection') is not None
    shared = xml(raw['xl/sharedStrings.xml']).findall(S+'si') if 'xl/sharedStrings.xml' in raw else []
    styles = xml(raw['xl/styles.xml']) if 'xl/styles.xml' in raw else None
    mac = canvas_display.date1904(raw)
    cells = []
    calculator = Calculator(root, calculation_cache, mac, namespace=item['id'], shared=shared)
    formula_family = 'xl/calcChain.xml' not in raw and not any(node.attrib for node in root.findall('.//'+S+'f'))
    formula_refs = set()
    for node in root.findall('./'+S+'sheetData/'+S+'row/'+S+'c'):
        try:
            r, c = address(node.get('r'))
        except GuardError:
            continue
        if r in visible_rows and c in visible_cols and node.find(S+'f') is not None:
            formula_refs.add(node.get('r'))
    calculations = calculator.results(formula_refs) if formula_refs else {}
    for node in root.findall('./'+S+'sheetData/'+S+'row/'+S+'c'):
        try:
            r, c = address(node.get('r'))
        except GuardError:
            continue
        if r not in visible_rows or c not in visible_cols:
            continue
        kind = node.get('t', 'n')
        value = node.findtext(S+'v', '')
        rich = False
        if kind == 's':
            shared_node = shared[int(value)]
            rich = any(node.tag != S+'t' for node in shared_node)
            value = ''.join(node.text or '' for node in shared_node.findall(S+'t')) if not rich else ''.join(node.text or '' for node in shared_node.findall('./'+S+'r/'+S+'t'))
        if kind == 'inlineStr':
            value = ''.join(node.find(S+'is').itertext())
            rich = any(child.tag != S+'t' for child in node.find(S+'is'))
        formula = node.find(S+'f')
        calculation = calculations.get(node.get('r')) if formula is not None else None
        display_value = calculation['value'] if formula is not None and calculation and calculation.get('status') == 'calculated' else value
        display = canvas_display.describe(styles, node, display_value, kind, mac) if formula is None or calculation and calculation.get('status') == 'calculated' else None
        if calculation and calculation.get('value_type') == 'text':
            display = {'text': calculation['value'], 'kind': 'text', 'format': '受限文本计算'}
        cell = {'id': node.get('r'), 'value': value, 'formula': formula.text if formula is not None else None, 'type': kind,
                'editable': not protected and formula is None and not rich and len(value) <= 4000,
                'formula_editable': formula_family and not protected and not rich and (formula is None or not formula.attrib and calculation['status'] == 'calculated'), 'calculation': calculation,
                'format_editable': format_allowed(root, node, styles),
                'style': _style(styles, node)}
        if display:
            cell['display'] = display
        cells.append(cell)
    merges = [node.get('ref') for node in root.findall('./'+S+'mergeCells/'+S+'mergeCell')]
    try:
        formula_graph = calculator.graph_report()
    except GuardError as error:
        formula_graph = {'status': 'unavailable', 'reason': str(error)}
    return {**common, 'row': row, 'col': col, 'rows': len(row_indices), 'cols': len(col_indices), 'cells': cells, 'merges': merges, 'protected': protected, 'formula_creation_allowed': formula_family and not protected,
            'formula_graph': formula_graph,
            'panes':panes, 'row_indices':row_indices, 'col_indices':col_indices, 'date_system': '1904' if mac else '1900',
            'display_scope': '只读日期/数值使用内置格式及明确白名单；未知自定义格式保留原值并标记不支持；保存后由Office重算/排版',
            'geometry': canvas_geometry.dimensions(root, row, col, len(row_indices), len(col_indices), row_indices=row_indices, col_indices=col_indices),
            'scope': '30行×16列窗口；同表受限算术、统计、日期与LEN/TRIM/CONCATENATE文本计算，NUMBERVALUE需明确分隔符；28位十进制预览，空白不当0，跨表/外链拒绝；保存后由Office重算。日期/数值显示近似，打印布局请查看页图。'}


def _format(payload):
    if not isinstance(payload, dict) or set(payload)-{'bold', 'italic', 'size', 'align'} or not payload:
        raise GuardError('格式参数无效')
    for key in ('bold', 'italic'):
        if key in payload and type(payload[key]) is not bool:
            raise GuardError('强调格式必须是布尔值')
    if 'size' in payload and (type(payload['size']) not in {int, float} or not 6 <= payload['size'] <= 72):
        raise GuardError('字号范围为6至72磅')
    if 'align' in payload and payload['align'] not in {'left', 'center', 'right'}:
        raise GuardError('对齐方式无效')


def transform(data, suffix, operation):
    if not isinstance(operation, dict) or set(operation) != {'item_id', 'target', 'action', 'value', 'expected'}:
        raise GuardError('操作字段不完整或含未允许字段')
    raw = load(data, suffix)
    item = item_for(raw, suffix, operation['item_id'])
    root = xml(raw[item['part']])
    action, value, target = operation['action'], operation['value'], operation['target']
    if action == canvas_panes.ACTION:
        if suffix != '.xlsx' or item.get('hidden'):
            raise GuardError('冻结窗格仅适用于可见工作表')
        views = xml(raw['xl/workbook.xml']).findall('./'+S+'bookViews/'+S+'workbookView')
        if not views:
            raise GuardError('缺少工作簿窗口身份，不能新增悬空的窗格引用')
        summary = canvas_panes.apply(root,operation)
        raw[item['part']] = encoded(root)
        return pack(raw), {'item':item['name'],**summary}
    if action in canvas_geometry.ACTIONS:
        if suffix != '.xlsx' or item.get('hidden'):
            raise GuardError('行列尺寸仅适用于可见工作表')
        r, c = address(target)
        summary = canvas_geometry.apply(root, operation, r, c)
        raw[item['part']] = encoded(root)
        return pack(raw), {'item': item['name'], **summary}
    if action not in {'text', 'number', 'clear', 'format', 'formula'}:
        raise GuardError('此操作未在当前编辑器开放')
    if action == 'format':
        _format(value)
    elif not isinstance(value, str) or len(value) > 4000 or any(ord(char) < 32 and char not in '\t\n\r' for char in value):
        raise GuardError('文字内容无效或超过4000字符')
    summary = {'item': item['name'], 'target': target, 'action': action}
    if suffix == '.pptx':
        if action not in {'text', 'format'} or not re.fullmatch(r'run-[0-9]{1,4}', target or ''):
            raise GuardError('演示稿首版仅支持普通文字片段和格式')
        options = {run['id']: run for run in view(data, suffix, item['id'])['runs']}
        if target not in options or options[target]['text'] != operation['expected']:
            raise GuardError('文字前置条件已改变，请重新选择')
        node = root.findall('.//'+A+'r')[int(target[4:])]
        text = node.find(A+'t')
        summary['before'] = text.text or ''
        summary['location'] = options[target].get('location', '普通文字')
        if action == 'text':
            if value == summary['before']: raise GuardError('没有实际变化')
            text.text = value
        else:
            if 'align' in value:
                raise GuardError('演示稿文字片段不支持段落对齐操作')
            prop = node.find(A+'rPr')
            if prop is None:
                prop = ET.Element(A+'rPr'); node.insert(0, prop)
            for key, attr in [('bold', 'b'), ('italic', 'i'), ('size', 'sz')]:
                if key in value:
                    prop.set(attr, str(round(value[key]*100)) if key == 'size' else str(int(value[key])))
        summary['after'] = value
    else:
        r, c = address(target)
        if root.find(S+'sheetProtection') is not None or item.get('hidden'):
            raise GuardError('受保护或隐藏工作表不可修改')
        for merge in root.findall('./'+S+'mergeCells/'+S+'mergeCell'):
            start, end = merge.get('ref').split(':') if ':' in merge.get('ref') else [merge.get('ref')]*2
            r1, c1 = address(start); r2, c2 = address(end)
            if r1 <= r <= r2 and c1 <= c <= c2 and target != start:
                raise GuardError('这是合并单元格的覆盖坐标，请编辑左上承载格')
        sheet_data = root.find(S+'sheetData')
        if sheet_data is None:
            raise GuardError('缺少普通工作表数据')
        row_node = next((node for node in sheet_data if node.get('r') == str(r)), None)
        cell = next((node for node in row_node if node.tag == S+'c' and node.get('r') == target), None) if row_node is not None else None
        existing = next((node for node in view(data, suffix, item['id'], r, c)['cells'] if node['id'] == target), {'value': '', 'editable': True})
        expected = '='+existing['formula'] if action == 'formula' and existing.get('formula') is not None else existing['value']
        if (not existing.get('formula_editable', existing['editable']) if action == 'formula' else not existing['editable']) or expected != operation['expected']:
            raise GuardError('公式、富文本或旧单元格前置条件不允许此修改')
        if action == 'formula' and ('xl/calcChain.xml' in raw or any(node.attrib for node in root.findall('.//'+S+'f'))):
            raise GuardError('含计算链或共享/数组等特殊公式，暂不新增或修改公式')
        calculation = calculator_for(raw, item, root).calculate(target, value) if action == 'formula' else None
        if action == 'formula' and expected == value:
            raise GuardError('没有实际变化')
        if action == 'text' and existing.get('type') in {'s', 'inlineStr'} and value == existing['value'] or action == 'clear' and existing['value'] == '':
            raise GuardError('没有实际变化')
        if row_node is None:
            row_node = ET.Element(S+'row', r=str(r))
            index = next((index for index, node in enumerate(sheet_data) if int(node.get('r')) > r), len(sheet_data)); sheet_data.insert(index, row_node)
        if cell is None:
            cell = ET.Element(S+'c', r=target)
            index = next((index for index, node in enumerate(row_node) if node.tag == S+'c' and address(node.get('r'))[1] > c), len(row_node)); row_node.insert(index, cell)
        summary['before'], summary['after'] = expected, value
        if calculation is not None:
            summary['calculation'] = calculation
        if action == 'format':
            if 'xl/styles.xml' not in raw:
                raise GuardError('缺少样式表，暂不创建新样式体系')
            styles = xml(raw['xl/styles.xml']); xfs, fonts = styles.find(S+'cellXfs'), styles.find(S+'fonts')
            if not format_allowed(root, cell, styles):
                raise GuardError('此格依赖行列/命名样式继承，首版暂不调整格式；文字和数值修改仍保留原样式')
            if len(xfs) >= 4096 or len(fonts) >= 1024:
                raise GuardError('样式数量达到编辑预算')
            xf = deepcopy(xfs[int(cell.get('s', '0'))]); font = deepcopy(fonts[int(xf.get('fontId', '0'))])
            for key, tag in [('bold', 'b'), ('italic', 'i'), ('size', 'sz')]:
                if key in value:
                    for old in font.findall(S+tag): font.remove(old)
                    ET.SubElement(font, S+tag, val=str(value[key]) if key == 'size' else str(int(value[key])))
            if set(value)-{'align'}:
                xf.set('fontId', str(len(fonts))); xf.set('applyFont', '1'); fonts.append(font); fonts.set('count', str(len(fonts)))
            if 'align' in value:
                align = xf.find(S+'alignment')
                if align is None: align = ET.SubElement(xf, S+'alignment')
                align.set('horizontal', value['align']); xf.set('applyAlignment', '1')
            cell.set('s', str(len(xfs))); xfs.append(xf); xfs.set('count', str(len(xfs))); raw['xl/styles.xml'] = encoded(styles)
        else:
            for tag in ('v', 'is', 'f'):
                for node in cell.findall(S+tag): cell.remove(node)
            cell.attrib.pop('t', None)
            if action == 'formula':
                ET.SubElement(cell, S+'f').text = value[1:]
            elif action == 'number':
                try:
                    number = Decimal(value)
                    if not number.is_finite() or len(number.as_tuple().digits) > 15 or abs(number.adjusted()) > 300: raise InvalidOperation
                except InvalidOperation:
                    raise GuardError('数值应为有限十进制且不超过15位有效数字') from None
                if existing.get('type') == 'n' and existing['value'] and number == Decimal(existing['value']):
                    raise GuardError('没有实际变化')
                ET.SubElement(cell, S+'v').text = str(number)
            elif action == 'text':
                cell.set('t', 'inlineStr'); text = ET.SubElement(ET.SubElement(cell, S+'is'), S+'t'); text.text = value; text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            dimension = root.find(S+'dimension')
            if dimension is not None:
                last = dimension.get('ref', 'A1').split(':')[-1]
                match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', last)
                if not match or not 1 <= int(match[2]) <= 1048576 or not 1 <= col_number(match[1]) <= 16384:
                    raise GuardError('原工作表范围标记无效，拒绝猜测或缩小范围')
                old_r, old_c = int(match[2]), col_number(match[1])
                dimension.set('ref', 'A1:'+column(max(c, old_c))+str(max(r, old_r)))
    raw[item['part']] = encoded(root)
    if suffix == '.xlsx' and action != 'format':
        for sheet in inventory(raw, suffix):
            tree = xml(raw[sheet['part']]); changed = False
            for cell in tree.findall('.//'+S+'c'):
                if cell.find(S+'f') is not None:
                    for cache in cell.findall(S+'v'): cell.remove(cache); changed = True
            if changed: raw[sheet['part']] = encoded(tree)
        tree = xml(raw['xl/workbook.xml']); calc = tree.find(S+'calcPr')
        if calc is None: calc = ET.SubElement(tree, S+'calcPr')
        calc.set('fullCalcOnLoad', '1'); calc.set('forceFullCalc', '1'); calc.set('calcMode', 'auto'); raw['xl/workbook.xml'] = encoded(tree)
        summary['derived'] = '保留公式并清除文件内旧缓存；本机仅预览已支持公式，保存后仍需Office重算'
    candidate = pack(raw)
    load(candidate, suffix)
    original = load(data, suffix)
    changed_parts = [name for name in raw if raw[name] != original[name]]
    if not changed_parts:
        raise GuardError('没有实际变化')
    summary['changed_parts'] = changed_parts
    return candidate, summary


def verify(source, candidate, suffix, operations):
    current = source
    for operation in operations:
        current, _ = transform(current, suffix, operation)
    expected, actual = load(current, suffix), load(candidate, suffix)
    if expected != actual:
        raise GuardError('候选中存在无法归因于批准操作的变化')
    original = load(source, suffix)
    return {'source_sha256': sha(source), 'candidate_sha256': sha(candidate), 'operations': len(operations),
            'changed_parts': [name for name in actual if actual[name] != original[name]],
            'non_target_parts_exact': True, 'coverage': '包部件精确回放；样式继承及Office排版正确性须另行核对'}
