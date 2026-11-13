from datetime import date, timedelta
from decimal import Decimal, DecimalException, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_EVEN, ROUND_HALF_UP, ROUND_UP, localcontext
from functools import reduce
import re

from .writes import GuardError


S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
TOKEN = re.compile(r'\s*("(?:[^"\r\n]|"")*"|\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,4}|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?|[A-Za-z]+|<=|>=|<>|[()+*/,:=<>-])')
CELL = re.compile(r'\$?([A-Z]{1,3})\$?([1-9][0-9]{0,4})')


def location(value):
    match = CELL.fullmatch(value.upper())
    if not match:
        raise GuardError('公式仅支持同表A1至IV10000引用')
    col = 0
    for char in match[1]:
        col = col * 26 + ord(char) - 64
    if col > 256 or int(match[2]) > 10000:
        raise GuardError('公式引用超过A1至IV10000范围')
    return match[1] + match[2], int(match[2]), col


def name(row, col):
    text = ''
    while col:
        col, rem = divmod(col - 1, 26)
        text = chr(65 + rem) + text
    return text + str(row)


def parse(value):
    if not isinstance(value, str) or not value.startswith('=') or not 1 < len(value) <= 256:
        raise GuardError('受限公式须以=开头，最多256字符')
    tokens = []
    offset = 1
    while offset < len(value.rstrip()):
        match = TOKEN.match(value, offset)
        if not match:
            raise GuardError('仅支持常见运算、比较、括号、同表引用及受限函数；不支持外链或其他函数')
        tokens.append(match[1] if match[1].startswith('"') else match[1].upper())
        offset = match.end()
        if len(tokens) > 64:
            raise GuardError('公式超过64个语法单元，请拆分计算')
    cursor = 0

    def take():
        nonlocal cursor
        if cursor >= len(tokens):
            raise GuardError('公式不完整，请核对括号和运算数')
        token = tokens[cursor]
        cursor += 1
        return token

    def expression(level=0, minimum=0):
        nonlocal cursor
        if level > 16:
            raise GuardError('公式嵌套超过16层')
        token = take()
        if token in {'+', '-'}:
            left = ('unary', token, expression(level + 1, 4))
        elif token == '(':
            left = expression(level + 1)
            if take() != ')':
                raise GuardError('公式括号不匹配')
        elif CELL.fullmatch(token):
            left = ('ref', location(token)[0])
        elif token.startswith('"'):
            left = ('text', token[1:-1].replace('""', '"'))
        elif token in {'SUM', 'AVERAGE', 'MIN', 'MAX', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'COUNT', 'PRODUCT', 'INT', 'MOD', 'POWER', 'SQRT', 'DATE', 'YEAR', 'MONTH', 'DAY', 'IF', 'AND', 'OR', 'NOT', 'MEDIAN', 'STDEV', 'STDEVP', 'SUMPRODUCT', 'LEN', 'TRIM', 'CONCATENATE', 'NUMBERVALUE'}:
            if take() != '(':
                raise GuardError('函数需要括号')
            args = []
            while True:
                arg = expression(level + 1)
                if cursor < len(tokens) and tokens[cursor] == ':':
                    cursor += 1
                    end = take()
                    if arg[0] != 'ref' or not CELL.fullmatch(end):
                        raise GuardError('范围必须由两个单元格引用组成')
                    start, r1, c1 = location(arg[1]); last, r2, c2 = location(end)
                    if r2 < r1 or c2 < c1 or (r2-r1+1)*(c2-c1+1) > 256:
                        raise GuardError('公式范围需正向且不超过256格')
                    arg = ('range', r1, c1, r2, c2)
                args.append(arg)
                separator = take()
                if separator == ')':
                    break
                if separator != ',':
                    raise GuardError('函数参数请使用英文逗号')
            if token in {'ABS', 'INT', 'SQRT', 'YEAR', 'MONTH', 'DAY'} and (len(args) != 1 or args[0][0] == 'range'):
                raise GuardError(token+'只接受一个数值')
            if token in {'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'MOD', 'POWER'} and (len(args) != 2 or any(arg[0] == 'range' for arg in args)):
                raise GuardError(token+'只接受两个数值')
            if token == 'DATE' and (len(args) != 3 or any(arg[0] == 'range' for arg in args)):
                raise GuardError('DATE只接受年、月、日三个数值')
            if token == 'IF' and (len(args) != 3 or any(arg[0] == 'range' for arg in args)):
                raise GuardError('IF只接受条件、真值和假值三个参数')
            if token in {'AND', 'OR'} and len(args) < 1:
                raise GuardError(token+'至少需要一个条件')
            if token == 'NOT' and (len(args) != 1 or args[0][0] == 'range'):
                raise GuardError('NOT只接受一个条件')
            if token in {'MEDIAN', 'STDEV', 'STDEVP'} and len(args) < 1:
                raise GuardError(token+'至少需要一个数值')
            if token == 'SUMPRODUCT' and len(args) < 1:
                raise GuardError('SUMPRODUCT至少需要一个区域')
            if token in {'LEN', 'TRIM'} and (len(args) != 1 or args[0][0] == 'range'):
                raise GuardError(token+'只接受一个文本参数')
            if token == 'CONCATENATE' and any(arg[0] == 'range' for arg in args):
                raise GuardError('CONCATENATE仅支持逐项文本参数')
            if token == 'NUMBERVALUE' and (len(args) != 3 or any(arg[0] == 'range' for arg in args)):
                raise GuardError('NUMBERVALUE需要文本、明确小数点和分组符三个参数')
            left = ('call', token, args)
        else:
            try:
                number = Decimal(token)
                if not number.is_finite() or len(number.as_tuple().digits) > 15 or abs(number.adjusted()) > 300:
                    raise ValueError
                left = ('number', number)
            except (DecimalException, ValueError):
                raise GuardError('公式常数应为有限十进制且不超过15位有效数字') from None
        precedence = {'=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '+': 2, '-': 2, '*': 3, '/': 3}
        while cursor < len(tokens) and precedence.get(tokens[cursor], 0) > minimum:
            operator = take()
            left = ('binary', operator, left, expression(level + 1, precedence[operator]))
        return left

    result = expression()
    if cursor != len(tokens):
        raise GuardError('公式含未支持语法或多余内容')
    return result


def references(tree):
    result = set()

    def walk(node):
        kind = node[0]
        if kind == 'ref':
            result.add(node[1])
        elif kind == 'range':
            if (node[3]-node[1]+1)*(node[4]-node[2]+1) > 256:
                raise GuardError('公式范围需正向且不超过256格')
            for row in range(node[1], node[3]+1):
                for col in range(node[2], node[4]+1):
                    result.add(name(row, col))
        elif kind == 'unary':
            walk(node[2])
        elif kind == 'binary':
            walk(node[2]); walk(node[3])
        elif kind == 'call':
            for arg in node[2]:
                walk(arg)

    walk(tree)
    return result


class Calculator:
    def __init__(self, sheet, cache=None, date1904=False, namespace=None, shared=()):
        self.cells = {node.get('r'): node for node in sheet.findall('./'+S+'sheetData/'+S+'row/'+S+'c')}
        self.remaining = 2048
        self.cache = cache if isinstance(cache, dict) else {}
        self._graph = None
        self.date1904 = bool(date1904)
        self.namespace = namespace
        self.shared = shared

    def results(self, refs=None):
        targets = [ref for ref, cell in self.cells.items() if cell.find(S+'f') is not None and (refs is None or ref in refs)]
        if len(targets) > 512:
            raise GuardError('公式依赖图最多512个公式，请拆分工作簿')
        return {ref: self.describe(ref) for ref in targets}

    def graph_report(self):
        if self._graph is None:
            formulas = [ref for ref, cell in self.cells.items() if cell.find(S+'f') is not None]
            if len(formulas) > 512:
                raise GuardError('公式依赖图最多512个公式，请拆分工作簿')
            edges = {}
            for ref in formulas:
                try:
                    tree = parse('='+str(self.cells[ref].findtext(S+'f') or ''))
                    edges[ref] = sorted(references(tree))
                except GuardError:
                    edges[ref] = []
            visiting, visited, order = set(), set(), []
            def visit(ref):
                if ref in visiting:
                    raise GuardError('公式依赖图存在循环引用')
                if ref in visited:
                    return
                visiting.add(ref)
                for child in edges.get(ref, []):
                    if child in formulas:
                        visit(child)
                visiting.remove(ref); visited.add(ref); order.append(ref)
            for ref in formulas:
                visit(ref)
            self._graph = {'formulas': len(formulas), 'edges': edges, 'topological': order, 'limited': False}
        return self._graph

    def calculate(self, target, formula=None):
        result = self.describe(target, formula)
        if result['status'] != 'calculated':
            raise GuardError(result['reason'])
        return result

    def describe(self, target, formula=None):
        try:
            return self._calculate(target, formula)
        except GuardError as error:
            message = str(error)
            code = 'calculation'
            for token, candidate in (('循环', 'cycle'), ('预算', 'budget'), ('深度', 'budget'),
                                     ('空白', 'reference'), ('文本', 'reference'), ('缺失', 'reference'),
                                     ('除零', 'division'), ('共享', 'unsupported_formula'), ('数组', 'unsupported_formula'),
                                     ('特殊', 'unsupported_formula'), ('不支持', 'unsupported_formula'), ('仅支持', 'unsupported_formula'),
                                     ('受限函数', 'unsupported_formula'), ('公式常数', 'unsupported_formula'),
                                     ('范围', 'range'), ('语法', 'syntax'), ('括号', 'syntax')):
                if token in message:
                    code = candidate
                    break
            return {'status': 'unavailable', 'code': code, 'reason': message}

    def _calculate(self, target, formula=None):
        active = set()
        reads = [-1]
        refs = set()

        def text_value(value):
            if not isinstance(value, str):
                raise GuardError('文本函数需要明确文本，不自动转换数字')
            if len(value) > 4000:
                raise GuardError('文本计算超过4000字符预算')
            return value

        def numeric(value):
            if not isinstance(value, Decimal):
                raise GuardError('数值运算引用含文本；请明确使用NUMBERVALUE，不自动按0计算')
            return value

        def resolve(ref):
            self.remaining -= 1
            reads[0] += 1
            refs.add(ref)
            if self.remaining < 0 or reads[0] > 256 or len(active) >= 16:
                raise GuardError('本机计算超过256次引用/16层依赖或本次计算预算')
            if ref in active:
                raise GuardError('公式存在循环引用')
            key = (self.namespace, ref) if self.namespace is not None else ref
            if formula is None and key in self.cache:
                return self.cache[key]
            cell = self.cells.get(ref)
            if formula is not None and ref == target:
                tree = parse(formula)
            elif cell is not None and cell.find(S+'f') is not None:
                source = cell.find(S+'f')
                if source.attrib:
                    raise GuardError('共享、数组或带特殊属性的公式不在本机计算范围')
                tree = parse('='+str(source.text or ''))
            else:
                if cell is None:
                    raise GuardError('引用含空白、文本或缺失值，不按0计算')
                if cell.get('t') in {'s', 'inlineStr'}:
                    if cell.get('t') == 's':
                        index = cell.findtext(S+'v', '')
                        if not index.isascii() or not index.isdigit() or int(index) >= len(self.shared):
                            raise GuardError('文本共享索引缺失或无效')
                        source = self.shared[int(index)]
                    else:
                        source = cell.find(S+'is')
                    if source is None or len(source) != 1 or source[0].tag != S+'t':
                        raise GuardError('富文本或特殊文本引用不在计算范围')
                    return text_value(source[0].text or '')
                if cell.get('t', 'n') != 'n' or not cell.findtext(S+'v'):
                    raise GuardError('引用含空白、文本或缺失值，不按0计算')
                value = Decimal(cell.findtext(S+'v'))
                if not value.is_finite() or abs(value.adjusted()) > 300:
                    raise GuardError('引用数值超出有限计算范围')
                return value
            active.add(ref)
            try:
                result = evaluate(tree)
            finally:
                active.remove(ref)
            if isinstance(result, str):
                text_value(result)
            elif not result.is_finite() or abs(result.adjusted()) > 300:
                raise GuardError('计算结果超过有限精度范围')
            if formula is None:
                self.cache[key] = result
            return result

        def evaluate(tree):
            kind = tree[0]
            if kind in {'number', 'text'}:
                return tree[1]
            if kind == 'ref':
                return resolve(tree[1])
            if kind == 'unary':
                value = numeric(evaluate(tree[2]))
                return -value if tree[1] == '-' else value
            if kind == 'binary':
                left, right = numeric(evaluate(tree[2])), numeric(evaluate(tree[3]))
                if tree[1] == '+': return left + right
                if tree[1] == '-': return left - right
                if tree[1] == '*': return left * right
                if tree[1] == '/': return left / right
                result = {'=': left == right, '<>': left != right, '<': left < right,
                          '>': left > right, '<=': left <= right, '>=': left >= right}[tree[1]]
                return Decimal(1 if result else 0)
            if tree[1] == 'IF':
                condition = numeric(evaluate(tree[2][0]))
                return evaluate(tree[2][1] if condition != 0 else tree[2][2])
            if tree[1] in {'LEN', 'TRIM', 'CONCATENATE', 'NUMBERVALUE'}:
                texts = [text_value(evaluate(arg)) for arg in tree[2]]
                if tree[1] == 'LEN':
                    if any(ord(char) > 0xffff for char in texts[0]):
                        raise GuardError('LEN的非BMP字符计数受Excel兼容版本影响，暂不支持')
                    return Decimal(len(texts[0]))
                if tree[1] == 'TRIM':
                    return ' '.join(part for part in texts[0].split(' ') if part)
                if tree[1] == 'CONCATENATE':
                    return text_value(''.join(texts))
                value, point, group = texts
                if point not in {'.', ','} or group not in {'.', ',', ' '} or point == group:
                    raise GuardError('NUMBERVALUE仅支持明确且不同的点/逗号小数点与点/逗号/空格分组符')
                value = value.replace(' ', '')
                percent = len(value) - len(value.rstrip('%'))
                value = value.rstrip('%')
                if not value or percent > 2 or value.count(point) > 1:
                    raise GuardError('NUMBERVALUE文本为空或超出受限数值格式')
                parts = value.split(point)
                if len(parts) == 2 and group in parts[1]:
                    raise GuardError('NUMBERVALUE小数部分不能含分组符')
                normalized = parts[0].replace(group, '') + ('.'+parts[1] if len(parts) == 2 else '')
                if not re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)', normalized):
                    raise GuardError('NUMBERVALUE文本不是受支持的区域小数')
                number = Decimal(normalized)
                if len(number.as_tuple().digits) > 15 or abs(number.adjusted()) > 300:
                    raise GuardError('NUMBERVALUE超过15位有效数字或有限范围')
                return number / (Decimal(100) ** percent)
            def vectors(arg):
                if arg[0] == 'range':
                    return [numeric(resolve(name(row, col))) for row in range(arg[1], arg[3]+1) for col in range(arg[2], arg[4]+1)]
                return [numeric(evaluate(arg))]
            if tree[1] == 'SUMPRODUCT':
                shapes = [(arg[3]-arg[1]+1, arg[4]-arg[2]+1) if arg[0] == 'range' else (1, 1) for arg in tree[2]]
                if any(shape != shapes[0] for shape in shapes):
                    raise GuardError('SUMPRODUCT区域行列形状必须一致')
                arrays = [vectors(arg) for arg in tree[2]]
                length = len(arrays[0])
                if any(len(array) != length for array in arrays):
                    raise GuardError('SUMPRODUCT参数长度必须一致')
                return sum((reduce(lambda a,b:a*b, [array[index] for array in arrays], Decimal(1)) for index in range(length)), Decimal(0))
            values = []
            for arg in tree[2]:
                if arg[0] == 'range':
                    values.extend(vectors(arg))
                else:
                    values.append(numeric(evaluate(arg)))
            if tree[1] == 'AND': return Decimal(1 if all(value != 0 for value in values) else 0)
            if tree[1] == 'OR': return Decimal(1 if any(value != 0 for value in values) else 0)
            if tree[1] == 'NOT': return Decimal(0 if values[0] != 0 else 1)
            if tree[1] == 'MEDIAN':
                ordered = sorted(values)
                middle = len(ordered) // 2
                return ordered[middle] if len(ordered) % 2 else (ordered[middle-1]+ordered[middle]) / 2
            if tree[1] in {'STDEV', 'STDEVP'}:
                count = len(values)
                if count < (2 if tree[1] == 'STDEV' else 1):
                    raise GuardError(tree[1]+'样本数量不足')
                mean = sum(values, Decimal(0)) / count
                variance = sum((value-mean)*(value-mean) for value in values) / (count-1 if tree[1] == 'STDEV' else count)
                return variance.sqrt()
            if tree[1] == 'MIN': return min(values)
            if tree[1] == 'MAX': return max(values)
            if tree[1] == 'ABS': return abs(values[0])
            if tree[1] == 'INT': return values[0].to_integral_value(rounding=ROUND_FLOOR)
            if tree[1] == 'SQRT':
                if values[0] < 0:
                    raise GuardError('SQRT不接受负数')
                return values[0].sqrt()
            if tree[1] == 'DATE':
                parts = []
                for value in values:
                    if value != value.to_integral_value():
                        raise GuardError('DATE只接受整数年月日')
                    parts.append(int(value))
                year, month, day = parts
                try:
                    moment = date(year, month, day)
                except ValueError:
                    raise GuardError('DATE年月日无效') from None
                base = date(1904, 1, 1) if self.date1904 else date(1899, 12, 31)
                serial = (moment-base).days
                if not self.date1904 and moment >= date(1900, 3, 1):
                    serial += 1
                return Decimal(serial)
            if tree[1] in {'YEAR', 'MONTH', 'DAY'}:
                serial = values[0]
                if serial < 0:
                    raise GuardError('日期序列不能为负')
                day = int(serial)
                if not self.date1904 and day == 60:
                    raise GuardError('1900闰日序列没有有效日期')
                base = date(1904, 1, 1) if self.date1904 else date(1899, 12, 31) if day < 60 else date(1899, 12, 30)
                try:
                    moment = base + timedelta(days=day)
                except OverflowError:
                    raise GuardError('日期序列超出范围') from None
                return Decimal(getattr(moment, tree[1].lower()))
            if tree[1] == 'MOD':
                if values[1] == 0:
                    raise GuardError('MOD除数不能为0')
                remainder = values[0] % values[1]
                return remainder + values[1] if remainder and (remainder < 0) != (values[1] < 0) else remainder
            if tree[1] == 'POWER':
                if values[1] != values[1].to_integral_value() or abs(values[1]) > 100:
                    raise GuardError('POWER指数必须为绝对不超过100的整数')
                return values[0] ** int(values[1])
            if tree[1] == 'COUNT': return Decimal(len(values))
            if tree[1] == 'PRODUCT':
                result = Decimal(1)
                for value in values:
                    result *= value
                return result
            if tree[1] in {'ROUND', 'ROUNDUP', 'ROUNDDOWN'}:
                digits = values[1]
                if digits != digits.to_integral_value() or abs(digits) > 15:
                    raise GuardError(tree[1]+'位数必须是绝对不超过15的整数')
                quantum = Decimal(1).scaleb(-int(digits))
                rounding = ROUND_HALF_UP if tree[1] == 'ROUND' else ROUND_UP if tree[1] == 'ROUNDUP' else ROUND_DOWN
                return values[0].quantize(quantum, rounding=rounding)
            total = sum(values, Decimal(0))
            return total / len(values) if tree[1] == 'AVERAGE' else total

        try:
            with localcontext() as context:
                context.prec = 28
                context.rounding = ROUND_HALF_EVEN
                result = resolve(target)
                text_result = isinstance(result, str)
                value = result if text_result else format(result.normalize(), 'f') if result else '0'
                source = formula if formula is not None else '='+str(self.cells.get(target).findtext(S+'f') or '') if target in self.cells and self.cells[target].find(S+'f') is not None else None
                dependencies = references(parse(source)) if source is not None else set()
                return {'status': 'calculated', 'value': value, 'precision': 28, 'rounding': str(context.rounding),
                        'value_type': 'text' if text_result else 'number',
                        'references': reads[0], 'depends_on': sorted(dependencies), 'dependency_scope': 'direct syntactic references', 'engine': 'bounded-decimal-v1'}
        except DecimalException:
            raise GuardError('除零或十进制计算失败；未采用公式') from None
