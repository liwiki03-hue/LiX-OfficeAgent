import re
from types import SimpleNamespace

from ..paged_documents import GuardError
from ..professional_edits import NUMBER, _protect
from ..rag import SECRET


YEAR_COMMAND = re.compile(r'(?:请)?(?:只|仅)?(?:把|将)?(?:选中句|选中的句子|选中段落|所选段落|当前选区|这句话)?(?:中)?(?:的)?(?:年份|(?P<before>(?:19|20)\d{2})年?)\s*(?:修改为|更改为|改为|改成|替换为)\s*(?P<after>(?:19|20)\d{2})年?')
YEAR = re.compile(r'(?<![\d.])(?:19|20)\d{2}(?=年)')
YEAR_CONDITIONAL = re.compile(r'如果|假如|假设|若|是否|能否|可否|例如|比如|示例|举例')
YEAR_VETO = re.compile(r'(?:不要|不得|禁止|不能|不许|暂不|无需|不必).{0,8}(?:年份|(?:19|20)\d{2}年)|(?:保留|保持)(?:原有|原|当前)?年份(?:不变)?')


class DraftNumberConflict(GuardError):
    pass


def validate_suggestion(before, after, reason, requirements):
    if SECRET.search(reason):
        raise GuardError('修改原因疑似包含凭据，未形成建议')
    policy = SimpleNamespace(preserve_numbers=False, preserve_citations=True, preserve_formulas=True,
                             protected_literals=[unit for unit in ('℃','℉','°C','°F') if unit in before])
    _protect(before, after, policy)
    years = list(YEAR.finditer(before))
    commands = []
    year_protected = any(YEAR_VETO.search(row['text']) for row in requirements)
    for row in requirements:
        if year_protected or YEAR_CONDITIONAL.search(row['text']):
            continue
        for clause in re.split(r'[，,。；;\n]', row['text']):
            match = YEAR_COMMAND.fullmatch(clause.strip())
            if match:
                commands.append((match['before'], match['after']))
    expected = before
    changes = []
    if len(years) == 1 and commands:
        original = years[0].group()
        targets = {new for old, new in commands if old is None or old == original}
        if len(targets) > 1:
            raise DraftNumberConflict('年份要求存在冲突，请保留一个目标年份后重新处理。')
        if targets:
            replacement = next(iter(targets))
            if replacement != original:
                expected = before[:years[0].start()] + replacement + before[years[0].end():]
                changes = [{'kind': 'year', 'before': original, 'after': replacement}]
    policy.preserve_numbers = True
    expected_numbers, proposed_numbers = NUMBER.findall(expected), NUMBER.findall(after)
    if expected_numbers != proposed_numbers and sorted(expected_numbers) == sorted(proposed_numbers):
        raise DraftNumberConflict('这次改写调换了受保护数字的顺序。可以保持原来的数字顺序，只调整文字吗？')
    _protect(expected, after, policy)
    if before == after:
        raise GuardError('建议没有实际变化')
    return changes


def model_failure(error):
    status = getattr(error, 'status_code', None)
    name = type(error).__name__.lower()
    if status in (401, 403):
        return 'authentication_failed', '接口鉴权失败，请检查当前连接配置后主动处理。'
    if status == 404:
        return 'model_not_found', '模型或接口地址不存在，请核对后主动处理。'
    if status == 429:
        return 'rate_limited', '接口限流，请稍后主动处理。'
    if status in (301, 302, 303, 307, 308):
        return 'redirect_refused', '接口返回重定向，请核对最终地址。'
    if isinstance(status, int) and status >= 500:
        return 'service_failed', '模型服务异常，请核对服务状态。'
    if 'timeout' in name:
        return 'timeout', '等待超时，服务端可能已受理，请先核对。'
    if 'connect' in name:
        return 'address_unreachable', '接口不可达，请检查地址和网络。'
    return 'response_incompatible', '模型响应未完成或格式不兼容，请核对接口。'


def validate_append(text, reason):
    from ..model_content import check_scope
    check_scope(text, reason)
    if (not text.strip() or len(text) > 6000 or len(text.splitlines()) > 20
            or any(ord(char) < 32 and char not in '\r\n' for char in text)):
        raise GuardError('追加内容限6000字、20段普通正文，不能包含控制字符')
    return []


class DraftModel:
    def __init__(self, model, failures):
        self.model = model
        self.failures = failures

    def bind_tools(self, tools):
        return DraftModel(self.model.bind_tools(tools), self.failures)

    def __getattr__(self, name):
        value = getattr(self.model, name)
        if name == 'invoke' and callable(value):
            def invoke(*args, **kwargs):
                try:
                    return value(*args, **kwargs)
                except Exception as error:
                    self.failures.append(model_failure(error))
                    raise ValueError(self.failures[-1][1]) from None
            return invoke
        if name == 'stream' and callable(value):
            def stream(*args, **kwargs):
                iterator = None
                try:
                    iterator = iter(value(*args, **kwargs))
                    yield from iterator
                except Exception as error:
                    self.failures.append(model_failure(error))
                    raise ValueError(self.failures[-1][1]) from None
                finally:
                    if callable(getattr(iterator, 'close', None)):
                        try:
                            iterator.close()
                        except Exception as error:
                            self.failures.append(model_failure(error))
                            raise ValueError(self.failures[-1][1]) from None
            return stream
        return value
