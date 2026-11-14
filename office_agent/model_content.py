import re

from .rag import SECRET
from .writes import GuardError


IDENTITY = re.compile(r'(?:身份证(?:号|号码)?|居民身份号码)\s*[:：]?\s*[1-9][0-9]{16}[0-9Xx]|(?:手机|电话)(?:号码|号)?\s*[:：]\s*1[3-9][0-9]{9}')


def check_scope(*texts):
    if any(not isinstance(text, str) for text in texts):
        raise GuardError('发送范围必须是明确文字')
    if any(SECRET.search(text) or IDENTITY.search(text) for text in texts):
        raise GuardError('发送范围疑似包含凭据或身份信息，请移除敏感内容或缩小引用')
