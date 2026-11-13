from collections import Counter
import hashlib
import json
import re
from typing import Literal

from pydantic import Field

from .rag import SECRET
from .writes import GuardError, Strict


class ProfileTaskRequest(Strict):
    request: str = Field(min_length=1, max_length=2000)
    editing_profile: Literal['general', 'technical_report_v1'] = 'general'
    protected_literals: list[str] = Field(default_factory=list, max_length=30)


PROFILE = 'technical_report_v1'
RULES = {
    'quantity': r'(?<![\d.A-Za-zµμ])[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?\s*(?:mmol|nmol|mol|mL|µL|μL|mg|µg|μg|kg|mM|µM|μM|ppm|ppb|kPa|MPa|Pa|kJ|mW|mV|mA|nm|µm|μm|mm|cm|km|ms|min|kHz|MHz|Hz|GB|MB|°C|L|g|J|W|V|A|m|s|h|K|%)(?:\s*[/·]\s*(?:mol|min|kg|mL|cm|mm|g|L|m|s|h)(?:\^?-?[1-3]|[²³])?){0,3}(?![A-Za-zµμ])',
    'condition': r'(?:不适用于|不得用于|仅限|仅在|只在|适用于)[\u3400-\u9fff]{1,16}|在[\u3400-\u9fffA-Za-z0-9 °%./_-]{1,24}条件下|\b(?:only under|only in|not applicable to|applicable to)\s+[^.;,\n]{1,60}',
    'qualifier': r'不得少于|不超过|不低于|不高于|无显著|不应|不得|不可|不能|未能|尚未|禁止|必须|应当|至少|至多|\b(?:not|no|never|without|must|shall|only|at least|at most)\b',
}
RULE_SHA256 = hashlib.sha256(json.dumps({'profile': PROFILE, 'flags': 'IGNORECASE', 'rules': RULES}, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()
PATTERNS = {kind: re.compile(pattern, re.IGNORECASE) for kind, pattern in RULES.items()}


def identity(profile, fingerprint):
    if profile == 'general':
        if fingerprint is not None:
            raise GuardError('professional_profile_version')
        return
    if profile != PROFILE or fingerprint != RULE_SHA256:
        raise GuardError('professional_profile_version')


def observations(text):
    if SECRET.search(text) or len(text) > 8000:
        raise GuardError('professional_sensitive_or_control_text')
    result = [(kind, match.group()) for kind, pattern in PATTERNS.items() for match in pattern.finditer(text)]
    if len(set(result)) > 30:
        raise GuardError('professional_profile_capacity')
    return result


def report(texts, literals):
    records = Counter(item for text in texts for item in observations(text))
    if len(records) > 30:
        raise GuardError('professional_profile_capacity')
    verify_literals(texts, literals)
    return {'profile': PROFILE, 'rule_sha256': RULE_SHA256,
            'items': [{'kind': kind, 'literal': literal, 'count': count} for (kind, literal), count in sorted(records.items())],
            'manual_literals': list(literals), 'scope': 'recognized exact quantities, conditions and qualifiers only; no unit conversion, semantic equivalence or fact verification'}


def verify_literals(texts, literals):
    if any(not literal or len(literal) > 200 or not any(literal in text for text in texts) for literal in literals):
        raise GuardError('professional_profile_literal_missing')


def verify(before, after, profile, fingerprint):
    identity(profile, fingerprint)
    if profile == 'general':
        return
    left, right = observations(before), observations(after)
    for kind in RULES:
        if Counter(value for key, value in left if key == kind) != Counter(value for key, value in right if key == kind):
            raise GuardError('professional_profile_changed_' + kind)
