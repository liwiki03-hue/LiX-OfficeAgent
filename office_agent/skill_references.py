import hashlib
import re
from pathlib import PurePosixPath
from urllib.parse import urlsplit


MAX_FILES = 32
MAX_SOURCE_BYTES = 2 * 1024 * 1024
MAX_REFERENCE_BYTES = 16384
MAX_DEPTH = 3
MAX_PROMPT_CHARS = 7800
RESERVED = {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(10)), *(f'LPT{i}' for i in range(10))}
LINK = re.compile(r'(!?)\[[^\]\n]*\]\(([^)\n]+)\)')


def safe_name(name):
    if not isinstance(name, str) or not name or len(name) > 300 or len(name.split('/')) > 8:
        raise ValueError('unsupported_reference_path')
    for part in name.split('/'):
        if not re.fullmatch(r'[A-Za-z0-9_-][A-Za-z0-9_.-]*', part) or part.endswith('.') or part.split('.')[0].upper() in RESERVED:
            raise ValueError('unsupported_reference_path')
    return name


def validate_resources(resources):
    if not isinstance(resources, dict) or len(resources) > MAX_FILES - 1:
        raise ValueError('package_budget')
    seen = set()
    for name, item in resources.items():
        safe_name(name)
        if PurePosixPath(name).suffix.lower() not in ('.md', '.txt') or name.lower() in seen or name.lower() == 'skill.md':
            raise ValueError('unsupported_reference_path')
        seen.add(name.lower())
        if not isinstance(item, dict) or set(item) != {'text', 'sha256'} or not isinstance(item['text'], str):
            raise ValueError('source_integrity')
        data = item['text'].encode('utf-8')
        if len(data) > MAX_REFERENCE_BYTES or '\x00' in item['text'] or any(ord(c) < 32 and c not in '\n\r\t' for c in item['text']):
            raise ValueError('reference_budget')
        if hashlib.sha256(data).hexdigest() != item['sha256']:
            raise ValueError('source_integrity')


def validate_links(body, resources):
    texts = {'SKILL.md': body, **{name: item['text'] for name, item in resources.items()}}
    graph = {}
    for name, text in texts.items():
        targets = []
        for image, target in LINK.findall(text):
            if image:
                raise ValueError('unsupported_reference_link')
            if target.startswith('#'):
                continue
            parsed = urlsplit(target)
            if parsed.scheme in ('http', 'https') and parsed.netloc and not parsed.username and not parsed.password:
                continue
            if parsed.scheme or parsed.netloc or parsed.query or ' ' in target:
                raise ValueError('unsupported_reference_link')
            if not parsed.path:
                raise ValueError('unsupported_reference_link')
            safe_name(parsed.path)
            resolved = str(PurePosixPath(name).parent / parsed.path)
            if resolved not in texts:
                raise ValueError('reference_missing')
            targets.append(resolved)
        if re.search(r'^\s*\[[^\]]+\]:|\]\[[^\]]*\]', text, re.MULTILINE):
            raise ValueError('unsupported_reference_link')
        graph[name] = list(dict.fromkeys(targets))
    def visit(name, chain):
        if name in chain:
            raise ValueError('reference_cycle')
        if len(chain) > MAX_DEPTH:
            raise ValueError('reference_depth')
        for target in graph[name]:
            visit(target, (*chain, name))
    for name in graph:
        visit(name, ())
    return graph


def expand(body, resources):
    validate_resources(resources)
    result = body + ''.join('\n\n包内参考：' + name + '\n' + resources[name]['text'] for name in sorted(resources))
    if len(result) > MAX_PROMPT_CHARS:
        raise ValueError('body_budget')
    validate_links(body, resources)
    return result


def select_references(body, resources, goal, budget=MAX_PROMPT_CHARS):
    from .skill_context import _terms
    validate_resources(resources)
    graph = validate_links(body, resources)
    names = sorted(resources)
    index = '\n\n包内参考目录（未列为已读取的内容尚未读取）：\n'+'\n'.join(names)
    note = '\n仅使用已读取资料；若所需文件未加载，请明确请求其名称；不要推断缺失内容。'
    base = body+index+note
    if len(base)>budget:
        raise ValueError('body_budget')
    terms = _terms(goal) if isinstance(goal,str) else set()
    ranked = sorted(names,key=lambda name:(-(len(terms&_terms(name+' '+resources[name]['text'].split('\n',1)[0])) + (100 if name in goal else 0)),name))
    selected, sections = set(), []
    def dependencies(name, found):
        if name in found:
            return
        found.add(name)
        for target in graph[name]:
            if target!='SKILL.md':
                dependencies(target,found)
    for name in ranked:
        if not (terms&_terms(name+' '+resources[name]['text'].split('\n',1)[0]) or name in goal):
            continue
        required = set()
        dependencies(name,required)
        additions = [(key,'\n\n已读取包内参考：'+key+'\n'+resources[key]['text']) for key in sorted(required-selected)]
        if len(base)+sum(len(text) for text in sections)+sum(len(text) for _,text in additions)>budget:
            continue
        selected.update(required)
        sections.extend(text for _,text in additions)
    return {'body':base+''.join(sections), 'selected_references':sorted(selected),
            'unloaded_references':[name for name in names if name not in selected],
            'reference_sha256':{name:resources[name]['sha256'] for name in sorted(selected)}}
