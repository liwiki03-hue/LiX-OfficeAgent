import json
import re

from .plugin_oauth import PluginError


URL = 'https://mcp.openalex.org/mcp'
TOOL = 'search_works'
COUNTS = {'一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5}
COUNT = re.compile(r'(\d+|[一二两三四五])\s*篇')
PARAMS = {'query', 'limit', 'include_abstracts'}


class SearchRejected(PluginError):
    pass


def supported(raw):
    schema = raw.get('inputSchema') or {}
    fields = schema.get('properties') or {}
    hints = raw.get('annotations') or {}
    if (raw.get('name') != TOOL or schema.get('type') != 'object' or not isinstance(fields, dict)
            or set(schema.get('required') or ()) - PARAMS or not PARAMS <= fields.keys()
            or hints.get('readOnlyHint') is not True or hints.get('destructiveHint') is not False):
        return False
    query, limit, abstracts = (fields[key] for key in ('query', 'limit', 'include_abstracts'))
    if not all(isinstance(field, dict) for field in (query, limit, abstracts)):
        return False
    return (query.get('type') == 'string' and query.get('maxLength') == 2000
            and limit.get('type') == 'integer' and limit.get('minimum') == 1 and limit.get('maximum') == 50
            and abstracts.get('type') == 'boolean')


def preview(question):
    count = COUNT.search(question)
    limit = 5
    if count:
        token = count[1]
        limit = COUNTS[token] if token in COUNTS else int(token)
    if not 1 <= limit <= 5:
        raise PluginError('每次可查 1 至 5 篇，请调整篇数。')
    topic = re.sub(r'^(?:请|帮我|请帮我)*\s*(?:(?:使用|用)\s*OpenAlex\s*)?(?:查找|查询|检索|搜索|查|找)\s*', '', question.strip(), flags=re.I)
    topic = re.sub(r'^\s*(?:\d+|[一二两三四五])\s*篇\s*', '', topic)
    topic = re.sub(r'^(?:有关|关于)\s*', '', topic)
    topic = re.split(r'[，,；;]\s*(?:列|给|附|带|提供|显示|含|包括|返回)', topic, maxsplit=1)[0]
    topic = re.sub(r'(?:的(?:相关)?(?:论文|文献|研究)|相关(?:论文|文献))\s*[。？！?!]?$', '', topic).strip(' 。？！?!「」“”')
    if not topic or len(topic) > 300 or re.search(r'选区|选中|这段|所选|当前文档|联系方式|\n|\r', topic):
        raise PluginError('请在问题中写出公开的论文检索词，例如“找两篇关于 BM25 的研究”。')
    return json.dumps({'query': topic, 'limit': limit, 'include_abstracts': False}, ensure_ascii=False)


def arguments(text):
    try:
        args = json.loads(text)
    except (TypeError, ValueError):
        args = None
    if (not isinstance(args, dict) or set(args) != PARAMS or not isinstance(args['query'], str)
            or not 1 <= len(args['query'].strip()) <= 300 or type(args['limit']) is not int
            or not 1 <= args['limit'] <= 5 or args['include_abstracts'] is not False):
        raise PluginError('OpenAlex 的检索字段已改变，请重新提问。')
    return args


def output(response, limit):
    if response.isError:
        raise SearchRejected('OpenAlex 返回检索错误，请检查登录、查询词或账户可用额度后重试。')
    parts = response.content
    if len(parts) != 1 or parts[0].type != 'text':
        raise SearchRejected('OpenAlex 返回格式已改变，请重新检查工具。')
    try:
        payload = json.loads(parts[0].text)
    except (ValueError, RecursionError):
        payload = None
    if not isinstance(payload, dict) or not isinstance(payload.get('results'), list):
        raise SearchRejected('OpenAlex 返回的资料缺少可核对的文献记录。')
    works = []
    seen = set()
    for row in payload['results']:
        if not isinstance(row, dict):
            continue
        identifier, title = row.get('id'), row.get('title')
        if (not isinstance(identifier, str) or not re.fullmatch(r'W\d{1,16}', identifier)
                or identifier in seen or not isinstance(title, str) or not title.strip()):
            continue
        seen.add(identifier)
        url = 'https://openalex.org/' + identifier
        doi = row.get('doi')
        authors = row.get('authors')
        names = [a.strip()[:180] for a in authors[:6] if isinstance(a, str) and a.strip()] if isinstance(authors, list) else []
        year = row.get('year')
        lines = [title.strip()[:300], '、'.join(names) or '作者未提供', str(year) if type(year) is int and 1000 <= year <= 9999 else '年份未提供']
        if isinstance(doi, str) and re.fullmatch(r'10\.\d{4,9}/[^\s<>?#]{1,200}', doi):
            lines.append('DOI ' + doi)
        lines.append(identifier)
        if row.get('is_retracted') or row.get('is_pulled'):
            lines.append('来源标记此记录已撤回或移除')
        works.append((' · '.join(lines), {'title': title.strip()[:160], 'url': url}))
        if len(works) == limit:
            break
    if payload['results'] and not works:
        raise SearchRejected('OpenAlex 返回的文献均缺少有效题名或标识，请核对工具后重试。')
    return {'text': '\n\n'.join(line for line, _ in works) or 'OpenAlex 本次未查到匹配文献。',
            'sources': [source for _, source in works]}
