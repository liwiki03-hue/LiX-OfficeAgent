import re
import time
import unicodedata

from ..literature import LiteratureQuery, doi, group_records
from .turn_scope import TTL, mask_text


_DOI = re.compile(r'(?<![a-zA-Z0-9])10\.[0-9]{4,9}/[a-zA-Z0-9._():;/+-]+', re.I)
_TITLE = re.compile(r'《([^《》\n]{2,300})》|[“"]([^“”"\n]{2,300})[”"]|(?:题名|标题|篇名|title)\s*[:：]\s*(.+?)(?=[\n；;]|[,，]\s*(?:作者|年份|doi)|$)', re.I)
_YEAR = re.compile(r'(?<!\d)((?:18|19|20)\d{2})(?!\d)')
_AUTHORS = re.compile(r'(?:作者|authors?)\s*[:：]\s*([^;；\n]+?)(?=[;；\n]|[,，]\s*(?:题名|标题|年份|doi)|$)', re.I)
_CLAIM = re.compile(r'(?:它|该(?:文|研究)|这(?:一)?(?:篇|条|项)|上述(?:文献|论文)).{0,32}(?:证明|支持|证实|佐证)|(?:证明|支持|证实|佐证).{0,16}(?:我的|本文|这段|选段).{0,8}(?:观点|结论|论点)')
_LOOKUP_ONLY = '仅取得了书目信息。判断这篇文献能否支持该结论，还需要核对相应原文、研究条件和具体页码。'
_TARGET = ('crossref', 'https://api.crossref.org')


def _norm(text):
    return ''.join(char for char in unicodedata.normalize('NFKC', text).casefold() if char.isalnum())


def _doi_values(text):
    values = set()
    for match in _DOI.finditer(text):
        value = match[0].rstrip('.,;:')
        while value.endswith(')') and value.count(')') > value.count('('):
            value = value[:-1]
        values.add(doi(value))
    return values


def _candidate_title(text, standalone):
    match = _TITLE.search(text)
    if match:
        return next(value.strip() for value in match.groups() if value)
    record = re.search(r'(?:^|[.。]\s*)([^.。\n]{3,300}?)\s*\[(?:J|M|C|D|R)(?:/OL)?\]', text, re.I)
    if record:
        return record[1].strip()
    apa = re.search(r'\((?:18|19|20)\d{2}\)\s*[.。]?\s*(.{3,300}?)(?:\.\s|。|$)', text)
    if apa:
        return apa[1].strip().rstrip('.')
    if re.search(r'(?:请|帮我|核对|论文|文献)[^。\n]{0,28}[：:]\s*\S', text):
        match = re.search(r'(?:核对(?:这(?:条|篇))?(?:引用|论文|题录|文献)?|论文|文献)[：:]\s*(.+)$', text)
        if match and match[1] != text:
            return _candidate_title(match[1].strip(), standalone)
    if (standalone and len(text) <= 220 and not re.search(r'(?i)[\n？！?。；;《》]|(?:帮我|请|核对|作者|这条|年份|查询|验证|检索|选中|选区|这篇|这本|这项|检查|引用|题录|文献|能否|是否|怎么|什么|为啥|为什么|我的)|^(?:please|check|verify|can you|what|who|when)\b', text)
            and len(text.strip()) >= 3):
        return text.strip()


def _query_fields(question, selected=''):
    ids = _doi_values(question)
    material = question
    if not ids and selected and not _TITLE.search(question):
        material = selected
        ids = _doi_values(selected)
    if len(ids) > 1:
        raise ValueError('本轮请只选择一条文献，或输入要核对的一个 DOI。')
    title = _candidate_title(_DOI.sub('', material).strip(), not ids)
    if ids:
        query = LiteratureQuery(mode='doi', query=next(iter(ids)), sources=('crossref',))
    else:
        if not title:
            raise ValueError('请输入一个 DOI，或用《题名》标出要核对的论文。')
        query = LiteratureQuery(mode='title', query=title, sources=('crossref',))
    if (mask_text(query.query) != query.query or '已隐藏]' in query.query
            or re.search(r'https?://|file:|[\\@]|[a-z]:[/\\]', query.query, re.I)):
        raise ValueError('请只提供公开题名或 DOI，移除路径、凭据和联系方式。')
    supplied = _DOI.sub('', material)
    if title:
        supplied = supplied.replace(title, '')
    if material != question:
        supplied += '\n' + _DOI.sub('', question)
    author_match = _AUTHORS.search(supplied)
    if not author_match:
        author_match = re.search(r'^(?:\[\d+\]\s*)?(.{2,150}?)\s*\((?:18|19|20)\d{2}\)\s*[.。]', material)
    if not author_match and title and re.search(r'\[(?:J|M|C|D|R)(?:/OL)?\]', material, re.I):
        author_match = re.search(r'^(?:\[\d+\]\s*)?([^。\n]{2,150}?)\s*[.。]\s*' + re.escape(title), material)

    years = list(dict.fromkeys(_YEAR.findall(supplied)))
    return query, title, years, author_match[1].strip() if author_match else ''


def prior_citation(conversation, row):
    if not _CLAIM.search(row['question']) or _DOI.search(row['question']) or _TITLE.search(row['question']):
        return None
    with conversation.lock:
        for prior in reversed(list(conversation.turns.values())):
            if prior is row:
                continue
            if prior.get('created', 0) > row.get('created', float('inf')):
                continue
            if (prior.get('status') == 'complete' and prior.get('document_id') == row['document_id']
                    and prior.get('source_version') == row.get('source_version') and prior.get('citation_records')):
                return prior
            break
    return None


def _record_text(record):
    title = record['title'] or '题名缺失'
    authors = '、'.join(record['authors'][:20]) or '作者缺失'
    if len(record['authors']) > 20 or record.get('authors_truncated'):
        authors += ' 等'
    return f"{title}\n{authors}，{record['year'] or '年份缺失'} · {record.get('container_title') or '刊物缺失'}\nDOI {record['doi'] or '缺失'}"


def _present(query, supplied_title, years, supplied_authors, records):
    if not records:
        return '没有查到匹配的书目。请核对 DOI 或题名。', []
    exact = [record for record in records if _norm(record['title']) == _norm(query.query)]
    if query.mode == 'title' and exact:
        records = exact
    groups = group_records(records)
    matched = query.mode == 'doi' or len(groups) == 1 and bool(exact)
    if len(groups) > 1 or not matched:
        candidates = [group['variants'][0] for group in groups[:3]]
        message = '有多条候选，请用 DOI 或作者区分。' if len(groups) > 1 else '只找到相近题名，请核对后再比较字段。'
        return message + '\n\n' + '\n\n'.join(_record_text(record) for record in candidates), candidates
    variants = groups[0]['variants']
    if groups[0]['conflicts']:
        return '同一 DOI 的来源字段存在差异，请核对以下题录。\n\n' + '\n\n'.join(_record_text(r) for r in variants[:3]), variants[:3]
    record = variants[0]
    findings = []
    if supplied_title and query.mode == 'doi' and _norm(supplied_title) != _norm(record['title']):
        findings.append('题名不符。输入为「' + supplied_title + '」。')
    if years and record['year'] and any(year != str(record['year']) for year in years):
        findings.append('年份不符。输入为 ' + '、'.join(years) + '，Crossref 记录为 ' + str(record['year']) + '。')
    if supplied_authors and record['authors']:
        names = [_norm(name) for name in re.split(r'\s*(?:、|，|;|,|\band\b|&)\s*', supplied_authors, flags=re.I) if name]
        if names != [_norm(name) for name in record['authors']]:
            findings.append('作者写法或顺序有差异。输入为「' + supplied_authors + '」，请与下列署名核对。')
    if not findings:
        findings.append('已找到对应书目。' if not (supplied_title or years or supplied_authors) or record.get('missing_fields') else '已核对，所提供的字段与书目相符。')
    return '\n'.join(findings) + '\n' + _record_text(record), [record]


def handles_without_model(floating, question, resources):
    if not resources:
        return bool(prior_citation(floating.conversation, {'question': question, 'document_id': floating.document['id'],
                                                          'source_version': floating.document['version']}))
    if not isinstance(resources, dict) or set(resources) != {'skill'}:
        return False
    from ..skills_backend import SkillsBackend
    selected = SkillsBackend(floating.app.path.parent).turn_skill(resources['skill'], question)
    return selected.get('workflow') == 'citation-check'


def answer_with_citation(conversation, row, reference, metrics, current):
    floating = conversation.floating
    scopes = conversation.scopes
    identifier = row['id']
    binding = conversation._read_binding()

    def guard():
        with floating.lock, conversation.lock:
            if (not current() or row['status'] != 'running' or binding != conversation._read_binding()
                    or not scopes.selection_valid(reference)):
                raise ValueError('文档、选区或当前回合已改变，请重新核对题录。')
            floating.authenticated()
            floating.check_source()

    guard()
    prior = prior_citation(conversation, row)
    if prior:
        with floating.lock, conversation.lock:
            guard()
            metrics.mark('first_effective_answer')
            conversation.update(identifier, 'answer', _LOOKUP_ONLY, status='complete', answer=_LOOKUP_ONLY,
                                citation_sources=prior.get('citation_sources', []), citation_records=prior['citation_records'])
        return
    selected = ''
    if reference and not (_DOI.search(row['question']) or _TITLE.search(row['question'])):
        request = conversation.read_requests.get(identifier)
        conversation._material(row['question'], reference, read_binding=request[1] if request else binding,
                               with_sources=True, force_selection=True)
        selected = reference.get('old', '')
    query, title, years, authors = _query_fields(row['question'], selected)
    field = 'DOI' if query.mode == 'doi' else '题名'
    payload = {field: query.query}
    with floating.lock, conversation.lock:
        guard()
        approved = scopes.approved.pop(identifier, None)
        if approved is None:
            scopes.ask(row, 'citation', binding, _TARGET, reference, payload, payload, 'Crossref（api.crossref.org）', [field], False)
            return
        decision, choice = approved
        if (choice != 'allow' or decision.action != 'citation' or decision.target != _TARGET or decision.full != payload
                or decision.binding != scopes._binding(binding, identifier) or time.monotonic() - decision.created > TTL
                or row.get('scope_dispatched')):
            raise ValueError('本次书目查询许可已失效，请核对后重新发送。')
        from ..skills_backend import SkillsBackend
        selected_skill = scopes.selected_skills.get(identifier) or {}
        renewed = SkillsBackend(floating.app.path.parent).turn_skill(selected_skill.get('ticket'), row['question'])
        if renewed != selected_skill or renewed.get('workflow') != 'citation-check':
            raise ValueError('核对技能已变更，请重新选择后再查询。')
        row['scope_dispatched'] = True
        conversation.persist()
        if conversation.storage_error:
            row.pop('scope_dispatched', None)
            raise ValueError('无法记录本次许可，查询未发送。')
        conversation.update(identifier, 'citation_query', '正在向 Crossref 核对' + field + '。', skill_loaded=True)
    with metrics.span('tool'):
        metrics.count('tool_requests')
        result = floating.api._literature.lookup_for_chat(query, guard)
    with floating.lock, conversation.lock:
        guard()
        records, reports = result['result'] if result.get('ok') else ([], [])
        receipt = reports[0] if reports else {}
        status = receipt.get('status', result.get('reason', 'literature_response'))
        metrics.count('tool_reuse' if receipt.get('cached') else 'tool_executions')
        details = {k: receipt[k] for k in ('source', 'status', 'retrieved_at', 'cached', 'response_sha256') if k in receipt}
        if status not in ('completed', 'literature_not_found'):
            metrics.count('tool_failures')
            message = {'literature_rate_limited': 'Crossref 查询限流，请稍后重试。',
                       'literature_timeout': '未在时限内取得 Crossref 响应，请稍后重试。',
                       'busy': '文献服务正忙，请稍后再查询。'}.get(status, '未取得可核对的 Crossref 响应，请稍后重试。')
            conversation.update(identifier, 'failed', message, status='failed', citation_receipt=details)
            return
        answer, matches = _present(query, title, years, authors, records)
        if _CLAIM.search(row['question']) and matches:
            answer += '\n' + _LOOKUP_ONLY
        sources = [{'title': record['title'], 'url': record['source_url']} for record in matches if record.get('source_url')]
        metrics.mark('first_effective_answer')
        conversation.update(identifier, 'answer', answer, status='complete', answer=answer, citation_sources=sources,
                            citation_records=[record['id'] for record in matches], citation_receipt=details)
