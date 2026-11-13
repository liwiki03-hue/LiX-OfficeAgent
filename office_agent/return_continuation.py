import time
import uuid
from typing import Literal

from pydantic import Field

from .document_tasks import compile_task, load
from .output_dependencies import file_identity
from .range_edits import _islands
from .rag import SECRET
from .writes import GuardError, Strict, digest


class ContinuationRequest(Strict):
    comparison_id: str = Field(pattern=r'^[a-f0-9]{32}$')
    row_id: str = Field(pattern=r'^[a-f0-9]{32}$')
    find: str = Field(min_length=1, max_length=120)
    replacement: str = Field(min_length=1, max_length=200)
    reason: str = Field(min_length=1, max_length=400)
    mode: Literal['clean', 'tracked'] = 'clean'
    acknowledge_current: bool = False
    editing_profile: Literal['general', 'technical_report_v1'] = 'general'
    protected_literals: list[str] = Field(default_factory=list, max_length=30)


def prepare(editor, current, result, binding):
    paths = [editor.space.file(str(editor.space.root / row['name'])) for row in result['versions']]
    identities = [file_identity(editor, path) for path in paths]
    version, _, _, entries = load(editor, current)
    if version != result['versions'][2]['sha256']:
        raise GuardError('professional_source_changed')
    by_location = {entry['location']: entry for entry in entries}
    table_mappings = None
    if result.get('continuation_tables_requested') and any('/table/' in (row['current_location'] or '') for row in result['items']):
        table_mappings = [{entry['location']: entry for entry in load(editor, source)[3]} for source in paths[:2]] + [by_location]
    allowed = {}
    for row in result['items']:
        entry = by_location.get(row['current_location'])
        reason = ('comparison_requires_review' if row['requires_review'] and row['classification'] != 'user_content_change' else
                  'unmapped_or_unsupported' if entry is None or ('/table/' in entry['location'] and table_mappings is None) or
                  entry['refusal'] == 'merged_cell_continuation' or not _islands(entry) else
                  'sensitive_or_oversize' if len(entry['text']) > 8000 or SECRET.search(entry['text']) else None)
        row['continuation_blocked'] = reason
        row['continuation_family'] = 'ordinary_table_cell' if entry and '/table/' in entry['location'] else 'main_paragraph'
        if reason is None and row['continuation_family'] == 'ordinary_table_cell':
            from .return_tables import eligibility
            row['continuation_blocked'] = reason = eligibility(row, table_mappings)
        row['continuation_requires_current_review'] = row['classification'] == 'user_content_change'
        row['continuation_id'] = None
        if reason is None:
            row['continuation_id'] = uuid.uuid4().hex
            allowed[row['continuation_id']] = {'location': entry['location'], 'text_sha256': entry['text_sha256'], 'family': row['continuation_family']}
    snapshot = {'id': uuid.uuid4().hex, 'binding': binding, 'created': time.monotonic(),
                'paths': paths, 'identities': identities, 'versions': [row['sha256'] for row in result['versions']],
                'rows': allowed}
    validate(editor, current, snapshot, binding)
    result.update(comparison_id=snapshot['id'], continuation_expires_in_seconds=300,
                  continuation_scope='one exact unique substring in a verified current main paragraph or explicitly admitted ordinary single-paragraph table cell; no merge or restored authority')
    return snapshot


def validate(editor, current, snapshot, binding):
    if not snapshot or snapshot['binding'] != binding or time.monotonic() - snapshot['created'] > 300:
        raise GuardError('professional_return_expired')
    if snapshot['paths'][-1] != current:
        raise GuardError('professional_source_changed')
    checked = [editor.space.file(str(path)) for path in snapshot['paths']]
    if (checked != snapshot['paths'] or snapshot['identities'] != [file_identity(editor, path) for path in checked] or
            snapshot['versions'] != [digest(path) for path in checked]):
        raise GuardError('professional_source_changed')


def candidate(editor, current, snapshot, binding, raw):
    request = ContinuationRequest.model_validate(raw)
    validate(editor, current, snapshot, binding)
    if request.comparison_id != snapshot['id'] or request.row_id not in snapshot['rows']:
        raise GuardError('professional_return_target')
    if not request.acknowledge_current:
        raise GuardError('professional_return_acknowledgement')
    if SECRET.search(request.model_dump_json()) or request.find == request.replacement:
        raise GuardError('professional_task_request')
    version, _, _, entries = load(editor, current)
    row = snapshot['rows'][request.row_id]
    entry = next((entry for entry in entries if entry['location'] == row['location']), None)
    if entry is None or entry['text_sha256'] != row['text_sha256']:
        raise GuardError('professional_source_changed')
    from .technical_profile import verify_literals
    verify_literals([entry['text']], request.protected_literals)
    start = entry['text'].find(request.find)
    if start < 0 or entry['text'].find(request.find, start + 1) >= 0:
        raise GuardError('professional_return_unique_text')
    end = start + len(request.find)
    islands = [island for island in _islands(entry) if island['start'] <= start < end <= island['end']]
    if len(islands) != 1:
        raise GuardError('professional_protected_boundary')
    spec = {'source_sha256': version, 'request': '回流后对当前版本独立继续修改；不合并或恢复旧批准。' + request.reason +
            '\n三版本摘要：' + ', '.join(snapshot['versions']), 'mode': request.mode,
            'editing_profile': request.editing_profile, 'protected_literals': request.protected_literals,
            'requirements': [{'key': 'continuation', 'goal': request.reason, 'kind': 'ranges', 'scope': 'selected',
                              'locators': [entry['locator_id']], 'acceptance': '只修改当前核验范围；保留非目标内容、保护字面量和其他作者修订；语义仍待审阅',
                              'operations': [{'locator_id': entry['locator_id'], 'expected_text_sha256': entry['text_sha256'],
                                              'island_id': islands[0]['id'], 'action': 'replace', 'start': start, 'end': end,
                                              'quote': request.find, 'text': request.replacement, 'reason': request.reason}]}]}
    task = compile_task(editor, current, spec)
    if len(task['items']) != 1 or task['items'][0]['state'] != 'pending':
        raise GuardError(task['items'][0]['reason'] or 'professional_return_target')
    validate(editor, current, snapshot, binding)
    return task
