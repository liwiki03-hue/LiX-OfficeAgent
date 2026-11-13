from copy import deepcopy
from itertools import product
import re

from lxml import etree as ET
from pydantic import Field

from .document_tasks import load
from .professional_edits import W, _canonical, _text
from .tracked_edits import _coalesced
from .output_dependencies import file_identity
from .writes import Strict, GuardError, digest, _audit_record, _task_literal_audit
from .rag import SECRET


class ReturnRequest(Strict):
    baseline_name: str = Field(min_length=6, max_length=220)
    delivered_name: str = Field(min_length=6, max_length=220)
    receipts: list[str] = Field(min_length=1, max_length=16)
    approved_read: bool = False
    offset: int = Field(default=0, ge=0, le=4000)
    include_plain_tables: bool = False


def _filename(value):
    if not value.lower().endswith('.docx') or any(c in value for c in '/\\:') or any(ord(c) < 32 for c in value) or SECRET.search(value):
        raise GuardError('professional_return_filename')
    return value


def _normalized(node, *, ignore_comments=False):
    node = deepcopy(node)
    volatile = {W + name for name in ('rsidR', 'rsidRDefault', 'rsidRPr', 'rsidP', 'rsidDel')}
    for child in list(node.iter()):
        for key in list(child.attrib):
            if key in volatile:
                child.attrib.pop(key)
        if ignore_comments and child.tag in {W + 'commentRangeStart', W + 'commentRangeEnd', W + 'commentReference'}:
            child.getparent().remove(child)
    node = ET.fromstring(_coalesced(node))
    ET.cleanup_namespaces(node)
    return ET.tostring(node, method='c14n', exclusive=True)


def _masked(node):
    root = deepcopy(node)
    stack = []
    for element in root.iter():
        if element.tag == W + 'fldSimple':
            element.attrib.pop(W + 'dirty', None)
            for text in element.iter(W + 't'):
                text.text = ''
        elif element.tag == W + 'fldChar':
            element.attrib.pop(W + 'dirty', None)
            kind = element.get(W + 'fldCharType')
            if kind == 'begin':
                stack.append(False)
            elif kind == 'separate' and stack:
                stack[-1] = True
            elif kind == 'end' and stack:
                stack.pop()
        elif element.tag == W + 't' and any(stack):
            element.text = ''
    return root


def _project(paragraph, identities, actions):
    projected = deepcopy(paragraph)
    for wrapper in list(projected):
        key = wrapper.get(W + 'id')
        if wrapper.tag not in {W + 'ins', W + 'del'} or key not in identities:
            continue
        action = actions[identities.index(key)]
        if action == 'keep':
            continue
        keep = (wrapper.tag == W + 'ins') == (action == 'accept')
        index = projected.index(wrapper)
        if keep:
            for run in list(wrapper):
                for text in run.findall(W + 'delText'):
                    text.tag = W + 't'
                projected.insert(index, run)
                index += 1
        projected.remove(wrapper)
    return projected


def _revision_states(baseline, delivered, current):
    previous = {node.get(W + 'id') for node in baseline.iter() if node.tag in {W + 'ins', W + 'del'}} if baseline is not None else set()
    wrappers = [node for node in delivered if node.tag in {W + 'ins', W + 'del'} and node.get(W + 'id') not in previous]
    if not wrappers or len(wrappers) > 6 or any(node.get(W + 'author') != 'Office Agent' or any(child.tag != W + 'r' or any(n.tag not in {W + 'rPr', W + 't', W + 'delText'} for n in child) for child in node) for node in wrappers):
        return None
    ids = [node.get(W + 'id') for node in wrappers]
    if len(set(ids)) != len(ids):
        return None
    signature = _normalized(current, ignore_comments=True)
    candidates = []
    for actions in product(('keep', 'accept', 'reject'), repeat=len(ids)):
        if _normalized(_project(delivered, ids, actions), ignore_comments=True) == signature:
            candidates.append(dict(zip(ids, actions)))
            if len(candidates) > 1:
                return {'classification': 'revision_projection_ambiguous', 'requires_review': True}
    if not candidates:
        return None
    actions = candidates[0]
    values = set(actions.values())
    label = 'system_revision_still_present' if values == {'keep'} else 'system_revision_accepted' if values == {'accept'} else 'system_revision_rejected' if values == {'reject'} else 'system_revision_partially_resolved'
    return {'classification': label, 'revision_actions': actions, 'requires_review': False,
            'baseline_view': 'stored w:t; project only new direct Office Agent ins/del; preserve pre-existing wrappers'}


def compare_return(editor, current, raw):
    request = ReturnRequest.model_validate(raw)
    if not request.approved_read:
        return {'status': 'cancelled', 'wrote_nothing': True}
    baseline = editor.space.file(str(editor.space.root / _filename(request.baseline_name)))
    delivered = editor.space.file(str(editor.space.root / _filename(request.delivered_name)))
    paths = [baseline, delivered, current]
    if len(set(paths)) != 3:
        raise GuardError('professional_return_three_versions')
    identities = [file_identity(editor, path) for path in paths]
    loaded = [load(editor, path) for path in paths]
    versions = [row[0] for row in loaded]
    audits = []
    for receipt in request.receipts:
        if not re.fullmatch('[a-f0-9]{32}', receipt):
            raise GuardError('professional_return_receipt')
        audit, evidence = _audit_record(editor.space, editor.audit / receipt)
        if audit['status'] != 'verified' or (audit['kind'] not in {'range_docx', 'block_docx'} and not _task_literal_audit(audit)):
            raise GuardError('professional_return_receipt')
        audits.append(audit)
    if (audits[0]['source_sha256'] != versions[0] or audits[-1]['output_sha256'] != versions[1] or
            any(left['output_sha256'] != right['source_sha256'] for left, right in zip(audits, audits[1:]))):
        raise GuardError('professional_return_lineage')
    base, system, returned = [row[3] for row in loaded]
    mapping = {entry['location']: entry['location'] for entry in base}
    for audit in audits:
        stage = json_mapping(audit)
        if stage is not None:
            mapping = {old: stage.get(location) for old, location in mapping.items()}
    inverse = {value: key for key, value in mapping.items() if value is not None}
    by_base = {entry['location']: entry for entry in base}
    by_current = {entry['location']: entry for entry in returned}
    same_shape = [entry['location'] for entry in system] == [entry['location'] for entry in returned]
    signatures = [_normalized(entry['node'], ignore_comments=True) for entry in system]
    current_signatures = [_normalized(entry['node'], ignore_comments=True) for entry in returned]
    para_id = '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
    ids = [entry['node'].get(para_id) for entry in system]
    current_ids = [entry['node'].get(para_id) for entry in returned]
    stable_ids = all(ids) and all(current_ids) and len(set(ids)) == len(ids) and len(set(current_ids)) == len(current_ids)
    by_current_id = {entry['node'].get(para_id): entry for entry in returned} if stable_ids else {}
    rows = []
    for index, entry in enumerate(system):
        original = by_base.get(inverse.get(entry['location']))
        candidate = by_current_id.get(entry['node'].get(para_id)) if stable_ids else by_current.get(entry['location']) if same_shape else None
        row = {'delivered_location': entry['location'], 'baseline_location': original['location'] if original else None,
               'current_location': candidate['location'] if candidate else None, 'mapping': 'unique_paraId' if stable_ids else 'structural_position_with_context' if same_shape else 'unmapped',
               'baseline': _safe(original['text']) if original else '', 'delivered': _safe(entry['text']),
               'current': _safe(candidate['text']) if candidate else '', 'requires_review': True}
        if candidate is None:
            row['classification'] = 'structure_changed_unmapped'
        elif signatures[index] == _normalized(candidate['node'], ignore_comments=True):
            row.update(classification='system_suggestion_still_present' if original is None or _normalized(original['node'], ignore_comments=True) != signatures[index] else 'unchanged', requires_review=False)
        elif not stable_ids and not _context_unique(index, signatures, current_signatures):
            row['classification'] = 'ambiguous_position'
        elif _normalized(_masked(entry['node']), ignore_comments=True) == _normalized(_masked(candidate['node']), ignore_comments=True):
            row['classification'] = 'stored_field_result_changed_unverified'
        else:
            projected = _revision_states(original['node'] if original else None, entry['node'], candidate['node'])
            if projected:
                row.update(projected)
            elif original is not None and _normalized(original['node'], ignore_comments=True) == _normalized(candidate['node'], ignore_comments=True):
                row['classification'] = 'baseline_restored_resolution_unknown'
            elif entry['text'] == candidate['text']:
                row['classification'] = 'format_or_structure_changed'
            elif original is not None and original['text'] == entry['text']:
                row['classification'] = 'user_content_change'
            else:
                row['classification'] = 'same_position_conflict'
        if candidate is not None:
            old = {(node.tag, node.get(W + 'id'), _canonical(node)) for node in entry['node'].iter() if node.tag in {W + 'ins', W + 'del'}}
            authors = [node.get(W + 'author', '')[:100] for node in candidate['node'].iter() if node.tag in {W + 'ins', W + 'del'} and (node.tag, node.get(W + 'id'), _canonical(node)) not in old]
            row['additional_revision_authors'] = [_safe(author) for author in authors[:8]]
        rows.append(row)
    changes = []
    system_parts, current_parts = loaded[1][1], loaded[2][1]
    for name in sorted(set(system_parts) | set(current_parts)):
        if system_parts.get(name) == current_parts.get(name) or name == 'word/document.xml':
            continue
        kind = 'package_part_changed_requires_review'
        if name in system_parts and name in current_parts and name.endswith('.xml'):
            if _normalized(_masked(ET.fromstring(system_parts[name]))) == _normalized(_masked(ET.fromstring(current_parts[name]))):
                kind = 'stored_field_or_volatile_metadata_only_unverified'
        changes.append({'part': name, 'classification': kind})
    if identities != [file_identity(editor, path) for path in paths] or versions != [digest(path) for path in paths]:
        raise GuardError('professional_source_changed')
    return {'status': 'compared', 'continuation_tables_requested': request.include_plain_tables, 'versions': [{'role': role, 'name': path.name, 'sha256': version} for role, path, version in zip(('baseline', 'delivered', 'current'), paths, versions)],
            'items': rows[request.offset:request.offset + 20], 'next_offset': request.offset + 20 if request.offset + 20 < len(rows) else None,
            'total_targets': len(rows), 'classifications': dict(__import__('collections').Counter(row['classification'] for row in rows)),
            'registered_current_targets': len(returned), 'unmapped_current_targets': len(returned) if not same_shape and not stable_ids else 0,
            'package_changes': changes[:40], 'package_changes_total': len(changes), 'package_changes_truncated': len(changes) > 40,
            'write_approval_restored': False, 'rollback_authority_restored': False, 'read_only': True,
            'coverage': 'registered main paragraphs and known simple direct revisions only; ambiguous mapping preserves both snippets; no automatic merge, dynamic Word pagination or full OOXML semantic validation'}


def json_mapping(audit):
    import json
    if audit['kind'] != 'block_docx':
        return None
    return json.loads(audit['after'])['items'][0]['verification']['position_map']


def _safe(value):
    return '[已脱敏]' if SECRET.search(value) else value[:160]


def _context_unique(index, original, current):
    if len(original) != len(current):
        return False
    left = index == 0 or original[index - 1] == current[index - 1]
    right = index + 1 == len(original) or original[index + 1] == current[index + 1]
    if not left or not right or len(original) == 1:
        return False
    frame = tuple(original[max(0, index - 1):index + 2])
    if sum(tuple(original[max(0, i - 1):i + 2]) == frame for i in range(len(original))) != 1:
        return False
    return current[index] not in original[:index] + original[index + 1:]
