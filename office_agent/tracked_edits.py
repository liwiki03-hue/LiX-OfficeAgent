from copy import deepcopy
from difflib import SequenceMatcher
import hashlib
import re
from types import SimpleNamespace
from typing import Literal

from lxml import etree as ET
from pydantic import Field

from .assisted_edits import AssistedDocx, transform as assisted_transform
from .professional_edits import W, _inventory, _canonical
from .writes import GuardError, _REVISION_ALL_MARKS, _apply_revision_decision


class TrackedAssistedDocx(AssistedDocx):
    kind: Literal['tracked_assisted_docx'] = 'tracked_assisted_docx'
    revision_author: Literal['Office Agent'] = 'Office Agent'
    revision_date: str = Field(pattern=r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$')
    revision_scope: Literal['paragraph', 'word'] = 'paragraph'


def _run_text(run):
    if any(child.tag not in {W + 'rPr', W + 't'} for child in run):
        raise GuardError('professional_revision_run_unsupported')
    return ''.join(child.text or '' for child in run.findall(W + 't'))


def _slices(runs, start, end):
    cursor, result = 0, []
    for run in runs:
        value = _run_text(run)
        left, right = max(start - cursor, 0), min(end - cursor, len(value))
        if right > left:
            copied = deepcopy(run)
            if left != 0 or right != len(value):
                for child in list(copied):
                    if child.tag != W + 'rPr':
                        copied.remove(child)
                node = ET.SubElement(copied, W + 't')
                node.text = value[left:right]
                if node.text[:1].isspace() or node.text[-1:].isspace():
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            result.append(copied)
        cursor += len(value)
    return result


def _word_changes(old_runs, new_runs):
    old, new = (''.join(_run_text(run) for run in runs) for runs in (old_runs, new_runs))
    pattern = r'[\u3400-\u9fff]|\w+|\s+|[^\w\s]'
    tokens = [re.findall(pattern, value) for value in (old, new)]
    if sum(map(len, tokens)) > 3000:
        raise GuardError('professional_revision_capacity')
    offsets = []
    for values in tokens:
        positions = [0]
        for value in values:
            positions.append(positions[-1] + len(value))
        offsets.append(positions)
    result = []
    for action, left, right, new_left, new_right in SequenceMatcher(None, *tokens, autojunk=False).get_opcodes():
        if action == 'equal':
            result.append(('equal', _slices(old_runs, offsets[0][left], offsets[0][right])))
        else:
            if left < right:
                result.append(('del', _slices(old_runs, offsets[0][left], offsets[0][right])))
            if new_left < new_right:
                result.append(('ins', _slices(new_runs, offsets[1][new_left], offsets[1][new_right])))
    return result


def _coalesced(root):
    root = deepcopy(root)
    for paragraph in root.iter(W + 'p'):
        previous = None
        for run in list(paragraph):
            if run.tag != W + 'r':
                previous = None
                continue
            if any(child.tag not in {W + 'rPr', W + 't'} for child in run):
                previous = None
                continue
            text = _run_text(run)
            if not text:
                paragraph.remove(run)
                continue
            properties = run.find(W + 'rPr')
            signature = (tuple(sorted(run.attrib.items())), _canonical(properties) if properties is not None else None)
            if previous is not None and previous[0] == signature:
                target = previous[1]
                text = _run_text(target) + text
                paragraph.remove(run)
            else:
                target = run
                previous = (signature, target)
            for child in list(target):
                if child.tag == W + 't':
                    target.remove(child)
            node = ET.SubElement(target, W + 't')
            node.text = text
            if text[:1].isspace() or text[-1:].isspace():
                node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return _canonical(root)


def transform(parts, request):
    if request.review.get('revision_mode') != 'native_plain_paragraph':
        raise GuardError('professional_assist_review_invalid')
    for name, data in parts.items():
        if name.startswith('word/') and name.endswith('.xml'):
            root = ET.fromstring(data)
            if any(node.tag == W + mark for node in root.iter() for mark in _REVISION_ALL_MARKS):
                raise GuardError('professional_existing_revisions')
    clean_request = AssistedDocx(**request.model_dump(exclude={'kind', 'revision_author', 'revision_date', 'revision_scope'}))
    clean, before, after = assisted_transform(parts, clean_request)
    root, original = _inventory(parts, request.source_sha256)
    _, approved = _inventory(clean, request.source_sha256)
    originals = {row['location']: row for row in original}
    approved_by_location = {row['location']: row for row in approved}
    identifiers = []
    for index, change in enumerate(request.changes):
        row = next(row for row in original if row['locator_id'] == change.locator_id)
        if '/table/' in row['location'] or row['style']:
            raise GuardError('professional_unsupported_target')
        paragraph = row['node']
        old_runs = [deepcopy(run) for run in paragraph.findall(W + 'r')]
        new_runs = [deepcopy(run) for run in approved_by_location[row['location']]['node'].findall(W + 'r')]
        for run in list(paragraph.findall(W + 'r')):
            paragraph.remove(run)
        changes = _word_changes(old_runs, new_runs) if request.revision_scope == 'word' else [('del', old_runs), ('ins', new_runs)]
        for action, runs in changes:
            if action == 'equal':
                paragraph.extend(runs)
                continue
            if len(identifiers) >= 100:
                raise GuardError('professional_revision_capacity')
            revision_id = len(identifiers)
            wrapper = ET.SubElement(paragraph, W + action, {W + 'id': str(revision_id), W + 'author': request.revision_author, W + 'date': request.revision_date})
            for run in runs:
                if action == 'del':
                    for text in run.iter(W + 't'):
                        text.tag = W + 'delText'
                wrapper.append(run)
            identifiers.append(str(revision_id))
    tracked = dict(parts)
    tracked['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    for action, expected in (('accept_all', clean['word/document.xml']), ('reject_all', parts['word/document.xml'])):
        projected = deepcopy(root)
        _apply_revision_decision(SimpleNamespace(element=SimpleNamespace(body=projected.find(W + 'body'))), action)
        canonical = _coalesced if request.revision_scope == 'word' else _canonical
        if canonical(projected) != canonical(ET.fromstring(expected)):
            raise GuardError('professional_revision_projection_failed')
    after[0]['native_revisions'] = {'ids': identifiers, 'author': request.revision_author, 'date': request.revision_date,
                                  'accept_all_matches_approved': True, 'reject_all_matches_original': True,
                                  'approved_xml_sha256': hashlib.sha256(clean['word/document.xml']).hexdigest(),
                                  'scope': 'word_diff_with_character_style_projection' if request.revision_scope == 'word' else 'whole_run_content_replacement_in_existing_plain_paragraphs',
                                  'projection': 'adjacent_identical_run_properties_coalesced' if request.revision_scope == 'word' else 'exact_canonical_xml',
                                  'word_ui_verified': False}
    return tracked, before, after
