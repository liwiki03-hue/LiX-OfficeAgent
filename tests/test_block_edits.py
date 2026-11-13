from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from lxml import etree as ET
import pytest

from office_agent.advanced_writes import _parts
from office_agent.block_edits import BlockDocx, transform, verify, render_anchors
from office_agent.professional_edits import W, _canonical
from office_agent.range_edits import view
from office_agent.writes import Editor, GuardError, digest
from tests.test_range_fields import make_source


def sample(tmp_path):
    path = tmp_path / 'source.docx'
    doc = Document()
    doc.add_heading('第一章', 1)
    doc.add_paragraph('开始说明，数值20 ms [1]。')
    doc.add_paragraph('后续条件，仅限室内。')
    doc.add_heading('第二章', 1)
    doc.add_paragraph('第二章内容保持。')
    doc.add_heading('第三章', 1)
    doc.add_paragraph('结论保持。')
    doc.save(path)
    editor = Editor(tmp_path, require_confirmation=True)
    return path, editor, view(editor, path)


def req(path, data, action, rows, **kw):
    operation = {'action': action, 'locator_id': rows[0]['locator_id'], 'locator_ids': [row['locator_id'] for row in rows],
                 'reason': '明确结构修改', **kw}
    return BlockDocx(path=str(path), source_sha256=data['document']['source_sha256'], operations=[operation], revision_date='2026-10-21T00:00:00Z')


def test_cross_paragraph_partial_selection_keeps_boundaries_and_unrelated_parts(tmp_path):
    path, editor, data = sample(tmp_path)
    selected = data['items'][1:3]
    request = req(path, data, 'paragraphs_replace', selected, start=0, end=4,
                  new_texts=['实验说明，数值20 ms [1]。', '限定条件'])
    original = _parts(path, range_edit=True)
    changed, before, after = transform(original, request)
    verify(original, changed, request)
    texts = [p.text for p in Document(path).paragraphs]
    root = ET.fromstring(changed['word/document.xml'])
    actual = [''.join(n.text or '' for n in p.iter(W + 't')) for p in root.find(W + 'body').findall(W + 'p')]
    assert actual[1:3] == ['实验说明，数值20 ms [1]。', '限定条件，仅限室内。']
    assert actual[:1] + actual[3:] == texts[:1] + texts[3:]
    assert all(changed[name] == value for name, value in original.items() if name != 'word/document.xml')
    assert not after[0]['native_structure_revision']
    assert len(render_anchors(original, request)) == 1


def test_plain_chapter_moves_as_identity_preserving_block_and_source_unchanged(tmp_path):
    path, editor, data = sample(tmp_path)
    request = req(path, data, 'chapter_move', data['items'][:3], destination_locator=data['items'][5]['locator_id'])
    original = _parts(path, range_edit=True)
    changed, before, after = transform(original, request)
    verify(original, changed, request)
    old = ET.fromstring(original['word/document.xml']).find(W + 'body')
    new = ET.fromstring(changed['word/document.xml']).find(W + 'body')
    assert [_canonical(node) for node in new] == [_canonical(old[i]) for i in (3, 4, 0, 1, 2, 5, 6, 7)]
    assert after[0]['verification']['position_map']['body:0/paragraph'] == 'body:2/paragraph'
    assert editor.plan(request).source_sha256 == digest(path)
    incomplete = request.model_copy(deep=True)
    incomplete.operations[0].locator_ids = incomplete.operations[0].locator_ids[:2]
    with pytest.raises(GuardError, match='incomplete_chapter'):
        transform(original, incomplete)


def test_merge_refuses_style_loss_and_special_structure(tmp_path):
    path, editor, data = sample(tmp_path)
    request = req(path, data, 'paragraphs_merge', data['items'][1:3], new_texts=['开始说明，数值20 ms [1]。 后续条件，仅限室内。'])
    changed, _, after = transform(_parts(path, range_edit=True), request)
    assert after[0]['verification']['position_map']['body:2/paragraph'] is None
    doc = Document(path)
    doc.paragraphs[2].style = 'Heading 2'
    doc.save(path)
    data = view(editor, path)
    with pytest.raises(GuardError, match='style_conflict'):
        transform(_parts(path, range_edit=True), req(path, data, 'paragraphs_merge', data['items'][1:3], new_texts=['开始说明，数值20 ms [1]。 后续条件，仅限室内。']))


def test_bookmark_target_requires_separate_acknowledged_protocol_and_keeps_field(tmp_path):
    path = tmp_path / 'reference.docx'
    make_source(path)
    editor = Editor(tmp_path)
    data = view(editor, path)
    request = req(path, data, 'bookmark_text', [data['items'][0]], bookmark_name='_RefExperiment', new_texts=['仅限于干燥条件'])
    original = _parts(path, range_edit=True)
    with pytest.raises(GuardError, match='acknowledgment'):
        transform(original, request)
    request.operations[0].acknowledge_stale_fields = True
    changed, _, after = transform(original, request)
    verify(original, changed, request)
    for tag in ('bookmarkStart', 'bookmarkEnd', 'fldSimple', 'fldChar', 'instrText'):
        assert [_canonical(n) for n in ET.fromstring(original['word/document.xml']).iter(W + tag)] == [_canonical(n) for n in ET.fromstring(changed['word/document.xml']).iter(W + tag)]
    assert after[0]['verification']['reference_impact']['needs_review_total'] == 2
    assert not after[0]['verification']['stored_fields_refreshed']
