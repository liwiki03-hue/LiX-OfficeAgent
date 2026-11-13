from copy import deepcopy
import json
from pathlib import Path
import shutil

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.advanced_writes import _parts as original_parts
from office_agent.professional_edits import W, _canonical, _text
from office_agent.range_edits import RangeDocx, transform, verify, view
from office_agent.writes import Editor, GuardError, digest
from tests.test_professional_edits import fake_render


ROOT = Path(__file__).resolve().parents[1]


def _parts(path):
    return original_parts(path, range_edit=True)


@pytest.fixture
def sample(tmp_path):
    path = tmp_path / 'source.docx'
    shutil.copyfile(ROOT / 'deliverables/docx-fidelity-20261020/challenges/development.docx', path)
    editor = Editor(tmp_path, require_confirmation=True)
    return path, editor, view(editor, path)


def operation(item, start=0, end=None, text='changed', action='replace', **kwargs):
    end = len(item['text']) if end is None else end
    island = next((row for row in item.get('islands', []) if row['start'] <= start <= end <= row['end']), None)
    return {'locator_id': item['locator_id'], 'expected_text_sha256': item['text_sha256'], 'action': action,
            'start': start, 'end': end, 'quote': item['text'][start:end], 'text': text,
            'island_id': island['id'] if island else None, 'reason': 'Approved precise edit', **kwargs}


def request(sample, operations, **kwargs):
    path, editor, data = sample
    return RangeDocx(path=str(path), source_sha256=data['document']['source_sha256'], operations=operations,
                     revision_date='2026-10-20T00:00:00Z', **kwargs)


def test_cross_run_requires_policy_and_distributed_keeps_formats(sample):
    item = sample[2]['items'][0]
    op = operation(item, 10, 20, 'abcdefghij')
    with pytest.raises(GuardError, match='mixed_format_policy'):
        transform(_parts(sample[0]), request(sample, [op]))
    op['inheritance'] = 'distributed'
    before = _parts(sample[0])
    after, _, _ = transform(before, request(sample, [op]))
    verify(before, after, request(sample, [op]))
    paragraph = ET.fromstring(after['word/document.xml']).find('.//' + W + 'p')
    assert 'abcdefghij' in _text(paragraph)
    assert paragraph.find('.//' + W + 'b') is not None and paragraph.find('.//' + W + 'i') is not None
    assert all(value == after[name] for name, value in before.items() if name != 'word/document.xml')


def test_repeat_two_ranges_same_run_reverse_order(sample):
    item = sample[2]['items'][1]
    req = request(sample, [operation(item, 0, 6, 'first'), operation(item, 7, 13, 'second')])
    after, _, _ = transform(_parts(sample[0]), req)
    assert ET.fromstring(after['word/document.xml']).findall('.//' + W + 'body/' + W + 'p')[1].xpath('string()').startswith('first second')


@pytest.mark.parametrize('index,start,end,new', [(2, 0, 6, 'Outside '), (2, 7, 13, 'LINK'), (3, 0, 4, 'LEFT'), (4, 0, 6, 'PREFIX')])
def test_complex_islands_preserve_specials(sample, index, start, end, new):
    item = sample[2]['items'][index]
    before = _parts(sample[0])
    req = request(sample, [operation(item, start, end, new)])
    after, _, _ = transform(before, req)
    verify(before, after, req)
    original, output = (ET.fromstring(parts['word/document.xml']) for parts in (before, after))
    for name in ('hyperlink', 'bookmarkStart', 'bookmarkEnd', 'fldChar', 'instrText'):
        a, b = list(original.iter(W + name)), list(output.iter(W + name))
        assert len(a) == len(b)
        for left, right in zip(a, b):
            assert left.attrib == right.attrib
            if name != 'hyperlink':
                assert _canonical(left) == _canonical(right)


def test_cross_bookmark_and_field_result_refused(sample):
    for index, start, end in ((3, 0, 10), (4, 7, 8)):
        with pytest.raises(GuardError, match='protected_boundary'):
            transform(_parts(sample[0]), request(sample, [operation(sample[2]['items'][index], start, end)]))


@pytest.mark.parametrize('offset', [15, 18, 19])
def test_unicode_grapheme_boundary_refusal(sample, offset):
    item = sample[2]['items'][1]
    text = item['text']
    invalid = text.index('\u0301') if offset == 15 else (text.index('\u200d') + (offset == 19))
    op = operation(item, invalid, invalid, 'x', action='insert')
    with pytest.raises(GuardError, match='invalid_range'):
        transform(_parts(sample[0]), request(sample, [op]))


def test_partial_format_splits_only_selection(sample):
    item = sample[2]['items'][1]
    req = request(sample, [operation(item, 7, 13, '', action='format', character_format={'bold': True, 'size_pt': 14})])
    after, _, _ = transform(_parts(sample[0]), req)
    paragraph = ET.fromstring(after['word/document.xml']).findall('.//' + W + 'body/' + W + 'p')[1]
    runs = paragraph.findall(W + 'r')
    assert [_text(run) for run in runs] == ['repeat ', 'repeat', ' e\u0301 👩\u200d🔬 中文']
    assert runs[0].find(W + 'rPr') is None and runs[2].find(W + 'rPr') is None
    assert runs[1].find('./' + W + 'rPr/' + W + 'sz').get(W + 'val') == '28'


def test_paragraph_format_only_selected_style_definition_unchanged(sample):
    item = sample[2]['items'][5]
    req = request(sample, [operation(item, text='', action='paragraph_format', paragraph_format={'style': 'Heading2', 'after_pt': 8, 'alignment': 'center'})])
    before = _parts(sample[0])
    after, _, _ = transform(before, req)
    assert after['word/styles.xml'] == before['word/styles.xml']
    assert b'Heading2' in after['word/document.xml']


@pytest.mark.parametrize('action', ['split', 'merge_next', 'insert_before', 'insert_after', 'paragraph_delete'])
def test_paragraph_structure_operations(sample, action):
    item = sample[2]['items'][5]
    op = operation(item, 6, 6, 'New paragraph' if action.startswith('insert') else '', action=action)
    req = request(sample, [op])
    before = _parts(sample[0])
    after, _, _ = transform(before, req)
    verify(before, after, req)
    assert before['word/header1.xml'] == after['word/header1.xml']
    assert before['word/footer1.xml'] == after['word/footer1.xml']


@pytest.mark.parametrize('action', ['list_continue', 'list_restart'])
def test_numbering_is_real_and_style_definition_unchanged(sample, action):
    item = sample[2]['items'][8]
    req = request(sample, [operation(item, text='', action=action)])
    before = _parts(sample[0])
    after, _, _ = transform(before, req)
    verify(before, after, req)
    assert after['word/styles.xml'] == before['word/styles.xml']
    assert b'numPr' in after['word/document.xml']
    assert (before['word/numbering.xml'] != after['word/numbering.xml']) == (action == 'list_restart')


@pytest.mark.parametrize('action', ['row_insert', 'row_delete'])
def test_simple_table_row_keeps_grid(sample, action):
    item = next(row for row in sample[2]['items'] if '/table/row:1/physical-cell:0' in row['location'])
    req = request(sample, [operation(item, text='', action=action)], preserve_numbers=False)
    before = _parts(sample[0])
    after, _, _ = transform(before, req)
    tables = [ET.fromstring(parts['word/document.xml']).find('.//' + W + 'tbl') for parts in (before, after)]
    assert _canonical(tables[0].find(W + 'tblGrid')) == _canonical(tables[1].find(W + 'tblGrid'))
    assert len(tables[1].findall(W + 'tr')) == (4 if action == 'row_insert' else 2)


def test_existing_revision_kept_and_new_projection_verified(sample):
    item = sample[2]['items'][9]
    req = request(sample, [operation(item, 0, 6, 'BEFORE')], mode='tracked')
    before = _parts(sample[0])
    after, _, preview = transform(before, req)
    original = ET.fromstring(before['word/document.xml']).find('.//' + W + 'ins')
    output = ET.fromstring(after['word/document.xml'])
    assert _canonical(original) == _canonical(next(n for n in output.iter(W + 'ins') if n.get(W + 'id') == '42'))
    assert len(list(output.iter(W + 'ins'))) == 2
    assert preview[0]['verification']['mode'] == 'tracked'


@pytest.mark.parametrize('action', ['revision_accept', 'revision_reject'])
def test_selective_revision(sample, action):
    item = sample[2]['items'][9]
    op = operation(item, text='', action=action, revision_id=item['revisions'][0]['id'])
    after, _, _ = transform(_parts(sample[0]), request(sample, [op]))
    root = ET.fromstring(after['word/document.xml'])
    assert not list(root.iter(W + 'ins'))
    assert ('existing' in _text(root)) == (action == 'revision_accept')


def test_transaction_cancel_external_conflict_copy_and_undo(sample, monkeypatch):
    path, editor, data = sample
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    req = request(sample, [operation(data['items'][1], 7, 13, 'second')])
    original = digest(path)
    plan = editor.plan(req)
    with pytest.raises(GuardError):
        editor.execute(plan.id, confirmed=False)
    assert digest(path) == original
    plan = editor.plan(req)
    result = editor.execute(plan.id, confirmed=True)
    assert result['status'] == 'verified' and digest(path) == original
    assert 'repeat second' in Document(result['output']).paragraphs[1].text
    assert editor.rollback(plan.id, confirmed=True)['status'] == 'rolled_back'
    plan = editor.plan(req)
    path.write_bytes(path.read_bytes() + b' ')
    with pytest.raises(GuardError):
        editor.execute(plan.id, confirmed=True)


def test_independent_package_tamper_rejected(sample):
    req = request(sample, [operation(sample[2]['items'][1], 7, 13, 'second')])
    before = _parts(sample[0])
    after, _, _ = transform(before, req)
    after['word/styles.xml'] += b' '
    with pytest.raises(GuardError, match='non_target_part_changed'):
        verify(before, after, req)


def test_overlaps_and_false_unused_payload_refused(sample):
    item = sample[2]['items'][1]
    for ops, reason in [([operation(item, 0, 6, 'one'), operation(item, 2, 5, 'two')], 'overlapping_changes'),
                        ([operation(item, 0, 6, 'ignored', action='paragraph_delete')], 'invalid_request')]:
        with pytest.raises(GuardError, match=reason):
            transform(_parts(sample[0]), request(sample, ops))


def test_bookmark_complex_row_and_field_failures_preserved(sample):
    parts = _parts(sample[0])
    original = deepcopy(parts)
    item = sample[2]['items'][2]
    op = operation(item, 0, len(item['text']), 'Whole rewrite', inheritance='start')
    with pytest.raises(GuardError, match='protected_boundary'):
        transform(parts, request(sample, [op]))
    merged = next(row for row in sample[2]['items'] if row['text'] == 'merged heading')
    with pytest.raises(GuardError, match='table_scope'):
        transform(parts, request(sample, [operation(merged, text='', action='row_insert')]))
    assert parts == original


def test_native_paragraph_ids_not_duplicated_on_split(sample):
    from office_agent.range_edits import _new_paragraph_ids
    root = ET.fromstring(_parts(sample[0])['word/document.xml'])
    paragraph = root.find('.//' + W + 'p')
    key = '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
    paragraph.set(key, '12345678')
    copied = deepcopy(paragraph)
    _new_paragraph_ids(copied, root, 'host-bound-fixture')
    assert copied.get(key) != '12345678' and len(copied.get(key)) == 8
    other = deepcopy(paragraph)
    _new_paragraph_ids(other, root, 'host-bound-fixture')
    assert copied.get(key) == other.get(key)


@pytest.mark.parametrize('text,point', [('क\u093f', 1), ('가', 1), ('🇨🇳', 1)])
def test_unicode_spacing_marks_hangul_flags_use_graphemes(text, point):
    from office_agent.range_edits import _boundary
    assert not _boundary(text, point)
    assert _boundary(text, 0) and _boundary(text, len(text))
