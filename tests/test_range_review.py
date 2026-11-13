import json
from pathlib import Path
import shutil
from types import SimpleNamespace

from lxml import etree as ET
import pytest

from office_agent.advanced_writes import _parts
from office_agent.app import build_dispatcher, professional_tools, select_tools
from office_agent.professional_edits import W, _text
from office_agent.range_edits import RangeDocx, RangeIntent, now, transform, verify, view
from office_agent.webui.professional import Professional
from office_agent.writes import Editor, GuardError, digest
from tests.test_range_edits import operation


@pytest.fixture
def setup(tmp_path):
    path = tmp_path / 'source.docx'
    root = Path(__file__).resolve().parents[1]
    shutil.copyfile(root / 'deliverables/docx-fidelity-20261020/challenges/development.docx', path)
    editor = Editor(tmp_path, require_confirmation=True)
    import threading
    service = Professional(threading.RLock(), lambda: (True, False, 'session'), lambda: editor, lambda: path)
    return path, editor, service


def test_model_tools_no_paths_and_plan_only(setup):
    path, editor, _ = setup
    tools = select_tools(professional_tools((path,)), 'word')
    assert {'read_docx_ranges', 'plan_docx_ranges'} <= {tool.name for tool in tools}
    assert 'path' not in RangeIntent.model_json_schema()['properties']
    dispatch = build_dispatcher(editor, path)
    data = dispatch('read_docx_ranges', {})
    result = dispatch('plan_docx_ranges', {'source_sha256': digest(path), 'operations': [operation(data['items'][1], 7, 13, 'second')]})
    assert result['plan_id'] in editor.plans and not editor.audit.exists()
    with pytest.raises(Exception):
        dispatch('read_docx_ranges', {'path': 'other.docx'})


def test_agent_keeps_exact_path_approval_for_host_bound_tools(setup, tmp_path):
    from office_agent.agent import AgentState, OfficeAgent, RequiredState, ToolCall
    path, editor, _ = setup
    tools = professional_tools((path,))
    agent = OfficeAgent(None, build_dispatcher(editor, path), tools, tmp_path / 'state', persist=False)
    state = AgentState(required=RequiredState(goal='Edit selected text', file_versions={str(path): digest(path)}))
    spec, args, plan = agent._validate_plan(ToolCall(name='read_docx_ranges', args={}), state)
    assert spec.access == 'read' and args == {'offset': 0}
    assert plan.file_versions == {str(path): digest(path)}
    assert agent.confirm is None and not editor.audit.exists()
    other = tmp_path / 'other.docx'
    ambiguous = professional_tools((path, other))
    other_agent = OfficeAgent(None, build_dispatcher(editor, path), ambiguous, tmp_path / 'state2', persist=False)
    with pytest.raises(ValueError, match='one explicitly approved file'):
        other_agent._validate_plan(ToolCall(name='read_docx_ranges', args={}), state)


def test_model_review_cannot_approve_and_only_once(setup):
    path, editor, service = setup
    data = view(editor, path)
    payload = {'source_sha256': digest(path), 'operations': [operation(data['items'][1], 7, 13, 'second')]}
    plan = service.plan_ranges(payload)['result']
    calls = []
    def respond(*args):
        calls.append(args)
        return {'verdict': 'concerns', 'reasons': ['请核对替换位置。'], 'requires_human_review': True}
    service.model_factory = lambda: SimpleNamespace(_request=respond)
    reply = service.review_ranges(plan['id'])
    assert reply['result']['writes_approved'] is False and reply['result']['unreviewed']
    assert service.review_ranges(plan['id'])['reason'] == 'professional_review_attempted'
    assert len(calls) == 1 and not editor.audit.exists()


def test_model_late_reply_invalid_after_scope_change(setup):
    path, editor, service = setup
    data = view(editor, path)
    plan = service.plan_ranges({'source_sha256': digest(path), 'operations': [operation(data['items'][1], 7, 13, 'second')]})['result']
    def respond(*args):
        service.invalidate()
        return {'verdict': 'consistent', 'reasons': ['未发现问题。'], 'requires_human_review': True}
    service.model_factory = lambda: SimpleNamespace(_request=respond)
    assert service.review_ranges(plan['id'])['ok'] is False
    assert not editor.audit.exists()


def test_comment_and_replacement_distinct_ranges_same_group(setup):
    path, editor, _ = setup
    data = view(editor, path)
    ops = [operation(data['items'][1], 0, 6, 'first'), operation(data['items'][1], 7, 13, '核对重复术语。', action='comment')]
    req = RangeDocx(path=str(path), source_sha256=digest(path), operations=ops, revision_date=now())
    before = _parts(path, range_edit=True)
    after, _, preview = transform(before, req)
    verify(before, after, req)
    root = ET.fromstring(after['word/document.xml'])
    assert 'first repeat' in _text(root)
    assert len(list(root.iter(W + 'commentRangeStart'))) == 1
    assert 'word/comments.xml' in after
    assert set(preview[0]['verification']['changed_parts']) == {'word/document.xml', 'word/comments.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'}
    assert before['word/styles.xml'] == after['word/styles.xml']


def test_same_anchor_insertion_order_and_duplicates(setup):
    path, editor, _ = setup
    data = view(editor, path)
    item = data['items'][5]
    for side in ('insert_before', 'insert_after'):
        ops = [operation(item, text=value, action=side) for value in ('Alpha', 'Beta')]
        req = RangeDocx(path=str(path), source_sha256=digest(path), operations=ops, revision_date=now())
        after, _, _ = transform(_parts(path, range_edit=True), req)
        values = [_text(p) for p in ET.fromstring(after['word/document.xml']).findall('.//' + W + 'body/' + W + 'p')]
        assert values.index('Alpha') + 1 == values.index('Beta')
        req.operations[1].text = 'Alpha'
        with pytest.raises(GuardError, match='overlapping_changes'):
            transform(_parts(path, range_edit=True), req)


def test_deleted_only_revision_visible_and_explicit_review_group(setup):
    from office_agent.advanced_writes import _zip
    from copy import deepcopy
    path, editor, _ = setup
    parts = _parts(path, range_edit=True)
    root = ET.fromstring(parts['word/document.xml'])
    body = root.find(W + 'body')
    paragraph = ET.Element(W + 'p')
    body.insert(max(0, len(body) - 1), paragraph)
    for kind, identity, text in [('del', '71', 'removed text'), ('ins', '72', 'replacement')]:
        wrapper = ET.SubElement(paragraph, W + kind, {W + 'id': identity, W + 'author': 'Reviewer'})
        ET.SubElement(ET.SubElement(wrapper, W + 'r'), W + ('delText' if kind == 'del' else 't')).text = text
    parts['word/document.xml'] = ET.tostring(root)
    target = path.with_name('revisions.docx')
    _zip(parts, target)
    data = view(editor, target)
    row = next(item for item in data['items'] if item['text'] == 'replacement')
    ids = [item['id'] for item in row['revisions']]
    req = RangeDocx(path=str(target), source_sha256=digest(target), revision_date=now(), operations=[operation(row, text='', action='revision_reject', revision_ids=ids)])
    after, _, _ = transform(parts, req)
    assert 'removed text' in _text(ET.fromstring(after['word/document.xml']))
    assert 'replacement' not in _text(ET.fromstring(after['word/document.xml']))
    paragraph.remove(paragraph.find(W + 'ins'))
    parts['word/document.xml'] = ET.tostring(root)
    other = path.with_name('deleted-only.docx')
    _zip(parts, other)
    data = view(editor, other)
    assert any(not item['text'] and any(r['text'] == 'removed text' for r in item.get('revisions', [])) for item in data['items'])


def test_redacted_metadata_cannot_reveal_text_through_structure_plan(setup):
    from docx import Document
    path, editor, service = setup
    doc = Document()
    doc.add_paragraph('Ordinary neighboring paragraph.')
    doc.styles.add_style('sk-syntheticmetadata', 1)
    doc.add_paragraph('password=synthetic-fixture-only', style='sk-syntheticmetadata')
    doc.save(path)
    version = digest(path)
    data = service.read_ranges()['result']
    assert 'sk-syntheticmetadata' not in json.dumps(data)
    hidden = data['items'][1]
    assert hidden['text'] == '' and hidden['refusal'] == 'sensitive_or_oversize'
    for item, action, extra in [(hidden, 'paragraph_format', {'paragraph_format': {'alignment': 'center'}}),
                                (hidden, 'insert_after', {'text': 'New text'}),
                                (data['items'][0], 'merge_next', {})]:
        op = operation(item, text='', action=action)
        op.update(extra)
        reply = service.plan_ranges({'source_sha256': version, 'operations': [op]})
        assert reply['ok'] is False
        assert 'synthetic-fixture-only' not in json.dumps(reply)
        assert not editor.plans and not editor.audit.exists()
    assert digest(path) == version
