from copy import deepcopy
import json

from docx import Document
from docx.oxml import OxmlElement
from lxml import etree as ET
import pytest

from office_agent.advanced_writes import _parts
from office_agent.professional_edits import W, _canonical
from office_agent.range_edits import static_render_parts, transform, verify, view
from office_agent.range_fields import analyze
from office_agent.writes import Editor, GuardError, digest
from tests.test_range_edits import operation, request
from tests.test_professional_edits import fake_render


def make_source(path):
    doc = Document()
    paragraph = doc.add_paragraph('实验条件：')
    start = OxmlElement('w:bookmarkStart')
    start.set(W + 'id', '7')
    start.set(W + 'name', '_RefExperiment')
    paragraph._p.append(start)
    paragraph.add_run('仅适用于干燥条件')
    end = OxmlElement('w:bookmarkEnd')
    end.set(W + 'id', '7')
    paragraph._p.append(end)
    paragraph = doc.add_paragraph('原结论：')
    field = OxmlElement('w:fldSimple')
    field.set(W + 'instr', ' REF _RefExperiment \\h ')
    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.text = '仅适用于干燥条件'
    run.append(text)
    field.append(run)
    paragraph._p.append(field)
    paragraph.add_run('，不能推广。')
    paragraph = doc.add_paragraph('参见第 ')
    for kind, value in [('begin', None), ('code', ' PAGE'), ('code', 'REF _RefExperiment \\h '), ('separate', None), ('text', '1'), ('end', None)]:
        run = OxmlElement('w:r')
        child = OxmlElement('w:instrText' if kind == 'code' else 'w:t' if kind == 'text' else 'w:fldChar')
        if kind in {'code', 'text'}:
            child.text = value
        else:
            child.set(W + 'fldCharType', kind)
        run.append(child)
        paragraph._p.append(run)
    paragraph.add_run(' 页，核对条件。')
    doc.add_paragraph('本段保持。')
    doc.save(path)


@pytest.fixture
def field_sample(tmp_path):
    path = tmp_path / 'cross-reference.docx'
    make_source(path)
    editor = Editor(tmp_path, require_confirmation=True)
    return path, editor, view(editor, path)


@pytest.mark.parametrize('mode', ['clean', 'tracked'])
def test_cross_reference_neighbor_preserved_and_reported(field_sample, mode):
    path, editor, data = field_sample
    assert data['items'][0]['reference_target_locked']
    assert [row['type'] for item in data['items'] for row in item['fields']] == ['REF', 'PAGEREF']
    item = data['items'][1]
    before = _parts(path, range_edit=True)
    req = request(field_sample, [operation(item, 0, 3, '研究结论')], mode=mode)
    after, _, preview = transform(before, req)
    verify(before, after, req)
    for tag in ('fldSimple', 'instrText', 'fldChar', 'bookmarkStart', 'bookmarkEnd'):
        a, b = [list(ET.fromstring(parts['word/document.xml']).iter(W + tag)) for parts in (before, after)]
        assert list(map(_canonical, a)) == list(map(_canonical, b))
    assert all(value == after[name] for name, value in before.items() if name != 'word/document.xml')
    assert preview[0]['verification']['fields']['cross_references'] == 2
    assert preview[0]['verification']['fields']['results_refreshed'] is False
    projected = ET.fromstring(static_render_parts(after)['word/document.xml'])
    assert not list(projected.iter(W + 'instrText')) and not list(projected.iter(W + 'fldChar'))
    assert '仅适用于干燥条件' in ''.join(projected.itertext())


@pytest.mark.parametrize('action', ['replace', 'delete', 'revision_accept'])
def test_referenced_target_cannot_be_rewritten(field_sample, action):
    item = field_sample[2]['items'][0]
    op = operation(item, 0, 5, '' if action != 'replace' else '条件为：', action=action)
    with pytest.raises(GuardError, match='reference_target_locked'):
        transform(_parts(field_sample[0], range_edit=True), request(field_sample, [op]))


def test_field_display_is_not_an_editable_island(field_sample):
    for item, start, end in [(field_sample[2]['items'][1], 4, 6), (field_sample[2]['items'][2], 4, 5)]:
        with pytest.raises(GuardError, match='protected_boundary'):
            transform(_parts(field_sample[0], range_edit=True), request(field_sample, [operation(item, start, end)]))


@pytest.mark.parametrize('instruction', ['DDEAUTO server topic item', 'REF _RefExperiment \\* MERGEFORMAT', 'REF _RefExperiment \\p', 'REF missing', 'PAGEREF _RefExperiment \\h \\h'])
def test_unsupported_or_unresolved_instruction_refused(field_sample, instruction):
    parts = _parts(field_sample[0], range_edit=True)
    root = ET.fromstring(parts['word/document.xml'])
    root.find('.//' + W + 'fldSimple').set(W + 'instr', instruction)
    parts['word/document.xml'] = ET.tostring(root)
    with pytest.raises(GuardError, match='professional_(field_scope|reference_target)'):
        static_render_parts(parts)


@pytest.mark.parametrize('mutation', ['duplicate', 'missing_end', 'end_before_start', 'unclosed_field'])
def test_malformed_bookmarks_and_fields_refused(field_sample, mutation):
    parts = _parts(field_sample[0], range_edit=True)
    root = ET.fromstring(parts['word/document.xml'])
    start, end = root.find('.//' + W + 'bookmarkStart'), root.find('.//' + W + 'bookmarkEnd')
    if mutation == 'duplicate':
        start.getparent().append(deepcopy(start))
    elif mutation == 'missing_end':
        end.getparent().remove(end)
    elif mutation == 'end_before_start':
        end.getparent().remove(end)
        start.getparent().insert(0, end)
    else:
        mark = list(root.iter(W + 'fldChar'))[-1]
        mark.getparent().remove(mark)
    parts['word/document.xml'] = ET.tostring(root)
    with pytest.raises(GuardError, match='professional_(field_scope|reference_target)'):
        analyze(parts)


def test_field_edit_plan_conflict_and_undo(field_sample, monkeypatch):
    path, editor, data = field_sample
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    req = request(field_sample, [operation(data['items'][1], 0, 3, '研究结论')])
    original = path.read_bytes()
    plan = editor.plan(req)
    assert not editor.audit.exists()
    assert json.loads(plan.after)['items'][0]['verification']['fields']['results_refreshed'] is False
    path.write_bytes(original + b' ')
    with pytest.raises(GuardError):
        editor.execute(plan.id, confirmed=True)
    path.write_bytes(original)
    plan = editor.plan(req)
    result = editor.execute(plan.id, confirmed=True)
    assert digest(path) == data['document']['source_sha256']
    editor.rollback(plan.id, confirmed=True)
    assert path.read_bytes() == original


def test_field_browser_explains_target_and_static_results(tmp_path, monkeypatch):
    from tests.test_professional_browser import _browser, STATIC
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, receipts = next(fixture)
        make_source(app.path)
        page.expose_function('fieldHost', lambda method, args: getattr(api, method)(*args))
        page.evaluate("() => { window.pywebview.api = new Proxy({}, {get: (_, method) => (...args) => window.fieldHost(method, args)}); }")
        page.add_script_tag(path=str(STATIC / 'range_edits.js'))
        page.locator('#range-load').click()
        page.wait_for_function("document.getElementById('range-target').options.length === 4")
        assert '禁止改写正文' in page.locator('#range-structures').inner_text()
        page.locator('#range-target').select_option(index=1)
        assert 'REF → _RefExperiment' in page.locator('#range-structures').inner_text()
        page.locator('#range-original').evaluate('(node) => {node.focus(); node.setSelectionRange(0, 3);}')
        page.locator('#range-text').fill('研究结论')
        page.locator('#range-reason').fill('只调整域外措辞')
        page.locator('#range-add').click()
        page.locator('#range-plan').click()
        page.wait_for_function("!document.getElementById('range-confirm').disabled")
        assert '未刷新或核实显示结果' in page.locator('#range-preview').inner_text()
        assert not app.editor.audit.exists()
    finally:
        fixture.close()
