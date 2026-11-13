from copy import deepcopy
import json
import threading
from pathlib import Path
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree as ET
import pytest

from office_agent.professional_edits import ConstrainedDocx, snapshot, transform, verify
from office_agent.writes import Editor, GuardError, digest
from office_agent.webui.professional import Professional


@pytest.fixture
def sample(tmp_path):
    path = tmp_path / 'experiment.docx'
    document = Document()
    document.add_heading('实验结果', 1)
    paragraph = document.add_paragraph()
    paragraph.add_run('样本30个，误差为').bold = True
    paragraph.add_run('12.5% [1]，计算采用 $x=2$。').italic = True
    document.add_paragraph('The method is robust and repeatable.')
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).merge(table.cell(0, 1)).text = '合并表头'
    table.cell(1, 0).text = '温度 20 °C'
    table.cell(1, 1).text = 'Non-target cell'
    document.save(path)
    return path, Editor(tmp_path, require_confirmation=True)


def payload(editor, path, *, index=1, new_text=None):
    data = snapshot(editor, path)
    item = data['items'][index]
    return {'source_sha256': data['document']['source_sha256'],
            'changes': [{'locator_id': item['locator_id'], 'expected_text_sha256': item['text_sha256'],
                         'new_text': new_text or item['text'].replace('误差为', '测得误差为'),
                         'reason': 'Improve wording without changing data', 'evidence_ids': [item['evidence']['id']]}]}


def parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def fake_render(*args, **kwargs):
    return {'render': {'status': 'different', 'equivalent': False, 'source_pages': 1, 'output_pages': 1},
            'inputs_unchanged': True, 'evidence': 'synthetic renderer result, not actual visual validation'}


def test_source_records_and_distinct_locator_domains(sample):
    path, editor = sample
    data = snapshot(editor, path, limit=2)
    assert data['document']['parser_version'] == 'docx-structural-edit-v1'
    assert data['next_offset'] == 2
    item = data['items'][1]
    assert item['locator_id'].startswith('loc:')
    assert item['evidence']['id'].startswith('ev:')
    assert item['evidence']['source_sha256'] == digest(path)
    assert item['evidence']['page'] is None
    assert item['evidence']['heading_path'] == ['实验结果']
    assert item['evidence']['summary'] is None


@pytest.mark.parametrize('before,after,code', [
    ('12.5%', '13.5%', 'numbers'), ('30个', '31个', 'numbers'),
    ('[1]', '[2]', 'numbers'), ('$x=2$', '$x=3$', 'numbers'),
    ('[1]', '[Smith]', 'numbers'), ('$x=2$', '$y=2$', 'formulas')])
def test_protected_content_refuses_entire_plan(sample, before, after, code):
    path, editor = sample
    original = path.read_bytes()
    value = payload(editor, path)
    value['changes'][0]['new_text'] = snapshot(editor, path)['items'][1]['text'].replace(before, after)
    with pytest.raises(GuardError, match='protected_' + code):
        editor.plan(ConstrainedDocx(**value, path=str(path)))
    assert path.read_bytes() == original and not editor.audit.exists()


def test_exact_protection_literal_and_citation(sample):
    path, editor = sample
    data = payload(editor, path, index=2, new_text='The approach is robust and repeatable.')
    data['protected_literals'] = ['method']
    with pytest.raises(GuardError, match='protected_literal'):
        editor.plan(ConstrainedDocx(**data, path=str(path)))
    document = Document(path)
    document.add_paragraph('See [Smith] for the method.')
    document.save(path)
    data = payload(editor, path, index=len(snapshot(editor, path)['items']) - 1, new_text='See [Jones] for the method.')
    with pytest.raises(GuardError, match='protected_citations'):
        editor.plan(ConstrainedDocx(**data, path=str(path)))


def test_overlap_stale_locator_and_unknown_evidence(sample):
    path, editor = sample
    data = payload(editor, path)
    data['changes'].append(deepcopy(data['changes'][0]))
    with pytest.raises(GuardError, match='overlapping'):
        editor.plan(ConstrainedDocx(**data, path=str(path)))
    data = payload(editor, path)
    data['changes'][0]['evidence_ids'] = ['ev:' + 'a' * 64]
    with pytest.raises(GuardError, match='evidence_missing'):
        editor.plan(ConstrainedDocx(**data, path=str(path)))
    data = payload(editor, path)
    data['changes'][0]['expected_text_sha256'] = '0' * 64
    with pytest.raises(GuardError, match='stale_locator'):
        editor.plan(ConstrainedDocx(**data, path=str(path)))


@pytest.mark.parametrize('special', ['bookmarkStart', 'fldSimple', 'hyperlink', 'ins', 'drawing', 'footnoteReference'])
def test_special_targets_are_not_flattened(sample, special):
    path, editor = sample
    doc = Document(path)
    paragraph = doc.paragraphs[2]
    if special in {'drawing', 'footnoteReference'}:
        paragraph.runs[0]._r.append(OxmlElement('w:' + special))
    else:
        paragraph._p.append(OxmlElement('w:' + special))
    doc.save(path)
    try:
        data = snapshot(editor, path)
    except GuardError:
        assert not editor.audit.exists()
        return
    item = data['items'][2]
    assert not item['writable'] and item['evidence']['writable_locator_id'] is None
    request = payload(editor, path, index=2, new_text='An improved statement.')
    with pytest.raises(GuardError, match='unsupported_target'):
        editor.plan(ConstrainedDocx(**request, path=str(path)))
    assert not editor.audit.exists()


def test_group_edit_preserves_runs_table_and_rollback(sample, monkeypatch):
    from office_agent import advanced_writes
    monkeypatch.setattr(advanced_writes, 'render_gate', fake_render)
    path, editor = sample
    original = path.read_bytes()
    data = payload(editor, path)
    row = next(item for item in snapshot(editor, path)['items'] if item['text'] == '温度 20 °C')
    data['changes'].append({'locator_id': row['locator_id'], 'expected_text_sha256': row['text_sha256'],
                            'new_text': '测试温度 20 °C', 'reason': 'Clarify measurement', 'evidence_ids': [row['evidence']['id']]})
    plan = editor.plan(ConstrainedDocx(**data, path=str(path)))
    assert plan.confirmation_required and not editor.audit.exists()
    with pytest.raises(GuardError, match='confirmation'):
        editor.execute(plan.id, confirmed=False)
    result = editor.execute(plan.id, confirmed=True)
    output = Path(result['output'])
    assert path.read_bytes() == original
    before, after = parts(path), parts(output)
    assert {key for key in before if before[key] != after[key]} == {'word/document.xml'}
    source_doc, output_doc = Document(path), Document(output)
    assert source_doc.paragraphs[1].runs[0].bold == output_doc.paragraphs[1].runs[0].bold
    assert source_doc.paragraphs[1].runs[1].italic == output_doc.paragraphs[1].runs[1].italic
    assert output_doc.tables[0].cell(0, 0)._tc.xml == source_doc.tables[0].cell(0, 0)._tc.xml
    assert output_doc.tables[0].cell(1, 1).text == 'Non-target cell'
    assert output_doc.tables[0].cell(1, 0).text == '测试温度 20 °C'
    assert result['status'] == 'verified'
    assert editor.rollback(plan.id, confirmed=True)['status'] == 'rolled_back'
    assert not output.exists() and path.read_bytes() == original


def test_post_preview_source_change_refused(sample):
    path, editor = sample
    plan = editor.plan(ConstrainedDocx(**payload(editor, path), path=str(path)))
    doc = Document(path)
    doc.add_paragraph('Changed externally')
    doc.save(path)
    with pytest.raises(GuardError, match='source changed'):
        editor.execute(plan.id, confirmed=True)
    assert not editor.audit.exists()


def test_non_target_tampering_is_detected(sample):
    path, editor = sample
    request = ConstrainedDocx(**payload(editor, path), path=str(path))
    before = parts(path)
    after, _, _ = transform(before, request)
    after['word/styles.xml'] += b' '
    with pytest.raises(GuardError, match='non_target_part'):
        verify(before, after, request)


def test_authenticated_receipts_expiry_and_unknown_write(sample, monkeypatch):
    from office_agent import advanced_writes
    monkeypatch.setattr(advanced_writes, 'render_gate', fake_render)
    path, editor = sample
    logged = [True]
    facade = Professional(threading.RLock(), lambda: (logged[0], False, 'fixed'), lambda: editor, lambda: path)
    first = facade.plan(payload(editor, path))['result']
    assert facade.confirm(first['id'], False)['result']['wrote_nothing']
    assert not editor.audit.exists()
    second = facade.plan(payload(editor, path))['result']
    facade._pending[second['id']]['created'] -= 301
    assert facade.confirm(second['id'], True)['reason'] == 'professional_plan_expired'
    third = facade.plan(payload(editor, path))['result']
    logged[0] = False
    assert facade.confirm(third['id'], True)['reason'] == 'not_authenticated'
    logged[0] = True
    fourth = facade.plan(payload(editor, path))['result']
    monkeypatch.setattr(editor, 'execute', lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError('private request data')))
    outcome = facade.confirm(fourth['id'], True)
    assert outcome['result']['status'] == 'uncertain'
    assert 'private request' not in json.dumps(outcome)
    assert facade.read()['reason'] == 'professional_uncertain'
