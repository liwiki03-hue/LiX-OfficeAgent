from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.professional_edits import W, _rewrite
from office_agent.tracked_edits import _word_changes, _coalesced
from office_agent.writes import _apply_revision_decision, GuardError, digest
from test_professional_edits import parts
from test_research_assist import setup, request, complete, Model


@pytest.mark.parametrize('old,new,actions', [
    ('The method works.', 'The new method works.', ['ins']),
    ('The old method works.', 'The method works.', ['del']),
    ('The old method works.', 'The new method works.', ['del','ins']),
    ('导热性能稳定。', '导热性能较稳定。', ['ins']),
])
def test_word_diff_changes_only_local_tokens_and_roundtrips(old, new, actions):
    doc = Document()
    paragraph = doc.add_paragraph(old)
    approved = deepcopy(paragraph._p)
    _rewrite(approved, old, new)
    changes = _word_changes(paragraph._p.findall(W+'r'), approved.findall(W+'r'))
    assert [action for action, _ in changes if action != 'equal'] == actions
    tracked = deepcopy(paragraph._p)
    for run in list(tracked.findall(W+'r')):
        tracked.remove(run)
    for action, runs in changes:
        if action == 'equal':
            tracked.extend(runs)
        else:
            wrapper = ET.SubElement(tracked,W+action)
            for run in runs:
                if action == 'del':
                    for node in run.iter(W+'t'):
                        node.tag = W+'delText'
                wrapper.append(run)
    assert 'works' not in ''.join(node.text or '' for wrapper in tracked for node in wrapper.iter(W+'delText'))
    for decision, expected in [('accept_all', approved), ('reject_all', paragraph._p)]:
        body = ET.Element(W+'body', nsmap={'w':W[1:-1]})
        body.append(deepcopy(tracked))
        _apply_revision_decision(SimpleNamespace(element=SimpleNamespace(body=body)),decision)
        assert _coalesced(body[0]) == _coalesced(expected)


def test_split_word_across_formats_keeps_original_delete_properties():
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run('The con').bold = True
    paragraph.add_run('ductivity').italic = True
    paragraph.add_run(' remains stable.')
    approved = deepcopy(paragraph._p)
    _rewrite(approved, paragraph.text, 'The resistivity remains stable.')
    changes = _word_changes(paragraph._p.findall(W+'r'), approved.findall(W+'r'))
    removed = next(runs for action,runs in changes if action == 'del')
    assert ''.join(''.join(node.text or '' for node in run.iter(W+'t')) for run in removed) == 'conductivity'
    assert removed[0].find(W+'rPr').find(W+'b') is not None
    assert removed[1].find(W+'rPr').find(W+'i') is not None


def test_word_diff_refuses_nontext_runs_and_token_overflow():
    paragraph = Document().add_paragraph('plain')
    paragraph.add_run().add_break()
    with pytest.raises(GuardError, match='professional_revision_run_unsupported'):
        _word_changes(paragraph._p.findall(W+'r'), paragraph._p.findall(W+'r'))
    paragraph = Document().add_paragraph('词' * 1600)
    with pytest.raises(GuardError, match='professional_revision_capacity'):
        _word_changes(paragraph._p.findall(W+'r'), paragraph._p.findall(W+'r'))


def granular_plan(service):
    task = complete(service, request(service))
    suggestion = task['result']['suggestions'][0]
    return service.plan({'task_id': task['id'], 'tracked': True, 'granular': True,
        'edits':[{'locator_id':suggestion['locator_id'],'new_text':suggestion['new_text'],'reason':suggestion['reason'],'reviewed':True}]})


def test_actual_plan_word_revision_reopen_copy_and_undo(tmp_path, monkeypatch):
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch)
    original = digest(path)
    planned = granular_plan(service)
    assert planned['ok'], planned
    assert planned['result']['items'][0]['native_revisions']['scope'] == 'word_diff_with_character_style_projection'
    saved = service.confirm(planned['result']['id'], True)
    assert saved['result']['status'] == 'saved', saved
    output = Path(saved['result']['output'])
    root = ET.fromstring(parts(output)['word/document.xml'])
    assert len(list(root.iter(W+'ins'))) == 1 and not list(root.iter(W+'del'))
    original_parts = parts(path)
    assert all(data == parts(output)[name] for name,data in original_parts.items() if name != 'word/document.xml')
    assert digest(path) == original
    assert service.rollback(planned['result']['id'], True)['result']['status'] == 'rolled_back'


def test_over_100_markers_refuses_whole_plan_no_output(tmp_path, monkeypatch):
    def mutate(value):
        value['suggestions'][0]['new_text'] = ' '.join(f'new{index}' for index in range(60))
        return value
    service, grounded, path, record, auth, busy, model = setup(tmp_path, monkeypatch, Model(mutate))
    doc = Document(path)
    doc.paragraphs[0].text = ' '.join(f'old{index}' for index in range(60))
    doc.save(path)
    original = digest(path)
    result = granular_plan(service)
    assert not result['ok'] and result['reason'] == 'professional_revision_capacity'
    assert len(list(tmp_path.glob('*.docx'))) == 1 and digest(path) == original
