import json
from pathlib import Path
import threading
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.output_dependencies import OutputDependencies
from office_agent.webui.structured import Structured
from office_agent.writes import Editor, GuardError
from test_professional_edits import fake_render


def family(tmp_path, monkeypatch):
    source = tmp_path / 'original.docx'
    doc = Document()
    doc.add_paragraph('Original anchor.')
    doc.save(source)
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    selected = [source]
    lock = threading.RLock()
    scope = lambda: (True, False, (selected[0], tmp_path))
    manager = OutputDependencies()
    def service():
        result = Structured(lock, scope, lambda: editor, lambda: selected[0])
        result.dependencies = manager
        result.grounded = SimpleNamespace(literature=SimpleNamespace(pin=lambda *args: None, unpin=lambda *args: None, _dependencies={}))
        return result
    def save(result, text):
        row = result.read_structure()['result']
        plan = result.plan_structure({'source_sha256': row['document']['source_sha256'],
            'insertions': [{'anchor': row['anchors'][0]['locator_id'], 'side': 'after', 'text': text}]})
        assert plan['ok'], plan
        saved = result.confirm(plan['result']['id'], True)
        assert saved['result']['status'] == 'saved', saved
        return plan['result']['id'], Path(saved['result']['output'])
    first = service()
    parent, output = save(first, 'Parent supplement.')
    selected[0] = output
    second = service()
    child, child_output = save(second, 'Child supplement.')
    return editor, selected, source, manager, first, second, parent, child, output, child_output


def test_dependency_checked_before_rollback_and_child_first_unblocks(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    selected[0] = source
    preview = manager.preview(editor, parent, [])
    assert preview['blockers'] == [{'receipt_id': child, 'reason': 'source_output', 'status': 'verified'}]
    assert preview['known_receipts_only'] and not preview['complete_dependency_graph']
    assert not preview['write_approval_granted'] and not preview['restores_rollback_authority']
    assert first.rollback(parent, True)['reason'] == 'professional_output_in_use'
    assert output.exists() and child_output.exists() and not first._disabled
    selected[0] = output
    assert second.rollback(child, True)['result']['status'] == 'rolled_back'
    selected[0] = source
    assert not manager.preview(editor, parent, [])['blockers']
    assert first.rollback(parent, True)['result']['status'] == 'rolled_back'
    assert not output.exists() and source.exists()


@pytest.mark.parametrize('tamper', ['missing', 'source', 'fake_rollback'])
def test_unreadable_or_changed_child_retains_blocker(tmp_path, monkeypatch, tamper):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    path = editor.audit / child / 'audit.json'
    if tamper == 'missing':
        path.unlink()
    else:
        audit = json.loads(path.read_bytes())
        if tamper == 'source':
            audit['source'] = str(source)
        else:
            audit['status'] = 'rolled_back'
        path.write_text(json.dumps(audit), encoding='utf8')
    selected[0] = source
    reply = first.rollback(parent, True)
    assert reply['reason'] == 'professional_output_in_use', reply
    assert output.exists() and child_output.exists()


def test_explicit_receipts_fresh_manager_has_no_undo_authority_or_scanning(tmp_path, monkeypatch):
    editor, selected, source, manager, first, second, parent, child, output, child_output = family(tmp_path, monkeypatch)
    fresh = OutputDependencies()
    monkeypatch.setattr(Path, 'iterdir', lambda *args: pytest.fail('No directory enumeration permitted'))
    preview = fresh.preview(editor, parent, [child])
    assert preview['blockers'][0]['receipt_id'] == child and len(preview['records']) == 2
    assert not preview['restores_rollback_authority']
    with pytest.raises(GuardError, match='professional_output_in_use'):
        fresh.check(editor, parent)
    assert str(tmp_path) not in json.dumps(preview)
    for receipts in [[True], [child, child], ['../audit'], [str(i) for i in range(9)]]:
        with pytest.raises(GuardError, match='professional_dependency_invalid'):
            fresh.preview(editor, parent, receipts)
