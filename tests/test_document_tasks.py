from copy import deepcopy
from pathlib import Path
import shutil
from threading import RLock

from docx import Document
import pytest

from office_agent.document_tasks import compile_task, natural_spec, navigation, public_task
from office_agent.output_dependencies import OutputDependencies
from office_agent.webui.document_tasks import DocumentTasks
from office_agent.writes import Editor, digest
from tests.test_professional_edits import fake_render

ROOT = Path(__file__).resolve().parents[1]


def prepared(tmp_path, monkeypatch):
    source = tmp_path / 'source.docx'
    shutil.copyfile(ROOT / 'deliverables/docx-tasks-20261021/development.docx', source)
    selected, authenticated = [source], [True]
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    service = DocumentTasks(RLock(), lambda: (authenticated[0], False, str(selected[0])), lambda: editor, lambda: selected[0], OutputDependencies())
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    result = service.create('将全文所有“旧术语”替换为“规范术语”，保留数字、单位、公式和引文')
    assert result['ok'], result
    return source, selected, editor, service, result['result'], authenticated


def execute(service, identifier):
    plan = service.plan_group(identifier)
    assert plan['ok'], plan
    result = service.confirm_group(identifier, plan['result']['id'], True)
    assert result['ok'] and result['result']['status'] == 'saved', result
    return result['result']['task']


def test_all_targets_beyond_first_screen_three_groups_and_omission_check(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    original = digest(source)
    assert task['total_items'] == 31 and task['group_sizes'] == [12, 12, 7]
    assert not task['deterministic_complete'] and task['verification_checks'][0]['remaining'] == 31
    nav = service.navigate({'query': '旧术语', 'offset': 20, 'limit': 20})['result']
    assert nav['matching_targets'] == 31 and len(nav['items']) == 11
    assert len(nav['outline']) == 3 and nav['scan_complete']
    identifier = task['task_id']
    for count in (19, 7, 0):
        task = execute(service, identifier)
        assert task['verification_checks'][0]['remaining'] == count
    assert task['deterministic_complete'] and task['states'] == {'completed': 31}
    output = tmp_path / task['current_name']
    doc = Document(output)
    assert sum(p.text.count('规范术语') for p in doc.paragraphs) == 30
    assert doc.tables[0].cell(1, 0).text == '规范术语'
    assert digest(source) == original
    assert len(task['receipts']) == 3 and not task['multi_transaction_atomic']
    for _ in range(3):
        result = service.undo_last(identifier, True)
        assert result['ok'] and result['result']['status'] == 'rolled_back', result
    assert not output.exists() and digest(source) == original


def test_overlapping_requirements_refused_before_any_group(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    spec = natural_spec(editor, source, '将全文“旧术语”替换为“规范术语”').model_dump()
    second = deepcopy(spec['requirements'][0])
    second.update(key='contradiction', replacement='另一术语')
    spec['requirements'].append(second)
    reply = service.create(spec)
    assert reply['reason'] == 'professional_task_conflicting_requirements'
    assert not editor.audit.exists()


def test_math_text_is_accounted_as_unsupported_not_silent_success(tmp_path):
    from docx.oxml import OxmlElement
    path = tmp_path / 'math.docx'
    doc = Document()
    paragraph = doc.add_paragraph('正文保持。')
    math = OxmlElement('m:oMath')
    run = OxmlElement('m:r')
    text = OxmlElement('m:t')
    text.text = '旧术语'
    run.append(text)
    math.append(run)
    paragraph._p.append(math)
    doc.save(path)
    editor = Editor(tmp_path)
    spec = natural_spec(editor, path, '将全文“旧术语”替换为“规范术语”')
    task = public_task(compile_task(editor, path, spec.model_dump()))
    assert task['states'] == {'restricted': 1}
    assert task['verification_checks'][0]['remaining'] == 1 and not task['deterministic_complete']


def test_long_task_reserves_output_name_budget_before_first_write(tmp_path):
    from office_agent.writes import GuardError
    path = tmp_path / ('a' * 105 + '.docx')
    doc = Document()
    for _ in range(96):
        doc.add_paragraph('旧术语')
    doc.save(path)
    editor = Editor(tmp_path)
    spec = natural_spec(editor, path, '将全文“旧术语”替换为“规范术语”')
    with pytest.raises(GuardError, match='professional_task_filename_budget'):
        compile_task(editor, path, spec.model_dump())
    assert not editor.audit.exists()


def test_cancel_exclusion_conflict_and_single_use_plan(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    identifier = task['task_id']
    plan = service.plan_group(identifier)['result']
    assert service.confirm_group(identifier, plan['id'], False)['result']['status'] == 'cancelled'
    assert not editor.audit.exists()
    assert service.confirm_group(identifier, plan['id'], True)['reason'] == 'professional_plan_invalid'
    excluded = service.exclude(identifier, task['items'][0]['id'], '不修改此处')
    assert excluded['ok'] and excluded['result']['states']['excluded'] == 1
    assert not excluded['result']['deterministic_complete']
    plan = service.plan_group(identifier)['result']
    source.write_bytes(source.read_bytes() + b'\n')
    assert service.confirm_group(identifier, plan['id'], True)['reason'] == 'professional_source_changed'
    assert not editor.audit.exists()


def test_task_archive_resume_requires_current_admission_and_no_undo_grant(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    identifier = task['task_id']
    task = execute(service, identifier)
    plan = service.plan_archive(identifier)
    assert plan['ok'], plan
    saved = service.confirm_archive(plan['result']['plan_id'], True)
    assert saved['ok'], saved
    receipt = saved['result']['save_id']
    service.invalidate()
    assert service.restore(receipt)['reason'] == 'professional_task_reselect_current'
    selected[0] = tmp_path / task['current_name']
    restored = service.restore(receipt)
    assert restored['ok'], restored
    assert not restored['result']['write_approval_restored'] and not restored['result']['rollback_authority_restored']
    assert service.undo_last(identifier, True)['reason'] == 'professional_plan_invalid'
    task = execute(service, identifier)
    task = execute(service, identifier)
    assert task['deterministic_complete']
    auth[0] = False
    assert service.status(identifier)['reason'] == 'not_authenticated'
    assert not service.tasks and not service.writers


def test_read_only_matches_and_extra_request_are_not_silent_success(tmp_path):
    source = tmp_path / 'header.docx'
    doc = Document()
    doc.add_paragraph('正文旧术语')
    doc.sections[0].header.paragraphs[0].text = '页眉旧术语'
    doc.save(source)
    editor = Editor(tmp_path)
    spec = natural_spec(editor, source, '将全文“旧术语”替换为“规范术语”')
    task = compile_task(editor, source, spec.model_dump())
    assert public_task(task)['states'] == {'pending': 1, 'restricted': 1}
    assert task['checks'][0]['remaining'] == 2


def test_mixed_text_and_format_are_ordered_as_independent_groups(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    nav = navigation(editor, source)
    locator = nav['items'][2]['locator_id']
    spec = natural_spec(editor, source, '将全文“旧术语”替换为“规范术语”').model_dump()
    spec['requirements'].append({'key': 'format', 'goal': '指定段落段后间距', 'kind': 'format', 'scope': 'selected',
        'locators': [locator], 'paragraph_format': {'after_pt': 12}, 'acceptance': '仅指定段落段后12磅'})
    reply = service.create(spec)
    assert reply['ok'], reply
    identifier = reply['result']['task_id']
    while service.status(identifier)['result']['group_index'] < service.status(identifier)['result']['groups_total']:
        task = execute(service, identifier)
    assert task['deterministic_complete'], task
    assert Document(tmp_path / task['current_name']).paragraphs[2].paragraph_format.space_after.pt == 12


def test_task_block_mapping_and_independent_undo(tmp_path, monkeypatch):
    from tests.test_block_edits import sample
    path, editor, data = sample(tmp_path)
    service = DocumentTasks(RLock(), lambda: (True, False, str(path)), lambda: editor, lambda: path, OutputDependencies())
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    first = data['items'][0]
    spec = {'source_sha256': digest(path), 'request': '将第一章完整移到第三章前', 'requirements': [{
        'key': 'move', 'goal': '完整普通章节移动', 'kind': 'blocks', 'scope': 'document', 'acceptance': '节点身份和非目标内容保持',
        'block_operations': [{'action': 'chapter_move', 'locator_id': first['locator_id'],
            'locator_ids': [row['locator_id'] for row in data['items'][:3]], 'destination_locator': data['items'][5]['locator_id'], 'reason': '明确章节次序'}]}]}
    created = service.create(spec)
    assert created['ok'], created
    task = execute(service, created['result']['task_id'])
    assert task['deterministic_complete'], task
    assert Document(tmp_path / task['current_name']).paragraphs[0].text == '第二章'
    undone = service.undo_last(task['task_id'], True)
    assert undone['ok'], undone
    assert undone['result']['task']['current_name'] == path.name
    assert Document(path).paragraphs[0].text == '第一章'


def test_new_tools_reachable_but_preview_only(tmp_path):
    from office_agent.app import professional_tools, select_tools, build_dispatcher
    doc = Document()
    doc.add_paragraph('旧术语')
    path = tmp_path / 'agent.docx'
    doc.save(path)
    editor = Editor(tmp_path)
    selected = select_tools(professional_tools((path,)), 'word')
    assert {'read_docx_task_navigation', 'plan_docx_task', 'plan_docx_blocks'} <= {tool.name for tool in selected}
    dispatch = build_dispatcher(editor, path)
    data = dispatch('read_docx_task_navigation', {})
    spec = natural_spec(editor, path, '将全文“旧术语”替换为“规范术语”')
    result = dispatch('plan_docx_task', spec.model_dump())
    assert result['preview_only'] and not editor.audit.exists()
    assert len(editor._document_task_previews) == 1


def test_agent_execution_uses_exact_path_for_navigation_task_and_blocks(tmp_path):
    from office_agent.agent import AgentState, OfficeAgent, RequiredState, ToolCall
    from office_agent.app import professional_tools, build_dispatcher
    doc = Document()
    doc.add_paragraph('旧术语仅限室内。')
    doc.add_paragraph('普通条件保持。')
    path = tmp_path / 'agent.docx'
    doc.save(path)
    editor = Editor(tmp_path)
    agent = OfficeAgent(None, build_dispatcher(editor, path), professional_tools((path,)), tmp_path / 'state', persist=False, confirm=lambda plan: True)
    state = AgentState(required=RequiredState(goal='完整文档编辑', file_versions={str(path): digest(path)}))
    read = ToolCall(name='read_docx_task_navigation', args={})
    _, _, bound = agent._validate_plan(read, state)
    assert bound.file_versions == {str(path): digest(path)}
    agent._execute(read, state)
    assert state.required.evidence[-1].tool == read.name
    spec = natural_spec(editor, path, '将全文“旧术语”替换为“规范术语”').model_dump()
    agent._execute(ToolCall(name='plan_docx_task', args=spec), state)
    assert state.required.pending and len(editor._document_task_previews) == 1
    assert not editor.plans and not editor.audit.exists()
    nav = navigation(editor, path)
    operation = {'action': 'paragraphs_replace', 'locator_id': nav['items'][0]['locator_id'],
                 'locator_ids': [row['locator_id'] for row in nav['items']], 'new_texts': ['旧术语仅限室内使用。', '普通条件保持。'], 'reason': '明确普通跨段编辑'}
    agent._execute(ToolCall(name='plan_docx_blocks', args={'source_sha256': digest(path), 'operations': [operation]}), state)
    assert len(editor.plans) == 1 and not editor.audit.exists()
    ambiguous = OfficeAgent(None, build_dispatcher(editor, path), professional_tools((path, tmp_path / 'other.docx')), tmp_path / 'state2', persist=False)
    with pytest.raises(ValueError, match='one explicitly approved file'):
        ambiguous._validate_plan(read, state)
