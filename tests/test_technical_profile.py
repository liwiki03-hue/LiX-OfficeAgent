from copy import deepcopy
from pathlib import Path
import shutil

import pytest

from office_agent.document_tasks import compile_task, load, group_request
from office_agent.range_edits import _islands, transform
from office_agent.technical_profile import observations, report, verify, RULE_SHA256, PROFILE
from office_agent.writes import Editor, GuardError, digest
from docx import Document


@pytest.fixture
def material(tmp_path):
    source = tmp_path / 'technical.docx'
    shutil.copyfile(Path(__file__).resolve().parents[1] / 'deliverables/technical-editing-20261025/BASELINE-MATERIAL.docx', source)
    return source, Editor(tmp_path, require_confirmation=True)


def spec(editor, path, index, find, replacement, *, profile=PROFILE, mode='clean', **extra):
    version, _, _, entries = load(editor, path)
    entry = entries[index]
    start = entry['text'].index(find)
    island = next(row for row in _islands(entry) if row['start'] <= start < start + len(find) <= row['end'])
    return {'source_sha256': version, 'request': '只改措辞，保留单位条件和否定', 'mode': mode, 'editing_profile': profile,
            'requirements': [{'key': 'technical', 'kind': 'ranges', 'scope': 'selected', 'locators': [entry['locator_id']],
                              'goal': '受约束措辞修改', 'acceptance': '保留识别的字面项；语义仍需人工核对',
                              'operations': [{'locator_id': entry['locator_id'], 'expected_text_sha256': entry['text_sha256'],
                                              'island_id': island['id'], 'action': 'replace', 'start': start, 'end': start + len(find),
                                              'quote': find, 'text': replacement, 'reason': '仅调整措辞'}]}], **extra}


@pytest.mark.parametrize('index,find,new', [(0, '仅限室内', '适用于室内外'), (0, 'mL/min', 'mL/s'), (1, '不得', '可以'), (2, 'no', 'a')])
def test_profile_blocks_unsafe_proposals_that_general_profile_does_not_claim_to_check(material, index, find, new):
    path, editor = material
    raw = spec(editor, path, index, find, new)
    guarded = compile_task(editor, path, raw)
    assert guarded['items'][0]['state'] == 'restricted'
    assert guarded['items'][0]['reason'].startswith('professional_profile_changed_')
    ordinary = compile_task(editor, path, {**raw, 'editing_profile': 'general'})
    assert ordinary['items'][0]['state'] == 'pending'
    assert not editor.audit.exists()


@pytest.mark.parametrize('mode', ['clean', 'tracked'])
def test_safe_wording_carries_same_policy_through_plan_and_native_projection(material, mode):
    path, editor = material
    original = digest(path)
    task = compile_task(editor, path, spec(editor, path, 0, '结果', '测得结果', mode=mode))
    assert task['profile_report']['rule_sha256'] == RULE_SHA256
    assert any(row['literal'] == '20 mL/min' for row in task['profile_report']['items'])
    request = group_request(editor, path, task)
    assert request.editing_profile == PROFILE and request.profile_rule_sha256 == RULE_SHA256
    parts = load(editor, path)[1]
    after, before, changes = transform(parts, request)
    assert after != parts and digest(path) == original
    assert changes[0]['verification']['mode'] == mode
    tampered = request.model_copy(update={'profile_rule_sha256': '0' * 64})
    with pytest.raises(GuardError, match='professional_profile_version'):
        transform(parts, tampered)


def test_added_qualifier_capacity_unknown_units_and_manual_literals_are_explicit(material):
    with pytest.raises(GuardError, match='changed_qualifier'):
        verify('The result is available.', 'The result is not available.', PROFILE, RULE_SHA256)
    with pytest.raises(GuardError, match='profile_capacity'):
        observations('；'.join(f'{n} mL/min' for n in range(31)))
    with pytest.raises(GuardError, match='literal_missing'):
        report(['目标文字'], ['原文不存在的责任主体'])
    assert observations('20 furlongs/fortnight') == []
    path, editor = material
    raw = spec(editor, path, 0, '结果', '测得结果', protected_literals=['临床决策'])
    with pytest.raises(GuardError, match='literal_missing'):
        compile_task(editor, path, raw)
    raw['protected_literals'] = ['仅限室内']
    assert compile_task(editor, path, raw)['items'][0]['state'] == 'pending'
    raw['profile_rule_sha256'] = '0' * 64
    with pytest.raises(GuardError, match='profile_version'):
        compile_task(editor, path, raw)


def test_combined_individually_safe_edits_cannot_introduce_a_new_qualifier(tmp_path):
    path = tmp_path / 'combined.docx'
    document = Document(); document.add_paragraph('The pox is recorded.'); document.save(path)
    editor = Editor(tmp_path, require_confirmation=True)
    raw = spec(editor, path, 0, 'p', 'n')
    entry = load(editor, path)[3][0]
    second = deepcopy(raw['requirements'][0]['operations'][0])
    second.update(start=6, end=7, quote='x', text='t')
    raw['requirements'][0]['operations'].append(second)
    task = compile_task(editor, path, raw)
    assert all(row['state'] == 'pending' for row in task['items'])
    with pytest.raises(GuardError, match='professional_profile_changed_qualifier'):
        group_request(editor, path, task)
    assert not editor.audit.exists()


def test_profile_archive_restores_policy_without_approval(material):
    from threading import RLock
    from office_agent.output_dependencies import OutputDependencies
    from office_agent.webui.document_tasks import DocumentTasks
    path, editor = material
    service = DocumentTasks(RLock(), lambda: (True, False, str(path)), lambda: editor, lambda: path, OutputDependencies())
    created = service.create({'request': '将全文“结果”替换为“观测值”', 'editing_profile': PROFILE, 'protected_literals': ['仅限室内']})
    assert created['ok'], created
    identifier = created['result']['task_id']
    planned = service.plan_archive(identifier)
    assert planned['ok'], planned
    saved = service.confirm_archive(planned['result']['plan_id'], True)
    assert saved['ok'], saved
    service.invalidate()
    restored = service.restore(saved['result']['save_id'])
    assert restored['ok'], restored
    policy = restored['result']['protection_profile']
    assert policy['rule_sha256'] == RULE_SHA256 and policy['manual_literals'] == ['仅限室内']
    assert any(row['literal'] == '20 mL/min' for row in policy['items'])
    assert not restored['result']['write_approval_restored'] and not restored['result']['rollback_authority_restored']
    assert service.undo_last(identifier, True)['reason'] == 'professional_plan_invalid'


def test_profile_routes_only_to_policy_aware_transaction(material):
    from office_agent.document_tasks import natural_spec
    from office_agent.task_routing import candidate
    path, editor = material
    raw = natural_spec(editor, path, '将全文“结果”替换为“观测值”').model_dump()
    raw['editing_profile'] = PROFILE
    task = compile_task(editor, path, raw)
    route, native = candidate(editor, path, task)
    assert route['selected'] == 'bounded_ranges' and native is None
    assert '技术报告' in route['reason']
