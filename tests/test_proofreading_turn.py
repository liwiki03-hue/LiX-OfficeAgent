import pytest

from office_agent.paged_documents import GuardError
from office_agent.skills_backend import SkillsBackend
from office_agent.webui.draft_suggestions import check_proofreading
from office_agent.webui.office_workflows import workflow_package
from tests.test_document_conversation import Answers, attach
from tests.test_paged_editor import draft, sync


def test_explicit_proofreading_requires_reviewed_package_and_same_selection(draft):
    page, floating, source = draft
    attach(floating, Answers())
    sync(page)
    backend = SkillsBackend(source.parent)
    package = workflow_package('chinese-proofread')
    install = backend.plan_public(package, 'session')
    backend.confirm_change(install['plan_id'], 'session', True)
    choice, = backend.turn_choices()
    ticket = {key: choice[key] for key in ('id', 'sha256')}
    target = {'node_id': 'node-1', 'start': 0, 'end': 2}
    ref = page.reference(page.document_id, page.revision, target)['reference']
    row, _ = floating.conversation.receive('explicit', '把选中这段改得简洁一些', page.document_id, page.revision, ref,
                                           resources={'skill': ticket})
    assert row['intent'] == 'edit' and page._proofreading('explicit', ref['id'])['workflow']=='chinese-proofread'
    new_ref = page.reference(page.document_id, page.revision, {**target, 'node_id': 'node-2'})['reference']
    with pytest.raises(GuardError, match='原校对选区'):
        page.submit('explicit', row['question'], new_ref['id'], page.document_id, page.source_version, '', page.inbox_id)
    assert not floating.outputs and not page.proposals
    edit = backend.plan_change('edit', {'source':'user','name':package['slug'],
                                       'content':'---\nname: lix-chinese-proofread\ndescription: 本地补充\n---\n自行校对。'}, 'session')
    backend.confirm_change(edit['plan_id'], 'session', True)
    changed, = backend.turn_choices()
    with pytest.raises(ValueError, match='先移除'):
        floating.conversation.receive('edited-skill', row['question'], page.document_id, page.revision, new_ref,
                                      resources={'skill': {'id': changed['id'], 'sha256': changed['sha256']}})
    assert 'edited-skill' not in floating.conversation.turns


@pytest.mark.parametrize('before,after', [
    ('延迟为 120 ms。', '延迟为 120 s。'),
    ('本次未验证断网场景 [3]。', '本次已验证断网场景 [3]。'),
    ('只测量了 20 个样本。', '只测量了 30 个样本。'),
])
def test_proofreading_refuses_changed_facts(before, after):
    with pytest.raises(GuardError):
        check_proofreading(before, after)
