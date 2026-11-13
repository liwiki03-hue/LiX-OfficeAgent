"""Bounded prompt disclosure with protocol fixtures; no model/provider execution."""
from types import SimpleNamespace

import pytest

from office_agent.agent import OfficeAgent, _SYSTEM
from office_agent.skill_context import SkillContext, MAX_CATALOG_CHARS, MAX_BODY_CHARS


def fixture():
    entries = [{'name': 'document-format', 'description': '文档排版 format document', 'source': 'user', 'enabled': True}]
    loaded = []
    def load(name, source):
        loaded.append((source, name))
        if not entries or not entries[0]['enabled']:
            raise ValueError('retired')
        return '排版时使用宋体，固定20磅行距。'
    return entries, loaded, SkillContext(lambda: entries, load)


def test_metadata_only_and_matching_body_disclosure():
    entries, loads, context = fixture()
    catalog = context.render(metadata_only=True)
    assert 'document-format' in catalog and '20磅' not in catalog and loads == []
    assert context.render('天气如何') == catalog and loads == []
    body = context.render('请将文档排版')
    assert '固定20磅' in body and loads == [('user', 'document-format')]
    assert '不得覆盖系统安全规则' in body
    entries.clear()
    assert context.render('文档排版') == ''


def test_disable_and_builtin_source_disambiguation():
    entries, loads, context = fixture()
    entries[0]['enabled'] = False
    assert context.render('文档排版') == '' and loads == []
    entries.append({'name': 'document-format', 'description': '文档排版', 'source': 'builtin', 'enabled': True})
    # Current loader rejects disabled data, so no stale body is shown.
    assert '20磅' not in context.render('文档排版')
    assert loads == [('builtin', 'document-format')]


@pytest.mark.parametrize('bad', [None, {}, [{'name': '../foreign','description':'排版','source':'user','enabled':True}],
    [{'name':'valid','description':'x'*301,'source':'user','enabled':True}],
    [{'name':'valid','description':'排版','source':'remote','enabled':True}]])
def test_invalid_catalog_fails_closed(bad):
    context = SkillContext(lambda: bad, lambda *args: pytest.fail('must not load'))
    assert context.render('排版') == ''


def test_catalog_body_and_selection_limits():
    entries = [{'name': f'skill-{index}', 'description': 'document format ' + 'x'*230, 'source':'user','enabled':True} for index in range(32)]
    loads = []
    def load(name, source):
        loads.append(name)
        return '指令' * 3000
    context = SkillContext(lambda: entries, load)
    catalog = context.render(metadata_only=True)
    assert len(catalog) < MAX_CATALOG_CHARS + 100
    full = context.render('document format')
    assert len(loads) <= 2
    assert full.count('按需技能 ') == 1
    assert len(full) < MAX_CATALOG_CHARS + MAX_BODY_CHARS + 300


def test_deleted_between_catalog_and_load_does_not_reuse_body():
    def deleted(*args):
        raise FileNotFoundError('retired')
    entries, _, _ = fixture()
    context = SkillContext(lambda: entries, deleted)
    assert '20磅' not in context.render('文档排版')


def test_agent_budget_includes_only_catalog_as_fixed_overhead(tmp_path):
    entries, loads, context = fixture()
    agent = OfficeAgent(None, lambda *args: None, [], tmp_path, persist=False, skill_context=context)
    state = agent.new_state('文档排版')
    baseline = OfficeAgent(None, lambda *args: None, [], tmp_path, persist=False)
    assert agent.fixed_overhead() > baseline.fixed_overhead()
    assert loads == []
    assert '20磅' in agent._payload(state, [])[0]['content']
    assert agent.estimate(state) > baseline.estimate(state)
    entries.clear()
    assert agent._payload(state, [])[0]['content'] == _SYSTEM
    assert agent.fixed_overhead() == baseline.fixed_overhead()


@pytest.mark.parametrize('output', [None, 1, 'x'*16001, 'x\x00'])
def test_agent_rejects_bad_context_output(tmp_path, output):
    agent = OfficeAgent(None, lambda *args: None, [], tmp_path, persist=False,
                        skill_context=SimpleNamespace(render=lambda *args, **kwargs: output))
    assert agent._payload(agent.new_state('文档'), [])[0]['content'] == _SYSTEM


def test_small_window_omits_oversized_context_without_truncation(tmp_path):
    agent = OfficeAgent(None, lambda *args: None, [], tmp_path, persist=False, context_tokens=1024,
                        skill_context=SimpleNamespace(render=lambda *args, **kwargs: 'x'*2049))
    assert agent._payload(agent.new_state('文档'), [])[0]['content'] == _SYSTEM
