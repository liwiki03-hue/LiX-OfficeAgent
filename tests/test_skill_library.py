import copy
from io import BytesIO

from PIL import Image
import pytest

from office_agent.agent import OfficeAgent
from office_agent.public_skill_package import parse_public
from office_agent.skills_backend import SkillsBackend
from office_agent.webui.skill_discovery import SkillDiscovery
from office_agent.webui.skill_icons import thumbnail
from tests.test_agent import MockModel
from tests.test_public_skill_install import candidate
from tests.test_skill_references_updates import source_reader


def test_featured_sources_are_offline_distinct_and_fixed(monkeypatch):
    service = SkillDiscovery(fetcher=lambda query:pytest.fail('No automatic network search'))
    result = service.featured()
    assert result['ok'] and len(result['items'])==6 and result['network_requests']==0
    assert len({row['candidate_id'] for row in result['items']})==6
    assert len({row['source'] for row in result['items']})==2
    item = next(row for row in result['items'] if row['name']=='grill-me')
    calls = []
    def resolve(repository,skill,**options):
        calls.append((repository,skill,options['fixed']))
        return copy.deepcopy(candidate())
    monkeypatch.setattr('office_agent.webui.public_skill_resolver.resolve',resolve)
    checked = service.resolve(item['candidate_id'])
    assert checked['ok'] and calls==[(item['source'],item['skill_id'],(item['commit'],item['path']))]
    assert not service.resolve('0'*32)['ok']


def test_bounded_metadata_preserves_source_and_rejects_yaml_authority():
    raw=b'---\nname: report-style\ndescription: >-\n  Report writing\n  and formatting\nmetadata:\n  author: Example\n  version: "1.0"\nlicense: MIT\n---\nKeep report numbers unchanged.'
    fields,body = parse_public(raw)
    assert fields['description']=='Report writing and formatting'
    assert fields['metadata']=={'author':'Example','version':'1.0'} and body=='Keep report numbers unchanged.'
    long=b'---\nname: report-style\ndescription: '+b'x'*1024+b'\n---\nBody'
    assert len(parse_public(long)[0]['description'])==1024
    for metadata in ['metadata:\n  author: &anchor anything','metadata:\n  author: first\n  author: second',
                     'metadata:\n  nested:\n    tool: shell','allowed-tools: !execute shell','unknown: enabled']:
        with pytest.raises((ValueError,TypeError)):
            parse_public(('---\nname: report-style\ndescription: reports\n'+metadata+'\n---\nBody').encode())


def progressive_source():
    return {'SKILL.md':'---\nname: report-writing\ndescription: report writing guidelines\nmetadata:\n  author: LiX synthetic test\n---\nUse [newsletter](references/newsletter.md) or [quarterly](references/quarterly.md) as appropriate.',
            'references/newsletter.md':'# Newsletter\n'+('Newsletter facts to preserve. '*145),
            'references/quarterly.md':'# Quarterly\n'+('Quarterly figures to preserve. '*145)}


def test_progressive_opt_in_keeps_original_and_reports_unloaded_material(tmp_path):
    service = SkillDiscovery()
    checked = service.resolve(github_url='https://github.com/lix-tests/workflows/tree/'+'a'*40+'/report-writing',reader=source_reader(progressive_source(),'a'*40))
    assert checked['ok'] and not checked['compatible'] and checked['progressive_available']
    assert service.installation(checked['resolution_id'])['_content'] is None
    chosen = service.progressive(checked['resolution_id'])
    package = service.installation(chosen['resolution_id'])
    assert chosen['compatible'] and 'Package/3' in package['_content']
    backend = SkillsBackend(tmp_path)
    plan = backend.plan_public(package,'session')
    assert '按需' in plan['format']
    assert backend.confirm_change(plan['plan_id'],'session',True)['verified']
    assert backend.read(package['slug'],'user')['origin']['reference_loading']=='task-matched'
    events = []
    model = MockModel([{'content':'请提供本周简报内容。'}])
    agent = OfficeAgent(model,lambda *args:pytest.fail('No tool authority'),[],tmp_path,persist=False,
                        skill_context=SkillsBackend(tmp_path,on_load=events.append).context(),token_estimator=lambda text:100)
    assert agent.run(agent.new_state('report-writing newsletter')).status=='done'
    prompt = model.messages[0][0]['content']
    assert package['body'] in prompt and 'Newsletter facts to preserve.' in prompt
    assert 'Quarterly figures to preserve.' not in prompt
    assert events[-1]['selected_references']==['references/newsletter.md']
    assert events[-1]['unloaded_references']==['references/quarterly.md'] and len(model.messages)==1
    assert backend.read(package['slug'],'user')['origin']['source_sha256']==checked['origin']['source_sha256']


def test_source_icons_decode_to_safe_bounded_png():
    image=Image.new('RGB',(128,64),'orange')
    buffer=BytesIO();image.save(buffer,format='PNG')
    icon=thumbnail(buffer.getvalue())
    assert icon['data_url'].startswith('data:image/png;base64,') and icon['width']==96 and icon['height']==48
    for data in [b'<svg onload="alert(1)"></svg>',b'not an image',b'x'*(512*1024+1)]:
        with pytest.raises(ValueError):thumbnail(data)
    oversized=BytesIO();Image.new('RGB',(1025,1),'white').save(oversized,format='PNG')
    with pytest.raises(ValueError):thumbnail(oversized.getvalue())
