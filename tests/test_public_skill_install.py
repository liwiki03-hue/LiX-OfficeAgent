import base64
import ast
import hashlib
import json
from pathlib import Path
from threading import RLock
from types import SimpleNamespace

import pytest

from office_agent.agent import OfficeAgent
from office_agent.public_skill_package import parse_public, unpack
from office_agent.skills_backend import SkillsBackend
from office_agent.skills_store import SkillsError
from office_agent.webui.skills import Skills
from office_agent.webui.public_skill_resolver import github_source, resolve
from tests.test_agent import MockModel
from tests.test_office_canvas import canvas
from tests.test_paged_editor import draft


FIXTURE = Path(__file__).parent / 'fixtures/public_skills/grill-me-a6bdfd9.json'
OUT = Path(__file__).resolve().parents[1] / 'deliverables/skills-s3-20261113'


def candidate():
    return json.loads(FIXTURE.read_text('utf-8'))


def test_public_install_agent_disable_edit_delete_and_rollback(tmp_path):
    source = candidate()
    backend = SkillsBackend(tmp_path)
    target = tmp_path / 'skills' / source['slug'] / 'SKILL.md'
    denied = backend.plan_public(source, 'session')
    assert backend.confirm_change(denied['plan_id'], 'session', False)['wrote_nothing']
    assert not target.exists()
    plan = backend.plan_public(source, 'session')
    saved = backend.confirm_change(plan['plan_id'], 'session', True)
    assert saved['status'] == 'saved' and target.is_file()
    assert backend.confirm_change(plan['plan_id'], 'session', True)['status'] == 'rejected'
    restarted = SkillsBackend(tmp_path)
    item = restarted.read(source['slug'], 'user')
    assert item['body'] == source['body'] and item['origin']['commit'] == source['origin']['commit']
    assert not item['origin']['local_modified']
    events = []
    context = SkillsBackend(tmp_path, on_load=events.append).context()
    model = MockModel([{'content':'请先说明这份报告方案的目标与读者。'}])
    agent = OfficeAgent(model, lambda *args: pytest.fail('No document tool requested'), [], tmp_path,
                        persist=False, skill_context=context, token_estimator=lambda text:100)
    state = agent.run(agent.new_state('Please grill me about this report plan'))
    assert state.status == 'done' and len(model.messages) == 1
    prompt = model.messages[0][0]['content']
    assert source['body'] in prompt and source['origin']['commit'] in prompt
    assert source['origin']['source_sha256'] in prompt and events
    disabled = backend.plan_change('enabled', {'name':source['slug'],'source':'user','enabled':False}, 'session')
    assert backend.confirm_change(disabled['plan_id'],'session',True)['verified']
    assert source['body'] not in SkillsBackend(tmp_path).context().render('grill report plan')
    edit = backend.plan_change('edit', {'name':source['slug'],'source':'user','content':
        '---\nname: "grill-me"\ndescription: "grill plan"\n---\nLocal revised workflow.'}, 'session')
    assert backend.confirm_change(edit['plan_id'],'session',True)['verified']
    assert backend.read(source['slug'],'user')['origin']['local_modified']
    user_file = target.parent / 'user-notes.txt'
    user_file.write_text('Keep this user file.', 'utf-8')
    delete = backend.plan_change('delete', {'name':source['slug'],'source':'user'}, 'session')
    removed = backend.confirm_change(delete['plan_id'],'session',True)
    assert not target.exists() and user_file.is_file()
    undo = backend.plan_rollback(removed['save_id'],'session')
    assert backend.confirm_rollback(undo['rollback_id'],'session',True)['status']=='rolled_back'
    assert target.is_file() and backend.read(source['slug'],'user')['body']=='Local revised workflow.'
    OUT.mkdir(exist_ok=True)
    (OUT/'AGENT_FLOW.json').write_text(json.dumps({'model':'controlled MockModel; no real service',
        'agent_status':state.status,'actual_model_calls':len(model.messages),'source':source['origin'],
        'loaded_events':events,'disabled_new_context_omits_body':True,'restart_origin_preserved':True,
        'local_edit_marked':True,'delete_preserves_user_file':True,'rollback_verified':True},ensure_ascii=False,indent=2),'utf-8')


def test_public_package_rejects_unsafe_metadata_and_detects_corruption():
    for head in ['name: bad\nname: duplicate\ndescription: safe', 'name: bad\ndescription: &anchor hi',
                 'name: bad\ndescription: safe\nunknown: value']:
        with pytest.raises((ValueError,TypeError)):
            parse_public(('---\n'+head+'\n---\nBody').encode())
    for url in ['https://github.com/a/b/tree/main/skill','https://evil.test/a/b/tree/'+'a'*40+'/skill',
                'https://github.com/a/b/tree/'+'a'*40+'/../skill']:
        with pytest.raises(ValueError):
            github_source(url)
    content=candidate()['_content'].split('\n---\n',1)[1]
    with pytest.raises(SkillsError):
        unpack(content+'tampered')


def test_catalog_adaptation_preserves_valid_entries_and_reports_budget():
    records=[{'id':'user:skill-'+str(i).zfill(2),'description':'plan','source':'user','enabled':True} for i in range(64)]
    backend=SkillsBackend.__new__(SkillsBackend)
    backend._store=SimpleNamespace(catalog=lambda:records)
    assert len(backend.prompt_catalog())==32
    assert len(backend.agent_catalog_status()['excluded'])==32
    records[0]['description']='x'*500
    assert backend.prompt_catalog() and len(backend.prompt_catalog()[0]['description'])==300
    assert len(records[0]['description'])==500
    assert backend.context().render(metadata_only=True)


def test_workspace_epoch_invalidates_public_plan(tmp_path):
    facade=Skills(RLock(),lambda:(True,False),lambda:SkillsBackend(tmp_path))
    planned=facade.plan_public(candidate())
    assert planned['ok']
    facade.invalidate()
    result=facade.confirm_change(planned['plan']['plan_id'],True)
    assert not result['ok'] or result['result']['status']=='rejected'
    assert not (tmp_path/'skills').exists()


def test_resolver_blocks_unhandled_resource_without_installing():
    source=candidate()
    raw=base64.b64decode(json.loads(source['_content'].split('LiX-Public-Skill-Package/1\n',1)[1].split('\nLiX-Instructions/1\n',1)[0])['original'])
    license_data=b'MIT License\nPermission is hereby granted. Test fixture license.'
    blobs={}
    def item(path,data):
        digest=hashlib.sha1(b'blob '+str(len(data)).encode()+b'\0'+data).hexdigest()
        blobs[digest]=data
        return {'path':path,'sha':digest,'type':'blob','mode':'100644','size':len(data)}
    tree={'truncated':False,'tree':[item('skill/grill-me/SKILL.md',raw),item('LICENSE',license_data),
                                   item('skill/grill-me/scripts/run.py',b'print(1)')]}
    def reader(path,deadline,**options):
        if '/git/trees/' in path:return json.dumps(tree).encode()
        data=blobs[path.rsplit('/',1)[1]]
        return json.dumps({'encoding':'base64','content':base64.b64encode(data).decode()}).encode()
    result=resolve('example/workflows','grill-me',fixed=('a'*40,'skill/grill-me'),reader=reader)
    assert not result['compatible'] and result['_content'] is None
    assert any('scripts/run.py' in reason for reason in result['blockers'])


def test_canvas_ai_receives_installed_reference(canvas,monkeypatch):
    from tests.test_canvas_ai import prepare, finish
    service,floating,source,reference,connections=prepare(canvas,monkeypatch)
    package=candidate()
    backend=SkillsBackend(source.parent)
    plan=backend.plan_public(package,'test-session')
    assert backend.confirm_change(plan['plan_id'],'test-session',True)['verified']
    original=source.read_bytes()
    service.ai.submit(service.document_id,service.revision,reference['id'],'','public-skill-test','Please grill this plan; keep the selected change.')
    finish(service.ai)
    assert connections.contexts
    prompt=ast.literal_eval(connections.contexts[0])[0]['content']
    assert package['body'] in prompt and package['origin']['commit'] in prompt
    assert source.read_bytes()==original


def test_paged_ai_receives_installed_reference(draft,monkeypatch):
    from tests.test_release_draft_ai import armed,Model,run
    service,floating,source=draft
    package=candidate()
    backend=SkillsBackend(source.parent)
    plan=backend.plan_public(package,'test-session')
    assert backend.confirm_change(plan['plan_id'],'test-session',True)['verified']
    row,runtime,route,_=armed(service,floating,monkeypatch,text='Please grill this report plan. 把年份改为2026，保持温度和其他数字不变。')
    class Inspect(Model):
        def invoke(self,messages):
            assert package['body'] in messages[0]['content'] and package['origin']['commit'] in messages[0]['content']
            return super().invoke(messages)
    original=source.read_bytes()
    model=Inspect(row)
    run(service,floating,monkeypatch,model,row,runtime,route)
    assert model.calls==2 and source.read_bytes()==original
