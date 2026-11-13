import base64
import hashlib
import json
from threading import RLock

import pytest

from office_agent.agent import OfficeAgent
from office_agent.public_skill_package import unpack
from office_agent.skills_backend import SkillsBackend
from office_agent.skills_store import SkillsError
from office_agent.webui.public_skill_resolver import resolve
from office_agent.webui.skill_discovery import SkillDiscovery
from office_agent.webui.skills import Skills
from tests.test_agent import MockModel
from tests.test_public_skill_install import candidate as legacy_candidate


def source_reader(files, commit='b'*40, folder='report-writing'):
    contents = {folder+'/'+path:text.encode('utf-8') for path,text in files.items()}
    contents['LICENSE'] = 'MIT License\nPermission is hereby granted. LiX synthetic test fixture.'.encode()
    blobs, entries = {}, []
    for path,data in contents.items():
        sha = hashlib.sha1(b'blob '+str(len(data)).encode()+b'\x00'+data).hexdigest()
        blobs[sha] = data
        entries.append({'path':path, 'sha':sha, 'mode':'100644', 'type':'blob', 'size':len(data)})
    def read(path, deadline, **options):
        if path.endswith('/commits/HEAD'):
            return commit.encode()
        if '/git/trees/' in path:
            return json.dumps({'tree':entries,'truncated':False}).encode()
        return json.dumps({'encoding':'base64','content':base64.b64encode(blobs[path.rsplit('/',1)[1]]).decode()}).encode()
    return read


def files(body='Use the [report checklist](references/checklist.md).'):
    return {'SKILL.md':'---\nname: report-writing\ndescription: report writing guidelines\nlicense: MIT\n---\n'+body,
            'references/checklist.md':'Keep conditions, units and source attribution. <script>never execute</script>'}


def resolved(items=None, commit='b'*40):
    return resolve('lix-tests/workflows','report-writing',fixed=(commit,'report-writing'),
                   reader=source_reader(items or files(), commit))


def test_reference_package_install_load_edit_and_rollback(tmp_path):
    package = resolved()
    assert package['compatible'] and 'Package/2' in package['_content']
    backend = SkillsBackend(tmp_path)
    plan = backend.plan_public(package,'session')
    assert backend.confirm_change(plan['plan_id'],'session',True)['verified']
    read = SkillsBackend(tmp_path).read(package['slug'],'user')
    assert read['body'] == package['body'] and len(read['origin']['resources']) == 1
    assert not read['origin']['local_modified']
    loaded = []
    model = MockModel([{'content':'Keep the report conditions and units.'}])
    agent = OfficeAgent(model,lambda *args:pytest.fail('No tool authority'),[],tmp_path,persist=False,
                        skill_context=SkillsBackend(tmp_path,on_load=loaded.append).context(),token_estimator=lambda text:100)
    assert agent.run(agent.new_state('report writing guidelines')).status == 'done'
    prompt = model.messages[0][0]['content']
    assert package['body'] in prompt and files()['references/checklist.md'] in prompt and loaded
    assert len(model.messages) == 1
    edit = backend.plan_change('edit',{'name':package['slug'],'source':'user','content':
          '---\nname: report-writing\ndescription: report writing guidelines\n---\nLocal workflow.'},'session')
    saved = backend.confirm_change(edit['plan_id'],'session',True)
    assert backend.read(package['slug'],'user')['origin']['local_modified']
    undo = backend.plan_rollback(saved['save_id'],'session')
    assert backend.confirm_rollback(undo['rollback_id'],'session',True)['verified']
    assert backend.read(package['slug'],'user')['origin']['resources'] == read['origin']['resources']


def test_reference_guards_keep_unsupported_candidates_disabled():
    cases = [files('Use [escape](../secret.md).'), files('Use [missing](references/missing.md).'),
             files('Use ![image](https://example.test/x.png).'),
             {**files(), 'references/checklist.md':'See [self](checklist.md).'},
             {**files(), 'references/checklist.md':'x'*7900},
             {**files(),'run.py':'print(1)'}]
    for items in cases:
        result = resolved(items)
        assert not result['compatible'] and result['_content'] is None and result['blockers']
    with pytest.raises(ValueError,match='unsupported_reference_path'):
        resolved({**files(),'CON.txt':'unsafe'})
    package = resolved()
    body = package['_content'].split('\n---\n',1)[1]
    with pytest.raises(SkillsError):
        unpack(body.replace('Keep conditions','Alter conditions'))


def test_update_compares_and_preserves_local_changes_epoch_and_rollback(tmp_path):
    backend = SkillsBackend(tmp_path)
    facade = Skills(RLock(),lambda:(True,False),lambda:backend)
    original = resolved(commit='a'*40)
    install = facade.plan_public(original)
    assert facade.confirm_change(install['plan']['plan_id'],True)['result']['verified']
    snapshot = facade.public_snapshot(original['slug'],'user')['snapshot']
    discovery = SkillDiscovery()
    changed = files('Use the [report checklist](references/checklist.md). Keep numbers unchanged.')
    changed['references/checklist.md'] += '\nPreserve report conditions.'
    result = discovery.resolve_update(snapshot,reader=source_reader(changed))
    assert result['ok'] and result['compatible']
    assert any(row['path']=='SKILL.md' and row['status']=='修改' for row in result['update']['files'])
    assert any(row['path']=='references/checklist.md' and row['status']=='修改' for row in result['update']['files'])
    unsupported = discovery.resolve_update(snapshot,reader=source_reader({**changed,'scripts/run.py':'print(1)'}))
    assert unsupported['ok'] and not unsupported['compatible']
    assert any(row['path']=='scripts/run.py' and row['status']=='新增' for row in unsupported['update']['files'])
    pending = discovery.installation(result['resolution_id'])
    cancelled = facade.plan_public_update(pending)
    assert facade.confirm_change(cancelled['plan']['plan_id'],False)['result']['wrote_nothing']
    assert backend.read(original['slug'],'user')['origin']['commit']=='a'*40
    update = facade.plan_public_update(pending)
    saved = facade.confirm_change(update['plan']['plan_id'],True)['result']
    assert saved['verified'] and backend.read(original['slug'],'user')['origin']['commit']=='b'*40
    assert facade.confirm_change(update['plan']['plan_id'],True)['result']['status']=='rejected'
    undo = facade.plan_rollback(saved['save_id'])
    assert facade.confirm_rollback(undo['plan']['rollback_id'],True)['result']['verified']
    assert backend.read(original['slug'],'user')['origin']['commit']=='a'*40
    facade.invalidate()
    assert facade.plan_public_update(pending)['reason']=='skills_scope_changed'
    edit = backend.plan_change('edit',{'name':original['slug'],'source':'user','content':
          '---\nname: report-writing\ndescription: LOCAL metadata change\n---\n'+original['body']},'session')
    assert backend.confirm_change(edit['plan_id'],'session',True)['verified']
    local = facade.public_snapshot(original['slug'],'user')['snapshot']
    blocked = discovery.resolve_update(local,reader=source_reader(changed))
    assert blocked['ok'] and not blocked['compatible'] and blocked['update']['local_modified']
    assert facade.plan_public_update(discovery.installation(blocked['resolution_id']))['reason']=='skills_source_changed'


def test_legacy_origin_update_and_external_change_before_confirmation(tmp_path):
    backend = SkillsBackend(tmp_path)
    legacy = legacy_candidate()
    installed = backend.plan_public(legacy,'session')
    assert backend.confirm_change(installed['plan_id'],'session',True)['verified']
    snapshot = backend.public_snapshot(legacy['slug'],'user')
    assert snapshot['origin']['format'].endswith('v1')
    body = {'SKILL.md':'---\nname: grill-me\ndescription: grill report plan\n---\nAsk about the report plan.'}
    target = resolve(legacy['origin']['repository'],'grill-me',fixed=('b'*40,legacy['origin']['path']),
                     reader=source_reader(body,folder=legacy['origin']['path']))
    target['update'] = {'files':[{'path':'SKILL.md','status':'修改'}]}
    plan = backend.plan_public_update(snapshot,target,'session')
    path = tmp_path/'skills'/legacy['slug']/'SKILL.md'
    previous = path.read_bytes()
    path.write_bytes(previous+b'\nExternal change')
    assert backend.confirm_change(plan['plan_id'],'session',True)['status']=='rejected'
    assert path.read_bytes()==previous+b'\nExternal change'
