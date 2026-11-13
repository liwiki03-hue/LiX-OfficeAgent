from copy import deepcopy
import json

import pytest

from office_agent.public_skill_package import unpack
from office_agent.skills_backend import SkillsBackend
from office_agent.skills_store import SkillsError
from office_agent.webui.draft_suggestions import validate_suggestion
from office_agent.webui.office_workflows import candidate, catalog
from office_agent.webui.skill_discovery import SkillDiscovery
from tests.test_paged_editor import draft
from tests.test_release_draft_ai import armed, Model, run


def test_local_workflow_catalog_resolve_and_transaction_preserve_provenance(tmp_path):
    service = SkillDiscovery(fetcher=lambda query:pytest.fail('No network for bundled workflows'))
    rows = service.workflows('报告')
    assert rows['ok'] and len(rows['items'])==1 and rows['network_requests']==0
    assert service.workflows('not-a-workflow')['items']==[]
    item = rows['items'][0]
    checked = service.resolve(item['candidate_id'],reader=lambda *args,**kwargs:pytest.fail('No external source read'))
    assert checked['compatible'] and checked['origin']['provider']=='lix' and 'commit' not in checked['origin']
    package = service.installation(checked['resolution_id'])
    backend = SkillsBackend(tmp_path)
    target = tmp_path/'skills'/package['slug']/'SKILL.md'
    plan = backend.plan_public(package,'session')
    assert not target.exists()
    assert backend.confirm_change(plan['plan_id'],'session',False)['wrote_nothing'] and not target.exists()
    plan = backend.plan_public(package,'session')
    assert backend.confirm_change(plan['plan_id'],'session',True)['verified']
    read = SkillsBackend(tmp_path).read(package['slug'],'user')
    assert read['body']==package['body'] and read['origin']['version']=='1.0.0'
    assert read['origin']['license']=='AGPL-3.0-only' and not read['origin']['local_modified']
    assert 'GNU AFFERO GENERAL PUBLIC LICENSE' in read['origin']['license_text']
    assert service.resolve_update(backend.public_snapshot(package['slug'],'user'))['reason']=='workflow_updates_with_app'
    edit = backend.plan_change('edit',{'name':package['slug'],'source':'user','content':'---\nname: lix-report-polish\ndescription: 技术报告润色\n---\n本地补充要求。'},'session')
    saved = backend.confirm_change(edit['plan_id'],'session',True)
    changed = backend.read(package['slug'],'user')
    assert changed['origin']['local_modified'] and changed['origin']['source_sha256']==read['origin']['source_sha256']
    undo = backend.plan_rollback(saved['save_id'],'session')
    assert backend.confirm_rollback(undo['rollback_id'],'session',True)['verified']
    assert backend.read(package['slug'],'user')['sha256']==read['sha256']


def test_local_workflow_corrupt_license_or_source_is_refused():
    package = candidate()
    prefix,body = package['_content'].split('\n---\n',1)[1].split('\n',1)
    record,instructions = body.split('\nLiX-Instructions/1\n',1)
    original = json.loads(record)
    for change in ('license','provider','version','path','source'):
        data = deepcopy(original)
        if change=='license':data['license_text'] += '\nChanged'
        elif change=='provider':data['origin']['provider']='remote'
        elif change=='version':data['origin']['version']='main'
        elif change=='path':data['origin']['path']='../../other/SKILL.md'
        else:data['origin']['source_sha256']='0'*64
        with pytest.raises(SkillsError):
            unpack(prefix+'\n'+json.dumps(data,ensure_ascii=False)+'\nLiX-Instructions/1\n'+instructions)


def test_report_workflow_reads_into_actual_agent_request_and_disable_omits_it(draft,monkeypatch):
    service,floating,source = draft
    original = source.read_bytes()
    package = candidate()
    backend = SkillsBackend(source.parent)
    install = backend.plan_public(package,'session')
    assert backend.confirm_change(install['plan_id'],'session',True)['verified']
    row,runtime,route,_ = armed(service,floating,monkeypatch,text='技术报告润色：精简选中句，保留数字和单位。')
    messages = []
    class Polish(Model):
        def invoke(self,values):
            messages.append(deepcopy(values))
            self.calls += 1
            if self.calls==1:
                return {'tool_calls':[{'name':'propose_draft_edit','id':'polish','args':{'reference_id':row['id'],'old':row['old'],
                      'new':'实验完成于2025年，温度120 ℃。','reason':'调整语序，保留年份和温度。'}}]}
            return {'content':'建议待审阅，尚未修改文件。'}
    model = Polish(row)
    run(service,floating,monkeypatch,model,row,runtime,route)
    assert model.calls==2 and len(service.proposals)==1
    prompt = messages[0][0]['content']
    assert package['body'] in prompt and package['origin']['source_sha256'] in prompt and 'LiX@1.0.0' in prompt
    assert 'GNU AFFERO GENERAL PUBLIC LICENSE' not in prompt
    receipts = [payload for kind,payload in list(floating.notifications.queue) if kind=='skill_loaded']
    assert receipts and all(value['document_id']==service.document_id and value['display_name']=='技术报告润色' for value in receipts)
    assert all(value['source_version']==floating.document['version'] and value['session_id']==runtime.session_id for value in receipts)
    assert source.read_bytes()==original and not list(source.parent.glob('*.agent-*.docx'))
    toggle = backend.plan_change('enabled',{'name':package['slug'],'source':'user','enabled':False},'session')
    assert backend.confirm_change(toggle['plan_id'],'session',True)['verified']
    assert package['body'] not in SkillsBackend(source.parent).context().render('技术报告润色')
    with pytest.raises(ValueError):
        validate_suggestion(row['old'],row['old'].replace('120','121'),'改为新温度',[{'text':'只润色，保留数字。'}])


def test_workflow_candidate_identity_binds_full_package(monkeypatch):
    from office_agent.webui import office_workflows
    initial = catalog()['items'][0]
    original = office_workflows.candidate
    def changed():
        value = original()
        value['_content']+='\n'
        return value
    monkeypatch.setattr(office_workflows,'candidate',changed)
    following = catalog()['items'][0]
    assert initial['source_sha256']==following['source_sha256']
    assert initial['candidate_id']!=following['candidate_id'] and initial['package_sha256']!=following['package_sha256']
