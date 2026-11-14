"""Authenticated Skills API with temp fixtures; no provider or native window."""
from types import SimpleNamespace
from pathlib import Path
import pytest

from office_agent.webui.bridge import Api
from office_agent.skills_backend import SkillsBackend


TEXT='---\nname: "报告风格"\ndescription: "报告格式使用固定表格"\n---\nUse approved report formatting.\n'


class Auth:
    authenticated=True
    def status(self): return {'authenticated':self.authenticated}
    def logout(self): self.authenticated=False; return {'ok':True}
    def login(self,*args): self.authenticated=True; return {'ok':True,'authenticated':True}


def make(tmp_path):
    root=tmp_path/'document-workspace';root.mkdir()
    app=SimpleNamespace(path=root/'fixture.docx',editor=SimpleNamespace(space=SimpleNamespace(root=root)),
                        busy=False,window=None,_clear_private_state=lambda:None,_build_model=lambda:None)
    auth=Auth()
    return Api(app,auth=auth,settings_root=tmp_path),app,auth,root


def create(api):
    plan=api.plan_skill_change('create',{'slug':'report-style','content':TEXT})['plan']
    return api.confirm_skill_change(plan['plan_id'],True)['result']


def test_actual_plan_approval_prompt_match_disable_read_delete_and_rollback(tmp_path):
    api,app,auth,root=make(tmp_path)
    plan=api.plan_skill_change('create',{'slug':'report-style','content':TEXT})['plan']
    assert list(root.iterdir())==[]
    saved=api.confirm_skill_change(plan['plan_id'],True)['result']
    assert saved['applied'] is True
    catalog=api.list_skills()['skills']
    assert catalog[0]['id']=='user:report-style' and 'body' not in catalog[0]
    prompt=SkillsBackend(root).context()
    assert 'Use approved report formatting.' in prompt.render('报告格式')
    assert 'Use approved report formatting.' not in prompt.render(metadata_only=True)
    disabled=api.plan_skill_change('enabled',{'name':'report-style','source':'user','enabled':False})['plan']
    assert api.confirm_skill_change(disabled['plan_id'],True)['result']['applied'] is True
    assert api.get_skill('report-style','user')['skill']['body'].strip()=='Use approved report formatting.'
    assert prompt.render('报告格式')==''
    deletion=api.plan_skill_change('delete',{'name':'report-style','source':'user'})['plan']
    result=api.confirm_skill_change(deletion['plan_id'],True)['result']
    assert api.list_skills()['skills']==[]
    rollback=api.plan_skill_rollback(result['save_id'])['plan']
    assert 'rollback_id' in rollback and 'plan_id' not in rollback
    assert api.confirm_skill_rollback(rollback['rollback_id'],True)['result']['restored'] is True
    assert len(api.list_skills()['skills'])==1


def test_authentication_blocks_every_new_api_and_relogin_retires_plan(tmp_path):
    api,app,auth,root=make(tmp_path)
    plan=api.plan_skill_change('create',{'slug':'report-style','content':TEXT})['plan']
    api.logout()
    for result in (api.list_skills(),api.get_skill('x','user'),api.plan_skill_change('delete',{}),
                   api.plan_skill_import(),api.confirm_skill_change(plan['plan_id'],True),
                   api.plan_skill_rollback('a'*32),api.confirm_skill_rollback('a'*32,True)):
        assert result['reason']=='not_authenticated'
    api.login('fixture','fixture')
    assert api.confirm_skill_change(plan['plan_id'],True)['result']['applied'] is False
    assert list(root.iterdir())==[]


def test_missing_selected_document_is_not_a_sticky_failure(tmp_path):
    api,app,auth,root=make(tmp_path)
    app.path=None
    assert api.list_skills()['reason']=='no_file'
    assert not api._skills._disabled
    app.path=root/'fixture.docx'
    listed=api.list_skills()
    assert listed['ok'] and listed['skills']==[] and listed['agent_catalog']['available']==0


def test_busy_and_diagram_retire_authorities_without_writing(tmp_path):
    api,app,auth,root=make(tmp_path)
    for owner,field in ((app,'busy'),(api._diagram,'_slot')):
        proposal=api.plan_skill_change('create',{'slug':'report-style','content':TEXT})['plan']
        setattr(owner,field,True)
        assert api.list_skills()['reason']=='busy'
        setattr(owner,field,False if field=='busy' else None)
        assert api.confirm_skill_change(proposal['plan_id'],True)['result']['applied'] is False
    assert list(root.iterdir())==[]


def test_host_picker_only_import_and_post_modal_auth_scope_check(tmp_path):
    api,app,auth,root=make(tmp_path)
    folder=tmp_path/'imported-style';folder.mkdir()
    (folder/'SKILL.md').write_text(TEXT,encoding='utf-8')
    (folder/'do-not-copy.py').write_text('raise RuntimeError()',encoding='utf-8')
    api._skills._picker=lambda:folder
    proposal=api.plan_skill_import()['plan']
    assert api.confirm_skill_change(proposal['plan_id'],True)['result']['applied'] is True
    assert not (root/'skills'/'imported-style'/'do-not-copy.py').exists()
    with pytest.raises(TypeError):api.plan_skill_import(str(folder))
    def logout_in_picker():
        api.logout()
        return folder
    api._skills._picker=logout_in_picker
    result=api.plan_skill_import()
    assert result['cancelled'] is True and result['reason']=='skills_scope_changed'


def test_native_selection_retires_old_root_and_plan(tmp_path):
    api,app,auth,root=make(tmp_path)
    proposal=api.plan_skill_change('create',{'slug':'report-style','content':TEXT})['plan']
    second=tmp_path/'second';second.mkdir()
    def select():
        app.path=second/'other.docx'
        app.editor=SimpleNamespace(space=SimpleNamespace(root=second))
        return {'ok':True}
    app.select_file=select
    assert api.select_file()['ok']
    assert api.confirm_skill_change(proposal['plan_id'],True)['result']['applied'] is False
    assert api.list_skills()['skills']==[]
    assert list(root.iterdir())==[] and list(second.iterdir())==[]


def test_no_client_root_and_unknown_fields_refused(tmp_path):
    api,app,auth,root=make(tmp_path)
    assert not api.plan_skill_change('create',{'slug':'x','content':TEXT,'root':'outside'})['ok']
    assert list(root.iterdir())==[]
