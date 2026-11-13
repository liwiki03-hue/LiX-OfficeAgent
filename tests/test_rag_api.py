"""Actual authenticated API/temp sqlite-vec, fabricated vectors; no native GUI/provider."""
import json
from types import SimpleNamespace
import pytest

from office_agent.webui.bridge import Api
from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import Workspace
from office_agent.rag import RagError
from office_agent.embeddings import EmbeddingError


class Auth:
    authenticated = True
    def status(self): return {'authenticated':self.authenticated}
    def logout(self): self.authenticated=False; return {'ok':True}
    def login(self,*args): self.authenticated=True; return {'ok':True,'authenticated':True}


class SyntheticEmbedder:
    identity={'channel':'local','model':'api-fixture-synthetic','dimension':2}
    def __init__(self): self.calls=[]; self.hook=lambda:None
    def embed(self,texts):
        self.calls.append(list(texts)); self.hook()
        return [[1.,0.] for text in texts]


def make(tmp_path):
    root=tmp_path/'documents'; root.mkdir()
    source=root/'report.txt'; source.write_text('仅用于测试的报告文本。'*200,encoding='utf-8')
    app=SimpleNamespace(path=source,editor=SimpleNamespace(space=Workspace(root)),busy=False,
                        _clear_private_state=lambda:None,_build_model=lambda:None)
    auth=Auth(); api=Api(app,auth=auth,settings_root=tmp_path)
    embedder=SyntheticEmbedder(); factories=[]
    def factory(check):
        factories.append(check)
        return WorkspaceRag(app.editor.space,embedder,scope_check=check)
    api._rag._factory=factory
    return api,app,auth,root,source,embedder,factories


def index(api):
    plan=api.plan_rag_index()['plan']
    result=api.confirm_rag_index(plan['plan_id'],True)
    assert result['ok'] and result['result']['applied'] is True
    return result['result']


def test_actual_index_evidence_bounded_read_independent_rollbacks(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    before=set(root.iterdir()); plan=api.plan_rag_index()['plan']
    assert embedder.calls==[] and set(root.iterdir())==before
    saved=api.confirm_rag_index(plan['plan_id'],True)['result']
    proposal=api.plan_rag_query('fixture question',1)['plan']
    assert 'hits' not in proposal and 'text' not in proposal
    assert not list(root.glob('.office-rag-evidence-*'))
    result=api.confirm_rag_query(proposal['plan_id'],True)['result']
    assert 'text' not in result['hits'][0] and len(result['hits'][0]['excerpt'])==200
    evidence=root/result['evidence']
    assert len(json.loads(evidence.read_bytes())['hits'][0]['text'])>200
    assert len(api.get_rag_evidence(result['save_id'])['evidence']['text'])<=2000
    rollback=api.plan_rag_rollback('evidence',result['save_id'])['plan']
    assert api.confirm_rag_rollback('evidence',rollback['rollback_id'],True)['result']['restored'] is True
    assert not evidence.exists() and api.get_rag_snapshot()['snapshot']['chunk_count']==1
    rollback=api.plan_rag_rollback('index',saved['save_id'])['plan']
    assert api.confirm_rag_rollback('index',rollback['rollback_id'],True)['result']['restored'] is True
    assert api.get_rag_snapshot()['snapshot']['chunk_count']==0 and source.exists()


@pytest.mark.parametrize('approved',[False,1,'true',None])
def test_literal_index_denial_before_embeddings_or_writes(tmp_path,approved):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    plan=api.plan_rag_index()['plan']
    result=api.confirm_rag_index(plan['plan_id'],approved)['result']
    assert result['wrote_nothing'] is True and result['applied'] is False
    assert embedder.calls==[] and list(root.iterdir())==[source]


@pytest.mark.parametrize('approved',[False,1,'true',None])
def test_literal_evidence_denial_never_returns_hits(tmp_path,approved):
    api,app,auth,root,source,embedder,factories=make(tmp_path); index(api)
    plan=api.plan_rag_query('fixture')['plan']
    result=api.confirm_rag_query(plan['plan_id'],approved)['result']
    assert result['wrote_nothing'] is True and 'hits' not in result
    assert not list(root.glob('.office-rag-evidence-*'))


def test_authentication_blocks_every_api_and_relogin_retires_grants(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    plan=api.plan_rag_index()['plan']; api.logout()
    calls=[lambda:api.get_rag_snapshot(),lambda:api.plan_rag_index(),
           lambda:api.confirm_rag_index(plan['plan_id'],True),lambda:api.plan_rag_query('fixture'),
           lambda:api.confirm_rag_query('a'*32,True),lambda:api.get_rag_evidence('a'*32),
           lambda:api.plan_rag_rollback('index','a'*32),lambda:api.confirm_rag_rollback('index','a'*32,True)]
    for call in calls: assert call()['reason']=='not_authenticated'
    assert len(factories)==1 and embedder.calls==[]
    api.login('fixture','fixture')
    assert api.confirm_rag_index(plan['plan_id'],True)['result']['wrote_nothing'] is True
    assert list(root.iterdir())==[source]


def test_busy_diagram_and_no_file_nonsticky_revoke_index_plans(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    for obj,attribute,value,reason in [(app,'busy',True,'busy'),(api._diagram,'_slot',True,'busy'),(app,'path',None,'no_file')]:
        plan=api.plan_rag_index()['plan']; old=getattr(obj,attribute); setattr(obj,attribute,value)
        assert api.get_rag_snapshot()['reason']==reason
        setattr(obj,attribute,old)
        assert api.confirm_rag_index(plan['plan_id'],True)['result']['wrote_nothing']
    assert not api._rag._disabled and embedder.calls==[] and list(root.iterdir())==[source]


def test_native_selection_retires_old_workspace_and_plan(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    plan=api.plan_rag_index()['plan']; another=tmp_path/'second'; another.mkdir()
    new_source=another/'second.txt';new_source.write_text('Another fixture',encoding='utf-8')
    def select():
        app.path=new_source; app.editor=SimpleNamespace(space=Workspace(another)); return {'ok':True}
    app.select_file=select
    assert api.select_file()['ok']
    assert api.confirm_rag_index(plan['plan_id'],True)['result']['wrote_nothing']
    assert list(root.iterdir())==[source] and list(another.iterdir())==[new_source]


def test_logout_during_embedding_blocks_index_publication(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    plan=api.plan_rag_index()['plan']; embedder.hook=api.logout
    result=api.confirm_rag_index(plan['plan_id'],True)
    assert result['ok'] is False and result['reason']=='rag_scope_changed'
    assert list(root.iterdir())==[source] and not auth.authenticated
    assert factories[0]() is False


def test_logout_during_query_returns_no_evidence_or_authority(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path); index(api)
    embedder.hook=api.logout
    result=api.plan_rag_query('fixture')
    assert result['ok'] is False and 'plan' not in result and 'hits' not in result
    assert not list(root.glob('.office-rag-evidence-*')) and not auth.authenticated


def test_source_stale_index_or_query_never_returns_evidence(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    plan=api.plan_rag_index()['plan']; source.write_text('changed',encoding='utf-8')
    assert api.confirm_rag_index(plan['plan_id'],True)['result']['wrote_nothing']
    assert embedder.calls==[]; index(api)
    source.write_text('changed again',encoding='utf-8')
    result=api.plan_rag_query('fixture')
    assert result['ok'] is False and 'plan' not in result
    assert not list(root.glob('.office-rag-evidence-*'))


@pytest.mark.parametrize('kind',['other',None,1,[]])
def test_rollback_kind_is_exact_fixed_choice_not_dynamic_method(tmp_path,kind):
    api,*rest=make(tmp_path)
    assert api.plan_rag_rollback(kind,'a'*32)=={'ok':False,'reason':'rag_invalid'}
    assert api.confirm_rag_rollback(kind,'a'*32,True)=={'ok':False,'reason':'rag_invalid'}
    assert not api._rag._disabled


def test_browser_cannot_choose_roots_source_names_hits_sessions_or_provider(tmp_path):
    api,*rest=make(tmp_path)
    for call in [lambda:api.plan_rag_index('outside.txt'),lambda:api.plan_rag_query('q',1,[]),
                 lambda:api.confirm_rag_query('a'*32,True,[]),lambda:api.get_rag_snapshot('root')]:
        with pytest.raises(TypeError):call()


@pytest.mark.parametrize('error',[EmbeddingError('embedding_not_configured'),RagError('rag_sensitive_content')],ids=['configuration','sensitive'])
def test_safe_known_codes_do_not_disable_other_directions(tmp_path,error):
    api,*rest=make(tmp_path)
    def factory(check): raise error
    api._rag._factory=factory
    assert api.plan_rag_index()=={'ok':False,'reason':error.code}
    assert not api._rag._disabled and api.list_skills()['ok']


def test_unknown_factory_error_redacted_and_fuse_survives_logout(tmp_path):
    api,*rest=make(tmp_path)
    def factory(check):raise RuntimeError('private fixture value must not escape')
    api._rag._factory=factory
    result=api.plan_rag_index()
    assert result=={'ok':False,'reason':'rag_failed','outcome_uncertain':True}
    api.logout();api.login('fixture','fixture')
    assert api.plan_rag_index()['reason']=='rag_service_disabled'
    assert api.list_skills()['ok']


def test_uncertain_receipt_strips_hits_and_authority_and_sticks(tmp_path):
    api,*rest=make(tmp_path)
    backend=SimpleNamespace(invalidate=lambda:None,confirm_query=lambda *args:{'status':'uncertain','applied':None,
                'hits':[{'text':'unverified full text','excerpt':'unverified'}],'save_id':'a'*32})
    api._rag._factory=lambda check:backend
    result=api.confirm_rag_query('b'*32,True)['result']
    assert result=={'status':'uncertain','applied':None,'restored':None,'wrote_nothing':False}
    api.logout();api.login('fixture','fixture')
    assert api.get_rag_snapshot()['reason']=='rag_service_disabled'


def test_malformed_saved_receipt_with_full_hit_text_is_not_published(tmp_path):
    api,*rest=make(tmp_path)
    backend=SimpleNamespace(invalidate=lambda:None,confirm_query=lambda *args:{'status':'saved','applied':True,
                'wrote_nothing':False,'hits':[{'text':'must stay private','excerpt':'bounded'}],'save_id':'a'*32})
    api._rag._factory=lambda check:backend
    assert api.confirm_rag_query('b'*32,True)=={'ok':False,'reason':'rag_failed','outcome_uncertain':True}


@pytest.mark.parametrize('plan',[{'plan_id':'a'*32,'hits':[{'text':'full text'}]},
                                {'plan_id':'a'*32,'api_key':'must not escape'}],ids=['full-hits','private-config'])
def test_malformed_plan_never_returns_unpersisted_rows_or_configuration(tmp_path,plan):
    api,*rest=make(tmp_path)
    backend=SimpleNamespace(invalidate=lambda:None,plan_index=lambda *args:plan)
    api._rag._factory=lambda check:backend
    assert api.plan_rag_index()=={'ok':False,'reason':'rag_failed','outcome_uncertain':True}


def test_settings_save_and_rollback_retire_rag_configuration_and_grants(tmp_path):
    api,app,auth,root,source,embedder,factories=make(tmp_path)
    # Settings transaction is a separately approved real temp transaction; no actual .env.
    (tmp_path/'.env').write_text('MODEL_PROVIDER=deepseek\n',encoding='utf-8')
    plan=api.plan_rag_index()['plan']; check=factories[0]
    settings_plan=api.plan_settings_update({'context_window':16384})['plan']
    settings=api.confirm_settings_update(settings_plan['plan_id'],True)['result']
    assert settings['applied'] is True and check() is False
    assert api.confirm_rag_index(plan['plan_id'],True)['result']['wrote_nothing']
    plan=api.plan_rag_index()['plan']; check=factories[-1]
    rollback=api.plan_settings_rollback(settings['save_id'])['plan']
    assert api.confirm_settings_rollback(rollback['rollback_id'],True)['result']['restored']
    assert check() is False and api.confirm_rag_index(plan['plan_id'],True)['result']['wrote_nothing']
