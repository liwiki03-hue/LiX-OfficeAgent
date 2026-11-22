"""Adapter contracts, independent of a mutating filesystem backend."""
import pytest

from office_agent.skills_backend import SkillsBackend
from office_agent.skills_store import SkillsError


class Store:
    def __init__(self): self.calls=[]
    def catalog(self, *, include_sha=False):
        item = {'id': 'user:report-style', 'name': '报告风格', 'description': '报告格式保持表格',
                'source': 'user', 'enabled': True}
        if include_sha:
            item['sha256'] = 'a' * 64
        return [item]
    def read(self,identifier):
        self.calls.append(('read',identifier)); return {'body':'disabled content','enabled':False}
    def load(self,identifier):
        self.calls.append(('load',identifier)); return {'body':'Use approved report format.'}
    def __getattr__(self,name):
        def call(*args): self.calls.append((name,*args)); return {'plan_id':'a'*32}
        return call


@pytest.fixture
def backend():
    adapter=SkillsBackend.__new__(SkillsBackend)
    adapter._store=Store()
    return adapter


def test_catalog_does_not_load_body_and_prompt_names_are_stable_slugs(backend):
    assert backend.catalog()[0]['name']=='报告风格'
    assert 'sha256' not in backend.catalog()[0]
    rendered=backend.context().render('报告格式')
    assert 'report-style' in rendered and 'Use approved report format.' in rendered
    assert backend._store.calls==[('load','user:report-style')]


def test_metadata_only_never_loads_body(backend):
    assert 'report-style' in backend.context().render(metadata_only=True)
    assert backend._store.calls==[]


def test_management_read_can_show_disabled_without_prompt_load(backend):
    assert backend.read('report-style','user')['enabled'] is False
    assert backend._store.calls==[('read','user:report-style')]


@pytest.mark.parametrize('operation,changes,expected',[
    ('create',{'slug':'report-style','content':'text'},('plan_create','report-style','text','session')),
    ('edit',{'name':'report-style','source':'user','content':'text'},('plan_edit','user:report-style','text','session')),
    ('delete',{'name':'report-style','source':'user'},('plan_delete','user:report-style','session')),
    ('enabled',{'name':'report-style','source':'builtin','enabled':False},('plan_enabled','builtin:report-style',False,'session')),
])
def test_exact_management_mapping(backend,operation,changes,expected):
    backend.plan_change(operation,changes,'session')
    assert backend._store.calls==[expected]


@pytest.mark.parametrize('operation,changes',[
    ('create',{'slug':'x','content':'text','root':'outside'}),
    ('delete',{'name':'x','source':'user','confirmed':True}),
    ('import',{'directory':'outside'}),('run',{}),('edit',{}),
    ('delete',{'name':'../outside','source':'user'}),
    ('delete',{'name':'x','source':['user']}),
])
def test_client_path_authority_unknown_or_extra_fields_refused(backend,operation,changes):
    with pytest.raises(SkillsError): backend.plan_change(operation,changes,'session')
    assert backend._store.calls==[]


def test_host_picker_import_slug_and_confirmation_values_pass_unchanged(backend,tmp_path):
    selected=tmp_path/'report-style'
    backend.plan_import(selected,'session')
    assert backend._store.calls==[('plan_import','report-style',selected,'session')]
    backend.confirm_change('id','session','true')
    assert backend._store.calls[-1]==('confirm','id','session','true')
    backend.plan_rollback('save','session')
    backend.confirm_rollback('rollback','session',1)
    assert backend._store.calls[-1]==('confirm_rollback','rollback','session',1)


def test_invalidate_keeps_backend_and_does_not_reset_fuse(backend):
    store=backend._store
    backend.invalidate()
    assert backend._store is store and store.calls==[('invalidate',)]
    assert repr(backend)=='<SkillsBackend host-only>'
