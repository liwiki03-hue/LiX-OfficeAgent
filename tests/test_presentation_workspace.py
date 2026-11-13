from copy import deepcopy
from io import BytesIO
import json
from pathlib import Path
import threading
import time
from types import SimpleNamespace
import zipfile

from docx import Document
from pptx import Presentation
from pydantic import ValidationError
import pytest

from office_agent.presentation_drafts import PresentationError, PresentationSpec, Slide, build_deck, identifier, make_outline, render_deck, validate_deck
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingApi
from office_agent.webui.presentation_workspace import PresentationWorkspace
from tests.test_paged_editor import draft, sync


def material(**changes):
    return {'kind':'text','title':'实验报告','pages':3,'text':'实验条件：温度120 ℃。\n实验结果：平均耗时120 ms。','document_id':'','draft_revision':0,'replace_id':'',**changes}


@pytest.fixture
def workspace(tmp_path):
    api=FloatingApi(WebviewApp(),settings_root=tmp_path)
    api._presentations=PresentationWorkspace(api)
    assert api.presentation_create(material())['ok']
    return api._presentations,api


def fake_ready(service,monkeypatch):
    environment={'complete':True,'identity':'synthetic-environment'}
    generator={'identity':'synthetic-generator'}
    monkeypatch.setattr('office_agent.render_identity.environment_identity',lambda path:environment)
    monkeypatch.setattr('office_agent.webui.presentation_workspace.generator_identity',lambda:generator)
    service.rendered={'raw':build_deck(service.spec),'revision':service.spec['revision'],'previews':[b'not-used']*len(service.spec['slides']),
                      'report':{'environment':environment,'generator':generator,'renderer':'synthetic-only'}}


@pytest.mark.parametrize('field,value',[
    ('title','\x00'),('title','   '),('notes','sk-'+'a'*32),('points',['a'*121]),('rows',[['a'],['b','c']]),('id','untrusted-path'),('revision',True)
])
def test_slide_rejects_unsafe_or_ambiguous_inputs(field,value):
    row=Slide(id=identifier(),revision=1,layout='table' if field=='rows' else 'bullets',title='原题',points=[] if field=='rows' else ['正文']).model_dump()
    with pytest.raises(ValidationError):Slide.model_validate({**row,field:value})


def test_schema_versions_unknown_fields_and_repeated_ids_are_rejected(workspace):
    service,api=workspace
    for changed in [{**service.spec,'schema_version':True},{**service.spec,'path':'C:/other'},{**service.spec,'slides':[service.spec['slides'][0]]*2}]:
        with pytest.raises(ValidationError):PresentationSpec.model_validate(changed)
    old=deepcopy(service.spec)
    result=api.presentation_edit(old['id'],1,'update',old['slides'][1]['id'],{'notes':'sk-'+'a'*32})
    assert not result['ok'] and 'aaaaaa' not in result['error'] and 'input_value' not in result['error']
    assert service.spec==old


def test_source_contract_and_empty_material_do_not_replace_existing_draft(workspace):
    service,api=workspace;old=deepcopy(service.spec)
    for changes in [{'text':''},{'document_id':'other'},{'draft_revision':True},{'pages':True},{'text':'x\n'*41},{'path':'C:/other'}]:
        result=api.presentation_create(material(replace_id=old['id'],**changes))
        assert not result['ok'] and service.spec==old


def test_identity_edit_order_and_old_operations_fail_closed(workspace):
    service,api=workspace;spec=deepcopy(service.spec);last=spec['slides'][-1]['id']
    moved=service.edit(spec['id'],1,'move',last,{'direction':'up'})
    assert moved['spec']['slides'][1]['id']==last and moved['spec']['revision']==2
    with pytest.raises(PresentationError,match='已改变'):service.edit(spec['id'],1,'delete',last,{})
    added=service.edit(spec['id'],2,'add',last,{})['spec']
    assert len(added['slides'])==4 and added['slides'][-1]['id'] not in {row['id'] for row in spec['slides']}
    service.edit(spec['id'],3,'delete',added['slides'][-1]['id'],{})
    assert [row['id'] for row in service.spec['slides']]==[spec['slides'][0]['id'],last,spec['slides'][1]['id']]


def test_current_unsaved_docx_snapshot_is_source_and_source_changes_are_explicit(draft):
    paged,floating,path=draft;original=path.read_bytes();sync(paged)
    api=floating.api;api._floating=floating;service=PresentationWorkspace(api)
    state=service.create(material(kind='draft',pages=4,text='',document_id=paged.document_id,draft_revision=1))
    assert any('2026' in row['text'] for row in state['source']['segments'])
    assert state['source']['draft_revision']==1 and state['source_status']=='current' and path.read_bytes()==original
    paged.changed(paged.document_id,2)
    assert service.state()['source_status']=='old_draft'
    path.write_bytes(b'external change')
    assert service.state()['source_status']=='unavailable'
    with pytest.raises(ValueError):service.build(service.spec['id'],1)


def test_capture_ordinary_table_keeps_blanks_and_reports_unsupported_coverage(draft):
    paged,floating,path=draft
    book=Document(path);table=book.add_table(rows=2,cols=2);table.cell(0,0).text='条件';table.cell(1,1).text='120 ms'
    merged=book.add_table(rows=2,cols=2);merged.cell(0,0).merge(merged.cell(0,1))
    book.save(path)
    assert floating.admit(str(path))['ok'];assert floating.read(floating.read_token,True)['ok']
    paged.open(floating.document['id'],floating.document['version'])
    paged.snapshot_revision=paged.revision
    api=floating.api;api._floating=floating;service=PresentationWorkspace(api)
    source=service.capture(material(kind='draft',text='',document_id=paged.document_id,draft_revision=paged.revision))
    rows=[row['rows'] for row in source['segments'] if 'rows' in row]
    assert rows==[[['条件',''],['','120 ms']]]
    assert source['limited'] and len(source['omitted_blocks'])==1 and '脚注' in source['coverage']


@pytest.mark.parametrize('fails',[False,True])
def test_edit_during_render_discards_old_success_and_failure(workspace,monkeypatch,fails):
    service,api=workspace;entered,release=threading.Event(),threading.Event();calls=[]
    def render(spec):
        calls.append(deepcopy(spec));entered.set();assert release.wait(5)
        if fails:raise PresentationError('旧版本失败')
        return {'raw':b'old','report':{},'previews':[]}
    monkeypatch.setattr('office_agent.webui.presentation_workspace.render_deck',render)
    service.build(service.spec['id'],1);assert entered.wait(2)
    try:
        with pytest.raises(PresentationError,match='重复'):service.build(service.spec['id'],1)
        service.edit(service.spec['id'],1,'update',service.spec['slides'][1]['id'],{'title':'新标题'})
    finally:release.set()
    deadline=time.monotonic()+3
    while service.task and time.monotonic()<deadline:time.sleep(.01)
    assert service.task is None and service.rendered is None and service.status=='outline' and service.error=='' and len(calls)==1


def test_privacy_invalidation_drops_late_render_and_approvals(workspace,monkeypatch):
    service,api=workspace;entered,release=threading.Event(),threading.Event()
    def render(spec):entered.set();assert release.wait(5);return {'raw':b'old'}
    monkeypatch.setattr('office_agent.webui.presentation_workspace.render_deck',render)
    service.build(service.spec['id'],1);assert entered.wait(2);service.invalidate();release.set()
    deadline=time.monotonic()+3
    while service.task and time.monotonic()<deadline:time.sleep(.01)
    assert service.spec is None and service.rendered is None and not service.plans and service.task is None


def test_real_editable_objects_notes_and_tampering_checks(workspace):
    service,api=workspace;spec=deepcopy(service.spec)
    spec['slides'][1].update(layout='columns',points=['支持条件120 ℃'],right=['边界条件不可外推'],notes='讲稿保留限定条件。')
    spec['slides'][2].update(layout='table',points=[],rows=[['指标','数值'],['耗时','120 ms'],['缺失','']])
    raw=build_deck(spec);assert validate_deck(raw,spec)['editable_tables']
    book=Presentation(BytesIO(raw))
    assert book.slides[2].shapes[-1].has_table and '讲稿保留限定条件' in book.slides[1].notes_slide.notes_text_frame.text
    book.slides[1].shapes[1].text='未经批准的文字'
    output=BytesIO();book.save(output)
    with pytest.raises(PresentationError):validate_deck(output.getvalue(),spec)
    with BytesIO() as output:
        with zipfile.ZipFile(output,'w') as archive:
            with zipfile.ZipFile(BytesIO(raw)) as original:
                for row in original.infolist():archive.writestr(row,original.read(row.filename))
            archive.writestr('ppt/activeX/unapproved.xml','<x/>')
        with pytest.raises(PresentationError,match='部件'):validate_deck(output.getvalue(),spec)


def test_dense_or_missing_material_is_not_silently_published(workspace):
    service,api=workspace
    missing=make_outline({'segments':[{'id':identifier(),'text':'只有一段材料'}]},'短材料',8)
    with pytest.raises(PresentationError,match='缺材料'):build_deck(missing)
    spec=deepcopy(service.spec);spec['slides'][1].update(layout='columns',points=['中'*120]*5,right=['文'*120]*5)
    with pytest.raises(PresentationError,match='过密'):build_deck(spec)


def test_publish_cancel_single_use_expiry_and_old_plan(workspace,tmp_path,monkeypatch):
    service,api=workspace;fake_ready(service,monkeypatch)
    assert service.plan(service.spec['id'],1,False,lambda:None)['cancelled']
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan']
    assert plan['wrote_nothing'] and not Path(plan['path']).exists()
    assert service.confirm(plan['id'],False)['result']['wrote_nothing']
    with pytest.raises(PresentationError):service.confirm(plan['id'],True)
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan']
    archive=service.plans[plan['id']]['archive'];expired=archive._plans[plan['id']]['time']+301;archive._clock=lambda:expired
    result=service.confirm(plan['id'],True)['result']
    assert result['status']=='rejected' and result['code']=='evidence_expired' and not Path(plan['path']).exists()
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan']
    service.edit(service.spec['id'],1,'update',service.spec['slides'][1]['id'],{'title':'修改大纲'})
    with pytest.raises(PresentationError):service.confirm(plan['id'],True)
    assert not Path(plan['path']).exists()


def test_actual_publish_source_backup_audit_reopen_and_independent_rollback(workspace,tmp_path,monkeypatch):
    service,api=workspace;fake_ready(service,monkeypatch)
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan'];result=service.confirm(plan['id'],True)['result'];path=Path(result['path'])
    assert result['status']=='verified' and validate_deck(path.read_bytes(),service.spec)['slides']==3
    snapshot=json.loads((path.parent/'backup').read_text('utf8'))['snapshot']
    assert snapshot['spec']==service.spec and snapshot['source']==service.source and (path.parent/'audit-verified.json').is_file()
    with pytest.raises(PresentationError):service.confirm(plan['id'],True)
    service.saved_action(plan['id'],'rollback_plan')
    assert service.saved_action(plan['id'],'rollback',False)['result']['wrote_nothing'] and path.exists()
    service.saved_action(plan['id'],'rollback_plan')
    assert service.saved_action(plan['id'],'rollback',True)['result']['status']=='rolled_back' and not path.exists()
    assert (path.parent/'backup').is_file() and (path.parent/'audit-removed.json').is_file()


def test_output_replacement_cannot_be_opened_or_rolled_back(workspace,tmp_path,monkeypatch):
    service,api=workspace;fake_ready(service,monkeypatch)
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan'];result=service.confirm(plan['id'],True)['result'];path=Path(result['path'])
    path.write_bytes(b'user replacement')
    assert not api.presentation_saved_action(plan['id'],'rollback_plan')['ok']
    assert not api.presentation_saved_action(plan['id'],'open')['ok'] and path.read_bytes()==b'user replacement'


def test_uncertain_write_fuses_and_does_not_claim_no_file(workspace,tmp_path,monkeypatch):
    service,api=workspace;fake_ready(service,monkeypatch)
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan'];archive=service.plans[plan['id']]['archive']
    monkeypatch.setattr(archive,'_exclusive',lambda *args:(_ for _ in ()).throw(OSError('synthetic failure')))
    result=service.confirm(plan['id'],True)['result']
    assert result['status']=='uncertain' and result['wrote_nothing'] is False and service.uncertain
    with pytest.raises(PresentationError,match='不确定'):service.build(service.spec['id'],1)


def test_fresh_source_and_environment_must_match_before_publish(workspace,tmp_path,monkeypatch):
    service,api=workspace;fake_ready(service,monkeypatch)
    plan=service.plan(service.spec['id'],1,False,lambda:tmp_path)['plan']
    monkeypatch.setattr('office_agent.webui.presentation_workspace.generator_identity',lambda:{'changed':True})
    with pytest.raises(PresentationError,match='环境'):service.confirm(plan['id'],True)
    assert not Path(plan['path']).exists()


def test_old_source_requires_explicit_ack_and_plan_binds_current_source_state(workspace,tmp_path,monkeypatch):
    service,api=workspace;fake_ready(service,monkeypatch);version=['2']
    monkeypatch.setattr(service,'source_status',lambda:('old_draft',version[0]))
    with pytest.raises(PresentationError,match='旧快照'):service.plan(service.spec['id'],1,False,lambda:tmp_path)
    plan=service.plan(service.spec['id'],1,True,lambda:tmp_path)['plan'];version[0]='3'
    with pytest.raises(PresentationError,match='已改变'):service.confirm(plan['id'],True)
    assert not Path(plan['path']).exists()
