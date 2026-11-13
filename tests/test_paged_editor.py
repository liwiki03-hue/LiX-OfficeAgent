import base64
from copy import deepcopy
import io
import json
from types import SimpleNamespace
import zipfile

from docx import Document
import pytest

from office_agent.instruction_runtime import InstructionRuntime
from office_agent.paged_documents import CandidatePublisher, GuardError, changes, eligibility, normalize, parts, projection
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from office_agent.webui.paged_editor import utf16_slice
from office_agent.writes import digest


@pytest.fixture
def draft(tmp_path):
    source=tmp_path/'分页样例.docx'
    book=Document();book.add_heading('验收报告',0);book.add_paragraph('实验于2025年完成，温度120 ℃。');book.add_paragraph('非目标内容保持。');book.save(source)
    app=WebviewApp()
    api=SimpleNamespace(_app=app,_auth=SimpleNamespace(status=lambda:{'authenticated':True}),_memory_key=lambda:'self-created-test',_connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating=FloatingSession(api)
    assert floating.admit(str(source))['ok']
    assert floating.read(floating.read_token,True)['ok']
    service=floating.paged
    result=service.open(floating.document['id'],floating.document['version'])
    assert result['editable']
    return service,floating,source


def snapshot(service,text='实验于2026年完成，温度120 ℃。'):
    book=Document(service.source);book.paragraphs[1].runs[0].text=text
    book.paragraphs[1].runs[0].bold=True
    stream=io.BytesIO();book.save(stream)
    manifest=[{'id':'node-'+str(n),'text':p.text} for n,p in enumerate(book.paragraphs)]
    return base64.b64encode(stream.getvalue()).decode(),manifest


def sync(service,revision=1):
    service.changed(service.document_id,revision)
    data,manifest=snapshot(service)
    return service.sync(service.document_id,revision,data,manifest)


def test_summary_distinguishes_new_empty_deleted_and_reformatted_paragraphs():
    book = Document();book.add_paragraph('保留温度120 ℃。')
    before = io.BytesIO();book.save(before)
    added = book.add_paragraph('新增说明。')
    after = io.BytesIO();book.save(after)
    rows = changes(before.getvalue(), after.getvalue())
    assert len(rows) == 1 and rows[0]['operation'] == 'insert'
    assert rows[0]['after'] == '新增说明。' and not rows[0]['format_changed']
    assert 'format_before' not in rows[0]
    removed = changes(after.getvalue(), before.getvalue())
    assert removed[0]['operation'] == 'delete' and not removed[0]['format_changed']
    added.text = ''
    empty = io.BytesIO();book.save(empty)
    assert changes(before.getvalue(), empty.getvalue())[0]['operation'] == 'insert'
    added.text = '填写原来空白的段落'
    filled = io.BytesIO();book.save(filled)
    assert changes(empty.getvalue(), filled.getvalue())[0]['operation'] == 'edit'
    added.runs[0].bold = True
    formatted = io.BytesIO();book.save(formatted)
    row = changes(filled.getvalue(), formatted.getvalue())[0]
    assert row['operation'] == 'edit' and row['format_changed']
    assert row['before'] == row['after'] and '粗体' in row['format_after']


def test_exact_snapshot_manifest_and_utf16(draft):
    service,floating,source=draft
    sync(service)
    reference=service.reference(service.document_id,1,{'node_id':'node-1','start':0,'end':9})['reference']
    assert '2026' in reference['old']
    assert utf16_slice('中😀é文',1,3)=='😀'
    with pytest.raises(GuardError):utf16_slice('中😀文',1,2)
    data,manifest=snapshot(service)
    manifest[1]['text']='2025'
    with pytest.raises(GuardError,match='清单'):service.sync(service.document_id,1,data,manifest)
    assert digest(source)==floating.document['version']


def test_bootstrap_refuses_loss_and_external_source_conflict(draft):
    service,floating,source=draft
    data,manifest=snapshot(service)
    with pytest.raises(GuardError,match='只读'):service.sync(service.document_id,0,data,manifest,True)
    assert not service.editable
    source.write_bytes(b'changed')
    with pytest.raises(ValueError):service.reference(service.document_id,0,{'node_id':'n','start':0,'end':1})


def test_dirty_recovery_reopen_requires_same_source_and_does_not_restore_approval(draft):
    service,floating,source=draft
    sync(service)
    with pytest.raises(GuardError,match='未保存'):service.before_admit()
    service.permit_switch(service.document_id,1,'recover');service.before_admit()
    service.close();service.open(floating.document['id'],floating.document['version'])
    assert service.state()['recovery_available']
    book=Document(source);buffer=io.BytesIO();book.save(buffer)
    manifest=[{'id':'reopened-'+str(i),'text':p.text} for i,p in enumerate(book.paragraphs)]
    service.sync(service.document_id,0,base64.b64encode(buffer.getvalue()).decode(),manifest,True)
    assert service.state()['recovery_available'] and service.recovery['draft_revision']==1
    restored=service.restore(service.document_id)
    assert restored['dirty'] and b'2026' in parts(base64.b64decode(restored['data']))['word/document.xml']
    assert service.references=={} and service.publisher.plans=={} and service.proposals=={}


def test_old_proposal_token_and_approval_cannot_survive_manual_edit(draft):
    service,floating,source=draft
    sync(service)
    floating.runtime=InstructionRuntime(floating.document['version']);floating.runtime.submit('m','精简')
    proposal={'id':'p','status':'pending','draft_revision':1,'instruction_revision':1,'old':'原文','new':'建议'}
    service.proposals['p']=proposal
    grant=service.prepare_apply('p',service.document_id,1)
    plan=service.plan(service.document_id,1)['plan']
    service.changed(service.document_id,2)
    assert proposal['status']=='stale'
    with pytest.raises(GuardError):service.applied(grant['token'],service.document_id,1,2)
    with pytest.raises(GuardError):service.publisher.claim(plan['id'],document_id=service.document_id,draft_revision=2,confirmed=True)


def test_apply_ack_preserves_new_revision_and_is_one_use(draft):
    service,floating,source=draft
    sync(service)
    floating.runtime=InstructionRuntime(floating.document['version']);floating.runtime.submit('m','精简')
    service.proposals['p']={'id':'p','status':'pending','draft_revision':1,'instruction_revision':1}
    grant=service.prepare_apply('p',service.document_id,1)
    service.changed(service.document_id,2,grant['token'])
    assert service.applied(grant['token'],service.document_id,1,2)['draft_revision']==2
    assert service.proposals['p']['status']=='adopted'
    with pytest.raises(GuardError):service.applied(grant['token'],service.document_id,1,2)


def test_candidate_cancel_expiry_source_conflict_publish_and_rollback(draft,monkeypatch):
    service,floating,source=draft
    original=source.read_bytes();sync(service)
    publisher=service.publisher
    plan=service.plan(service.document_id,1)['plan']
    assert service.publish(plan['id'],service.document_id,1,False)['cancelled']
    assert not list(source.parent.glob('*.agent-*.docx'))
    plan=service.plan(service.document_id,1)['plan'];publisher.plans[plan['id']]['deadline']=0
    with pytest.raises(GuardError):publisher.claim(plan['id'],document_id=service.document_id,draft_revision=1,confirmed=True)
    plan=service.plan(service.document_id,1)['plan']
    monkeypatch.setattr('office_agent.fidelity.compare_documents',lambda *a,**kw:{'inputs_unchanged':True,'render':{'status':'different'}})
    result=publisher.execute(plan['id'],document_id=service.document_id,draft_revision=1,confirmed=True)
    assert source.read_bytes()==original
    output=source.parent/(source.stem+'.agent-'+plan['id'][:8]+'.docx')
    assert Document(output).paragraphs[1].text=='实验于2026年完成，温度120 ℃。'
    assert Document(output).paragraphs[1].runs[0].bold
    assert result['status']=='verified' and result['candidate_sha256']==digest(output)
    floating.app.editor.rollback(plan['id'],confirmed=True)
    assert not output.exists() and source.read_bytes()==original


def test_new_input_during_publish_remains_dirty(draft,monkeypatch):
    service,floating,source=draft
    sync(service);plan=service.plan(service.document_id,1)['plan']
    frozen=service.publisher.claim(plan['id'],document_id=service.document_id,draft_revision=1,confirmed=True)
    service.changed(service.document_id,2)
    monkeypatch.setattr(service.publisher,'execute',lambda *a,**kw:{'id':plan['id'],'status':'verified','output':'fixed-own-test-copy','output_sha256':'own-test'})
    service._publish(frozen,service.publisher)
    assert service.state()['published_revision']==1 and service.state()['draft_revision']==2 and service.state()['dirty']


def test_package_rejects_external_relationship_and_preserves_metadata(draft):
    service,floating,source=draft
    original=source.read_bytes();values=parts(original)
    values['word/_rels/document.xml.rels']=values['word/_rels/document.xml.rels'].replace(b'</Relationships>',b'<Relationship Id="bad" Type="image" Target="https://invalid.example/secret" TargetMode="External"/></Relationships>')
    stream=io.BytesIO()
    with zipfile.ZipFile(stream,'w') as archive:
        for name,value in values.items():archive.writestr(name,value)
    with pytest.raises(GuardError,match='外部'):parts(stream.getvalue())
    values=parts(original);values['customXml/item1.xml']=b'<changed/>'
    stream=io.BytesIO()
    with zipfile.ZipFile(stream,'w') as archive:
        for name,value in values.items():archive.writestr(name,value)
    assert parts(normalize(stream.getvalue(),original))['customXml/item1.xml']==parts(original)['customXml/item1.xml']
    assert projection(normalize(original,original))==projection(original)


def test_retained_drop_after_explicit_recovery_does_not_require_picker(draft,monkeypatch):
    import threading
    service,floating,source=draft;sync(service)
    target=source.with_name('下一份.docx');book=Document();book.add_paragraph('第二份自建材料');book.save(target)
    current=floating.document['id']
    floating.dropped((str(target),))
    assert floating.pending_draft_drop and floating.document['id']==current and floating.drop[1]==str(target)
    opened=threading.Event()
    def controlled_open(request):
        assert floating.admit(request['path'],_opening_request=request)['ok'];opened.set()
    monkeypatch.setattr(floating,'_open_native',controlled_open)
    service.permit_switch(service.document_id,1,'recover')
    assert floating.admit_drop()['opening']
    assert opened.wait(10)
    worker=floating._open_worker
    if worker:worker.join(10);assert not worker.is_alive()
    assert floating.document['name']=='下一份.docx' and floating.document['id']!=current
    assert floating.app.path==target and not floating.pending_draft_drop


def test_stop_invalidates_proposal_lease_and_unclaimed_publish(draft):
    service,floating,source=draft;sync(service)
    floating.runtime=InstructionRuntime(floating.document['version']);floating.runtime.submit('m','精简')
    service.proposals['p']={'id':'p','status':'pending','draft_revision':1,'instruction_revision':1}
    grant=service.prepare_apply('p',service.document_id,1)
    service.plan(service.document_id,1)
    floating.stop()
    assert service.proposals['p']['status']=='stale' and service.publisher.plans=={}
    with pytest.raises(GuardError):service.prepare_apply('p',service.document_id,1)
    with pytest.raises(GuardError):service.applied(grant['token'],service.document_id,1,2)
