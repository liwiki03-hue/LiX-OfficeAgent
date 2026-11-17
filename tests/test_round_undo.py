import base64
import io
from pathlib import Path
import time

from docx import Document
import pytest

from office_agent.paged_documents import GuardError
from office_agent.webui.round_undo import RoundUndo
from tests.test_paged_editor import draft


def export(document):
    stream=io.BytesIO();document.save(stream)
    manifest=[{'id':'round-node-'+str(index),'text':paragraph.text} for index,paragraph in enumerate(document.paragraphs)]
    return base64.b64encode(stream.getvalue()).decode(),manifest


def make_round(draft, monkeypatch):
    page,floating,source=draft
    monkeypatch.setattr('office_agent.fidelity.compare_documents',lambda *args,**kwargs:{'inputs_unchanged':True,'render':{'status':'different'}})
    baseline=Document(source);baseline.paragraphs[2].text='此前人工编辑保留。'
    encoded,manifest=export(baseline)
    page.changed(page.document_id,1);page.sync(page.document_id,1,encoded,manifest)
    turn,_=floating.conversation.receive('round-a','改写这一段，让意思更清楚',page.document_id,1)
    assert turn['intent']=='edit'
    floating.round_undo.capture('round-a',1)
    modified=Document(io.BytesIO(page.data));modified.paragraphs[1].text='本轮AI改稿，温度120 ℃。'
    encoded,manifest=export(modified)
    page.changed(page.document_id,2);floating.round_undo.applied('round-a',2);page.sync(page.document_id,2,encoded,manifest)
    plan=page.publisher.plan(page.data,document_id=page.document_id,draft_revision=2)
    receipt=page.publisher.execute(plan['id'],document_id=page.document_id,draft_revision=2,confirmed=True)
    floating.outputs[receipt['id']]={**receipt,'source_version':page.source_version}
    page.published_revision=2;page.published_id=receipt['id']
    floating.round_undo.published('round-a',receipt)
    assert not floating.round_undo.public('round-a')['available']
    floating.conversation.update('round-a','complete','已更新这一段。',status='complete')
    assert floating.round_undo.public('round-a')['available']
    return page,floating,source,receipt,baseline


def wait_deleted(floating):
    deadline=time.monotonic()+5
    while floating.round_undo.records['round-a']['phase']=='deleting':
        if time.monotonic()>deadline:raise AssertionError('Own rollback worker did not finish')
        time.sleep(.01)


def test_round_undo_restores_manual_baseline_and_only_removes_final(draft,monkeypatch):
    page,floating,source,receipt,baseline=make_round(draft,monkeypatch)
    original=source.read_bytes()
    assert floating.round_undo.public('round-a')['available']
    prepared=floating.round_undo.begin('round-a',page.document_id,2,True)
    assert Path(receipt['output']).exists()
    snapshot=Document(io.BytesIO(base64.b64decode(prepared['data'])))
    encoded,manifest=export(snapshot)
    floating.round_undo.restored('round-a',prepared['token'],page.document_id,3,encoded,manifest)
    wait_deleted(floating)
    assert floating.round_undo.records['round-a']['phase']=='files_done'
    assert '已撤销这次修改' not in floating.conversation.turns['round-a']['text']
    result=floating.round_undo.finish('round-a',prepared['token'],page.document_id,3)
    assert result['complete'] and not floating.round_undo.busy
    restored=Document(io.BytesIO(page.data))
    assert [paragraph.text for paragraph in restored.paragraphs]==[paragraph.text for paragraph in baseline.paragraphs]
    assert not Path(receipt['output']).exists()
    assert source.read_bytes()==original and page.state()['dirty']
    assert floating.round_undo.public('round-a')['available'] is False
    assert floating.conversation.turns['round-a']['stage']=='undone'
    assert '已撤销这次修改' in floating.conversation.turns['round-a']['text']
    assert floating.round_undo.begin('round-a',page.document_id,3,True)['complete']


@pytest.mark.parametrize('condition',['cancel','manual','file','dependency','wrong-document'])
def test_round_undo_preflight_preserves_everything_on_refusal(draft,monkeypatch,condition):
    page,floating,source,receipt,baseline=make_round(draft,monkeypatch)
    target=Path(receipt['output'])
    if condition=='manual':
        document=Document(io.BytesIO(page.data));document.paragraphs[1].text+=' 用户后来又改了这里。'
        encoded,manifest=export(document);page.changed(page.document_id,3);page.sync(page.document_id,3,encoded,manifest)
    elif condition=='file':
        document=Document(target);document.add_paragraph('外部修改保留。');document.save(target)
    elif condition=='dependency':
        floating.outputs['f'*32]={'output':'unused','source_version':receipt['output_sha256']}
    before=page.data;file_before=target.read_bytes()
    if condition=='cancel':
        assert floating.round_undo.begin('round-a',page.document_id,page.revision,False)['cancelled']
    else:
        with pytest.raises((GuardError,ValueError)):
            floating.round_undo.begin('round-a','other-document' if condition=='wrong-document' else page.document_id,page.revision,True)
    assert page.data==before and target.read_bytes()==file_before and not floating.round_undo.busy


def test_round_undo_invalid_restore_never_deletes_and_can_reconcile(draft,monkeypatch):
    page,floating,source,receipt,baseline=make_round(draft,monkeypatch)
    prepared=floating.round_undo.begin('round-a',page.document_id,2,True)
    document=Document(io.BytesIO(page.data));encoded,manifest=export(document)
    with pytest.raises(GuardError):
        floating.round_undo.restored('round-a',prepared['token'],page.document_id,3,encoded,manifest)
    assert Path(receipt['output']).exists() and page.revision==2
    floating.round_undo.abort('round-a',prepared['token'],page.document_id)
    assert not floating.round_undo.busy and floating.round_undo.public('round-a')['recoverable']
    retry=floating.round_undo.begin('round-a',page.document_id,2,True)
    assert retry['token']!=prepared['token']
    restored=Document(io.BytesIO(base64.b64decode(retry['data'])));encoded,manifest=export(restored)
    floating.round_undo.restored('round-a',retry['token'],page.document_id,3,encoded,manifest)
    wait_deleted(floating)
    assert floating.round_undo.finish('round-a',retry['token'],page.document_id,3)['complete']


def test_round_undo_group_reverses_all_final_outputs(draft,monkeypatch):
    page,floating,source,first,baseline=make_round(draft,monkeypatch)
    original=source.read_bytes()
    plan=page.publisher.plan(page.data,document_id=page.document_id,draft_revision=2)
    second=page.publisher.execute(plan['id'],document_id=page.document_id,draft_revision=2,confirmed=True)
    floating.outputs[second['id']]={**second,'source_version':page.source_version}
    floating.round_undo.published('round-a',second)
    calls=[];rollback=floating.app.editor.rollback
    def record(identifier,**kwargs):
        calls.append(identifier)
        return rollback(identifier,**kwargs)
    monkeypatch.setattr(floating.app.editor,'rollback',record)
    prepared=floating.round_undo.begin('round-a',page.document_id,2,True)
    encoded,manifest=export(Document(io.BytesIO(base64.b64decode(prepared['data']))))
    floating.round_undo.restored('round-a',prepared['token'],page.document_id,3,encoded,manifest)
    wait_deleted(floating)
    assert floating.round_undo.finish('round-a',prepared['token'],page.document_id,3)['complete']
    assert calls==[second['id'],first['id']]
    assert all(not Path(row['output']).exists() for row in (first,second))
    assert source.read_bytes()==original


def test_round_undo_reload_reconciles_audits_without_repeat_deletion(draft,monkeypatch):
    page,floating,source,receipt,baseline=make_round(draft,monkeypatch)
    prepared=floating.round_undo.begin('round-a',page.document_id,2,True)
    encoded,manifest=export(Document(io.BytesIO(base64.b64decode(prepared['data']))))
    floating.round_undo.restored('round-a',prepared['token'],page.document_id,3,encoded,manifest)
    wait_deleted(floating)
    assert not Path(receipt['output']).exists()
    floating.round_undo=RoundUndo(floating);floating.app.busy=False
    floating.round_undo.synced()
    assert floating.round_undo.public('round-a')['recoverable']
    with pytest.raises(GuardError):
        floating.round_undo.finish('round-a',prepared['token'],page.document_id,3)
    def no_repeat(*args,**kwargs):
        raise AssertionError('An already rolled-back file must not be deleted twice')
    monkeypatch.setattr(floating.app.editor,'rollback',no_repeat)
    retry=floating.round_undo.begin('round-a',page.document_id,3,True)
    assert retry['token']!=prepared['token']
    encoded,manifest=export(Document(io.BytesIO(base64.b64decode(retry['data']))))
    floating.round_undo.restored('round-a',retry['token'],page.document_id,4,encoded,manifest)
    wait_deleted(floating)
    assert floating.round_undo.finish('round-a',retry['token'],page.document_id,4)['complete']


def test_round_undo_actor_switch_invalidates_old_controls(draft,monkeypatch):
    page,floating,source,receipt,baseline=make_round(draft,monkeypatch)
    prepared=floating.round_undo.begin('round-a',page.document_id,2,True)
    monkeypatch.setattr(floating.api._auth,'account_key',lambda:'another-account',raising=False)
    assert floating.round_undo.public('round-a') is None
    for action in (lambda:floating.round_undo.status('round-a',prepared['token'],page.document_id),
                   lambda:floating.round_undo.abort('round-a',prepared['token'],page.document_id),
                   lambda:floating.round_undo.restored('round-a',prepared['token'],page.document_id,3,'',[])):
        with pytest.raises(GuardError):action()
    assert Path(receipt['output']).exists() and page.revision==2


def test_round_undo_blocks_new_publication_until_visible_finish(draft,monkeypatch):
    page,floating,source,receipt,baseline=make_round(draft,monkeypatch)
    prepared=floating.round_undo.begin('round-a',page.document_id,2,True)
    with pytest.raises(GuardError,match='正在撤销'):
        page.plan(page.document_id,2)
    floating.round_undo.abort('round-a',prepared['token'],page.document_id)
    assert Path(receipt['output']).exists() and page.revision==2
