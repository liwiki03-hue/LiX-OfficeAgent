import json
import os
from pathlib import Path
import shutil
import threading

from docx import Document

from tests.test_paged_browser import paged_browser
from tests.test_paged_inbox_browser import Suggestion, select_model, select_reference, wait


def send(page,text):
    page.locator('#floating-input').fill(text);page.locator('#floating-send').click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")


def test_edit_cancel_then_only_effective_goal_generates_verified_copy(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser;original=source.read_bytes();calls=[]
    monkeypatch.setattr(api._connections,'build_model',lambda route:Suggestion(next(iter(api._floating.paged.references.values())),calls))
    send(page,'旧要求：将主题改成无人机导航。')
    send(page,'待取消要求：加入没有依据的结论。')
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    dialog=page.get_by_role('dialog');items=dialog.locator('section[data-requirement]')
    items.nth(0).get_by_role('button',name='修改要求',exact=True).click()
    dialog.get_by_role('textbox',name='修改要求 1',exact=True).fill('精简所选实验描述，保留年份、温度和单位。')
    dialog.get_by_role('button',name='保存修改',exact=True).click()
    wait(page,lambda:api._floating.runtime.revision==3)
    items.nth(1).get_by_role('button',name='取消要求',exact=True).click()
    dialog.get_by_role('button',name='确认取消要求',exact=True).click()
    wait(page,lambda:api._floating.runtime.revision==4)
    assert not calls
    items.nth(1).get_by_role('heading').filter(has_text='已取消后续处理').wait_for(state='visible')
    assert items.nth(1).get_by_role('heading').inner_text().endswith('已取消后续处理')
    items.nth(0).get_by_text('原文与变更记录',exact=True).click()
    assert items.nth(0).get_by_text('已被替代：旧要求：将主题改成无人机导航。',exact=True).is_visible()
    evidence=Path(os.environ['OFFICE_AGENT_CHANGE_EVIDENCE']) if os.getenv('OFFICE_AGENT_CHANGE_EVIDENCE') else None
    if evidence:
        evidence.mkdir(parents=True,exist_ok=True);page.screenshot(path=str(evidence/'01-changed-and-cancelled.png'))
    dialog.get_by_role('button',name='关闭',exact=True).click()
    select_model(page);select_reference(page)
    page.get_by_role('button',name='处理已保存要求',exact=True).click()
    page.get_by_role('button',name='采用到草稿',exact=True).wait_for(state='visible')
    assert calls and all('无人机导航' not in text and '没有依据的结论' not in text for text in calls)
    assert all('精简所选实验描述' in text for text in calls)
    page.get_by_role('button',name='采用到草稿',exact=True).click()
    wait(page,lambda:any(row['status']=='adopted' for row in api._floating.paged.proposals.values()))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).click()
    assert not list(source.parent.glob('*.agent-*.docx'))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('button',name='确认生成副本',exact=True).click()
    wait(page,lambda:bool(api._floating.outputs),180)
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('副本已验证')")
    receipt=next(iter(api._floating.outputs.values()));book=Document(receipt['output'])
    assert receipt['status']=='verified' and receipt['render_diff']['inputs_unchanged']
    assert book.paragraphs[1].text=='实验完成于2025年，温度120 ℃。'
    assert book.paragraphs[2].text=='非目标段落不变。' and book.tables[0].cell(1,1).text=='120 ms'
    assert source.read_bytes()==original and errors==[]
    if evidence:
        page.screenshot(path=str(evidence/'02-verified-copy.png'))
        shutil.copy2(source,evidence/'source.docx');shutil.copy2(receipt['output'],evidence/'verified-copy.docx')
        (evidence/'RESULT.json').write_text(json.dumps({'status':receipt['status'],'source_unchanged':True,'model_invocations':len(calls),'obsolete_goals_absent':True,'scope':'controlled model; real browser, DOCX engine and render gate; not cloud or native','render':receipt['render_diff']},ensure_ascii=False,indent=2),encoding='utf8')


def test_cancel_inflight_never_dispatches_old_action_or_automatic_retry(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser;original=source.read_bytes()
    entered=threading.Event();release=threading.Event();calls=[]
    monkeypatch.setattr(api._connections,'build_model',lambda route:Suggestion(next(iter(api._floating.paged.references.values())),calls,(entered,release)))
    select_model(page);select_reference(page);send(page,'精简这句实验描述。');wait(page,entered.is_set)
    try:
        page.get_by_role('button',name='查看已收到要求',exact=True).click()
        dialog=page.get_by_role('dialog');dialog.get_by_role('button',name='取消要求',exact=True).click()
        dialog.get_by_role('button',name='确认取消要求',exact=True).click()
        wait(page,lambda:api._floating.runtime.revision==2)
        assert api._app.busy and len(calls)==1
        dialog.get_by_role('alert').filter(has_text='已取消后续处理').wait_for(state='visible')
        assert '不保证撤回' in dialog.get_by_role('alert').inner_text()
        dialog.get_by_role('button',name='关闭',exact=True).click()
    finally:release.set()
    wait(page,lambda:not api._app.busy)
    assert not api._floating.paged.proposals and len(calls)==1
    assert not api._floating.paged.state()['pending_requirements']
    assert source.read_bytes()==original and errors==[]


def test_failed_change_preserves_text_and_escape_then_manual_retry(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser;send(page,'保留这一条要求。')
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    dialog=page.get_by_role('dialog');dialog.get_by_role('button',name='修改要求',exact=True).click()
    field=dialog.get_by_role('textbox',name='修改要求 1',exact=True);field.fill('尚未保存的新要求。')
    persist=api._floating.runtime.persist
    def fail(_):raise OSError('simulated disk failure')
    monkeypatch.setattr(api._floating.runtime,'persist',fail)
    dialog.get_by_role('button',name='保存修改',exact=True).click()
    wait(page,lambda:bool(dialog.get_by_role('alert').inner_text()))
    assert field.input_value()=='尚未保存的新要求。' and api._floating.runtime.revision==1
    page.evaluate('window.confirm=()=>false');page.keyboard.press('Escape')
    assert dialog.is_visible() and field.input_value()=='尚未保存的新要求。'
    monkeypatch.setattr(api._floating.runtime,'persist',persist)
    dialog.get_by_role('button',name='保存修改',exact=True).click()
    wait(page,lambda:api._floating.runtime.revision==2)
    assert len(api._floating.runtime.requirement_changes)==1 and not api._app.busy
    page.evaluate('window.confirm=()=>true');dialog.get_by_role('button',name='关闭',exact=True).click()
    assert page.evaluate('document.activeElement.id')=='paged-requirement-details' and errors==[]


def test_stale_editor_requires_refresh_and_preserves_unsaved_draft(paged_browser):
    page,api,source,errors=paged_browser;send(page,'原始要求。')
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    dialog=page.get_by_role('dialog');dialog.get_by_role('button',name='修改要求',exact=True).click()
    field=dialog.get_by_role('textbox',name='修改要求 1',exact=True);field.fill('当前窗口的未保存文字。')
    service=api._floating.paged;state=service.state();row=state['requirements'][0]
    service.change_requirement(service.document_id,service.source_version,service.inbox_id,state['requirement_session_id'],'other-window',row['id'],row['version'],state['instruction_revision'],'replace','另一窗口已保存的要求。')
    dialog.get_by_role('button',name='保存修改',exact=True).click()
    wait(page,lambda:'列表已更新' in dialog.get_by_role('alert').inner_text())
    assert field.input_value()=='当前窗口的未保存文字。'
    page.evaluate('window.confirm=()=>false');dialog.get_by_role('button',name='重新读取要求',exact=True).click()
    assert field.input_value()=='当前窗口的未保存文字。'
    page.evaluate('window.confirm=()=>true');dialog.get_by_role('button',name='重新读取要求',exact=True).click()
    dialog.get_by_text('另一窗口已保存的要求。',exact=True).wait_for(state='visible')
    assert not dialog.get_by_role('textbox').count() and api._floating.runtime.revision==2 and errors==[]


def test_late_change_ack_keeps_newer_requirement_version(paged_browser):
    page,api,source,errors=paged_browser;send(page,'第一版要求。')
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    dialog=page.get_by_role('dialog');dialog.get_by_role('button',name='修改要求',exact=True).click()
    dialog.get_by_role('textbox',name='修改要求 1',exact=True).fill('当前窗口保存的第二版。')
    page.evaluate("""()=>{const previous=window.pywebview.api;window.pywebview.api=new Proxy(previous,{get:(target,name)=>name==='paged_change_requirement'?async(...args)=>{const result=await target[name](...args);const delivered=await new Promise(resolve=>{window.releaseChangeReceipt=()=>resolve(result);});window.changeReceiptDelivered=true;return delivered;}:target[name]});}""")
    dialog.get_by_role('button',name='保存修改',exact=True).click()
    page.wait_for_function("typeof window.releaseChangeReceipt==='function'")
    service=api._floating.paged;state=service.state();row=state['requirements'][0]
    service.change_requirement(service.document_id,service.source_version,service.inbox_id,state['requirement_session_id'],'newer',row['id'],row['version'],2,'replace','另一窗口已保存的第三版。')
    page.wait_for_function("document.getElementById('floating-revisions').textContent.includes('收到 3')")
    page.evaluate('window.releaseChangeReceipt()')
    dialog.get_by_text('另一窗口已保存的第三版。',exact=True).wait_for(state='visible')
    assert '随后有更新' in dialog.get_by_role('alert').inner_text()
    assert api._floating.runtime.revision==3 and errors==[]


def test_private_clear_drops_late_change_ui_without_claiming_server_rollback(paged_browser):
    page,api,source,errors=paged_browser;send(page,'需要编辑的要求。')
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    dialog=page.get_by_role('dialog');dialog.get_by_role('button',name='修改要求',exact=True).click()
    dialog.get_by_role('textbox',name='修改要求 1',exact=True).fill('已送达宿主的新要求。')
    page.evaluate("""()=>{const previous=window.pywebview.api;window.pywebview.api=new Proxy(previous,{get:(target,name)=>name==='paged_change_requirement'?async(...args)=>{const result=await target[name](...args);const delivered=await new Promise(resolve=>{window.releaseChangeReceipt=()=>resolve(result);});window.changeReceiptDelivered=true;return delivered;}:target[name]});}""")
    dialog.get_by_role('button',name='保存修改',exact=True).click()
    page.wait_for_function("typeof window.releaseChangeReceipt==='function'")
    assert api._floating.runtime.revision==2
    page.evaluate("window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    assert not page.get_by_role('dialog').count()
    page.locator('#floating-input').fill('清空后新填写的本地草稿。')
    page.evaluate('window.releaseChangeReceipt()')
    page.wait_for_function('window.changeReceiptDelivered===true')
    assert not page.get_by_role('dialog').count()
    assert page.locator('#floating-input').input_value()=='清空后新填写的本地草稿。'
    assert api._floating.runtime.revision==2 and not api._app.busy and errors==[]


def test_old_receive_receipt_cannot_restore_superseded_requirement(paged_browser):
    page,api,source,errors=paged_browser
    send(page,'待修改的原要求。')
    page.evaluate("""()=>{const original=window.pywebview.api;window.pywebview.api=new Proxy(original,{get:(target,name)=>name==='paged_submit'?async(...args)=>{const result=await target[name](...args);return new Promise(resolve=>{window.releaseOldReceive=()=>resolve(result);});}:target[name]});}""")
    page.locator('#floating-input').fill('第二条保留要求。')
    page.locator('#floating-send').click()
    page.wait_for_function("typeof window.releaseOldReceive==='function'")
    service=api._floating.paged;runtime=api._floating.runtime
    row=service.state()['requirements'][0]
    service.change_requirement(service.document_id,service.source_version,service.inbox_id,runtime.session_id,'other-client',row['id'],row['version'],runtime.revision,'replace','另一窗口已确认的新要求。')
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    dialog=page.get_by_role('dialog')
    if not dialog.get_by_text('另一窗口已确认的新要求。',exact=True).count():
        dialog.get_by_role('button',name='重新读取要求',exact=True).click()
    dialog.get_by_text('另一窗口已确认的新要求。',exact=True).wait_for(state='visible')
    dialog.get_by_role('button',name='关闭',exact=True).click()
    page.evaluate('window.OfficePaged.update=async()=>{}')
    page.evaluate('window.releaseOldReceive()')
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    assert page.get_by_role('dialog').get_by_text('另一窗口已确认的新要求。',exact=True).is_visible()
    assert runtime.revision==3 and errors==[]
