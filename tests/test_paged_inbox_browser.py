from copy import deepcopy
import json
import os
from pathlib import Path
import shutil
import threading
import time

from docx import Document

from tests.test_paged_browser import engine_frame, paged_browser


def wait(page, predicate, seconds=15):
    deadline=time.monotonic()+seconds
    while not predicate():
        if time.monotonic()>deadline:raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(50)


def select_model(page):
    page.locator('#simple-model-menu > summary').click()
    page.locator('#model-selector').select_option(json.dumps(['official','same-model','agent'],separators=(',',':')))
    page.keyboard.press('Escape')


def select_reference(page, old='实验于2025年完成，温度120 ℃。', new=None):
    frame=engine_frame(page)
    frame.locator('.superdoc-page').get_by_text(old,exact=True).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End')
    if new:
        page.keyboard.insert_text(new)
        page.keyboard.press('Home');page.keyboard.press('Shift+End')
    page.get_by_role('button',name='引用选区',exact=True).click()
    page.wait_for_function("document.getElementById('paged-reference-state').textContent.startsWith('已引用')")


class Suggestion:
    def __init__(self, row, calls, blocked=None):
        self.row=deepcopy(row);self.calls=calls;self.blocked=blocked;self.count=0

    def bind_tools(self, tools):return self

    def invoke(self, messages):
        self.count+=1;self.calls.append(str(messages))
        if self.blocked and self.count==1:
            self.blocked[0].set()
            assert self.blocked[1].wait(15)
        if self.count==1:
            return {'tool_calls':[{'name':'propose_draft_edit','id':'inbox-call','args':{
                'reference_id':self.row['id'],'old':self.row['old'],
                'new':self.row['old'].replace('实验于','实验完成于').replace('年完成','年'),
                'reason':'精简措辞，保留年份、温度和单位。'}}]}
        return {'content':'建议待审阅，尚未改动文件。'}


def test_receive_without_model_then_review_apply_and_verified_copy(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    original=source.read_bytes();calls=[]
    monkeypatch.setattr(api._connections,'build_model',lambda route:Suggestion(next(iter(api._floating.paged.references.values())),calls))
    page.locator('#floating-input').fill('精简所选实验描述，保留年份、温度和单位。')
    page.locator('#floating-send').click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    assert not calls and api._floating.runtime.revision==1
    assert api._floating.paged.state()['pending_reason']=='reference'
    page.wait_for_function("document.getElementById('paged-requirement-state').textContent.includes('待引用')")
    for width in [1280,520]:
        page.set_viewport_size({'width':width,'height':820})
        for theme in ['light','dark','highcontrast']:
            page.evaluate('(theme)=>document.documentElement.dataset.theme=theme',theme)
            for identity in ['floating-input','floating-send','paged-requirement-state','paged-requirement-details']:
                box=page.locator('#'+identity).bounding_box()
                assert box and 0<=box['x']<box['x']+box['width']<=width and box['y']+box['height']<=820
    page.set_viewport_size({'width':1280,'height':820})
    page.evaluate("document.documentElement.dataset.theme='dark'")
    evidence=Path(os.environ['OFFICE_AGENT_INBOX_EVIDENCE']) if os.getenv('OFFICE_AGENT_INBOX_EVIDENCE') else None
    if evidence:
        evidence.mkdir(parents=True,exist_ok=True)
        page.screenshot(path=str(evidence/'01-received-waiting-reference.png'))
    page.get_by_role('button',name='查看已收到要求',exact=True).click()
    assert page.get_by_role('dialog').get_by_text('精简所选实验描述，保留年份、温度和单位。',exact=True).is_visible()
    page.get_by_role('dialog').get_by_role('button',name='关闭',exact=True).click()
    assert page.evaluate('document.activeElement.id')=='paged-requirement-details'
    select_model(page);select_reference(page)
    assert not calls
    page.get_by_role('button',name='处理已保存要求',exact=True).click()
    page.get_by_role('button',name='采用到草稿',exact=True).wait_for(state='visible')
    assert '精简所选实验描述' in calls[0] and api._floating.runtime.revision==1
    if evidence:page.screenshot(path=str(evidence/'02-proposal.png'))
    page.get_by_role('button',name='采用到草稿',exact=True).click()
    wait(page,lambda:any(row['status']=='adopted' for row in api._floating.paged.proposals.values()))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).click()
    assert not list(source.parent.glob('*.agent-*.docx'))
    page.get_by_role('button',name='保存副本',exact=True).click()
    page.get_by_role('button',name='确认生成副本',exact=True).click()
    wait(page,lambda:bool(api._floating.outputs),180)
    receipt=next(iter(api._floating.outputs.values()))
    book=Document(receipt['output'])
    assert receipt['status']=='verified' and receipt['render_diff']['inputs_unchanged']
    assert book.paragraphs[1].text=='实验完成于2025年，温度120 ℃。'
    assert book.paragraphs[2].text=='非目标段落不变。'
    assert book.tables[0].cell(1,1).text=='120 ms'
    assert source.read_bytes()==original and errors==[]
    page.wait_for_function("document.getElementById('paged-save-state').textContent.includes('副本已验证')")
    if evidence:
        page.screenshot(path=str(evidence/'03-verified-copy.png'))
        shutil.copy2(source,evidence/'source.docx');shutil.copy2(receipt['output'],evidence/'verified-copy.docx')
        (evidence/'RESULT.json').write_text(json.dumps({'status':receipt['status'],'source_unchanged':True,'model_invocations':len(calls),'scope':'controlled model; real DOCX engine and render gate; browser, not native or cloud','render':receipt['render_diff']},ensure_ascii=False,indent=2),encoding='utf8')


def test_inflight_invalid_reference_receives_then_requires_explicit_processing(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    calls=[];entered=threading.Event();release=threading.Event();built=[]
    def build(route):
        model=Suggestion(next(iter(api._floating.paged.references.values())),calls,(entered,release) if not built else None)
        built.append(model);return model
    monkeypatch.setattr(api._connections,'build_model',build)
    select_model(page);select_reference(page)
    page.locator('#floating-input').fill('精简文字。');page.locator('#floating-send').click()
    wait(page,entered.is_set)
    try:
        frame=engine_frame(page)
        frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
        page.keyboard.press('Home');page.keyboard.press('Shift+End');page.keyboard.insert_text('实验于2027年完成，温度120 ℃。')
        page.locator('#floating-input').fill('保留2027年和120 ℃，使用最新草稿。');page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
        assert api._floating.runtime.revision==2
    finally:release.set()
    wait(page,lambda:not api._app.busy)
    assert len(built)==1 and not api._floating.paged.proposals
    select_reference(page,'实验于2027年完成，温度120 ℃。')
    page.wait_for_timeout(800)
    assert len(built)==1
    page.get_by_role('button',name='处理已保存要求',exact=True).click()
    page.get_by_role('button',name='采用到草稿',exact=True).wait_for(state='visible')
    assert len(built)==2 and '2027' in calls[-1]
    assert all('2027' in row['old'] for row in api._floating.paged.proposals.values())
    assert errors==[]


def test_receipt_failure_preserves_input_and_retry_is_user_initiated(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    persist=api._floating.persist
    def fail(_):raise OSError('simulated write failure')
    monkeypatch.setattr(api._floating,'persist',fail)
    page.locator('#floating-input').fill('保留这一条待处理要求。');page.locator('#floating-send').click()
    page.wait_for_function("document.getElementById('paged-feedback').textContent.length>0")
    assert page.locator('#floating-input').input_value()=='保留这一条待处理要求。'
    assert not api._floating.runtime.messages
    monkeypatch.setattr(api._floating.runtime,'persist',persist)
    page.locator('#floating-send').click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    assert len(api._floating.runtime.messages)==1 and not api._app.busy
    assert errors==[]


def test_stop_then_explicit_new_requirement_waits_for_old_worker(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    calls=[];entered=threading.Event();release=threading.Event();built=[]
    def build(route):
        model=Suggestion(next(iter(api._floating.paged.references.values())),calls,(entered,release) if not built else None)
        built.append(model);return model
    monkeypatch.setattr(api._connections,'build_model',build)
    select_model(page);select_reference(page)
    page.locator('#floating-input').fill('第一轮旧要求。');page.locator('#floating-send').click()
    wait(page,entered.is_set)
    old_session=api._floating.runtime.session_id
    try:
        page.locator('#floating-stop').click()
        wait(page,lambda:api._floating.runtime.closed)
        page.wait_for_function("id=>document.getElementById('paged-requirement-state').dataset.inbox===id",arg=api._floating.paged.inbox_id)
        page.locator('#floating-input').fill('停止后主动提出的新要求，保留所有数字。');page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
        assert api._floating.runtime.session_id!=old_session and len(built)==1
    finally:release.set()
    wait(page,lambda:len(built)==2)
    page.get_by_role('button',name='采用到草稿',exact=True).wait_for(state='visible')
    assert '第一轮旧要求' not in calls[-1]
    assert '停止后主动提出' in calls[-1] and len(api._floating.runtime.messages)==1
    assert errors==[]


def test_late_receipt_after_new_round_cannot_clear_new_input(paged_browser):
    page,api,source,errors=paged_browser
    page.evaluate("""()=>{const previous=window.pywebview.api;window.pywebview.api=new Proxy(previous,{get:(target,name)=>name==='paged_submit'?async(...args)=>{const result=await target[name](...args);return new Promise(resolve=>{window.releaseInboxReceipt=()=>resolve(result);});}:target[name]});}""")
    page.locator('#floating-input').fill('上一轮待引用的要求。');page.locator('#floating-send').click()
    page.wait_for_function("typeof window.releaseInboxReceipt==='function'")
    page.locator('#floating-input').fill('下一轮尚未发送的新草稿。')
    page.get_by_label('更多',exact=True).click()
    page.get_by_role('button',name='开始新一轮要求',exact=True).click()
    wait(page,lambda:api._floating.runtime is None)
    page.evaluate('window.releaseInboxReceipt()')
    page.wait_for_function("!document.getElementById('floating-send').disabled")
    assert page.locator('#floating-input').input_value()=='下一轮尚未发送的新草稿。'
    assert api._floating.runtime is None and not api._app.busy
    assert errors==[]


def test_stop_resets_requirement_counter_and_rejects_retired_session_snapshot(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    original=source.read_bytes()
    calls=[];entered=threading.Event();release=threading.Event();built=[]
    def build(route):
        model=Suggestion(next(iter(api._floating.paged.references.values())),calls,(entered,release) if not built else None)
        built.append(model);return model
    monkeypatch.setattr(api._connections,'build_model',build)
    for text in ['旧要求一。','旧要求二。','旧要求三。']:
        page.locator('#floating-input').fill(text);page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
    assert api._floating.runtime.revision==3 and not calls
    select_model(page);select_reference(page)
    page.get_by_role('button',name='处理已保存要求',exact=True).click()
    wait(page,entered.is_set)
    old_session=api._floating.runtime.session_id
    old_state=deepcopy(api.floating_snapshot(0))
    try:
        page.locator('#floating-stop').click()
        wait(page,lambda:api._floating.runtime.closed)
        page.wait_for_function("document.getElementById('floating-status').textContent.includes('已停止后续派发')")
        page.locator('#floating-input').fill('新一轮精简文字，保留数字和单位。');page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
        page.wait_for_function("document.getElementById('paged-requirement-state').textContent.startsWith('已收到 1 条')")
        page.wait_for_function("!document.getElementById('floating-status').textContent.includes('已停止后续派发')")
        assert api._floating.runtime.session_id!=old_session and len(built)==1
        page.evaluate('state=>window.OfficePaged.update(state)',old_state)
        assert page.locator('#paged-requirement-state').text_content().startswith('已收到 1 条')
        page.locator('#floating-input').fill('追加：保留表格和结论。');page.locator('#floating-send').click()
        page.wait_for_function("document.getElementById('floating-input').value === ''")
        page.wait_for_function("document.getElementById('paged-requirement-state').textContent.startsWith('已收到 2 条')")
        assert api._floating.runtime.revision==2 and len(built)==1
    finally:release.set()
    page.get_by_role('button',name='采用到草稿',exact=True).wait_for(state='visible')
    assert len(built)==2 and len(api._floating.paged.proposals)==1
    assert '新一轮精简文字' in calls[-1] and '保留表格和结论' in calls[-1] and '旧要求一' not in calls[-1]
    assert source.read_bytes()==original and not list(source.parent.glob('*.agent-*.docx')) and errors==[]


def test_late_processing_ack_cannot_restore_previous_round(paged_browser,monkeypatch):
    page,api,source,errors=paged_browser
    calls=[]
    monkeypatch.setattr(api._connections,'build_model',lambda route:Suggestion(next(iter(api._floating.paged.references.values())),calls))
    page.locator('#floating-input').fill('先保存，稍后处理。');page.locator('#floating-send').click()
    page.wait_for_function("document.getElementById('floating-input').value === ''")
    select_model(page);select_reference(page)
    page.evaluate("""()=>{const previous=window.pywebview.api;window.pywebview.api=new Proxy(previous,{get:(target,name)=>name==='paged_start_pending'?async(...args)=>{const result=await target[name](...args);return new Promise(resolve=>{window.releaseProcessingReceipt=()=>resolve(result);});}:target[name]});}""")
    page.get_by_role('button',name='处理已保存要求',exact=True).click()
    page.wait_for_function("typeof window.releaseProcessingReceipt==='function'")
    wait(page,lambda:not api._app.busy)
    page.locator('#floating-input').fill('新一轮输入保持。')
    page.get_by_label('更多',exact=True).click()
    page.get_by_role('button',name='开始新一轮要求',exact=True).click()
    wait(page,lambda:api._floating.runtime is None)
    page.evaluate('window.releaseProcessingReceipt()')
    page.wait_for_function("document.getElementById('paged-requirement-state').textContent===''")
    assert page.locator('#floating-input').input_value()=='新一轮输入保持。'
    assert not api._floating.paged.proposals and errors==[]
