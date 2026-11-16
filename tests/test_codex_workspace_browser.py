from pathlib import Path
from types import SimpleNamespace
import importlib
import json

import pytest
from tests.ui_confirmation import wait_for_browser_event

from tests.test_professional_browser import _browser, STATIC
from tests.test_daily_browser import setup_daily
from tests.test_model_connections import connection, payload, save
from tests.test_ui_simplify_browser import simplify, mode, pick_model
from tests.test_ui_simplify_browser import open_document


OUT=STATIC.parents[2]/'deliverables/codex-workspace-20261103/browser'


def test_history_entry_refreshes_once_without_losing_task_draft(tmp_path, monkeypatch):
    fixture = _browser(tmp_path, monkeypatch)
    try:
        page, api, original, calls, receipts = next(fixture)
        setup_daily(page, api, original, calls, receipts, monkeypatch)
        simplify(page)
        mode(page, 'chat')
        page.locator('#prompt').fill('返回后保留这段草稿')
        page.wait_for_function('!historyLoading')
        listings = []
        def latest():
            listings.append(True)
            return {'ok': True, 'sessions': [], 'warning': '本次已读取最新历史'}
        monkeypatch.setattr(api, 'list_sessions', latest)
        page.locator('#nav-history').click()
        page.wait_for_function('!historyLoading && historyWarning==="本次已读取最新历史"')
        assert len(listings) == 1
        assert page.evaluate('historyWarning') == '本次已读取最新历史'
        page.locator('#simple-history-documents').click()
        assert len(listings) == 1
        page.locator('#nav-settings').click()
        page.locator('#settings-close').click()
        assert len(listings) == 1
        page.locator('#nav-work').click()
        assert page.locator('#prompt').input_value() == '返回后保留这段草稿'
        page.locator('#nav-history').click()
        page.wait_for_function('!historyLoading')
        assert len(listings) == 2
        assert not [x for x in calls if x in {'ask', 'confirm_image_generation'}]
    finally:
        fixture.close()


@pytest.mark.parametrize('kind',['documents','chat','image'])
def test_history_settings_focus_preserves_one_view(tmp_path,monkeypatch,kind):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        setup_daily(page,api,original,calls,receipts,monkeypatch); simplify(page); mode(page,'chat')
        page.locator('#prompt').fill('历史切换后应保留的草稿')
        page.locator('#nav-history').click()
        page.locator('#simple-history-'+kind).click()
        page.wait_for_function('document.getElementById("workspace").dataset.uiScreen==="history"')
        title=page.locator('#page-title').inner_text()
        for _ in range(2):
            page.locator('#nav-settings').click()
            page.locator('#settings-close').click()
            page.locator('#nav-settings').focus()
            assert page.locator('#page-title').inner_text()==title=='历史记录'
            assert page.locator('#simple-session-menu').is_hidden()
            assert page.locator('#history-reader').is_visible()
            assert page.locator('#nav-history').get_attribute('aria-current')=='page'
        page.locator('#nav-work').click()
        assert page.locator('#prompt').input_value()=='历史切换后应保留的草稿'
        assert not [x for x in calls if x in {'ask','confirm_image_generation'}]
    finally:
        fixture.close()


@pytest.mark.parametrize('width,height',[(1440,900),(1160,760),(820,620),(520,800)])
@pytest.mark.parametrize('theme',['dark','light','highcontrast'])
def test_codex_layout_controls_popups_and_scroll(tmp_path,monkeypatch,width,height,theme):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        save(api._connections,payload([connection()]))
        setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page); pick_model(page,'chat','same-model')
        page.set_viewport_size({'width':width,'height':height})
        page.evaluate('(t)=>document.documentElement.dataset.theme=t',theme)
        page.locator('#prompt').fill('请解释技术报告摘要的结构。\n保留适用条件、数字与单位。📄 e\u0301')
        assert page.locator('#prompt').evaluate('n=>n.tagName')=='TEXTAREA'
        assert page.locator('#prompt').bounding_box()['height']>=64
        assert page.locator('#btn-ask').bounding_box()['y']+page.locator('#btn-ask').bounding_box()['height']<=height
        page.locator('#simple-model-menu > summary').click()
        box=page.locator('#simple-model-menu > .simple-menu-body').bounding_box()
        assert 0<=box['x'] and box['x']+box['width']<=width
        assert 0<=box['y'] and box['y']+box['height']<=height
        page.keyboard.press('Escape')
        assert page.locator('#simple-model-menu > summary').evaluate('n=>n===document.activeElement')
        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
        folder=OUT/'matrix'; folder.mkdir(parents=True,exist_ok=True)
        page.screenshot(path=str(folder/f'chat-{width}-{theme}.png'))
        mode(page,'agent')
        assert page.locator('#simple-start-document').is_visible()
        assert page.locator('#workspace-stage').is_hidden()
        open_document(page,'task')
        page.locator('#workspace-task-advanced > summary').click()
        page.locator('#workspace-see-task-review').scroll_into_view_if_needed()
        assert page.locator('#workspace-see-task-review').is_visible()
        assert page.locator('#workspace-task').evaluate('n=>n.scrollHeight>n.clientHeight')
        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
        page.screenshot(path=str(folder/f'task-{width}-{theme}.png'))
        assert not [x for x in calls if x in {'ask','confirm_image_generation','confirm_professional_edit'}]
    finally:
        fixture.close()


def test_multiline_shortcuts_modal_and_next_draft(tmp_path,monkeypatch):
    import threading
    fixture=_browser(tmp_path,monkeypatch)
    ready=threading.Event()
    release=threading.Event()
    try:
        page,api,original,calls,receipts=next(fixture)
        save(api._connections,payload([connection()]))
        setup_daily(page,api,original,calls,receipts,monkeypatch)
        def invoke(messages):
            ready.set()
            if not release.wait(15): raise TimeoutError('isolated test deadline')
            return SimpleNamespace(content='摘要应对应目的、方法、结果与适用条件。',response_metadata={'finish_reason':'stop'})
        monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=invoke))
        simplify(page); pick_model(page,'chat','same-model')
        page.locator('#prompt').fill('请润色以下摘要：')
        page.locator('#prompt').press('Shift+Enter')
        assert '\n' in page.locator('#prompt').input_value()
        page.locator('#prompt').dispatch_event('keydown',{'key':'Enter','isComposing':True})
        assert calls.count('ask')==0
        page.locator('#nav-settings').click()
        assert page.locator('#workspace').evaluate('n=>n.inert')
        page.locator('#settings-close').press('Control+Enter')
        assert calls.count('ask')==0
        page.locator('#settings-close').click()
        assert not page.locator('#workspace').evaluate('n=>n.inert')
        page.locator('#prompt').press('Control+Enter')
        page.wait_for_function('document.getElementById("btn-stop").disabled===false')
        wait_for_browser_event(page,ready,5)
        assert ready.wait(5)
        page.locator('#prompt').fill('下一条尚未发送的草稿 📄')
        release.set()
        page.wait_for_function('document.getElementById("chat-transcript").textContent.includes("摘要应对应")')
        assert page.locator('#prompt').input_value()=='下一条尚未发送的草稿 📄'
        assert calls.count('ask')==1
    finally:
        release.set()
        fixture.close()


def test_settings_research_and_history_search_remain_reachable(tmp_path,monkeypatch):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        setup_daily(page,api,original,calls,receipts,monkeypatch); simplify(page)
        page.locator('#nav-resources').click()
        page.locator('#workspace-tool-workbench-panel > summary').click()
        assert page.locator('#simple-project-management').is_visible()
        assert page.locator('[data-research-step="0"]').is_visible()
        page.locator('#simple-research-steps button').nth(1).click()
        assert page.locator('[data-research-step="0"]').is_hidden()
        assert page.locator('[data-research-step="1"]').is_visible()
        page.locator('#nav-settings').click()
        assert page.locator('#settings-save').is_visible()
        page.locator('#simple-retrieval-settings > summary').click()
        assert page.locator('#settings-embedding-channel').is_visible()
        page.locator('#settings-close').click()
        assert page.locator('#page-title').inner_text()=='资料库'
        page.locator('#simple-session-list > summary').click()
        page.locator('#chat-search').fill('不存在的会话')
        assert '没有匹配' in page.locator('#simple-history-chat-list').inner_text()
        assert page.locator('#page-title').inner_text()=='历史记录'
        assert not [x for x in calls if x in {'ask','test_model_connection','confirm_image_generation'}]
    finally:
        fixture.close()


@pytest.mark.parametrize('revision',['before','after'])
def test_same_state_visual_evidence(tmp_path,monkeypatch,revision):
    base=STATIC.parents[2]/'deliverables/codex-workspace-20261103'
    if revision=='before':
        frozen=base/'before/static'
        for name in ['professional','range','workspace_layout','model_connections','daily','ui_simplify']:
            monkeypatch.setattr(importlib.import_module('tests.test_'+name+'_browser'),'STATIC',frozen)
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        save(api._connections,payload([connection()]))
        setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page); mode(page,'chat')
        page.set_viewport_size({'width':1160,'height':760})
        page.evaluate('document.documentElement.dataset.theme="dark"')
        folder=base/revision/'comparison'; folder.mkdir(parents=True,exist_ok=True)
        density={}
        def capture(name):
            page.screenshot(path=str(folder/(name+'.png')))
            density[name]=page.locator('button,input,textarea,select,summary').evaluate_all('nodes=>nodes.filter(n=>n.checkVisibility()&&n.getBoundingClientRect().top>=0&&n.getBoundingClientRect().bottom<=innerHeight).length')
        capture('chat-empty')
        mode(page,'agent'); capture('document-empty')
        mode(page,'chat')
        page.locator('#nav-history').click()
        page.locator('#nav-settings').click(); capture('settings')
        page.locator('#settings-close').click(); capture('history-after-settings')
        page.locator('#nav-resources').click(); capture('resources')
        page.locator('#workspace-tool-workbench-panel > summary').click(); capture('research')
        (folder/'controls.json').write_text(json.dumps({'revision':revision,'width':1160,'height':760,'theme':'dark','visible_controls':density,'scope':'isolated host fixtures; counts are density evidence, not a usability score'},ensure_ascii=False,indent=2),encoding='utf8')
        assert not [x for x in calls if x in {'ask','confirm_image_generation','confirm_professional_edit'}]
    finally:
        fixture.close()
