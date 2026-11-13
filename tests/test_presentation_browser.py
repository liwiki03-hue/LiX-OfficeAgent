from io import BytesIO
import json
import os
from pathlib import Path
import shutil
import time

from docx import Document
from pptx import Presentation
import pytest

from tests.test_paged_browser import paged_browser, engine_frame
from office_agent.presentation_drafts import validate_deck


def open_source(page):
    page.get_by_label('更多',exact=True).first.click()
    page.get_by_role('button',name='生成演示稿',exact=True).click()
    return page.get_by_role('dialog').filter(has=page.get_by_role('heading',name='新建演示稿',exact=True))


def wait_ready(page,api):
    deadline=time.monotonic()+45
    while api._presentations.task:
        if time.monotonic()>deadline:raise AssertionError('bounded presentation render exceeded deadline')
        page.wait_for_timeout(100)
    assert api._presentations.status=='ready',(api._presentations.error,api._presentations.diagnostic)
    page.locator('[data-presentation-image]').wait_for(state='visible',timeout=10000)


def test_current_docx_draft_real_preview_publish_cancel_and_rollback(paged_browser):
    page,api,source,errors=paged_browser;original=source.read_bytes();frame=engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('实验于2025年完成，温度120 ℃。',exact=True).first.click()
    page.keyboard.press('Home');page.keyboard.press('Shift+End');page.keyboard.insert_text('实验于2026年完成，温度120 ℃。')
    node=open_source(page)
    node.get_by_label('演示标题',exact=True).fill('技术报告演示')
    node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for()
    captured=api._presentations.source
    assert captured['kind']=='draft' and captured['draft_revision']>0 and any('2026' in row['text'] for row in captured['segments'])
    assert len(api._presentations.spec['slides'])==8
    work.get_by_role('button',name='2. 技术报告',exact=True).click()
    work.get_by_label('页面版式',exact=True).select_option('section')
    work.get_by_label('要点（每行一条）',exact=True).fill('')
    work.get_by_role('button',name='保存本页到大纲',exact=True).click()
    page.wait_for_function("document.querySelector('[data-presentation-title]').textContent.includes('草稿 2')")
    work.locator('[data-presentation-slides] button').nth(2).click()
    work.get_by_label('页面版式',exact=True).select_option('columns')
    work.get_by_label('页面标题',exact=True).fill('实验条件与边界')
    work.get_by_label('右栏要点',exact=True).fill('仅为本次实验结果，不外推至其他温度。')
    work.get_by_role('button',name='保存本页到大纲',exact=True).click()
    page.wait_for_function("document.querySelector('[data-presentation-title]').textContent.includes('草稿 3')")
    assert {row['layout'] for row in api._presentations.spec['slides']}=={'cover','section','bullets','columns','table'}
    work.get_by_role('button',name='生成实际预览',exact=True).click();wait_ready(page,api)
    assert all(row['expected_blocks']==row['matched_blocks'] for row in api._presentations.rendered['report']['text_checks'])
    page.screenshot(path=str(source.parent/'presentation-preview.png'))
    api._app.window=type('FolderWindow',(),{'create_file_dialog':lambda self,*args,**kwargs:[str(source.parent)]})()
    work.get_by_role('button',name='发布PPTX',exact=True).click()
    confirmation=page.get_by_role('dialog').filter(has=page.get_by_role('heading',name='确认发布演示稿',exact=True))
    confirmation.get_by_role('button',name='取消',exact=True).click()
    confirmation.wait_for(state='detached')
    assert not list(source.parent.glob('.office-presentation-*'))
    work.get_by_role('button',name='发布PPTX',exact=True).click()
    confirmation.get_by_role('button',name='确认发布PPTX',exact=True).click()
    confirmation.wait_for(state='detached',timeout=20000)
    work.get_by_role('strong').filter(has_text='已发布 presentation.pptx').wait_for()
    saved=next(iter(api._presentations.saved.values()));path=Path(saved['path'])
    assert path.is_file() and validate_deck(path.read_bytes(),api._presentations.spec)['slides']==8
    deck=Presentation(path)
    assert any(shape.has_table for slide in deck.slides for shape in slide.shapes)
    assert any('2026' in shape.text for slide in deck.slides for shape in slide.shapes if shape.has_text_frame)
    assert all(slide.has_notes_slide for slide in deck.slides) and source.read_bytes()==original
    page.screenshot(path=str(source.parent/'presentation-published.png'))
    evidence=os.environ.get('OFFICE_PRESENTATION_EVIDENCE')
    if evidence:
        folder=Path(evidence);folder.mkdir(parents=True,exist_ok=True)
        for name in ['presentation-preview.png','presentation-published.png']:shutil.copyfile(source.parent/name,folder/name)
        shutil.copyfile(source,folder/'自建源稿.docx');shutil.copyfile(path,folder/'技术报告演示.pptx')
        shutil.copyfile(path.parent/'backup',folder/'SOURCE_SPEC_BACKUP.json')
        shutil.copyfile(path.parent/'audit-verified.json',folder/'PUBLISH_AUDIT.json')
        for index,image in enumerate(api._presentations.rendered['previews']):(folder/f'slide-{index+1}.png').write_bytes(image)
        (folder/'SPEC.json').write_text(json.dumps(api._presentations.spec,ensure_ascii=False,indent=2),encoding='utf8')
        (folder/'VERIFICATION.json').write_text(json.dumps({'source_unchanged':source.read_bytes()==original,'source':captured,'report':api._presentations.rendered['report'],'saved_path':str(path),'native':False,'cloud_calls':0},ensure_ascii=False,indent=2),encoding='utf8')
    work.get_by_role('button',name='撤销本次发布',exact=True).click()
    deadline=time.monotonic()+10
    while path.exists():
        if time.monotonic()>deadline:raise AssertionError(page.locator('[data-presentation-status]').text_content())
        page.wait_for_timeout(100)
    assert (path.parent/'audit-removed.json').is_file() and source.read_bytes()==original and not errors
    if evidence:shutil.copyfile(path.parent/'audit-removed.json',Path(evidence)/'ROLLBACK_AUDIT.json')


def test_outline_edit_responsive_focus_and_reopen_preserve_drafts(paged_browser):
    page,api,source,errors=paged_browser
    page.locator('#floating-input').fill('尚未发送的文档修改要求')
    node=open_source(page);node.get_by_label('材料来源',exact=True).select_option('text')
    node.get_by_label('页数（2—12）',exact=True).fill('3')
    node.get_by_label('材料（按段落换行）',exact=True).fill('研究目的：保留限定条件。\n实验结果：耗时120 ms。')
    node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for()
    first=api._presentations.spec['slides'][0]['id']
    matrix=[]
    for width in [1440,1280,1024,520]:
        page.set_viewport_size({'width':width,'height':800})
        for theme in ['light','dark','highcontrast']:
            page.evaluate('(theme)=>document.documentElement.dataset.theme=theme',theme)
            assert not work.evaluate('node=>node.scrollWidth>node.clientWidth+2')
            save=work.get_by_role('button',name='保存本页到大纲',exact=True);save.scroll_into_view_if_needed()
            assert save.is_visible() and save.bounding_box()['y']<800
            matrix.append({'width':width,'height':800,'theme':theme,'save_reachable':True,'no_dialog_horizontal_overflow':True})
            if os.environ.get('OFFICE_PRESENTATION_EVIDENCE'):
                folder=Path(os.environ['OFFICE_PRESENTATION_EVIDENCE']);folder.mkdir(parents=True,exist_ok=True)
                page.screenshot(path=str(folder/f'workspace-{width}-{theme}.png'))
    if os.environ.get('OFFICE_PRESENTATION_EVIDENCE'):(Path(os.environ['OFFICE_PRESENTATION_EVIDENCE'])/'RESPONSIVE.json').write_text(json.dumps(matrix,indent=2),encoding='utf8')
    work.get_by_label('页面标题',exact=True).fill('已编辑封面')
    assert work.get_by_role('button',name='发布PPTX',exact=True).is_disabled()
    work.get_by_role('button',name='保存本页到大纲',exact=True).click()
    page.wait_for_function("document.querySelector('[data-presentation-title]').textContent.includes('草稿 2')")
    assert api._presentations.spec['slides'][0]['title']=='已编辑封面' and api._presentations.spec['slides'][0]['id']==first
    work.get_by_role('button',name='新增一页',exact=True).click()
    page.wait_for_function("document.querySelector('[data-presentation-editor] [name=title]').value==='新增页面'")
    assert len(api._presentations.spec['slides'])==4
    work.get_by_role('button',name='关闭',exact=True).click();work.wait_for(state='detached')
    assert page.locator('#floating-input').input_value()=='尚未发送的文档修改要求'
    page.get_by_label('更多',exact=True).first.click();page.get_by_role('button',name='生成演示稿',exact=True).click()
    work.wait_for();assert len(api._presentations.spec['slides'])==4
    page.keyboard.press('Escape');work.wait_for(state='detached')
    assert not errors and not list(source.parent.glob('.office-presentation-*'))


def test_late_create_response_after_privacy_clear_cannot_restore_source_ui(paged_browser):
    page,api,source,errors=paged_browser
    node=open_source(page);node.get_by_label('材料来源',exact=True).select_option('text')
    node.get_by_label('页数（2—12）',exact=True).fill('2');node.get_by_label('材料（按段落换行）',exact=True).fill('仅用于迟到响应隔离。')
    page.evaluate("""()=>{const original=window.testHost;window.testHost=(name,args)=>name==='presentation_create'?new Promise(resolve=>{window.delayedPresentation=()=>original(name,args).then(resolve)}):original(name,args)}""")
    node.get_by_role('button',name='整理大纲',exact=True).click()
    page.wait_for_function("typeof window.delayedPresentation==='function'")
    page.evaluate("window.dispatchEvent(new Event('officeauth:unauthenticated'))")
    page.evaluate('window.delayedPresentation()');page.wait_for_timeout(200)
    assert page.locator('.presentation-workspace,.presentation-dialog').count()==0 and not errors


def test_uncertain_publish_result_can_close_without_repeating_confirmation(paged_browser,monkeypatch):
    from tests.test_presentation_workspace import fake_ready
    page,api,source,errors=paged_browser
    node=open_source(page);node.get_by_label('材料来源',exact=True).select_option('text')
    node.get_by_label('页数（2—12）',exact=True).fill('2');node.get_by_label('材料（按段落换行）',exact=True).fill('发布异常交互样例。')
    node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for()
    fake_ready(api._presentations,monkeypatch)
    page.wait_for_function("!document.querySelector('[data-presentation-export]').disabled")
    api._app.window=type('FolderWindow',(),{'create_file_dialog':lambda self,*args,**kwargs:[str(source.parent)]})()
    work.get_by_role('button',name='发布PPTX',exact=True).click()
    entry=next(iter(api._presentations.plans.values()));calls=[]
    def fail(*args):calls.append(1);raise OSError('synthetic failure')
    monkeypatch.setattr(entry['archive'],'_exclusive',fail)
    confirmation=page.get_by_role('dialog').filter(has=page.get_by_role('heading',name='确认发布演示稿',exact=True))
    confirmation.get_by_role('button',name='确认发布PPTX',exact=True).click()
    confirmation.get_by_text('结果不确定，请核对目标文件；不会自动重试',exact=True).wait_for()
    confirmation.get_by_role('button',name='关闭',exact=True).click();confirmation.wait_for(state='detached')
    assert calls==[1] and work.get_by_role('button',name='发布PPTX',exact=True).is_disabled()
    work.get_by_role('button',name='关闭',exact=True).click();work.wait_for(state='detached')
    assert not errors
