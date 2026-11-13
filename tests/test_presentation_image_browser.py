from copy import deepcopy
from io import BytesIO
import json
import os
from pathlib import Path
import shutil
import time

from PIL import Image,ImageDraw
from pptx import Presentation

from office_agent.presentation_drafts import validate_deck
from tests.test_presentation_chart_browser import chart_browser
from tests.test_presentation_browser import open_source,wait_ready


def fixture_images(folder):
    paths=[]
    for name,size,alpha,orientation in [('横向实验示意.png',(1000,600),False,None),('旋转竖向实验.jpg',(900,500),False,6),('透明实验标记.png',(500,700),True,None)]:
        logical=size[::-1] if orientation else size
        image=Image.new('RGBA' if alpha else 'RGB',logical,(0,0,0,0) if alpha else '#e8eef5');draw=ImageDraw.Draw(image)
        draw.rectangle((35,40,logical[0]-35,logical[1]-40),fill='#246384',outline='#162334',width=5)
        draw.rectangle((70,80,logical[0]//2,logical[1]//2),fill='#e6ba5c')
        draw.text((85,100),'SELF-CREATED / IMAGE CHECK',fill='#162334')
        draw.text((85,130),'TOP / 2026',fill='#162334')
        exif=Image.Exif()
        if orientation:
            rotated=image.transpose(Image.Transpose.ROTATE_90);image.close();image=rotated
            exif[274]=orientation;exif[315]='QA fixture'
        path=folder/name;image.save(path,**({'exif':exif} if orientation else {}));image.close();paths.append(path)
    return paths


def wait_revision(page,revision):
    page.wait_for_function('(revision)=>document.querySelector("[data-presentation-title]").textContent.endsWith("草稿 "+revision)',arg=revision)


def set_page(page,api,work,index,layout,**fields):
    work.locator('[data-presentation-slides] button').nth(index).click()
    work.get_by_label('页面版式',exact=True).select_option(layout)
    labels={'title':'页面标题','points':'要点（每行一条）','right':'右栏要点','rows':'表格（列用Tab分隔，行用换行分隔）','chart_unit':'图表单位'}
    for key,value in fields.items():work.get_by_label(labels[key],exact=True).fill(value)
    if layout=='chart':work.get_by_role('checkbox',name='我已核对本页分类、数值和单位',exact=True).check()
    revision=api._presentations.spec['revision']+1
    work.get_by_role('button',name='保存本页到大纲',exact=True).click();wait_revision(page,revision)


def choose_picture(page,api,work,path,adopt=True,variant='normalized'):
    api._app.window=type('ImageWindow',(),{'create_file_dialog':lambda self,*args,**kwargs:[str(path)] if path else None})()
    work.get_by_role('button',name='替换图片' if work.get_by_role('button',name='替换图片',exact=True).count() else '选择图片',exact=True).click()
    if path is None:
        page.wait_for_function('!document.querySelector("[data-presentation-build]").disabled');return
    dialog=page.get_by_role('dialog',name='核对本机图片',exact=True);dialog.wait_for()
    if not adopt:dialog.get_by_role('button',name='取消',exact=True).click()
    else:
        dialog.locator('input[value="'+variant+'"]').check()
        dialog.get_by_role('checkbox',name='我确认有权使用，并已核对所选副本的朝向、清晰度和透明度',exact=True).check()
        dialog.get_by_role('button',name='采用到本页草稿',exact=True).click()
    dialog.wait_for(state='detached')


def test_real_ten_page_images_chart_publish_cancel_and_rollback(chart_browser):
    page,api,source,errors=chart_browser;original=source.read_bytes();paths=fixture_images(source.parent);original_images=[path.read_bytes() for path in paths]
    node=open_source(page);node.get_by_label('页数（2—12）',exact=True).fill('10');node.get_by_label('演示标题',exact=True).fill('实验材料与数据说明')
    node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for()
    evidence=os.environ.get('OFFICE_PRESENTATION_IMAGE_EVIDENCE');folder=Path(evidence) if evidence else None
    if folder:folder.mkdir(parents=True,exist_ok=True);page.screenshot(path=str(folder/'before-images.png'))
    for index,path in enumerate(paths,1):
        set_page(page,api,work,index,'image',title=['实验材料横图','竖向材料朝向','透明素材说明'][index-1],points='自建验收素材，不代表真实实验照片。\n完整等比显示，原文件保持。')
        before=deepcopy(api._presentations.spec)
        if index==1:
            choose_picture(page,api,work,None);assert api._presentations.spec==before
            choose_picture(page,api,work,path,False);assert api._presentations.spec==before
        choose_picture(page,api,work,path)
        assert api._presentations.spec['slides'][index]['asset_id'] and len(api._presentations.assets)==index
    set_page(page,api,work,4,'image',title='同一素材跨页复用',points='复用当前已授权图片快照，不扩大文件权限。')
    first=api._presentations.spec['slides'][1]['asset_id'];work.get_by_label('复用本演示稿授权图片',exact=True).select_option(first)
    revision=api._presentations.spec['revision']+1;work.get_by_role('button',name='用于本页',exact=True).click();wait_revision(page,revision)
    set_page(page,api,work,5,'section',title='实验数据与边界',points='以下数据仅用于自建功能验收。')
    set_page(page,api,work,6,'columns',title='条件和适用边界',points='温度120 ℃\n相同样本与测量方式',right='不外推到其他温度。\n图中小字未自动认定为证据。')
    set_page(page,api,work,7,'chart',title='耗时对照',rows='方案\t耗时\n基线\t120\n优化\t105.5\n复核\t98.25',chart_unit='ms')
    set_page(page,api,work,8,'table',title='原始数据表',rows='方案\t耗时(ms)\n基线\t120\n优化\t105.5\n复核\t98.25')
    set_page(page,api,work,9,'bullets',title='结论与未验范围',points='本次仅验证图文和数据发布链。\nPowerPoint人工质量另行核对。')
    work.locator('[data-presentation-slides] button').nth(1).click()
    work.get_by_role('button',name='生成实际预览',exact=True).click();wait_ready(page,api)
    approved=api._presentations.fresh_assets();spec=deepcopy(api._presentations.spec);rendered=api._presentations.rendered
    assert rendered['report']['editable_pictures']==4 and rendered['report']['editable_charts']==1 and len(approved)==3
    assert len(rendered['previews'])==10
    if folder:page.screenshot(path=str(folder/'image-preview.png'))
    api._app.window=type('FolderWindow',(),{'create_file_dialog':lambda self,*args,**kwargs:[str(source.parent)]})()
    confirmation=page.get_by_role('dialog',name='确认发布演示稿',exact=True)
    work.get_by_role('button',name='发布PPTX',exact=True).click();confirmation.get_by_role('button',name='取消',exact=True).click();confirmation.wait_for(state='detached')
    assert not list(source.parent.glob('.office-presentation-*'))
    work.get_by_role('button',name='发布PPTX',exact=True).click();confirmation.get_by_role('button',name='确认发布PPTX',exact=True).click();confirmation.wait_for(state='detached',timeout=20000)
    work.get_by_role('strong').filter(has_text='已发布 presentation.pptx').wait_for()
    saved=next(iter(api._presentations.saved.values()));path=Path(saved['path']);raw=path.read_bytes();report=validate_deck(raw,spec,approved)
    assert report['editable_pictures']==4 and source.read_bytes()==original and [path.read_bytes() for path in paths]==original_images
    book=Presentation(BytesIO(raw));assert sum(shape.shape_type==13 for slide in book.slides for shape in slide.shapes)==4
    backup=json.loads((path.parent/'backup').read_text('utf8'));assert len(backup['snapshot']['assets'])==3 and 'base64' not in json.dumps(backup)
    if folder:
        page.screenshot(path=str(folder/'published.png'))
        for asset in paths:shutil.copyfile(asset,folder/asset.name)
        for item,name in [(source,'自建实验原稿.docx'),(path,'实验材料与数据说明.pptx'),(path.parent/'backup','SOURCE_SPEC_BACKUP.json'),(path.parent/'audit-verified.json','PUBLISH_AUDIT.json')]:shutil.copyfile(item,folder/name)
        for index,data in enumerate(rendered['previews']):(folder/f'slide-{index+1}.png').write_bytes(data)
        (folder/'SPEC.json').write_text(json.dumps(spec,ensure_ascii=False,indent=2),encoding='utf8')
        (folder/'VERIFICATION.json').write_text(json.dumps({'report':rendered['report'],'source_unchanged':True,'images_unchanged':True,'publish_cancel_created_nothing':True,'native':False,'cloud_calls':0,'asset_processing':backup['snapshot']['asset_processing'],'raw_bytes':len(raw)},ensure_ascii=False,indent=2),encoding='utf8')
    work.get_by_role('button',name='撤销本次发布',exact=True).click()
    deadline=time.monotonic()+10
    while path.exists():
        if time.monotonic()>deadline:raise AssertionError('rollback deadline')
        page.wait_for_timeout(100)
    if folder:shutil.copyfile(path.parent/'audit-removed.json',folder/'ROLLBACK_AUDIT.json')
    assert source.read_bytes()==original and [path.read_bytes() for path in paths]==original_images and not errors


def test_image_form_responsive_oversized_points_and_late_private_response(chart_browser):
    page,api,source,errors=chart_browser;paths=fixture_images(source.parent)
    node=open_source(page);node.get_by_label('材料来源',exact=True).select_option('text');node.get_by_label('页数（2—12）',exact=True).fill('2');node.get_by_label('材料（按段落换行）',exact=True).fill('本机图片接入与清晰度核对。');node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for();work.locator('[data-presentation-slides] button').nth(1).click()
    work.get_by_label('页面版式',exact=True).select_option('image');points=work.get_by_label('要点（每行一条）',exact=True);points.fill('一\n二\n三\n四')
    revision=api._presentations.spec['revision'];work.get_by_role('button',name='保存本页到大纲',exact=True).click();page.locator('[data-presentation-status]').filter(has_text='校验未通过').wait_for()
    assert points.input_value()=='一\n二\n三\n四' and api._presentations.spec['revision']==revision
    points.fill('保留图文说明');work.get_by_role('button',name='保存本页到大纲',exact=True).click();wait_revision(page,revision+1)
    choose_picture(page,api,work,paths[0]);matrix=[];folder=Path(os.environ['OFFICE_PRESENTATION_IMAGE_EVIDENCE']) if os.environ.get('OFFICE_PRESENTATION_IMAGE_EVIDENCE') else None
    for width in [1440,1280,1024,520]:
        page.set_viewport_size({'width':width,'height':800})
        for theme in ['light','dark','highcontrast']:
            page.evaluate('(theme)=>document.documentElement.dataset.theme=theme',theme)
            choose=work.get_by_role('button',name='替换图片',exact=True);choose.scroll_into_view_if_needed()
            assert choose.is_visible() and not work.evaluate('node=>node.scrollWidth>node.clientWidth+2')
            matrix.append({'width':width,'theme':theme,'image_controls_reachable':True})
            if folder:page.screenshot(path=str(folder/f'images-{width}-{theme}.png'))
    if folder:(folder/'RESPONSIVE.json').write_text(json.dumps(matrix,indent=2),encoding='utf8')
    page.evaluate("""()=>{const host=window.testHost;window.testHost=(name,args)=>name==='presentation_image_choose'?new Promise((resolve,reject)=>{window.releaseImage=()=>Promise.resolve(host(name,args)).then(resolve,reject)}):host(name,args)}""")
    api._app.window=type('ImageWindow',(),{'create_file_dialog':lambda self,*args,**kwargs:[str(paths[1])]})()
    work.get_by_role('button',name='替换图片',exact=True).click();page.wait_for_function("typeof window.releaseImage==='function'")
    api._presentations.invalidate();page.evaluate("window.dispatchEvent(new Event('officeauth:unauthenticated'))");page.evaluate('window.releaseImage()');page.wait_for_timeout(200)
    assert not api._presentations.assets and page.locator('.presentation-dialog,.presentation-workspace').count()==0 and not errors


def test_repeated_pick_and_late_candidate_preserve_new_text_without_duplicate_requests(chart_browser):
    page,api,source,errors=chart_browser;paths=fixture_images(source.parent)
    node=open_source(page);node.get_by_label('材料来源',exact=True).select_option('text');node.get_by_label('页数（2—12）',exact=True).fill('2');node.get_by_label('材料（按段落换行）',exact=True).fill('实验材料图文说明。');node.get_by_role('button',name='整理大纲',exact=True).click()
    work=page.get_by_role('dialog',name='演示稿工作区',exact=True);work.wait_for();set_page(page,api,work,1,'image',points='原说明')
    api._app.window=type('ImageWindow',(),{'create_file_dialog':lambda self,*args,**kwargs:[str(paths[0])]})()
    page.evaluate("""()=>{const host=window.testHost;window.imageCalls=[];window.testHost=(name,args)=>{if(name.startsWith('presentation_image_'))window.imageCalls.push(name);return name==='presentation_image_choose'?host(name,args).then(value=>new Promise(resolve=>{window.releaseCandidate=()=>resolve(value)})):host(name,args)}}""")
    work.get_by_role('button',name='选择图片',exact=True).dblclick();page.wait_for_function("typeof window.releaseCandidate==='function'")
    points=work.get_by_label('要点（每行一条）',exact=True);points.fill('读取期间追加的文字，必须保留。')
    page.evaluate('window.releaseCandidate()');dialog=page.get_by_role('dialog',name='核对本机图片',exact=True);dialog.wait_for()
    dialog.get_by_role('checkbox',name='我确认有权使用，并已核对所选副本的朝向、清晰度和透明度',exact=True).check()
    dialog.get_by_role('button',name='采用到本页草稿',exact=True).dblclick();dialog.wait_for(state='detached')
    assert points.input_value()=='读取期间追加的文字，必须保留。' and api._presentations.spec['slides'][1]['points']==['原说明']
    assert page.evaluate('window.imageCalls')==['presentation_image_choose','presentation_image_use']
    revision=api._presentations.spec['revision']+1;work.get_by_role('button',name='保存本页到大纲',exact=True).click();wait_revision(page,revision)
    assert api._presentations.spec['slides'][1]['points']==['读取期间追加的文字，必须保留。'] and not errors and not list(source.parent.glob('.office-presentation-*'))
