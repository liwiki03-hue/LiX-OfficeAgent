from copy import deepcopy
from io import BytesIO
import os
from pathlib import Path
import random
import sys
import threading
import time
import zipfile

from PIL import Image
from pptx import Presentation
from pydantic import ValidationError
import pytest

from office_agent import presentation_assets as assets
from office_agent.presentation_drafts import PresentationError,Slide,build_deck,validate_deck
from tests.test_presentation_workspace import workspace,fake_ready
from tests.test_presentation_charts import rewrite
from office_agent.diagram_process import _run_owned


def picture(path,size=(640,360),format='PNG',alpha=False,orientation=None):
    image=Image.new('RGBA' if alpha else 'RGB',size,(24,80,136,100) if alpha else (24,80,136))
    kwargs={}
    if orientation:
        exif=Image.Exif();exif[274]=orientation;exif[315]='self-created-fixture';kwargs['exif']=exif
    try:image.save(path,format=format,**kwargs)
    finally:image.close()
    return path


def image_page(service,index=1):
    slide=service.spec['slides'][index]
    service.edit(service.spec['id'],service.spec['revision'],'update',slide['id'],{'layout':'image','points':['自建图片说明'],'rows':[],'right':[]})
    return slide['id']


def attach(service,path,slide_id,variant='normalized'):
    chosen=service.choose_image(service.spec['id'],service.spec['revision'],slide_id,lambda:path)
    return service.use_image(service.spec['id'],service.spec['revision'],slide_id,chosen['ticket'],variant,True)


@pytest.mark.parametrize('size,format,alpha,orientation',[
    ((640,360),'PNG',False,None),((300,600),'PNG',True,None),((600,300),'JPEG',False,6)
])
def test_real_worker_normalizes_exif_and_metadata_and_preserves_transparency(tmp_path,size,format,alpha,orientation):
    path=picture(tmp_path/('fixture.png' if format=='PNG' else 'fixture.jpg'),size,format,alpha,orientation)
    original=path.read_bytes();result=assets.prepare_image(path)
    assert path.read_bytes()==original and result['process']['empty_confirmed'] and result['process']['memory_limit_bytes']==512*1024*1024
    value=result['variants']['normalized'];expected=size[::-1] if orientation==6 else size
    assert (value['width'],value['height'])==expected
    with Image.open(BytesIO(value['raw'])) as image:
        assert not image.getexif() and 'exif' not in image.info
        if alpha:assert image.mode=='RGBA' and image.getpixel((0,0))[3]==100
    assert result['original']['width']==size[0] and result['thumbnail'].startswith('data:image/png;base64,')


@pytest.mark.parametrize('kind',['wrong_extension','multiframe','corrupt','large'])
def test_reject_unsupported_or_misidentified_images_without_changing_source(tmp_path,kind):
    path=tmp_path/'image.png'
    if kind=='wrong_extension':picture(path,format='JPEG')
    elif kind=='multiframe':
        one=Image.new('RGB',(20,20),'red');two=Image.new('RGB',(20,20),'blue')
        one.save(path,format='PNG',save_all=True,append_images=[two]);one.close();two.close()
    elif kind=='corrupt':path.write_bytes(b'not an image')
    else:path.write_bytes(b'x'*(assets.MAX_INPUT+1))
    original=path.read_bytes()
    with pytest.raises(PresentationError):assets.prepare_image(path)
    assert path.read_bytes()==original


def test_hardlink_and_modified_source_are_rejected(tmp_path):
    path=picture(tmp_path/'image.png');_,stamp=assets.read_source(path)
    path.write_bytes(path.read_bytes()+b'changed')
    with pytest.raises(PresentationError):assets.check_source(path,stamp)
    other=tmp_path/'hardlink.png';os.link(path,other)
    with pytest.raises(PresentationError):assets.read_source(other)


@pytest.mark.parametrize('path',[r'\\untrusted.invalid\share\image.png',r'Z:\image.png'])
def test_remote_image_paths_are_rejected_before_any_filesystem_read(monkeypatch,path):
    import win32file
    touched=[]
    monkeypatch.setattr(win32file,'GetDriveType',lambda drive:4)
    monkeypatch.setattr(assets,'_audit_path',lambda *args,**kwargs:touched.append('filesystem'))
    with pytest.raises(PresentationError,match='网络路径不支持'):assets.read_source(path)
    assert touched==[]


def test_image_worker_job_applies_real_512_mib_limit_before_spawn():
    job=assets.ImageJob()
    try:
        info=job.job.QueryInformationJobObject(job.handle,job.job.JobObjectExtendedLimitInformation)
        assert info['JobMemoryLimit']==assets.MEMORY_LIMIT and info['BasicLimitInformation']['ActiveProcessLimit']==2
        assert info['BasicLimitInformation']['LimitFlags']&job.job.JOB_OBJECT_LIMIT_JOB_MEMORY
    finally:assert job.close()


def test_image_worker_deadline_terminates_only_owned_job():
    result=_run_owned((sys.executable,'-c','import time; time.sleep(10)'),timeout_ms=150,_job_factory=assets.ImageJob)
    assert result.code=='render_timeout' and result.empty_confirmed and result.elapsed_seconds<3


def test_asset_count_and_deduplicated_media_budget_remain_bounded():
    values={}
    for index in range(4):
        aid=f'{index+1:032x}'
        values[aid]=assets.Asset(aid,bytes([index])*90000,400,300,'PNG','a'*64,'b'*64)
    ids=list(values)
    spec={'slides':[{'layout':'image','asset_id':aid} for aid in ids]}
    with pytest.raises(PresentationError,match='3个'):assets.checked_assets(spec,values)
    spec['slides']=spec['slides'][:3]
    with pytest.raises(PresentationError,match='256 KiB'):assets.checked_assets(spec,values)
    same={key:assets.Asset(key,b'x'*90000,400,300,'PNG','a'*64,'b'*64) for key in ids[:3]}
    assert len(assets.checked_assets(spec,same))==3


def test_equal_bytes_in_another_file_do_not_inherit_original_grant(workspace,tmp_path):
    service,api=workspace;sid=image_page(service);path=picture(tmp_path/'one.png');attach(service,path,sid)
    original=service.spec['slides'][1]['asset_id'];other=tmp_path/'two.png';other.write_bytes(path.read_bytes())
    second=image_page(service,2);attach(service,other,second)
    assert len(service.assets)==2 and service.spec['slides'][2]['asset_id']!=original
    assert service.assets[original]['fingerprint']!=service.assets[service.spec['slides'][2]['asset_id']]['fingerprint']


def test_dimension_and_processor_changes_fail_closed(tmp_path,workspace,monkeypatch):
    too_wide=picture(tmp_path/'wide.png',(33000,1))
    with pytest.raises(PresentationError):assets.prepare_image(too_wide)
    service,api=workspace;sid=image_page(service);path=picture(tmp_path/'source.png');attach(service,path,sid)
    monkeypatch.setattr('office_agent.webui.presentation_workspace.processor_identity',lambda:{'changed':True})
    with pytest.raises(PresentationError,match='处理器'):service.fresh_assets()


def test_high_entropy_photo_is_previewable_but_cannot_silently_exceed_media_budget(tmp_path):
    path=tmp_path/'noise.jpg';image=Image.frombytes('RGB',(1800,1200),random.Random(20261108).randbytes(1800*1200*3));image.save(path,quality=95);image.close()
    original=path.read_bytes();candidate=assets.prepare_image(path)
    assert candidate['variants']['optimized']['width']==1600 and candidate['variants']['optimized']['bytes']>assets.MAX_ASSET
    with pytest.raises(PresentationError,match='192 KiB'):assets.approved_asset(candidate,'optimized')
    assert path.read_bytes()==original and all(value['preview'].startswith('data:image/png;base64,') for value in candidate['variants'].values())


def test_image_budget_no_silent_downscale_and_explicit_optimized_adoption(workspace,tmp_path):
    service,api=workspace;sid=image_page(service)
    path=picture(tmp_path/'wide.png',(2400,800));chosen=service.choose_image(service.spec['id'],service.spec['revision'],sid,lambda:path)
    assert chosen['variants']['normalized']['width']==2400 and chosen['variants']['optimized']['width']==1600
    before=deepcopy(service.spec)
    with pytest.raises(PresentationError,match='1600px'):service.use_image(service.spec['id'],service.spec['revision'],sid,chosen['ticket'],'normalized',True)
    assert service.spec==before
    attach(service,path,sid,'optimized');assert next(iter(service.assets.values()))['asset'].width==1600


def test_pick_and_adopt_cancel_do_not_change_draft_or_write(workspace,tmp_path):
    service,api=workspace;sid=image_page(service);before=deepcopy(service.spec)
    assert service.choose_image(service.spec['id'],service.spec['revision'],sid,lambda:None)['cancelled']
    path=picture(tmp_path/'source.png');chosen=service.choose_image(service.spec['id'],service.spec['revision'],sid,lambda:path)
    assert service.use_image(service.spec['id'],service.spec['revision'],sid,chosen['ticket'],'',False)['cancelled']
    assert service.spec==before and not service.assets and not list(tmp_path.glob('.office-presentation-*'))
    with pytest.raises(PresentationError):service.use_image(service.spec['id'],service.spec['revision'],sid,chosen['ticket'],'normalized',True)


def test_asset_authorization_reuse_order_and_removal_are_host_bound(workspace,tmp_path):
    service,api=workspace;sid=image_page(service);path=picture(tmp_path/'source.png');original=path.read_bytes();attach(service,path,sid)
    aid=service.spec['slides'][1]['asset_id'];second=image_page(service,2)
    with pytest.raises(PresentationError):service.edit(service.spec['id'],service.spec['revision'],'image_reference',second,{'asset_id':'a'*32})
    service.edit(service.spec['id'],service.spec['revision'],'image_reference',second,{'asset_id':aid})
    assert len(service.assets)==1
    raw=build_deck(service.spec,service.fresh_assets());report=validate_deck(raw,service.spec,service.fresh_assets())
    assert report['editable_pictures']==2 and len(report['assets'])==1
    with zipfile.ZipFile(BytesIO(raw)) as package:assert len([name for name in package.namelist() if name.startswith('ppt/media/')])==1
    service.edit(service.spec['id'],service.spec['revision'],'image_reference',sid,{'asset_id':''})
    assert len(service.assets)==1
    service.edit(service.spec['id'],service.spec['revision'],'image_reference',second,{'asset_id':''})
    assert not service.assets and path.read_bytes()==original


def test_expired_wrong_page_and_changed_draft_candidates_cannot_be_adopted(workspace,tmp_path):
    service,api=workspace;sid=image_page(service);path=picture(tmp_path/'source.png')
    for condition in ['page','expiry','revision']:
        chosen=service.choose_image(service.spec['id'],service.spec['revision'],sid,lambda:path)
        revision=service.spec['revision'];target=sid
        if condition=='page':target=service.spec['slides'][0]['id']
        elif condition=='expiry':service.image_pending['created']-=301
        else:service.edit(service.spec['id'],revision,'update',sid,{'title':'新标题'})
        with pytest.raises(PresentationError):service.use_image(service.spec['id'],revision,target,chosen['ticket'],'normalized',True)
        assert not service.assets


@pytest.mark.parametrize('privacy',[False,True])
def test_late_image_worker_cannot_attach_or_overwrite_new_state(workspace,tmp_path,monkeypatch,privacy):
    service,api=workspace;sid=image_page(service);path=picture(tmp_path/'source.png');candidate=assets.prepare_image(path)
    entered,release=threading.Event(),threading.Event();failures=[]
    def prepare(path):entered.set();assert release.wait(4);return candidate
    monkeypatch.setattr('office_agent.webui.presentation_workspace.prepare_image',prepare)
    def choose():
        try:service.choose_image(service.spec['id'],service.spec['revision'],sid,lambda:path)
        except PresentationError as error:failures.append(str(error))
    thread=threading.Thread(target=choose);thread.start();assert entered.wait(2)
    try:
        with pytest.raises(PresentationError,match='重复'):service.choose_image(service.spec['id'],service.spec['revision'],sid,lambda:path)
        if privacy:service.invalidate()
        else:service.edit(service.spec['id'],service.spec['revision'],'update',sid,{'title':'保留的新标题'})
        service.error='保留新错误'
    finally:release.set();thread.join(5)
    assert not thread.is_alive() and failures and not service.assets and service.image_pending is None and service.error=='保留新错误'


def test_image_data_fields_cannot_be_smuggled_through_generic_edit(workspace):
    service,api=workspace;sid=image_page(service)
    assert not api.presentation_edit(service.spec['id'],service.spec['revision'],'update',sid,{'asset_id':'a'*32})['ok']
    row=deepcopy(service.spec['slides'][1]);row['points']=['保留']*4
    with pytest.raises(ValidationError):Slide.model_validate(row)
    with pytest.raises(PresentationError):build_deck(service.spec)


@pytest.mark.parametrize('kind',['bytes','orphan','crop','external','geometry','extra_link','flip','effect'])
def test_image_package_tampering_is_rejected(workspace,tmp_path,kind):
    service,api=workspace;sid=image_page(service);attach(service,picture(tmp_path/'source.png'),sid)
    approved=service.fresh_assets();raw=build_deck(service.spec,approved)
    with zipfile.ZipFile(BytesIO(raw)) as archive:
        media=next(name for name in archive.namelist() if name.startswith('ppt/media/'))
        rel='ppt/slides/_rels/slide2.xml.rels';xml=archive.read(rel)
        if kind=='bytes':tampered=rewrite(raw,{media:b'not approved'})
        elif kind=='orphan':tampered=rewrite(raw,{'ppt/media/orphan.png':b'orphan'})
        elif kind=='external':tampered=rewrite(raw,{rel:xml.replace(b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"',b'TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"')})
        elif kind=='extra_link':tampered=rewrite(raw,{rel:xml.replace(b'</Relationships>',b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/'+Path(media).name.encode()+b'"/></Relationships>')})
        else:
            deck=Presentation(BytesIO(raw));shape=next(shape for shape in deck.slides[1].shapes if shape.name=='image-'+sid)
            if kind=='crop':shape.crop_left=.1
            elif kind=='flip':shape._element.spPr.xfrm.set('flipH','1')
            elif kind=='effect':
                from lxml import etree
                etree.SubElement(shape._element.blipFill.blip,'{http://schemas.openxmlformats.org/drawingml/2006/main}grayscl')
            else:shape.width+=100
            output=BytesIO();deck.save(output);tampered=output.getvalue()
    with pytest.raises((PresentationError,ValueError,OSError)):validate_deck(tampered,service.spec,approved)


def test_source_change_rejects_old_plan_even_when_old_docx_allowed(workspace,tmp_path,monkeypatch):
    service,api=workspace;sid=image_page(service);path=picture(tmp_path/'source.png');attach(service,path,sid)
    monkeypatch.setattr('tests.test_presentation_workspace.build_deck',lambda spec:build_deck(spec,service.fresh_assets()))
    fake_ready(service,monkeypatch)
    plan=service.plan(service.spec['id'],service.spec['revision'],True,lambda:tmp_path)['plan']
    path.write_bytes(path.read_bytes()+b'replaced')
    with pytest.raises(PresentationError):service.confirm(plan['id'],True)
    assert not Path(plan['path']).exists()
