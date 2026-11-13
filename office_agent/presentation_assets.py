import base64
from copy import deepcopy
from dataclasses import asdict, dataclass
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import re
import stat
import sys
import tempfile

from .diagram_process import _WindowsJob, _run_owned
from .imagetransactions import _stamp
from .presentation_drafts import PresentationError, identifier, identity
from .rag_evidence import _dir
from .writes import Workspace, _audit_path


MAX_INPUT=10_000_000
MAX_ASSETS=3
MAX_ASSET=192*1024
MAX_MEDIA=256*1024
MEMORY_LIMIT=512*1024*1024
TIMEOUT_MS=15000


def processor_identity():
    import PIL
    from importlib.metadata import version
    package=Path(PIL.__file__).parent
    files=[Path(__file__),Path(__file__).with_name('presentation_image_worker.py'),*sorted(package.glob('*.py')),*sorted(package.glob('_imaging*.pyd'))]
    return {'version':'local-image-v1','pillow':version('Pillow'),'sha256':identity([hashlib.sha256(file.read_bytes()).hexdigest() for file in files])}


class ImageJob(_WindowsJob):
    def __init__(self):
        super().__init__()
        try:
            info=self.job.QueryInformationJobObject(self.handle,self.job.JobObjectExtendedLimitInformation)
            info['JobMemoryLimit']=MEMORY_LIMIT
            info['BasicLimitInformation']['ActiveProcessLimit']=2
            self.job.SetInformationJobObject(self.handle,self.job.JobObjectExtendedLimitInformation,info)
        except BaseException:
            self.handle.Close()
            raise

    def close(self):
        try:
            info=self.job.QueryInformationJobObject(self.handle,self.job.JobObjectExtendedLimitInformation)
            self.peak_job_memory=info['PeakJobMemoryUsed']
        finally:closed=super().close()
        return closed


def read_source(path):
    try:
        path=Path(path)
        if not path.is_absolute() or path.suffix.lower() not in {'.png','.jpg','.jpeg'}:raise ValueError()
        if os.name=='nt':
            import win32file
            if not re.fullmatch(r'[A-Za-z]:',path.drive) or win32file.GetDriveType(path.anchor) not in {2,3,6}:raise ValueError()
        _audit_path(path.parent,directory=True);_audit_path(path,directory=False)
        if Workspace(path.parent).file(str(path))!=path:raise ValueError()
        parent=_dir(path.parent);before=path.lstat()
        if not stat.S_ISREG(before.st_mode) or before.st_nlink!=1 or not 0<before.st_size<=MAX_INPUT:raise ValueError()
        with path.open('rb') as stream:
            first=os.fstat(stream.fileno())
            if _stamp(first)[:-1]!=_stamp(before)[:-1] or first.st_nlink!=1:raise ValueError()
            data=stream.read(MAX_INPUT+1);last=os.fstat(stream.fileno())
            if _stamp(last)!=_stamp(first) or last.st_nlink!=1:raise ValueError()
        after=path.lstat()
        if _stamp(after)!=_stamp(before) or after.st_nlink!=1 or _dir(path.parent)!=parent or len(data)!=before.st_size:raise ValueError()
        return data,{'stamp':_stamp(after),'parent':parent,'sha256':hashlib.sha256(data).hexdigest()}
    except (ValueError,OSError,TypeError):raise PresentationError('图片路径、身份或容量不符，请选择本机磁盘静态PNG/JPEG；网络路径不支持，不会读取其他文件') from None


def check_source(path,fingerprint):
    if read_source(path)[1]!=fingerprint:raise PresentationError('图片原文件已改变，请重新关联；旧预览和批准不可发布')


def prepare_image(path):
    from .fidelity import _temporary_root
    data,fingerprint=read_source(path);processor=processor_identity()
    try:
        with tempfile.TemporaryDirectory(prefix='oai-',dir=_temporary_root()) as directory:
            root=Path(directory);(root/'input').write_bytes(data)
            jobs=[]
            def factory():
                job=ImageJob();jobs.append(job);return job
            run=_run_owned((sys.executable,str(Path(__file__).with_name('presentation_image_worker.py')),str(root)),timeout_ms=TIMEOUT_MS,_job_factory=factory)
            if run.code!='process_completed' or run.exit_code!=0 or not run.empty_confirmed:raise ValueError()
            report=json.loads((root/'report.json').read_text('utf8'))
            expected='PNG' if Path(path).suffix.lower()=='.png' else 'JPEG'
            if report['original']['format']!=expected:raise ValueError()
            variants={}
            for key,details in report['variants'].items():
                if key not in {'normalized','optimized'}:raise ValueError()
                blob=(root/key).read_bytes()
                if len(blob)>MAX_INPUT or len(blob)!=details['bytes'] or hashlib.sha256(blob).hexdigest()!=details['sha256']:raise ValueError()
                preview=(root/(key+'-thumbnail')).read_bytes()
                if len(preview)>MAX_ASSET:raise ValueError()
                variants[key]={'raw':blob,**details,'preview':'data:image/png;base64,'+base64.b64encode(preview).decode('ascii')}
            thumbnail=(root/'thumbnail').read_bytes()
            if len(thumbnail)>MAX_ASSET:raise ValueError()
    except Exception:raise PresentationError('图片解码未完成：仅支持静态PNG/JPEG；请检查格式或改用更小图片（15秒/512 MiB）') from None
    check_source(path,fingerprint)
    if processor_identity()!=processor:raise PresentationError('图片处理器已改变，请重新选择')
    return {'path':Path(path),'fingerprint':fingerprint,'original':report['original'],'processor':processor,'variants':variants,
            'thumbnail':'data:image/png;base64,'+base64.b64encode(thumbnail).decode('ascii'),
            'process':{**asdict(run),'memory_limit_bytes':MEMORY_LIMIT,'peak_job_memory_bytes':jobs[0].peak_job_memory,'timeout_ms':TIMEOUT_MS,'active_process_limit':2}}


@dataclass(frozen=True)
class Asset:
    id: str
    raw: bytes
    width: int
    height: int
    format: str
    processor_sha256: str
    source_sha256: str
    source_stamp: tuple = ()
    source_parent: tuple = ()

    def manifest(self):
        return {'id':self.id,'sha256':hashlib.sha256(self.raw).hexdigest(),'bytes':len(self.raw),'width':self.width,'height':self.height,'format':self.format,
                'processor_sha256':self.processor_sha256,'source_sha256':self.source_sha256,'source_identity':{'stamp':list(self.source_stamp),'parent':list(self.source_parent)},'provenance':'用户选择的本机素材，许可未独立核验'}


def approved_asset(candidate,variant):
    if variant not in candidate['variants']:raise PresentationError('素材候选已失效，请重新选择')
    value=candidate['variants'][variant]
    if len(value['raw'])>MAX_ASSET or max(value['width'],value['height'])>1600:raise PresentationError('此副本超过192 KiB或1600px预算，请核对优化候选或选择更小图片')
    return Asset(identifier(),value['raw'],value['width'],value['height'],value['format'],identity(candidate['processor']),candidate['fingerprint']['sha256'],tuple(candidate['fingerprint']['stamp']),tuple(candidate['fingerprint']['parent']))


def checked_assets(spec,assets):
    ids={slide['asset_id'] for slide in spec['slides'] if slide.get('layout')=='image' and slide.get('asset_id')}
    if not ids<=set(assets or {}) or len(ids)>MAX_ASSETS:raise PresentationError('图片未获本演示稿授权或超过3个素材，请重新关联')
    selected={key:assets[key] for key in ids}
    if any(not isinstance(value,Asset) or value.id!=key or not 0<len(value.raw)<=MAX_ASSET or max(value.width,value.height)>1600 for key,value in selected.items()):raise PresentationError('素材快照无效')
    if sum(len(blob) for blob in {value.raw for value in selected.values()})>MAX_MEDIA:raise PresentationError('去重图片超过256 KiB预算，请减少图片或拆分演示稿')
    return selected


def geometry(asset):
    from pptx.util import Inches
    left,top,width,height=map(Inches,(.8,2,7.1,4.5))
    scale=min(width/asset.width,height/asset.height)
    w,h=round(asset.width*scale),round(asset.height*scale)
    return left+(width-w)//2,top+(height-h)//2,w,h


def add_picture(page,item,asset):
    shape=page.shapes.add_picture(BytesIO(asset.raw),*geometry(asset))
    shape.name='image-'+item.id


def validate_picture(shape,asset):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if shape.shape_type!=MSO_SHAPE_TYPE.PICTURE or tuple([shape.left,shape.top,shape.width,shape.height])!=geometry(asset):raise PresentationError('图片对象或等比位置不符')
    if any(value!=0 for value in [shape.crop_left,shape.crop_right,shape.crop_top,shape.crop_bottom]) or shape.rotation!=0:raise PresentationError('图片发生未批准裁切或旋转')
    part=shape.part.related_part(shape._element.blipFill.blip.rEmbed)
    if part.blob!=asset.raw or shape.image.size!=(asset.width,asset.height) or part.content_type!=('image/png' if asset.format=='PNG' else 'image/jpeg'):raise PresentationError('图片内容、尺寸或类型与授权快照不符')
    from lxml import etree
    from pptx import Presentation
    book=Presentation();page=book.slides.add_slide(book.slide_layouts[6]);expected=page.shapes.add_picture(BytesIO(asset.raw),*geometry(asset));expected.name=shape.name
    def normalized(element):
        node=deepcopy(element)
        for item in node.iter():
            if item.tag.rsplit('}',1)[-1]=='cNvPr':item.set('id','0')
            if item.tag.rsplit('}',1)[-1]=='blip':item.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed','approved-image')
        return etree.tostring(node,method='c14n',exclusive=True)
    if normalized(shape._element)!=normalized(expected._element):raise PresentationError('图片含未批准的效果、翻转或结构变化')
    return str(part.partname).lstrip('/')
