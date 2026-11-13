from copy import deepcopy
from io import BytesIO
import hashlib
import json
from pathlib import Path
import re
import sys
import tempfile
import time
import unicodedata
import uuid
import zipfile

from pydantic import BaseModel, ConfigDict, Field, StrictInt, field_validator, model_validator
from typing import Literal

from .rag import SECRET
from .rag_evidence import MAX_BYTES, _read


MAX_SLIDES = 12
LAYOUTS = ('cover', 'section', 'bullets', 'columns', 'table', 'chart', 'image')
FONT = 'Microsoft YaHei'
THEME = 'office-paper-v1'


class PresentationError(ValueError):
    pass


def identity(value):
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()


def identifier():
    return uuid.uuid4().hex


def generator_identity():
    import pptx
    package=Path(pptx.__file__).parent
    files=[Path(__file__),Path(__file__).with_name('presentation_preview_worker.py'),Path(__file__).with_name('presentation_charts.py'),Path(__file__).with_name('presentation_assets.py'),Path(__file__).with_name('presentation_image_worker.py'),package/'templates'/'default.pptx',*sorted(package.rglob('*.py'))]
    return {'version':THEME,'library':pptx.__version__,'source_files':len(files),'sha256':identity([hashlib.sha256(path.read_bytes()).hexdigest() for path in files])}


def speaker_notes(item):
    provenance='材料摘录，未独立核验' if item.origin=='source_excerpt' else '用户编辑内容，来源关系未核验'
    note=item.notes+'\n'+provenance
    if item.source_ids:note+='\n来源片段：'+', '.join(item.source_ids)
    if item.layout=='chart':note+='\n图表单位：'+item.chart_unit+'；数据由用户核对，不代表事实验证'
    if item.layout=='image':note+='\n图片：用户选择的本机素材，许可未独立核验；未作为事实证据自动核验'
    return note.strip()


class Slide(BaseModel):
    model_config = ConfigDict(extra='forbid', strict=True)
    id: str = Field(pattern=r'^[a-f0-9]{32}$')
    revision: StrictInt = Field(ge=1,le=1000000)
    layout: Literal['cover','section','bullets','columns','table','chart','image']
    title: str = Field(min_length=1,max_length=44)
    points: list[str] = Field(default_factory=list,max_length=5)
    right: list[str] = Field(default_factory=list,max_length=5)
    rows: list[list[str]] = Field(default_factory=list,max_length=6)
    chart_unit: str = Field(default='',max_length=20)
    chart_reviewed: bool = False
    asset_id: str = Field(default='',pattern=r'^(?:[a-f0-9]{32})?$')
    notes: str = Field(default='',max_length=1000)
    source_ids: list[str] = Field(default_factory=list,max_length=12)
    origin: Literal['source_excerpt','user_edited','needs_material'] = 'source_excerpt'

    @field_validator('title','notes','chart_unit')
    @classmethod
    def safe_text(cls,value):
        if '\x00' in value or any((ord(c)<32 and c not in '\n\t') or 0xD800<=ord(c)<=0xDFFF for c in value) or SECRET.search(value):
            raise ValueError('文字含不支持的控制符或疑似凭据')
        return value

    @field_validator('title')
    @classmethod
    def nonblank_title(cls,value):
        if not value.strip():raise ValueError('页面标题不能为空白')
        return value

    @model_validator(mode='after')
    def contents(self):
        for text in [*self.points,*self.right]:
            self.safe_text(text)
            if not text.strip() or len(text)>120:
                raise ValueError('每条内容应为1—120字符')
        for text in [cell for row in self.rows for cell in row]:
            self.safe_text(text)
            if len(text)>120:raise ValueError('表格单元格最多120字符')
        if self.rows and (not 1<=len(self.rows[0])<=4 or any(len(row)!=len(self.rows[0]) for row in self.rows)):
            raise ValueError('普通表格限6行4列，行列必须完整')
        if any(not re.fullmatch('[a-f0-9]{32}',value) for value in self.source_ids) or len(set(self.source_ids))!=len(self.source_ids):
            raise ValueError('来源标识无效')
        if self.layout in {'table','chart'} and (self.points or self.right):raise ValueError('数据页只使用表格内容')
        if self.layout not in {'table','chart'} and self.rows:raise ValueError('当前版式不接收表格')
        if self.layout!='chart' and (self.chart_unit or self.chart_reviewed):raise ValueError('当前版式不接收图表参数')
        if self.layout=='chart':
            if self.rows:
                from .presentation_charts import chart_values
                chart_values(self.rows)
            if self.chart_unit.lstrip().startswith(('=','+','-','@','{')):raise ValueError('单位需明确文字，不接收公式')
        if self.layout!='columns' and self.right:raise ValueError('当前版式没有右栏')
        if self.layout!='image' and self.asset_id:raise ValueError('当前版式不接收图片')
        if self.layout=='image' and len(self.points)>3:raise ValueError('图文页最多3条要点，请先减少内容；不会丢弃原输入')
        return self


class PresentationSpec(BaseModel):
    model_config=ConfigDict(extra='forbid',strict=True)
    schema_version: Literal[1] = 1
    id: str = Field(pattern=r'^[a-f0-9]{32}$')
    revision: StrictInt = Field(ge=1,le=1000000)
    theme: Literal['office-paper-v1'] = THEME
    title: str = Field(min_length=1,max_length=44)
    slides: list[Slide] = Field(min_length=1,max_length=MAX_SLIDES)

    @field_validator('schema_version',mode='before')
    @classmethod
    def version(cls,value):
        if type(value) is not int or value!=1:raise ValueError('不支持的版本')
        return value

    @model_validator(mode='after')
    def unique(self):
        Slide.safe_text(self.title)
        if not self.title.strip():raise ValueError('演示标题不能为空白')
        if len({slide.id for slide in self.slides})!=len(self.slides):raise ValueError('页面身份重复')
        return self


def make_outline(source,title,count):
    if type(count) is not int or not 2<=count<=MAX_SLIDES:raise PresentationError('页数应为2—12；较长材料请拆分演示稿')
    content=source['segments']
    if not content:raise PresentationError('没有可用材料，请先填写文字或选择已授权草稿')
    slides=[Slide(id=identifier(),revision=1,layout='cover',title=title,points=[],source_ids=[],origin='user_edited')]
    groups=[[] for _ in range(count-1)]
    for index,segment in enumerate(content):groups[min(index*(count-1)//len(content),count-2)].append(segment)
    for index,group in enumerate(groups,1):
        if not group:
            slides.append(Slide(id=identifier(),revision=1,layout='bullets',title=f'待补材料 {index}',origin='needs_material'))
            continue
        first=group[0]
        heading=first['text'].split('：',1)[0].split('\n',1)[0]
        heading=heading if 1<=len(heading)<=36 else f'材料 {index}'
        tables=[row for row in group if row.get('rows')]
        if len(group)==1 and tables:
            slides.append(Slide(id=identifier(),revision=1,layout='table',title=heading,rows=first['rows'],source_ids=[first['id']]))
            continue
        points=[]
        for segment in group:
            if segment.get('rows'):raise PresentationError('文字和表格落在同页，请增加页数或选择更少材料')
            text=segment['text']
            chunks=re.findall(r'.{1,100}(?:[。！？；]|$)',text) if len(text)>120 else [text]
            if not chunks or ''.join(chunks)!=text:raise PresentationError('材料段落过长，请先分段或选择明确节选')
            points.extend(chunks)
        if len(points)>5:raise PresentationError('每页最多5条要点，请增加页数或缩小材料范围')
        slides.append(Slide(id=identifier(),revision=1,layout='bullets',title=heading,points=points,source_ids=[row['id'] for row in group]))
    return PresentationSpec(id=identifier(),revision=1,title=title,slides=slides).model_dump()


def _lines(text,width,font):
    columns=max(1,width/font)
    return sum(max(1,__import__('math').ceil(sum(0 if unicodedata.combining(c) else 1 if unicodedata.east_asian_width(c) in 'WF' else .58 for c in line)/columns)) for line in text.split('\n'))


def build_deck(value,assets=None):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    spec=PresentationSpec.model_validate(value)
    from .presentation_assets import checked_assets,add_picture
    assets=checked_assets(spec.model_dump(),assets)
    book=Presentation();book.slide_width=Inches(13.333333);book.slide_height=Inches(7.5)
    def box(page,name,text,x,y,width,height,size=24,color='162334',bold=False):
        if sum(_lines(part,width*72,size) for part in text.split('\n'))*size*1.35>height*72-12:
            raise PresentationError('内容过密，请缩短文字、增加页面或换用单栏；不会自动缩小到不可读字号')
        shape=page.shapes.add_textbox(Inches(x),Inches(y),Inches(width),Inches(height));shape.name=name
        shape.text_frame.word_wrap=True
        shape.text_frame.margin_left=shape.text_frame.margin_right=Inches(.02)
        shape.text_frame.margin_top=shape.text_frame.margin_bottom=Inches(.02)
        for index,line in enumerate(text.split('\n')):
            paragraph=shape.text_frame.paragraphs[0] if index==0 else shape.text_frame.add_paragraph()
            paragraph.text=line;paragraph.font.name=FONT;paragraph.font.size=Pt(size);paragraph.font.bold=bold;paragraph.font.color.rgb=RGBColor.from_string(color)
            paragraph.space_after=Pt(10)
        return shape
    for index,item in enumerate(spec.slides):
        if item.origin=='needs_material' or (item.layout not in {'cover','section','table','chart'} and not item.points and not item.right) or (item.layout in {'table','chart'} and not item.rows):
            raise PresentationError(f'第{index+1}页尚缺材料，请补充或删除后预览')
        page=book.slides.add_slide(book.slide_layouts[6]);page.background.fill.solid();page.background.fill.fore_color.rgb=RGBColor(250,251,253)
        box(page,'oa-page',f'{index+1:02d}',12.0,6.8,.6,.45,12,'64748B')
        if item.layout in {'cover','section'}:
            box(page,'title-'+item.id,item.title,.9,2.0,11.5,1.7,36,'162334',True)
            if item.points:box(page,'body-'+item.id,'\n'.join(item.points),.95,4,11.2,1.6,24,'475569')
        else:
            box(page,'title-'+item.id,item.title,.75,.55,11.8,1.4,30,'162334',True)
            if item.layout=='image':
                if not item.asset_id:raise PresentationError('图文页尚未关联图片，请主动选择并核对素材')
                add_picture(page,item,assets[item.asset_id])
                box(page,'body-'+item.id,'\n'.join(item.points),8.35,2,4.1,4.5,24)
            elif item.layout=='chart':
                from .presentation_charts import add_chart,unit_text
                add_chart(page,item)
                box(page,'unit-'+item.id,unit_text(item.chart_unit),.85,6.35,11.6,.45,12,'475569')
            elif item.layout=='table':
                rows=item.rows;shape=page.shapes.add_table(len(rows),len(rows[0]),Inches(.8),Inches(2),Inches(11.7),Inches(4.2));shape.name='table-'+item.id
                for r,row in enumerate(rows):
                    for c,text in enumerate(row):
                        if _lines(text,11.7*72/len(row)-16,20)*27>4.2*72/len(rows)-12:raise PresentationError(f'第{index+1}页表格文字过密，请减少行列或缩短单元格')
                        cell=shape.table.cell(r,c);cell.text=text;cell.fill.solid();cell.fill.fore_color.rgb=RGBColor.from_string('E8EEF5' if r==0 else 'FFFFFF')
                        for p in cell.text_frame.paragraphs:p.font.name=FONT;p.font.size=Pt(20);p.font.color.rgb=RGBColor.from_string('162334');p.font.bold=r==0
            elif item.layout=='columns':
                box(page,'body-'+item.id,'\n'.join(item.points),.8,2,5.55,4.6,24)
                box(page,'right-'+item.id,'\n'.join(item.right),7,2,5.5,4.6,24)
            else:box(page,'body-'+item.id,'\n'.join(item.points),.8,2,11.6,4.6,24)
        page.notes_slide.notes_text_frame.text=speaker_notes(item)
    output=BytesIO();book.save(output);raw=output.getvalue()
    if len(raw)>MAX_BYTES:raise PresentationError('演示稿超过现有512 KiB发布预算，请减少内容')
    validate_deck(raw,value,assets)
    return raw


def validate_deck(raw,value,assets=None):
    from pptx import Presentation
    from .presentation_charts import checked_zip,validate_chart,unit_text
    spec=PresentationSpec.model_validate(value)
    from .presentation_assets import checked_assets,validate_picture
    assets=checked_assets(spec.model_dump(),assets)
    if len(raw)>MAX_BYTES:raise PresentationError('演示稿超过容量')
    with checked_zip(raw) as archive:
        names=archive.namelist()
        if len(names)!=len(set(names)) or len(names)>500 or sum(item.file_size for item in archive.infolist())>8*1024*1024:raise PresentationError('演示稿包结构或容量无效')
        for name in names:
            if name.startswith(('/','\\')) or '..' in Path(name).parts or any(word in name.lower() for word in ('vba','activex','oleobject')):raise PresentationError('不支持的演示稿部件')
            if name.endswith(('.xml','.rels')):
                from xml.etree.ElementTree import fromstring
                markup=archive.read(name)
                if b'<!doctype' in markup.lower() or b'<!entity' in markup.lower():raise PresentationError('不允许XML实体')
                root=fromstring(markup)
                if any(node.get('TargetMode','').lower()=='external' for node in root.iter()):raise PresentationError('演示稿含外部关系')
    book=Presentation(BytesIO(raw))
    allowed_workbooks=set();allowed_charts=set();allowed_images=set();image_links={}
    if len(book.slides)!=len(spec.slides) or abs(book.slide_width/914400-13.333333)>.001 or book.slide_height!=int(7.5*914400):raise PresentationError('页面数量或尺寸不符')
    for page,item in zip(book.slides,spec.slides):
        by_name={shape.name:shape for shape in page.shapes}
        expected={'oa-page','title-'+item.id}
        if item.layout=='image':expected.update({'image-'+item.id,'body-'+item.id})
        elif item.layout=='chart':expected.update({'chart-'+item.id,'unit-'+item.id})
        elif item.layout=='table':expected.add('table-'+item.id)
        elif item.layout not in {'cover','section'} or item.points:expected.add('body-'+item.id)
        if item.layout=='columns':expected.add('right-'+item.id)
        if len(by_name)!=len(page.shapes) or set(by_name)!=expected:raise PresentationError('含未解释的对象变化')
        for shape in page.shapes:
            if shape.left<0 or shape.top<0 or shape.left+shape.width>book.slide_width+2 or shape.top+shape.height>book.slide_height+2:raise PresentationError('对象超出页面')
            if shape.has_text_frame:
                size=12 if shape.name=='oa-page' or shape.name.startswith('unit-') else (36 if item.layout in {'cover','section'} else 30) if shape.name.startswith('title-') else 24
                if any(p.font.name!=FONT or p.font.size is None or p.font.size.pt!=size for p in shape.text_frame.paragraphs):raise PresentationError('文字格式不符')
        if by_name['title-'+item.id].text!=item.title:raise PresentationError('标题或页面顺序不符')
        image_rels=[rel for rel in page.part.rels.values() if rel.reltype.endswith('/image')]
        if item.layout=='image':
            if not item.asset_id or len(image_rels)!=1:raise PresentationError('图片关系或素材身份不符')
            name=validate_picture(by_name['image-'+item.id],assets[item.asset_id]);allowed_images.add(name)
            image_links[str(page.part.partname).lstrip('/')]=name
            if image_rels[0].is_external or str(image_rels[0].target_part.partname).lstrip('/')!=name:raise PresentationError('图片关系目标不符')
            if by_name['body-'+item.id].text!='\n'.join(item.points):raise PresentationError('图文说明不符')
        elif image_rels:raise PresentationError('非图文页含未批准图片')
        if item.layout=='chart':
            shape=by_name['chart-'+item.id]
            if not shape.has_chart or by_name['unit-'+item.id].text!=unit_text(item.chart_unit):raise PresentationError('图表或单位不符')
            workbook_name,chart_name=validate_chart(shape.chart,item)
            if workbook_name in allowed_workbooks or chart_name in allowed_charts:raise PresentationError('图表数据关系重复')
            allowed_workbooks.add(workbook_name);allowed_charts.add(chart_name)
        elif item.layout=='table':
            table=by_name['table-'+item.id].table
            if [[cell.text for cell in row.cells] for row in table.rows]!=item.rows:raise PresentationError('表格数据不符')
            if any(p.font.name!=FONT or p.font.size is None or p.font.size.pt!=20 or p.font.bold!=(index==0) for index,row in enumerate(table.rows) for cell in row.cells for p in cell.text_frame.paragraphs):raise PresentationError('表格文字格式不符')
        elif item.layout not in {'cover','section'} or item.points:
            if by_name['body-'+item.id].text!='\n'.join(item.points):raise PresentationError('正文文字不符')
        if item.layout=='columns' and by_name['right-'+item.id].text!='\n'.join(item.right):raise PresentationError('右栏文字不符')
        if page.notes_slide.notes_text_frame.text!=speaker_notes(item):raise PresentationError('讲稿或来源说明不符')
    embedded={name for name in names if 'embeddings/' in name.lower()}
    chart_parts={name for name in names if name.startswith('ppt/charts/') and not '/_rels/' in name}
    chart_rels={name for name in names if name.startswith('ppt/charts/') and '/_rels/' in name}
    if embedded!=allowed_workbooks or chart_parts!=allowed_charts or chart_rels!={str(Path(name).parent).replace('\\','/')+'/_rels/'+Path(name).name+'.rels' for name in allowed_charts}:raise PresentationError('含未批准的图表或内嵌部件')
    if {name for name in names if name.startswith('ppt/media/')}!=allowed_images:raise PresentationError('含未批准或孤立的图片部件')
    from xml.etree.ElementTree import fromstring
    import posixpath
    with zipfile.ZipFile(BytesIO(raw)) as media_zip:
        for rel_name in [name for name in names if name.endswith('.rels')]:
            base=posixpath.dirname(posixpath.dirname(rel_name))
            source=posixpath.join(base,Path(rel_name).name[:-5])
            for rel in fromstring(media_zip.read(rel_name)):
                if rel.get('Type','').endswith('/image') and source!='docProps/thumbnail.jpeg':
                    target=posixpath.normpath(posixpath.join(base,rel.get('Target','')))
                    if image_links.get(source)!=target:raise PresentationError('含未批准的图片引用')
    if any(not name.endswith(('.xml','.rels')) and name not in allowed_workbooks and name not in allowed_images and name not in {'docProps/thumbnail.jpeg','ppt/printerSettings/printerSettings1.bin'} for name in names):raise PresentationError('含未批准的二进制部件')
    import pptx
    with zipfile.ZipFile(BytesIO(raw)) as archive,zipfile.ZipFile(Path(pptx.__file__).parent/'templates/default.pptx') as template:
        name='ppt/printerSettings/printerSettings1.bin'
        if name in names and archive.read(name)!=template.read(name):raise PresentationError('打印设置与固定模板不符')
    return {'slides':len(book.slides),'editable_text':True,'editable_tables':True,'editable_charts':len(allowed_charts),'editable_pictures':len(image_links),
            'assets':[assets[key].manifest() for key in sorted(assets)],'notes':True,'imported_pptx':False}


def render_deck(value,assets=None):
    from .fidelity import runtime_path,_temporary_root,_run,_convert,package_projection
    from .render_identity import environment_identity
    runtime=runtime_path(None);environment=environment_identity(runtime);generator=generator_identity()
    if not environment['complete']:raise PresentationError('无法完整识别渲染环境，暂不发布验证副本')
    with tempfile.TemporaryDirectory(prefix='oap-',dir=_temporary_root()) as directory:
        root=Path(directory);(root/'spec.json').write_text(json.dumps(value,ensure_ascii=False),encoding='utf8')
        runs=[];deadline=time.monotonic()+30
        (root/'presentation.pptx').write_bytes(build_deck(value,assets))
        if not any(row['layout'] in {'chart','image'} for row in value['slides']):package_projection(root/'presentation.pptx',for_render=True)
        try:
            _convert(root/'presentation.pptx',root/'pdf',runtime,25000,runs)
            remaining=int((deadline-time.monotonic())*1000)
            if remaining<100:raise PresentationError('预览超时，未发布')
            _run((sys.executable,str(Path(__file__).with_name('presentation_preview_worker.py')),str(root)),min(8000,remaining),runs)
        except ValueError:
            error=PresentationError('本机生成或渲染未完成，未发布文件；请核对材料与渲染器')
            error.details=json.loads(_read(root/'failure.json')[0]) if (root/'failure.json').is_file() else {'stage':'worker','code':'timeout_or_cleanup'}
            error.details['processes']=runs
            raise error from None
        raw,_=_read(root/'presentation.pptx');report=json.loads(_read(root/'report.json')[0])
        checked=validate_deck(raw,value,assets)
        if report['pages']!=len(value['slides']) or environment_identity(runtime)!=environment or generator_identity()!=generator:raise PresentationError('页数、模板或渲染环境改变，请重新预览')
        previews=[_read(root/f'page-{i+1}.png')[0] for i in range(report['pages'])]
        return {'raw':raw,'previews':previews,'report':{**report,**checked,'environment':environment,'generator':generator,'processes':runs,'sha256':hashlib.sha256(raw).hexdigest(),'spec_sha256':identity(value),'renderer':'LibreOffice/PDFium; not Microsoft PowerPoint'}}
