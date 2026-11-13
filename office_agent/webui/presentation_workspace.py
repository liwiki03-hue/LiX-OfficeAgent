import base64
from copy import deepcopy
from io import BytesIO
import os
from pathlib import Path
import threading
import time

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from pydantic import ValidationError

from ..presentation_drafts import PresentationSpec, Slide, PresentationError, identifier, identity, make_outline, render_deck, validate_deck, generator_identity
from ..rag_evidence import RagEvidence, EvidenceError, _json, _hash, _dir, TTL
from ..writes import Workspace
from ..presentation_assets import prepare_image,approved_asset,checked_assets,check_source,processor_identity


class PresentationArchive(RagEvidence):
    artifact_name='presentation.pptx'
    directory_prefix='.office-presentation-'

    def __init__(self,workspace,snapshot,fresh,assets=None):
        super().__init__(workspace,lambda key:None)
        self.snapshot,self.fresh=deepcopy(snapshot),fresh
        self.assets=dict(assets or {})

    def _fresh(self,hits):
        self.fresh()

    def _save_audit(self,plan):
        return _json({'version':1,'operation':'publish_presentation','spec_sha256':identity(self.snapshot['spec']),
                      'source_sha256':self.snapshot['source']['sha256'],'output_sha256':_hash(plan['payload']),
                      'verification':self.snapshot['report']})

    def _exclusive(self,path,data):
        if path.name=='backup':data=_json({'kind':'new_presentation_source_and_approved_spec','snapshot':self.snapshot})
        if path.name=='stage':validate_deck(data,self.snapshot['spec'],self.assets)
        return super()._exclusive(path,data)


class PresentationWorkspace:
    def __init__(self,api):
        self.api=api
        self.lock=threading.RLock()
        self.spec=self.source=self.rendered=None
        self.task=None
        self.status='empty'
        self.error=''
        self.diagnostic={}
        self.plans={}
        self.saved={}
        self.uncertain=False
        self.epoch=0
        self.materials_pending=None
        self.assets={}
        self.image_pending=None
        self.image_task=None

    def invalidate(self):
        with self.lock:
            self.epoch+=1
            self.spec=self.source=self.rendered=None
            for item in self.plans.values():item['archive'].invalidate()
            self.plans.clear()
            for item in self.saved.values():item['archive'].invalidate()
            self.saved.clear()
            self.status='empty'
            self.error='';self.diagnostic={}
            self.materials_pending=None
            self.assets.clear();self.image_pending=None

    def guard(self,draft_id=None,revision=None):
        self.api._floating.authenticated()
        if self.uncertain:raise PresentationError('上次发布结果不确定，请核对产物并重启；不会自动重试')
        if draft_id is not None and (not self.spec or self.spec['id']!=draft_id or type(revision) is not int or revision!=self.spec['revision']):
            raise PresentationError('演示稿已改变，请读取最新大纲；旧操作未应用')

    def source_status(self,source=None):
        source=self.source if source is None else source
        if source['kind']=='text':return ('current',source['sha256'])
        floating=self.api._floating
        floating.check_source()
        if not floating.document or floating.document['id']!=source['document_id'] or not floating.paged.active():raise PresentationError('来源已切换，请重新准入并关联当前草稿')
        current=floating.paged.revision
        return ('current' if current==source['draft_revision'] else 'old_draft',str(current))

    def state(self):
        with self.lock:
            self.api._floating.authenticated()
            source_status='none'
            if self.source:
                try:source_status=self.source_status()[0]
                except ValueError:source_status='unavailable'
            return {'ok':True,'spec':deepcopy(self.spec),'source':deepcopy(self.source),'source_status':source_status,
                    'status':self.status,'error':self.error,'diagnostic':deepcopy(self.diagnostic),'busy':self.task is not None,'uncertain':self.uncertain,
                    'preview_revision':self.rendered['revision'] if self.rendered else None,
                    'image_busy':self.image_task is not None,
                    'assets':[{**row['asset'].manifest(),'name':row['name'],'thumbnail':row['thumbnail']} for row in self.assets.values()],
                    'outputs':[{key:value for key,value in row.items() if key in {'id','path','revision','status'}} for row in self.saved.values()]}

    def capture(self,payload):
        if type(payload) is not dict or set(payload)!={'kind','title','pages','text','document_id','draft_revision','replace_id'}:
            raise PresentationError('来源参数不完整')
        if payload['kind'] not in {'text','draft'} or type(payload['text']) is not str or type(payload['title']) is not str or not payload['title'].strip() or len(payload['title'])>44:
            raise PresentationError('请输入44字以内标题并选择有效来源')
        segments=[];limited=False;total=0;omitted=[]
        if payload['kind']=='text':
            if payload['document_id']!='' or type(payload['draft_revision']) is not int or payload['draft_revision']!=0:raise PresentationError('直接文字不能携带文档身份')
            text=payload['text'].strip()
            if not text or len(text)>12000:raise PresentationError('直接材料应为1—12000字符；较长内容请分批')
            blocks=[line.strip() for line in text.splitlines() if line.strip()]
            total=len(blocks)
            if total>40:raise PresentationError('本次最多40段材料，请缩小范围')
            for index,line in enumerate(blocks):segments.append({'id':identifier(),'position':index+1,'text':Slide.safe_text(line)})
            source={'kind':'text','name':'用户提供文字','sha256':identity(text)}
        else:
            if payload['text']!='':raise PresentationError('草稿来源不接收额外文字，请分开新建')
            paged=self.api._floating.paged
            with paged.lock:
                paged.guard(payload['document_id'],payload['draft_revision'])
                if paged.snapshot_revision!=paged.revision:raise PresentationError('当前草稿尚未同步，请稍后重新捕获')
                raw=bytes(paged.data);book=Document(BytesIO(raw));budget=12000
                nodes=list(book.element.body.iterchildren());total=sum(node.tag.rsplit('}',1)[-1] in {'p','tbl'} for node in nodes)
                position=0
                for node in nodes:
                    kind=node.tag.rsplit('}',1)[-1]
                    if kind not in {'p','tbl'}:continue
                    position+=1
                    if len(segments)>=40:limited=True;break
                    if kind=='p':
                        text=Paragraph(node,book).text.strip()
                        if not text:continue
                        rows=None
                    else:
                        table=Table(node,book)
                        if len(table.rows)>6 or len(table.columns)>4 or any(item.tag.rsplit('}',1)[-1] in {'gridSpan','vMerge'} for item in node.iter()) or len(list(node.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')))>1:
                            omitted.append(position);limited=True;continue
                        rows=[[cell.text.strip() for cell in row.cells] for row in table.rows]
                        if any(len(cell)>120 for row in rows for cell in row):omitted.append(position);limited=True;continue
                        text='普通表格 '+str(position)
                    cost=len(text)+sum(len(cell) for row in (rows or []) for cell in row)
                    if cost>1500 or cost>budget:omitted.append(position);limited=True;continue
                    budget-=cost;Slide.safe_text(text)
                    segment={'id':identifier(),'position':position,'text':text}
                    if rows:segment['rows']=rows
                    segments.append(segment)
                source={'kind':'draft','name':paged.source.name,'sha256':__import__('hashlib').sha256(raw).hexdigest(),
                        'document_id':paged.document_id,'source_version':paged.source_version,'draft_revision':paged.revision}
        return {**source,'segments':segments,'limited':limited,'total_blocks':total,'included_blocks':len(segments),'omitted_blocks':omitted,
                'coverage':'选定文字，或当前草稿正文的有界普通段落/普通表格；不含页眉、脚注、图片与未纳入范围'}

    def create(self,payload):
        with self.lock:
            self.guard()
            if type(payload) is not dict:raise PresentationError('来源参数无效')
            if self.task:raise PresentationError('正在生成预览，请等待后再新建')
            if set(payload)=={'material_id','selected_ids','title','pages','replace_id','replace_revision'}:
                pending=self.materials_pending
                if not pending or type(payload['material_id']) is not str or payload['material_id']!=pending['id'] or pending['epoch']!=self.epoch or time.monotonic()-pending['created']>=TTL:raise PresentationError('材料选择已过期，请重新读取材料')
                base_id=self.spec['id'] if self.spec else '';base_revision=self.spec['revision'] if self.spec else 0
                if payload['replace_id']!=pending['base_id'] or type(payload['replace_revision']) is not int or payload['replace_revision']!=pending['base_revision'] or (base_id,base_revision)!=(pending['base_id'],pending['base_revision']):raise PresentationError('现有大纲已改变，请核对后重新选择材料')
                if self.source_status(pending['source'])[0]!='current':raise PresentationError('来源草稿已改变，请重新读取材料')
                ids=payload['selected_ids']
                if type(ids) is not list or not ids or len(ids)>40 or any(type(value) is not str for value in ids) or len(set(ids))!=len(ids) or not set(ids)<={row['id'] for row in pending['source']['segments']}:raise PresentationError('请选择有效且不重复的宿主材料')
                source=deepcopy(pending['source']);original=source['segments'];source['segments']=[row for row in original if row['id'] in ids]
                source.update(available_blocks=len(original),included_blocks=len(ids),unselected_blocks=[row['position'] for row in original if row['id'] not in ids],selection_sha256=identity(source['segments']),limited=source['limited'] or len(ids)!=len(original))
            else:
                if self.spec and payload.get('replace_id')!=self.spec['id']:raise PresentationError('已有演示稿草稿，请明确确认替换后新建')
                source=self.capture(payload)
            spec=make_outline(source,payload['title'].strip(),payload['pages'])
            self.source,self.spec=source,spec
            self.epoch+=1;self.assets.clear();self.image_pending=None
            self.materials_pending=None
            self.rendered=None;self.plans.clear();self.status='outline';self.error='';self.diagnostic={}
            return self.state()

    def materials(self,payload):
        with self.lock:
            self.guard();source=self.capture(payload)
            token=identifier()
            self.materials_pending={'id':token,'epoch':self.epoch,'created':time.monotonic(),'source':source,
                                    'base_id':self.spec['id'] if self.spec else '','base_revision':self.spec['revision'] if self.spec else 0}
            return {'ok':True,'material_id':token,'source':deepcopy(source),'replace_id':self.materials_pending['base_id'],
                    'replace_revision':self.materials_pending['base_revision'],'expires_in_seconds':TTL}

    def edit(self,draft_id,revision,action,slide_id,fields):
        with self.lock:
            self.guard(draft_id,revision)
            if type(fields) is not dict:raise PresentationError('修改参数无效')
            spec=deepcopy(self.spec);slides=spec['slides']
            index=next((i for i,row in enumerate(slides) if row['id']==slide_id),None)
            if action=='add':
                if fields or len(slides)>=12:raise PresentationError('最多12页；较长任务请拆稿')
                slides.append(Slide(id=identifier(),revision=1,layout='bullets',title='新增页面',origin='needs_material').model_dump())
            elif index is None:raise PresentationError('页面身份失效，未按旧序号执行')
            elif action=='update':
                if not fields or not set(fields)<= {'title','layout','points','right','rows','notes','chart_unit','chart_reviewed'}:raise PresentationError('仅支持文字、版式、普通表格、图表参数和讲稿字段')
                if any(key in fields and fields[key]!=slides[index][key] for key in ['rows','chart_unit']) and 'chart_reviewed' not in fields:fields={**fields,'chart_reviewed':False}
                item=slides[index];item.update(fields)
                if item['layout']!='image':item['asset_id']=''
                item['revision']+=1;item['origin']='user_edited'
                slides[index]=Slide.model_validate(item).model_dump()
            elif action=='image_reference':
                if set(fields)!={'asset_id'} or type(fields['asset_id']) is not str or slides[index]['layout']!='image':raise PresentationError('请选择图文版式并使用已授权素材')
                if fields['asset_id'] and fields['asset_id'] not in self.assets:raise PresentationError('素材授权已失效，请重新选择')
                slides[index]['asset_id']=fields['asset_id'];slides[index]['revision']+=1
            elif action=='delete':
                if fields or len(slides)<=1:raise PresentationError('至少保留1页')
                slides.pop(index)
            elif action=='move':
                if set(fields)!={'direction'} or fields['direction'] not in {'up','down'}:raise PresentationError('移动方向无效')
                target=index+(-1 if fields['direction']=='up' else 1)
                if not 0<=target<len(slides):raise PresentationError('页面已在边界')
                slides[index],slides[target]=slides[target],slides[index]
            else:raise PresentationError('不支持的演示稿操作')
            spec['revision']+=1
            checked_assets(spec,{key:row['asset'] for key,row in self.assets.items()})
            self.spec=PresentationSpec.model_validate(spec).model_dump()
            self.image_pending=None
            used={row['asset_id'] for row in self.spec['slides'] if row.get('asset_id')}
            self.assets={key:row for key,row in self.assets.items() if key in used}
            self.plans.clear();self.rendered=None;self.status='building' if self.task else 'outline';self.error='';self.diagnostic={}
            return self.state()

    def fresh_assets(self):
        assets=checked_assets(self.spec,{key:row['asset'] for key,row in self.assets.items()})
        for key,asset in assets.items():
            row=self.assets[key];check_source(row['path'],row['fingerprint'])
            if identity(processor_identity())!=asset.processor_sha256:raise PresentationError('图片处理器已改变，请重新关联素材')
        return assets

    def choose_image(self,draft_id,revision,slide_id,choose):
        with self.lock:
            self.guard(draft_id,revision)
            slide=next((row for row in self.spec['slides'] if row['id']==slide_id),None)
            if not slide or slide['layout']!='image':raise PresentationError('请先保存本页图文版式')
            if self.image_task:raise PresentationError('正在读取图片，请勿重复选择')
            token=identifier();self.image_task=token;epoch=self.epoch;source_state=self.source_status()
        try:
            path=choose()
            if not path:return {'ok':True,'cancelled':True}
            with self.lock:
                self.guard(draft_id,revision)
                if epoch!=self.epoch or self.source_status()!=source_state:raise PresentationError('选择期间来源已改变，图片未关联')
            candidate=prepare_image(path)
            with self.lock:
                self.guard(draft_id,revision)
                if epoch!=self.epoch or self.source_status()!=source_state:raise PresentationError('解码期间来源已改变，图片未关联')
                ticket=identifier()
                self.image_pending={'id':ticket,'epoch':epoch,'revision':revision,'slide_id':slide_id,'source_state':source_state,'created':time.monotonic(),'candidate':candidate}
                return {'ok':True,'ticket':ticket,'original':candidate['original'],'name':Path(path).name,'thumbnail':candidate['thumbnail'],
                        'variants':{key:{field:value for field,value in item.items() if field!='raw'} for key,item in candidate['variants'].items()},
                        'scope':'仅本次明确选择的文件；采用后进入当前页草稿，不是发布批准'}
        finally:
            with self.lock:
                if self.image_task==token:self.image_task=None

    def use_image(self,draft_id,revision,slide_id,ticket,variant,approved):
        with self.lock:
            self.guard(draft_id,revision)
            if type(approved) is not bool:raise PresentationError('素材采用参数无效')
            pending=self.image_pending
            if not pending or type(ticket) is not str or ticket!=pending['id'] or pending['epoch']!=self.epoch or pending['revision']!=revision or pending['slide_id']!=slide_id or time.monotonic()-pending['created']>=TTL:raise PresentationError('素材候选已过期或页面已改变，请重新选择')
            self.image_pending=None
            if not approved:return {**self.state(),'cancelled':True}
            if self.source_status()!=pending['source_state']:raise PresentationError('来源已改变，请重新选择图片')
            candidate=pending['candidate'];check_source(candidate['path'],candidate['fingerprint'])
            if processor_identity()!=candidate['processor']:raise PresentationError('图片处理器已改变，请重新选择')
            if type(variant) is not str:raise PresentationError('素材候选参数无效')
            asset=approved_asset(candidate,variant)
            existing=next((key for key,row in self.assets.items() if row['path']==candidate['path'] and row['fingerprint']==candidate['fingerprint'] and row['asset'].raw==asset.raw),None)
            if existing:asset=self.assets[existing]['asset']
            spec=deepcopy(self.spec);slide=next((row for row in spec['slides'] if row['id']==slide_id),None)
            if not slide or slide['layout']!='image':raise PresentationError('图文页已失效')
            slide['asset_id']=asset.id;slide['revision']+=1;spec['revision']+=1
            assets={**self.assets,asset.id:{'asset':asset,'path':candidate['path'],'fingerprint':candidate['fingerprint'],'name':candidate['path'].name,'thumbnail':candidate['variants'][variant]['preview'],'process':candidate['process']}}
            selected=checked_assets(spec,{key:row['asset'] for key,row in assets.items()})
            self.spec=PresentationSpec.model_validate(spec).model_dump();self.assets={key:assets[key] for key in selected}
            self.plans.clear();self.rendered=None;self.status='building' if self.task else 'outline';self.error='';self.diagnostic={}
            return self.state()

    def build(self,draft_id,revision):
        with self.lock:
            self.guard(draft_id,revision)
            self.source_status()
            assets=self.fresh_assets()
            if self.task:raise PresentationError('正在生成，请勿重复提交')
            if self.api._app.busy:raise PresentationError('文档任务正在运行，请结束后再生成预览')
            token=identifier();self.task=token;self.status='building';self.error='';self.diagnostic={};spec=deepcopy(self.spec);epoch=self.epoch
            def run():
                try:
                    result=render_deck(spec,assets) if assets else render_deck(spec)
                    with self.lock:
                        if self.epoch==epoch and self.spec and self.spec['id']==draft_id and self.spec['revision']==revision:
                            self.rendered={**result,'revision':revision};self.status='ready'
                except Exception as error:
                    with self.lock:
                        if self.epoch==epoch and self.spec and self.spec['id']==draft_id and self.spec['revision']==revision:
                            self.status='failed';self.error=str(error)[:180] if isinstance(error,PresentationError) else '生成或渲染失败，未发布文件；请核对材料、版式和本机渲染器后主动重试'
                            self.diagnostic=getattr(error,'details',{})
                finally:
                    with self.lock:
                        if self.task==token:
                            self.task=None
                            if self.status=='building':self.status='outline'
            threading.Thread(target=run,daemon=True).start()
            return self.state()

    def preview(self,draft_id,revision,slide_id):
        with self.lock:
            self.guard(draft_id,revision)
            if not self.rendered or self.rendered['revision']!=revision:raise PresentationError('预览已过期，请重新生成')
            index=next((i for i,row in enumerate(self.spec['slides']) if row['id']==slide_id),None)
            if index is None:raise PresentationError('页面已移除')
            return {'ok':True,'slide_id':slide_id,'revision':revision,'image':'data:image/png;base64,'+base64.b64encode(self.rendered['previews'][index]).decode(),
                    'renderer':self.rendered['report']['renderer']}

    def plan(self,draft_id,revision,allow_old,choose):
        with self.lock:
            self.guard(draft_id,revision)
            if type(allow_old) is not bool or not self.rendered or self.rendered['revision']!=revision or self.task:raise PresentationError('请先生成当前大纲的实际预览')
            source_state=self.source_status()
            assets=self.fresh_assets()
            if source_state[0]=='old_draft' and not allow_old:raise PresentationError('来源草稿已有新修改，请明确选择继续旧快照，或重新关联来源')
            snapshot={'spec':deepcopy(self.spec),'source':deepcopy(self.source),'report':deepcopy(self.rendered['report']),
                      'assets':[assets[key].manifest() for key in sorted(assets)],'asset_processing':[self.assets[key]['process'] for key in sorted(assets)]}
            raw=self.rendered['raw'];epoch=self.epoch
        folder=choose()
        if not folder:return {'ok':True,'cancelled':True}
        with self.lock:
            self.guard(draft_id,revision)
            self.plans={key:row for key,row in self.plans.items() if time.monotonic()-row['created']<TTL}
            if len(self.saved)>=8:raise PresentationError('本窗口发布回执已达8份，请先核对现有产物')
            if epoch!=self.epoch or self.source_status()!=source_state or self.task:raise PresentationError('选择目录期间来源或草稿改变，请重新预览')
            self.fresh_assets()
            folder=Path(folder)
            if not folder.is_absolute():raise PresentationError('请选择真实本机目录')
            _dir(folder);space=Workspace(folder)
            if space.root!=folder:raise PresentationError('保存目录身份不符')
            def fresh():
                self.guard(draft_id,revision)
                if self.source_status()!=source_state:raise PresentationError('来源已改变，旧发布计划失效')
                if [asset.manifest() for key,asset in sorted(self.fresh_assets().items())]!=snapshot['assets']:raise PresentationError('图片授权已改变，请重新预览')
                from ..fidelity import runtime_path
                from ..render_identity import environment_identity
                if environment_identity(runtime_path(None))!=snapshot['report']['environment'] or generator_identity()!=snapshot['report']['generator']:raise PresentationError('渲染环境或模板改变，请重新生成预览')
            archive=PresentationArchive(space,snapshot,fresh,assets);session=identifier();archive.ensure_available(session)
            token=archive._id();archive._plans[token]={'session':session,'time':archive._clock(),'operation':'save','payload':raw,'hits':[]}
            if len(self.plans)>=8:raise PresentationError('待确认计划过多，请取消旧计划')
            self.plans[token]={'archive':archive,'session':session,'draft_id':draft_id,'revision':revision,'epoch':epoch,'created':time.monotonic()}
            path=folder/(archive.directory_prefix+token)/archive.artifact_name
            return {'ok':True,'plan':{'id':token,'path':str(path),'revision':revision,'pages':len(self.spec['slides']),'source_status':source_state[0],
                                      'confirmation_required':True,'expires_in_seconds':TTL,'wrote_nothing':True,'sha256':_hash(raw)}}

    def confirm(self,token,approved):
        with self.lock,self.api._app._lock:
            self.guard()
            entry=self.plans.pop(token,None) if type(token) is str else None
            if not entry or entry['epoch']!=self.epoch:raise PresentationError('发布计划已失效，请重新预览')
            if type(approved) is not bool:raise PresentationError('确认参数无效')
            if approved:
                if self.api._app.busy:raise PresentationError('已有文档事务，请结束后重新确认')
                self.guard(entry['draft_id'],entry['revision']);entry['archive'].fresh()
            result=entry['archive'].confirm(token,entry['session'],approved)
            if result['status']=='uncertain':self.uncertain=True
            if result['status']=='saved':
                path=str(entry['archive'].root/result['evidence'])
                self.saved[token]={**entry,'id':token,'path':path,'status':'verified'};result.update(path=path,revision=entry['revision'],status='verified')
            return {'ok':True,'result':result,'state':self.state()}

    def saved_action(self,save_id,action,approved=False):
        with self.lock,self.api._app._lock:
            self.guard();entry=self.saved.get(save_id) if type(save_id) is str else None
            if not entry:raise PresentationError('没有本窗口可用的发布回执')
            archive,session=entry['archive'],entry['session']
            if action=='open':
                saved=archive._saved_for(save_id,session);os.startfile(str(saved['directory']));return {'ok':True}
            if action=='rollback_plan':
                plan=archive.plan_rollback(save_id,session);entry['rollback']=plan['rollback_id'];return {'ok':True,'plan':plan}
            if action=='rollback':
                if self.api._app.busy:raise PresentationError('已有事务在执行，暂不能撤销')
                result=archive.confirm_rollback(entry.pop('rollback',None),session,approved)
                if result['status']=='uncertain':self.uncertain=True
                if result['status']=='rolled_back':entry['status']='rolled_back'
                return {'ok':True,'result':result}
            raise PresentationError('不支持的产物操作')


class PresentationApi:
    def _presentation_call(self,method,*args):
        def action():
            with self._lock:
                if not hasattr(self,'_presentations'):self._presentations=PresentationWorkspace(self)
                service=self._presentations
            try:return getattr(service,method)(*args)
            except ValidationError:
                raise PresentationError('页面字段校验未通过，请核对标题、版式、要点长度和表格行列；原大纲未改变') from None
            except EvidenceError as error:
                raise PresentationError({'evidence_expired':'计划已过期，请重新预览','evidence_path_changed':'保存目录或产物已改变，请重新选择并核对',
                                         'evidence_uncertain':'结果不确定，请核对产物后重启；不会自动重试','evidence_plan_invalid':'计划或撤销入口已失效，请重新核对',
                                         'evidence_capacity':'超出现有发布容量，请减少内容'}.get(error.code,'文件校验未通过，请核对目录和已有产物')) from None
        return self._floating_call(action)

    def presentation_state(self):return self._presentation_call('state')
    def presentation_create(self,payload):return self._presentation_call('create',payload)
    def presentation_materials(self,payload):return self._presentation_call('materials',payload)
    def presentation_image_choose(self,draft_id,revision,slide_id):
        def choose():
            if self._app.window is None:return None
            from webview import FileDialog
            values=self._app.window.create_file_dialog(FileDialog.OPEN,allow_multiple=False,file_types=('图片 (*.png;*.jpg;*.jpeg)',))
            return values[0] if values else None
        return self._presentation_call('choose_image',draft_id,revision,slide_id,choose)
    def presentation_image_use(self,draft_id,revision,slide_id,ticket,variant,approved=False):return self._presentation_call('use_image',draft_id,revision,slide_id,ticket,variant,approved)
    def presentation_edit(self,draft_id,revision,action,slide_id,fields):return self._presentation_call('edit',draft_id,revision,action,slide_id,fields)
    def presentation_build(self,draft_id,revision):return self._presentation_call('build',draft_id,revision)
    def presentation_preview(self,draft_id,revision,slide_id):return self._presentation_call('preview',draft_id,revision,slide_id)
    def presentation_plan(self,draft_id,revision,allow_old=False):
        def choose():
            if self._app.window is None:return None
            from webview import FileDialog
            values=self._app.window.create_file_dialog(FileDialog.FOLDER)
            return values[0] if values else None
        return self._presentation_call('plan',draft_id,revision,allow_old,choose)
    def presentation_confirm(self,token,approved=False):return self._presentation_call('confirm',token,approved)
    def presentation_saved_action(self,save_id,action,approved=False):return self._presentation_call('saved_action',save_id,action,approved)
