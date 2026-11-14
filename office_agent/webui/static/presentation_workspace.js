(() => {
  'use strict';
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  let root=null,state=null,selected=null,dirty=false,operation=false,epoch=0,timer=null,formRevision=0,previewKey='',polling=false,sourceNode=null,opening=false,returnFocus=null;
  const dialogs=new Set();
  const names={cover:'封面',section:'章节',bullets:'标题与要点',columns:'双栏对比',table:'普通表格',chart:'数据柱状图',image:'图文'};
  const make=(tag,text,parent)=>{const node=document.createElement(tag);if(text)node.textContent=text;parent?.append(node);return node;};
  const button=(text,parent,action)=>{const node=make('button',text,parent);node.type='button';node.addEventListener('click',async()=>{try{await action();}catch(error){if(!error.stale)message(error.message);}});return node;};
  const api=()=>window.OfficeAuth?.wrapApi(window.pywebview?.api);
  const message=text=>{const node=root?.querySelector('[data-presentation-status]');if(node)node.textContent=text;};
  async function call(name,...args){
    const mine=epoch,result=await api()[name](...args);
    if(mine!==epoch){const error=Error('此面板已关闭；请重新打开查看实际状态，勿重复确认');error.stale=true;throw error;}
    if(!result?.ok)throw Error(result?.error||'此步骤未完成，请核对状态后主动重试');
    return result;
  }
  function modal(title,previous){
    const node=make('dialog','',document.body);node.className='presentation-dialog';node.setAttribute('aria-label',title);make('h2',title,node);
    dialogs.add(node);node.addEventListener('close',()=>{dialogs.delete(node);node.remove();if(previous?.isConnected)previous.focus();});node.showModal();return node;
  }
  function field(label,parent,value,tag='input'){
    const wrapper=make('label','',parent);make('span',label,wrapper);const node=make(tag,'',wrapper);node.setAttribute('aria-label',label);node.value=value||'';return node;
  }
  async function createSource(previous){
    if(sourceNode?.isConnected){sourceNode.focus();return;}
    const node=modal('新建演示稿',previous),kind=field('材料来源',node,'','select');sourceNode=node;
    node.addEventListener('close',()=>{if(sourceNode===node)sourceNode=null;});
    for(const [value,label] of [['text','直接提供文字'],['draft','当前未发布 DOCX 草稿']]){const option=make('option',label,kind);option.value=value;}
    kind.value=window.OfficePaged?.active()?'draft':'text';
    const title=field('演示标题',node,'技术报告');title.maxLength=44;
    const count=field('页数（2—12）',node,'8');count.type='number';count.min=2;count.max=12;
    const text=field('材料（按段落换行）',node,'','textarea');text.maxLength=12000;text.rows=8;text.parentElement.hidden=kind.value!=='text';
    let sourceEpoch=0,prepared=null,selectedMaterials=new Set();
    const materialList=make('section','',node);materialList.className='presentation-materials';
    const resetMaterials=()=>{sourceEpoch++;prepared=null;selectedMaterials.clear();materialList.replaceChildren();};
    kind.onchange=()=>{text.parentElement.hidden=kind.value!=='text';resetMaterials();};text.oninput=resetMaterials;
    make('p','先整理大纲，不调用模型；事实与来源需自行核对。',node).className='presentation-muted';
    const error=make('p','',node);error.setAttribute('role','alert');
    let busy=false;
    const actions=make('div','',node);actions.className='presentation-actions';
    const cancel=button('取消',actions,()=>{if(!busy)node.close();});
    async function payload(){
      const data={kind:kind.value,title:title.value,pages:Number(count.value),text:kind.value==='text'?text.value:'',document_id:'',draft_revision:0,replace_id:state?.spec?.id||''};
      if(data.kind==='draft'){
        if(!window.OfficePaged?.active())throw Error('请先打开并授权DOCX，或选择直接提供文字');
        Object.assign(data,await window.OfficePaged.presentationSource());
      }
      return data;
    }
    const choose=button('选择材料',actions,async()=>{
      if(busy)return;busy=true;create.disabled=choose.disabled=cancel.disabled=true;const version=sourceEpoch;
      try{
        const result=await call('presentation_materials',await payload());
        if(version!==sourceEpoch){error.textContent='材料输入已改变，请重新选择材料。';return;}
        prepared=result;selectedMaterials=new Set(result.source.segments.map(row=>row.id));materialList.replaceChildren();
        const status=make('p','',materialList);
        const countSelected=()=>{status.textContent='已选 '+selectedMaterials.size+' / '+result.source.included_blocks+' 块 · 原正文共 '+result.source.total_blocks+' 块'+(result.source.limited?' · 仍有材料未纳入':'');};countSelected();
        for(const row of result.source.segments){
          const label=make('label','',materialList);label.className='presentation-material-option';
          const checkbox=make('input','',label);checkbox.type='checkbox';checkbox.checked=true;checkbox.dataset.materialId=row.id;
          make('span','第'+row.position+'块 · '+(row.rows?'普通表格':row.text.slice(0,70)),label);
          checkbox.onchange=()=>{if(checkbox.checked)selectedMaterials.add(row.id);else selectedMaterials.delete(row.id);countSelected();};
          const details=make('details','',materialList);make('summary','查看材料',details);make('pre',row.rows?row.rows.map(cells=>cells.join(' | ')).join('\n'):row.text,details);
        }
        error.textContent='未选材料不会用于本次大纲；不代表不存在证据。';
      }catch(failure){if(!failure.stale)error.textContent=failure.message;}
      finally{busy=false;create.disabled=choose.disabled=cancel.disabled=false;}
    });
    const create=button('整理大纲',actions,async()=>{
      if(busy)return;busy=true;create.disabled=choose.disabled=cancel.disabled=true;for(const input of [kind,title,count,text])input.disabled=true;
      try{
        const data=prepared?{material_id:prepared.material_id,selected_ids:[...selectedMaterials],title:title.value,pages:Number(count.value),replace_id:prepared.replace_id,replace_revision:prepared.replace_revision}:await payload();
        const result=await call('presentation_create',data);
        state=result;selected=result.spec.slides[0].id;dirty=false;node.close();mount();
      }catch(failure){if(!failure.stale)error.textContent=failure.message;}
      finally{busy=false;create.disabled=choose.disabled=cancel.disabled=false;for(const input of [kind,title,count,text])input.disabled=false;}
    });create.className='presentation-primary';
    node.addEventListener('cancel',event=>{if(busy)event.preventDefault();});title.focus();
  }
  const currentSlide=()=>state?.spec?.slides.find(row=>row.id===selected);
  const imageInfo=value=>value.format+' · '+value.width+' × '+value.height+' · '+(value.bytes/1024).toFixed(1)+' KiB';
  async function chooseImage(){
    if(operation)return;
    if(dirty){await edit('update',formValues());if(dirty)return;}
    if(currentSlide()?.layout!=='image')throw Error('请先将本页保存为图文版式');
    operation=true;controls();const id=state.spec.id,revision=state.spec.revision,slide=selected;
    try{
      message('正在读取图片…');
      const result=await call('presentation_image_choose',id,revision,slide);
      if(result.cancelled){message('');operation=false;controls();return;}
      if(!root||state.spec.id!==id||state.spec.revision!==revision||selected!==slide)throw Error('当前页面已改变，请重新选择图片');
      message('');
      const node=modal('核对本机图片',root.querySelector('[data-presentation-choose-image]'));
      make('p',result.name+' · '+imageInfo(result.original),node);
      make('p','优化只用于演示稿，原图片保持不变。',node).className='presentation-muted';
      const list=make('div','',node);list.className='presentation-image-candidates';
      let choice='',pending=true,inflight=false;
      const options=[];
      for(const [key,value] of Object.entries(result.variants)){
        const label=make('label','',list),radio=make('input','',label);radio.type='radio';radio.name='presentation-image-variant';radio.value=key;
        const allowed=value.bytes<=192*1024&&Math.max(value.width,value.height)<=1600;
        radio.disabled=!allowed;if(!choice&&allowed){choice=key;radio.checked=true;}
        make('strong',key==='normalized'?'去除元数据的副本':'等比优化候选',label);
        make('span',imageInfo(value)+(allowed?'':' · 超过本页素材预算'),label);
        if(/^data:image\/png;base64,/.test(value.preview)){const image=make('img','',label);image.src=value.preview;image.alt=key==='normalized'?'归一化副本预览':'优化候选预览';}
        radio.onchange=()=>{choice=key;};options.push(radio);
      }
      const license=make('label','',node);license.className='presentation-material-option';const reviewed=make('input','',license);reviewed.type='checkbox';
      make('span','我确认有权使用，并已核对所选副本的朝向、清晰度和透明度',license);
      const status=make('p',choice?'许可未独立核验；不会自动读取图片中的数字或文字。':'候选仍超预算，请取消并选择更小图片。',node);status.setAttribute('role','alert');
      const actions=make('div','',node);actions.className='presentation-actions';
      const cancel=button('取消',actions,async()=>{
        if(inflight||!pending)return;inflight=true;pending=false;cancel.disabled=adopt.disabled=true;
        try{await call('presentation_image_use',id,revision,slide,result.ticket,'',false);node.close();}
        catch(error){if(!error.stale)status.textContent=error.message;cancel.textContent='关闭';cancel.onclick=()=>node.close();}
        finally{inflight=false;cancel.disabled=false;}
      });
      const adopt=button('采用到本页草稿',actions,async()=>{
        if(!pending||inflight||!choice||!reviewed.checked)return;
        inflight=true;pending=false;adopt.disabled=cancel.disabled=true;
        try{
          const changed=dirty,values=JSON.stringify(formValues());
          const next=await call('presentation_image_use',id,revision,slide,result.ticket,choice,true);
          state=next;node.close();renderList();
          if(!changed&&JSON.stringify(formValues())===values){dirty=false;renderEditor();}
          else{formRevision=state.spec.revision;renderImageControls();message('图片已采用；新的文字输入仍待保存。');}
          await showPreview();
        }catch(error){if(!error.stale)status.textContent=error.message;cancel.textContent='关闭';cancel.onclick=()=>node.close();}
        finally{inflight=false;cancel.disabled=false;}
      });adopt.className='presentation-primary';adopt.disabled=true;
      reviewed.onchange=()=>{adopt.disabled=!reviewed.checked||!choice||!pending;};
      node.addEventListener('cancel',event=>{event.preventDefault();if(!inflight)cancel.click();});
      node.addEventListener('close',()=>{operation=false;controls();},{once:true});cancel.focus();
    }catch(error){operation=false;controls();throw error;}
  }
  function renderImageControls(){
    const pane=root?.querySelector('[data-presentation-assets]');if(!pane)return;pane.replaceChildren();
    const slide=currentSlide(),asset=(state.assets||[]).find(row=>row.id===slide?.asset_id);
    if(asset){
      if(/^data:image\/png;base64,/.test(asset.thumbnail)){const image=make('img','',pane);image.src=asset.thumbnail;image.alt='当前页授权图片';}
      make('p',asset.name+' · '+imageInfo(asset),pane);
    }else make('p','为本页选择一张图片，完整等比显示。',pane);
    const actions=make('div','',pane);actions.className='presentation-actions';
    button(asset?'替换图片':'选择图片',actions,chooseImage).dataset.presentationChooseImage='';
    if(asset)button('移除图片',actions,()=>{if(dirty)throw Error('请先保存文字输入到大纲');return edit('image_reference',{asset_id:''});});
    const available=(state.assets||[]).filter(row=>row.id!==slide?.asset_id);
    if(available.length){
      const reuse=field('复用本演示稿授权图片',pane,'','select');const empty=make('option','选择已授权素材',reuse);empty.value='';
      for(const row of available){const option=make('option',row.name,reuse);option.value=row.id;}
      button('用于本页',pane,()=>{if(dirty)throw Error('请先保存文字输入到大纲');if(!reuse.value)throw Error('请选择已授权素材');return edit('image_reference',{asset_id:reuse.value});});
    }
  }
  function formValues(){
    const node=root.querySelector('[data-presentation-editor]');
    const layout=node.querySelector('[name=layout]').value;
    const lines=value=>value.split('\n').filter(row=>row.trim()).map(row=>row.trim());
    return {title:node.querySelector('[name=title]').value,layout,notes:node.querySelector('[name=notes]').value,
      points:['table','chart'].includes(layout)?[]:lines(node.querySelector('[name=points]').value),right:layout==='columns'?lines(node.querySelector('[name=right]').value):[],
      rows:['table','chart'].includes(layout)?node.querySelector('[name=rows]').value.split('\n').map(row=>row.split('\t')):[],
      chart_unit:layout==='chart'?node.querySelector('[name=chart_unit]').value:'',chart_reviewed:layout==='chart'&&node.querySelector('[name=chart_reviewed]').checked};
  }
  function controls(){
    if(!root||!state?.spec)return;
    root.querySelector('[data-presentation-title]').textContent=state.spec.title+' · 草稿 '+state.spec.revision;
    root.querySelector('[data-presentation-build]').disabled=operation||state.busy||state.image_busy||state.uncertain;
    root.querySelector('[data-presentation-export]').disabled=operation||state.busy||state.uncertain||dirty||state.preview_revision!==state.spec.revision;
    root.querySelector('[data-presentation-build]').classList.toggle('presentation-primary',state.preview_revision!==state.spec.revision);
    root.querySelector('[data-presentation-export]').classList.toggle('presentation-primary',state.preview_revision===state.spec.revision&&!dirty);
    const source=root.querySelector('[data-presentation-source]');source.textContent=state.source_status==='old_draft'?'来源草稿已改变 · 当前使用旧快照':state.source_status==='unavailable'?'来源已失效 · 请重新关联':state.source.name+' · '+state.source.included_blocks+' / '+state.source.total_blocks+' 段';
    if(state.source.limited)source.textContent+=' · 部分材料未纳入';
    const stages={outline:'大纲待预览',building:'正在生成并渲染',ready:'实际页面已生成',failed:state.error,empty:'请选择材料'};
    root.querySelector('[data-presentation-stage]').textContent=state.uncertain?'结果不确定，请核对目标文件后重启':stages[state.status]||state.status;
  }
  async function showPreview(){
    if(!root||!state?.spec)return;
    const image=root.querySelector('[data-presentation-image]'),caption=root.querySelector('[data-presentation-caption]');
    if(state.preview_revision!==state.spec.revision){image.hidden=true;image.removeAttribute('src');caption.textContent=state.busy?'正在生成实际页面…':'大纲已改变，请生成实际预览';previewKey='';return;}
    const key=state.spec.id+':'+state.spec.revision+':'+selected;
    if(previewKey===key)return;
    const mine=epoch,revision=state.spec.revision,id=selected;
    try{
      const result=await call('presentation_preview',state.spec.id,revision,id);
      if(mine!==epoch||!root||state.spec.revision!==revision||selected!==id)return;
      if(!/^data:image\/png;base64,/.test(result.image))throw Error('预览格式未通过校验');
      image.src=result.image;image.alt='第 '+(state.spec.slides.findIndex(row=>row.id===id)+1)+' 页实际渲染';image.hidden=false;
      caption.textContent='实际 PPTX 页面 · LibreOffice / PDFium';previewKey=key;
    }catch(error){if(!error.stale)message(error.message);}
  }
  async function edit(action,fields={}){
    if(operation)return;
    if(action!=='update'&&dirty&&!await ask('放弃本页未保存的输入？'))return;
    operation=true;controls();const previousInput=action==='update'?JSON.stringify(formValues()):null;
    try{
      const previousIds=new Set(state.spec.slides.map(row=>row.id));
      const result=await call('presentation_edit',state.spec.id,action==='update'?formRevision:state.spec.revision,action,selected,fields);
      if(!root)return;
      const newerInput=action==='update'&&JSON.stringify(formValues())!==previousInput;
      state=result;if(!state.spec.slides.some(row=>row.id===selected))selected=state.spec.slides[0].id;
      if(action==='add')selected=state.spec.slides.find(row=>!previousIds.has(row.id)).id;
      dirty=newerInput;renderList();if(!newerInput)renderEditor();else{formRevision=state.spec.revision;message('本页已保存，仍有新的输入未保存。');}
      await showPreview();
    }finally{operation=false;controls();}
  }
  function renderList(){
    const list=root.querySelector('[data-presentation-slides]');list.replaceChildren();
    state.spec.slides.forEach((slide,index)=>{
      const choose=button((index+1)+'. '+slide.title,list,async ()=>{
        if(operation)return;if(dirty&&!await ask('放弃本页未保存的输入？'))return;
        selected=slide.id;dirty=false;renderList();renderEditor();showPreview();
      });choose.dataset.slide=slide.id;choose.setAttribute('aria-current',selected===slide.id?'page':'false');
    });
  }
  function renderEditor(){
    const pane=root.querySelector('[data-presentation-editor]');pane.replaceChildren();
    const slide=currentSlide();if(!slide)return;formRevision=state.spec.revision;
    make('h3','编辑本页',pane);
    const title=field('页面标题',pane,slide.title);title.name='title';title.maxLength=44;
    const layout=field('页面版式',pane,'','select');layout.name='layout';
    for(const [value,label] of Object.entries(names)){const item=make('option',label,layout);item.value=value;}layout.value=slide.layout;
    const points=field('要点（每行一条）',pane,slide.points.join('\n'),'textarea');points.name='points';points.rows=5;
    const right=field('右栏要点',pane,slide.right.join('\n'),'textarea');right.name='right';right.rows=5;
    const rows=field('表格（列用Tab分隔，行用换行分隔）',pane,slide.rows.map(row=>row.join('\t')).join('\n'),'textarea');rows.name='rows';rows.rows=6;
    const unit=field('图表单位',pane,slide.chart_unit||'');unit.name='chart_unit';unit.maxLength=20;unit.placeholder='例如 ms、% 或百分点';
    const reviewLabel=make('label','',pane);reviewLabel.className='presentation-material-option';const review=make('input','',reviewLabel);review.type='checkbox';review.name='chart_reviewed';review.checked=Boolean(slide.chart_reviewed);make('span','我已核对本页分类、数值和单位',reviewLabel);
    const chartHint=make('p','单系列柱状图：两列含表头，最多5项。单位单独填写；缺失值不按0绘制。',pane);chartHint.className='presentation-muted';
    const assets=make('section','',pane);assets.className='presentation-assets';assets.dataset.presentationAssets='';renderImageControls();
    rows.oninput=unit.oninput=()=>{review.checked=false;};
    const extras=make('details','',pane);make('summary','讲稿与来源',extras);
    const notes=field('讲稿',extras,slide.notes,'textarea');notes.name='notes';notes.maxLength=1000;notes.rows=4;
    make('p',slide.origin==='source_excerpt'?'材料摘录 · 未独立核验':'用户编辑 · 来源关系未核验',extras);
    for(const id of slide.source_ids){const source=state.source.segments.find(row=>row.id===id);if(source){make('p',source.text,extras);if(source.rows)make('pre',source.rows.map(cells=>cells.join(' | ')).join('\n'),extras);}}
    const visibility=()=>{points.parentElement.hidden=['table','chart'].includes(layout.value);right.parentElement.hidden=layout.value!=='columns';rows.parentElement.hidden=!['table','chart'].includes(layout.value);assets.hidden=layout.value!=='image';for(const item of [unit.parentElement,reviewLabel,chartHint])item.hidden=layout.value!=='chart';};visibility();
    pane.oninput=event=>{if(!event.target.name)return;dirty=true;controls();};
    layout.onchange=()=>{dirty=true;review.checked=false;visibility();controls();};
    button('保存本页到大纲',pane,()=>edit('update',formValues()));
    const actions=make('div','',pane);actions.className='presentation-actions';
    button('上移',actions,()=>edit('move',{direction:'up'}));button('下移',actions,()=>edit('move',{direction:'down'}));
    button('删除本页',actions,async ()=>{if(await ask('从大纲移除此页？不会删除已有PPTX。'))return edit('delete');});
    controls();
  }
  async function build(){
    if(operation||state.busy)return;if(dirty)throw Error('请先保存本页输入到大纲');
    operation=true;controls();
    try{state=await call('presentation_build',state.spec.id,state.spec.revision);controls();await poll();}
    finally{operation=false;controls();}
  }
  async function exportDeck(){
    if(operation||dirty)return;
    operation=true;controls();
    try{
      const allowOld=state.source_status==='old_draft'&&await ask('来源草稿已改变。是否明确继续使用本次捕获的旧快照？');
      if(state.source_status==='old_draft'&&!allowOld)return;
      const result=await call('presentation_plan',state.spec.id,state.spec.revision,Boolean(allowOld));if(result.cancelled)return;
      const plan=result.plan,node=modal('确认发布演示稿',root.querySelector('[data-presentation-export]'));
      let pending=true,inflight=false,finished=false;
      node.addEventListener('close',()=>{operation=false;controls();},{once:true});
      make('p',plan.pages+' 页 · 草稿 '+plan.revision+' · 可编辑 PPTX 新文件',node);
      make('p','备份来源与大纲、校验并记录审计。原文档保持不变。',node);
      const details=make('details','',node);make('summary','目标位置与限制',details);make('p',plan.path,details);make('p','已通过指定渲染器；不代表PowerPoint人工质量通过。',details);
      const status=make('p','',node);status.setAttribute('role','alert');
      const actions=make('div','',node);actions.className='presentation-actions';
      const cancel=button('取消',actions,async()=>{
        if(inflight)return;if(finished){node.close();return;}
        inflight=true;pending=false;confirm.disabled=cancel.disabled=true;
        try{await call('presentation_confirm',plan.id,false);node.close();message('已取消发布，未创建文件。');}
        catch(error){if(!error.stale)status.textContent=error.message;finished=true;cancel.textContent='关闭';}
        finally{inflight=false;cancel.disabled=false;}
      });
      const confirm=button('确认发布PPTX',actions,async()=>{
        if(!pending||inflight||finished)return;inflight=true;pending=false;confirm.disabled=cancel.disabled=true;
        try{
          const receipt=await call('presentation_confirm',plan.id,true);
          state=receipt.state;
          if(receipt.result.status!=='verified'){status.textContent=receipt.result.status==='uncertain'?'结果不确定，请核对目标文件；不会自动重试':'发布被拒绝，请重新核对计划与目录';return;}
          node.close();if(root)renderOutputs();message('');
        }catch(error){if(!error.stale)status.textContent=error.message;}
        finally{inflight=false;finished=true;cancel.disabled=false;cancel.textContent='关闭';controls();}
      });confirm.className='presentation-primary';
      node.addEventListener('cancel',event=>{event.preventDefault();if(!inflight)cancel.click();});
    }finally{if(!document.querySelector('.presentation-dialog[open]'))operation=false;controls();}
  }
  function renderOutputs(){
    const node=root.querySelector('[data-presentation-outputs]'),signature=JSON.stringify(state.outputs||[]);
    if(node.dataset.signature===signature)return;node.dataset.signature=signature;node.replaceChildren();
    for(const output of state.outputs||[]){
      const item=make('section','',node);make('strong',(output.status==='verified'?'已发布 presentation.pptx':'此副本已撤销')+' · 草稿 '+output.revision,item);
      if(output.status!=='verified')continue;
      button('打开位置',item,()=>call('presentation_saved_action',output.id,'open'));
      button('撤销本次发布',item,async()=>{
        if(operation||!root)return;
        const mine=epoch,id=state.spec.id;
        operation=true;controls();
        try{
          await call('presentation_saved_action',output.id,'rollback_plan');
          if(!root||mine!==epoch||state.spec.id!==id)return;
          const confirm=await ask('仅移除本次生成的PPTX，保留来源和审计；确认撤销？')===true;
          if(!root||mine!==epoch||state.spec.id!==id)return;
          const result=await call('presentation_saved_action',output.id,'rollback',confirm);
          message(result.result.status==='rolled_back'?'本次发布已撤销':result.result.status==='uncertain'?'撤销结果不确定，请核对目录':'已取消或拒绝撤销，请核对记录');await poll();
        }finally{if(mine===epoch){operation=false;controls();}}
      });
      const details=make('details','',item);make('summary','文件路径',details);make('p',output.path,details);
    }
  }
  async function poll(){
    if(!root||polling||operation&&!state?.busy)return;
    polling=true;const mine=epoch,id=state?.spec?.id;
    try{
      const result=await call('presentation_state');if(!root||mine!==epoch||id!==state?.spec?.id||result.spec?.id!==id)return;
      if(result.spec.revision<state.spec.revision)return;
      const changed=result.spec.revision!==state.spec.revision;state=result;
      if(changed&&!dirty){renderList();renderEditor();}else if(changed&&dirty)message('大纲已有更新，当前输入保留；保存时将重新核对版本。');
      controls();renderOutputs();await showPreview();
    }finally{polling=false;}
  }
  async function close(){
    if(operation)return;if(dirty&&!await ask('放弃本页未保存的输入？已保存大纲仍在本窗口。'))return;
    epoch++;clearInterval(timer);timer=null;root?.close();root=null;dirty=false;previewKey='';
  }
  function mount(){
    if(root?.isConnected){renderList();renderEditor();controls();showPreview();return;}
    const previous=returnFocus||document.activeElement;root=make('dialog','',document.body);const node=root;root.className='presentation-workspace';root.setAttribute('aria-label','演示稿工作区');
    const header=make('header','',root);make('strong','',header).dataset.presentationTitle='';
    button('新建',header,async ()=>{if(dirty&&!await ask('放弃本页未保存输入并新建？'))return;createSource(header.querySelector('button'));});
    button('新增一页',header,()=>edit('add'));
    const preview=button('生成实际预览',header,build);preview.dataset.presentationBuild='';preview.className='presentation-primary';
    const publish=button('发布PPTX',header,exportDeck);publish.dataset.presentationExport='';
    button('关闭',header,close);
    const info=make('div','',root);info.className='presentation-info';make('span','',info).dataset.presentationSource='';make('span','',info).dataset.presentationStage='';
    const work=make('div','',root);work.className='presentation-work';
    make('nav','',work).dataset.presentationSlides='';
    const center=make('section','',work);center.className='presentation-preview';const image=make('img','',center);image.dataset.presentationImage='';image.hidden=true;make('p','',center).dataset.presentationCaption='';
    make('aside','',work).dataset.presentationEditor='';
    const footer=make('footer','',root);make('p','',footer).dataset.presentationStatus='';root.querySelector('[data-presentation-status]').setAttribute('role','status');
    make('div','',footer).dataset.presentationOutputs='';
    const details=make('details','',footer);make('summary','来源与草稿说明',details);make('p','本轮是窗口内临时大纲；关闭面板保留，重启不恢复。已发布PPTX与来源/大纲快照保存在所选目录。资料摘录或人工编辑不代表事实核验。',details);
    root.addEventListener('cancel',event=>{event.preventDefault();close();});root.addEventListener('close',()=>{node.remove();if(root===node)root=null;if(previous?.isConnected)previous.focus();});root.showModal();
    renderList();renderEditor();controls();renderOutputs();showPreview();
    timer=setInterval(()=>poll().catch(error=>{if(!error.stale)message(error.message);}),1000);
  }
  async function open(){
    if(root?.isConnected){root.focus();return;}
    if(sourceNode?.isConnected){sourceNode.focus();return;}if(opening)return;opening=true;
    try{
      returnFocus=document.activeElement?.closest('details')?.querySelector('summary')||document.activeElement;
      state=await call('presentation_state');
      if(!state.spec)return createSource(returnFocus);
      selected=state.spec.slides.some(row=>row.id===selected)?selected:state.spec.slides[0].id;mount();
    }finally{opening=false;}
  }
  window.OfficePresentations={open};
  window.addEventListener('officeauth:unauthenticated',()=>{epoch++;clearInterval(timer);timer=null;root?.remove();for(const node of dialogs)node.remove();dialogs.clear();root=null;sourceNode=null;state=null;dirty=false;operation=false;previewKey='';returnFocus=null;});
})();
