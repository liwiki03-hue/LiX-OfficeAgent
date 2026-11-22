(() => {
  'use strict';
  let engine, port, revision = 0, instruction = 0, ready = false, applying = null, totalPages = 0, editable = false;
  let lastBodyDocument, bodyStyle = null, tableAction = null, composing = false, tableFailure = null;
  let unsafeEdit=null;
  const $ = id => document.getElementById(id);
  const send = (type, value) => port?.postMessage({type, value});
  const editor = () => engine?.activeEditor;
  function paragraphText(node){
    let text='';
    node.descendants(child=>{
      if(child.isText){text+=child.text;return false;}
      if(child.type.name==='tab'){text+='\t';return false;}
      if(['hardBreak','lineBreak'].includes(child.type.name)){if(!['page','column'].includes(child.attrs.lineBreakType))text+='\n';return false;}
    });
    return text;
  }
  function rows() {
    const result = [];
    editor().state.doc.descendants((node, position) => {
      if (node.type.name === 'paragraph') {
        let from=null,to=null;
        node.descendants((child,offset)=>{if(child.isText||child.type.name==='tab'||child.type.name==='lineBreak'){if(from===null)from=position+1+offset;to=position+1+offset+child.nodeSize;return false;}});
        result.push({id:node.attrs.sdBlockId,text:paragraphText(node),position,from:from??position+1,to:to??position+1,style:node.attrs.paragraphProperties?.styleId||'Normal'});
      }
    });
    return result;
  }
  function state() {
    const values = rows();
    const pages = [...document.querySelectorAll('.superdoc-page[data-page-number]')];
    const scroll = $('page-scroll').getBoundingClientRect();
    const visible = pages.find(page => page.getBoundingClientRect().bottom > scroll.top + 50);
    return {revision, instruction, pages:totalPages || pages.length, page:Number(visible?.dataset.pageNumber || 1),
      characters:values.reduce((sum,row) => sum + [...row.text].length,0),
      outline:values.filter(row => /heading|title/i.test(row.style)).slice(0,100).map(row => ({id:row.id,text:row.text.slice(0,100)}))};
  }
  function selectionFormat(){
    const e=editor(),selection=e.state.selection,styles=new Set(),families=new Set(),sizes=new Set();
    e.state.doc.nodesBetween(selection.from,Math.max(selection.to,selection.from),(node)=>{
      if(node.type.name==='paragraph')styles.add(node.attrs.paragraphProperties?.styleId||'Normal');
      if(node.isText)for(const mark of node.marks){if(mark.type.name==='textStyle'){families.add(mark.attrs.fontFamily||'继承');sizes.add(mark.attrs.fontSize||'继承');}}
    });
    let mixed=$('paragraph-style').querySelector('option[value=""]');
    if(styles.size>1){if(!mixed){mixed=document.createElement('option');mixed.value='';mixed.disabled=true;mixed.textContent='混合样式';$('paragraph-style').prepend(mixed);}$('paragraph-style').value='';}
    else if(styles.size===1)$('paragraph-style').value=[...styles][0];
    $('selection-format').textContent=[families.size>1?'混合字体':'',sizes.size>1?'混合字号':''].filter(Boolean).join(' · ');
  }
  function layoutSelection(){
    if(!editable||composing)throw Error('请先进入可编辑状态并完成当前输入');
    const segments=editor().doc.selection.current().target?.segments||[];
    const available=rows(),ids=segments.map(item=>item.blockId);
    if(available.some(row=>ids.includes(row.id)&&editor().state.doc.resolve(row.position).depth!==0))throw Error('本入口仅调整普通正文段落，表格或特殊结构请保留原布局');
    const result=OfficeParagraphLayout.capture(editor().state.doc,available,ids,revision);
    result.selection={from:editor().state.selection.from,to:editor().state.selection.to};
    return result;
  }
  function layoutApply(selected,input){
    if(!editable||composing)throw Error('当前不能调整段落布局');
    if(selected.selection.from!==editor().state.selection.from||selected.selection.to!==editor().state.selection.to)throw Error('选区已改变，请重新打开布局设置');
    const transaction=OfficeParagraphLayout.prepare(editor(),selected,revision,input);
    editor().dispatch(transaction);
  }
  function layoutDialog(){
    const selected=layoutSelection(),dialog=document.createElement('dialog');
    dialog.className='paragraph-layout-dialog';dialog.setAttribute('aria-label','段落布局');
    const title=document.createElement('h2');title.textContent='段落布局';dialog.append(title);
    const note=document.createElement('p');note.textContent='所选 '+selected.targets.length+' 段 · 留空保留原值 · 单位：磅';dialog.append(note);
    const inputs={};
    for(const [key,name,min,max] of [['before','段前',0,96],['after','段后',0,96],['left','左缩进',0,144],['right','右缩进',0,144],['firstLine','首行（负数为悬挂）',-72,72]]){
      const label=document.createElement('label'),input=document.createElement('input');
      label.textContent=name;input.type='number';input.min=min;input.max=max;input.step='0.05';input.placeholder='保留';input.setAttribute('aria-label',name+'（磅）');
      label.append(input);dialog.append(label);inputs[key]=input;
    }
    const feedback=document.createElement('p');feedback.setAttribute('role','status');dialog.append(feedback);
    const adopt=document.createElement('button');adopt.type='button';adopt.textContent='采用到草稿';
    adopt.onclick=()=>{try{const input={};for(const [key,node] of Object.entries(inputs))if(node.value!=='')input[key]=Number(node.value);layoutApply(selected,input);dialog.close();}catch(error){feedback.textContent=error.message;}};
    const cancel=document.createElement('button');cancel.type='button';cancel.textContent='取消';cancel.onclick=()=>dialog.close();dialog.append(adopt,cancel);
    dialog.addEventListener('close',()=>{dialog.remove();$('paragraph-layout').focus();},{once:true});
    document.body.append(dialog);dialog.showModal();inputs.before.focus();
  }
  function tableContext(){
    const e=editor(),s=e?.state.selection;
    if(!s)return null;
    for(let depth=s.$from.depth;depth>0;depth--){
      if(s.$from.node(depth).type.name!=='table')continue;
      if(depth!==1||s.to>s.$from.after(depth))return null;
      return {node:s.$from.node(depth),position:s.$from.before(depth)};
    }
    if(s.empty&&s.$from.depth===0){
      if(s.$from.nodeBefore?.type.name==='table')return {node:s.$from.nodeBefore,position:s.from-s.$from.nodeBefore.nodeSize};
      if(s.$from.nodeAfter?.type.name==='table')return {node:s.$from.nodeAfter,position:s.from};
    }
    return null;
  }
  function tableActions(){
    const context=editable&&tableContext();
    tableAction=context?{document:editor().state.doc,revision,position:context.position,id:context.node.attrs.sdBlockId}:null;
    $('table-body-actions').hidden=!context;
  }
  function columnWidthSelection(){
    if(!editable||composing||unsafeEdit)throw Error('当前不能调整表格列宽');
    const e=editor(),context=tableContext(),selection=e.state.selection;
    if(!context)throw Error('请先将光标放入目标列的单元格');
    let cellDepth=0;
    for(let depth=selection.$from.depth;depth>0;depth--)if(selection.$from.node(depth).type.name==='tableCell'){cellDepth=depth;break;}
    if(cellDepth!==3||selection.to>selection.$from.after(cellDepth))throw Error('请在单个普通单元格中选择列宽位置');
    const selected=OfficeTableColumnLayout.capture(e.state.doc,context.position,selection.$from.index(cellDepth-1),revision);
    selected.selection={from:selection.from,to:selection.to};return selected;
  }
  function columnWidthApply(selected,value){
    if(!editable||composing||unsafeEdit)throw Error('当前不能调整表格列宽');
    const e=editor();
    if(selected.selection.from!==e.state.selection.from||selected.selection.to!==e.state.selection.to)throw Error('选区已改变，请重新打开列宽设置');
    const size=OfficeTableColumnLayout.validate(e,selected,revision,value),prior=revision;
    try{
      const result=e.doc.tables.setColumnWidth({nodeId:selected.node.attrs.sdBlockId,columnIndex:selected.column,widthPt:size.pt});
      if(result?.success!==true||revision!==prior+1)throw Error('列宽执行结果需要核对');
      return OfficeTableColumnLayout.verify(selected,e.state.doc,size);
    }catch(error){
      if(e.state.doc!==selected.doc&&!e.state.doc.eq(selected.doc)){
        unsafeEdit='列宽校验发现未解释变化，已阻止导出；保留当前草稿，请核对后重新打开。';throw Error(unsafeEdit);
      }
      throw Error(error.message+'；未发布副本，请重新核对草稿');
    }
  }
  function columnWidthDialog(){
    const selected=columnWidthSelection(),dialog=document.createElement('dialog');dialog.className='paragraph-layout-dialog';dialog.setAttribute('aria-label','表格列宽');
    const title=document.createElement('h2');title.textContent='第'+(selected.column+1)+'列宽度';dialog.append(title);
    const summary=document.createElement('p');summary.textContent=selected.rows+'行 · 调整整列并同步表格总宽，可能改变换行与分页';dialog.append(summary);
    const label=document.createElement('label');label.textContent='列宽（磅）';const input=document.createElement('input');input.type='number';input.min=24;input.max=360;input.step=.75;input.value=selected.current??'';input.setAttribute('aria-label','表格列宽（磅）');label.append(input);dialog.append(label);
    const feedback=document.createElement('p');feedback.setAttribute('role','status');dialog.append(feedback);
    input.oninput=()=>{try{const size=OfficeTableColumnLayout.width(Number(input.value)),old=selected.json.attrs.grid.reduce((sum,col)=>sum+col.col,0);feedback.textContent='表格总宽：'+old/20+' → '+(old-selected.json.attrs.grid[selected.column].col+size.twips)/20+'磅';}catch(error){feedback.textContent=error.message;}};
    const adopt=document.createElement('button');adopt.type='button';adopt.textContent='采用到草稿';adopt.onclick=()=>{try{if(input.value==='')throw Error('请填写明确列宽');const proof=columnWidthApply(selected,Number(input.value));dialog.close();$('frame-error').textContent=proof.summary;send('notice',{text:proof.summary});}catch(error){feedback.textContent=error.message;}};
    const cancel=document.createElement('button');cancel.type='button';cancel.textContent='取消';cancel.onclick=()=>dialog.close();dialog.append(adopt,cancel);
    dialog.addEventListener('close',()=>{dialog.remove();$('table-column-width').focus();},{once:true});document.body.append(dialog);dialog.showModal();input.focus();
  }
  function continueBody(side,expected){
    const e=editor();
    if(composing)throw Error('请先完成当前输入，再切换正文');
    if(!editable||!expected||expected.revision!==revision||expected.document!==e.state.doc)throw Error('表格目标已改变，请重新将光标放入当前表格');
    const context=tableContext(),node=e.state.doc.nodeAt(expected.position);
    if(!context||context.position!==expected.position||node!==context.node||node.attrs.sdBlockId!==expected.id)throw Error('表格位置已改变，请重新选择');
    const boundary=side==='after'?context.position+node.nodeSize:context.position;
    const point=e.state.doc.resolve(boundary),adjacent=side==='after'?point.nodeAfter:point.nodeBefore;
    let transaction=e.state.tr,target;
    if(adjacent?.type.name==='paragraph')target=side==='after'?boundary+1:boundary-adjacent.nodeSize+1;
    else{
      const paragraph=e.state.schema.nodes.paragraph.createAndFill({sdBlockId:crypto.randomUUID(),paragraphProperties:bodyStyle?{styleId:bodyStyle}:null});
      if(!paragraph)throw Error('当前文档不能创建普通正文段落');
      transaction=transaction.insert(boundary,paragraph);
      target=boundary+1;
    }
    if(!transaction.doc.resolve(target).parent.inlineContent)throw Error('当前正文位置不可编辑，请保留草稿并重新核对');
    transaction.setStoredMarks([]);
    e.dispatch(transaction.scrollIntoView());
    if(e.commands.setTextSelection({from:target,to:target})===false)throw Error('正文已准备，光标定位需要核对；请勿重复插入');
    e.focus();
    tableActions();
  }
  function changed(event) {
    if (!ready||!editable) return;
    if(event?.surface==='body'&&lastBodyDocument===editor().state.doc)return;
    lastBodyDocument=editor().state.doc;
    revision++;
    tableActions();
    send('changed', {revision, apply_token:applying});
    send('state', state());
  }
  async function open(value) {
    if (engine) throw Error('此编辑实例已经打开文档');
    revision = value.revision || 0; editable = false;
    document.body.dataset.theme=['light','dark'].includes(value.theme)?value.theme:'dark';
    const bytes = Uint8Array.from(atob(value.data), char => char.charCodeAt(0));
    const blob = new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    const Engine = typeof SuperDoc === 'function' ? SuperDoc : SuperDoc.SuperDoc;
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(Error('文档引擎打开超时；原稿未变')),60000);
      engine = new Engine({selector:'#editor', document:blob, documentMode:'viewing', allowSelectionInViewMode:true,
        telemetry:{enabled:false}, modules:{toolbar:{selector:'#engine-toolbar', responsiveToContainer:false,
          groups:{left:['undo','redo','fontFamily','fontSize','bold','italic','underline','color','highlight','textAlign','list','numberedlist','lineHeight']},
          texts:{bold:'加粗',italic:'斜体',underline:'下划线',undo:'撤销',redo:'重做'}}},
        onReady:() => { clearTimeout(timer); ready = true; resolve(); },
        onEditorUpdate:changed, onPaginationUpdate:event => {totalPages=event.totalPages; if(ready)send('state',state());},
        onContentError:() => {clearTimeout(timer);reject(Error('文档包含引擎无法解释的内容，不能编辑保存'));}
      });
    });
    const styleNames={Normal:'正文',Title:'标题',Subtitle:'副标题',ListParagraph:'列表段落'};
    bodyStyle=(value.styles||[]).some(style=>style.id==='Normal')?'Normal':null;
    for (const style of value.styles || []) { const option = document.createElement('option');option.value=style.id;option.textContent=styleNames[style.id]||(/^Heading[1-9]$/.test(style.id)?'标题 '+style.id.slice(7):style.name);$('paragraph-style').append(option); }
    const unconfirmed=(value.fonts?.fonts||[]).filter(item=>item.status==='unconfirmed').map(item=>item.family);
    if(unconfirmed.length)$('frame-error').textContent='未确认本机字体：'+unconfirmed.join('、')+'；可能使用替代字体，分页可能不同。没有自动下载字体。';
    $('paragraph-style').disabled = $('paragraph-after').disabled = $('paragraph-layout').disabled = !editable;
    editor().on('selectionUpdate',()=>{selectionFormat();tableActions();localizeToolbar();});
    selectionFormat();
    send('state', state());
    return {revision,capabilities:{engine:'SuperDoc',version:'1.47.2',editable}};
  }
  function localizeToolbar(){
    const names={undo:'撤销',redo:'重做',fontFamily:'字体',fontSize:'字号（磅）',bold:'加粗',italic:'斜体',underline:'下划线',color:'文字颜色',highlight:'突出显示',textAlign:'段落对齐',list:'项目符号',numberedlist:'编号列表',lineHeight:'行距'};
    for(const [id,label] of Object.entries(names)){
      const element=document.querySelector('[data-item="btn-'+id+'"]');
      if(!element)continue;
      const trigger=element.closest('[role="button"]')||element;
      trigger.setAttribute('aria-label',label);trigger.title=label;
      const input=element.querySelector('input');if(input)input.setAttribute('aria-label',label);
    }
  }
  new MutationObserver(localizeToolbar).observe($('engine-toolbar'),{childList:true,subtree:true});
  async function snapshot(value) {
    if(unsafeEdit)throw Error(unsafeEdit);
    if (value.revision !== revision) throw Error('草稿已更新，请重新核对');
    const captured = revision, manifest = rows().map(({id,text})=>({id,text}));
    const blob = await engine.export({triggerDownload:false});
    if (blob.size > 8*1024*1024) throw Error('导出超过8 MiB预算，请缩小文档');
    const data = await new Promise((yes,no) => {const reader=new FileReader();reader.onload=()=>yes(reader.result.split(',')[1]);reader.onerror=no;reader.readAsDataURL(blob);});
    if (revision !== captured) throw Error('导出时草稿已更新；没有保存旧内容，请重新预览');
    return {data,manifest,revision:captured};
  }
  function locate(id) {
    const row = rows().find(item=>item.id===id);
    if (!row) throw Error('目标段落不在当前草稿中');
    editor().commands.setTextSelection({from:row.from,to:row.to});
    editor().dispatch(editor().state.tr.scrollIntoView());
    return row;
  }
  function selection(options={}) {
    const context=tableContext(), active=editor().state.selection;
    if(context&&(options.table===true||!active.empty&&rows().filter(row=>row.from<active.to&&row.to>active.from).length>1)){
      const picked=rows().filter(row=>row.position>context.position&&row.position<context.position+context.node.nodeSize&&row.text&&(options.table===true||row.from<active.to&&row.to>active.from));
      if(!picked.length||picked.length>8)throw Error('表格引用限1—8个非空段落，请缩小范围');
      return {nodes:picked.map(row=>{const positions=batchPositions(row),start=options.table===true?0:positions.findIndex(position=>position>=active.from),end=options.table===true?row.text.length:positions.findIndex(position=>position>=active.to);return {node_id:row.id,start:start<0?row.text.length:start,end:end<0?row.text.length:end};}).filter(row=>row.end>row.start),whole_table:options.table===true};
    }
    if(options.table===true)throw Error('请先将光标放入表格，再引用整表');
    const value = editor().doc.selection.current();
    const segments = value.target?.segments;
    let segment=segments?.length===1?segments[0]:null;
    if(value.empty&&options.allowParagraph===true){
      const caret=editor().state.selection.from,row=rows().find(item=>caret>=item.from&&caret<=item.to);
      if(row&&row.text.trim())segment={blockId:row.id,range:{start:0,end:row.text.length}};
      else throw Error('请在要修改的段落中放置光标，或选择一段文字');
    }else if(value.empty||!segment)throw Error('首版AI修改请选择一个段落内的明确文字范围');
    const row=rows().find(item=>item.id===segment.blockId);
    if (!row) throw Error('首版AI仅支持正文或普通表格内文字；请重新选择');
    if(options.source===true){
      const values=rows();
      if(values.length>3000||values.reduce((size,item)=>size+item.text.length,0)>1000000)throw Error('请从原稿摘录选取');
      return {source_index:values.findIndex(item=>item.id===row.id),paragraphs:values.map(item=>item.text),
        start:segment.range.start,end:segment.range.end,expected:editor().state.doc.textBetween(active.from,active.to,'\n')};
    }
    const styles=new Set(),fonts=new Set(),sizes=new Set(),marks=new Set();
    editor().state.doc.nodesBetween(editor().state.selection.from,Math.max(editor().state.selection.from,editor().state.selection.to),node=>{
      if(node.type.name==='paragraph')styles.add(node.attrs.paragraphProperties?.styleId||'Normal');
      if(node.isText)for(const mark of node.marks){
        if(mark.type.name==='textStyle'){fonts.add(mark.attrs.fontFamily||'inherit');sizes.add(mark.attrs.fontSize||'inherit');}
        else marks.add(mark.type.name);
      }
    });
    return {node_id:segment.blockId,start:segment.range.start,end:segment.range.end,
      formats:{paragraph_styles:[...styles],fonts:[...fonts],sizes:[...sizes],marks:[...marks].sort()}};
  }
  const tableCommands=new Set(['addRowBefore','addRowAfter','deleteRow']);
  function runTableCommand(name) {
    if(!editable||composing)throw Error('当前文档只读或正在输入，不能修改表格结构');
    if(!tableCommands.has(name))throw Error('不支持的表格结构操作');
    const context=tableContext();
    if(!context)throw Error('请先将光标放入当前表格');
    const prior=revision,tableId=context.node.attrs.sdBlockId,command=editor()?.commands?.[name];
    if(typeof command!=='function')throw Error('当前引擎不支持'+name);
    if(command()===false||revision===prior)throw Error('表格结构未改变，请重新选择');
    if(revision!==prior+1)throw Error('表格结构结果需要核对；保留当前草稿，未写磁盘');
    tableActions();
    return {revision,prior_revision:prior,command:name,table_id:tableId};
  }
  function historyState() {
    const can=editor()?.can?.();
    return {canUndo:Boolean(can?.undo?.()),canRedo:Boolean(can?.redo?.())};
  }
  function historyStep(name) {
    if(!editable)throw Error('当前文档只读，不能撤销或重做');
    const prior=revision,command=editor()?.commands?.[name];
    if(typeof command!=='function')throw Error('当前引擎不支持'+name);
    command();
    if(revision===prior)throw Error(name==='undo'?'没有可撤销的编辑':'没有可重做的编辑');
    return {revision,prior_revision:prior};
  }
  function apply(value) {
    if (!editable || value.revision!==revision || value.instruction!==instruction) throw Error('草稿或要求已改变，旧建议未采用');
    if(Array.isArray(value.proposal.edits)){
      const cells=new Map(value.proposal.reference.cells.map(row=>[row.node_id,row]));
      const operations=value.proposal.edits.map(edit=>{const cell=cells.get(edit.node_id);if(!cell?.editable||edit.old!==cell.old)throw Error('表格建议目标无效');return {target_id:edit.node_id,start:cell.start,end:cell.end,expected_before:edit.old,after:edit.new};});
      return applyBatch({...value,operations});
    }
    const reference=value.proposal.reference, row=rows().find(item=>item.id===reference.node_id);
    if (!row || row.text.slice(reference.start,reference.end)!==value.proposal.old) throw Error('选区原文不匹配，未采用');
    const prior=revision;
    applying=value.token;
    try {
      if(reference.kind==='append'){
        const e=editor(),paragraph=e.state.doc.nodeAt(row.position),text=value.proposal.new.trim();
        if(e.state.doc.resolve(row.position).depth!==0||paragraph?.type.name!=='paragraph'||!text||text.length>6000)throw Error('追加位置或内容无效，未修改正文');
        const lines=text.split(/\r?\n+/).filter(line=>line.trim());
        if(lines.length>20)throw Error('一次追加最多20段，请缩小范围');
        const nodes=lines.map(line=>e.state.schema.nodes.paragraph.create(
          {sdBlockId:crypto.randomUUID(),paragraphProperties:bodyStyle?{styleId:bodyStyle}:null},e.state.schema.text(line)));
        e.dispatch(e.state.tr.insert(e.state.doc.content.size,nodes).scrollIntoView());
        if(revision!==prior+1)throw Error('追加结果需要核对；保留草稿，未写磁盘');
        const last=nodes[nodes.length-1];
        return {revision,prior_revision:prior,reference:{node_id:last.attrs.sdBlockId,start:0,end:paragraphText(last).length}};
      }
      const result=editor().doc.replace({target:{kind:'selection',start:{kind:'text',blockId:row.id,offset:reference.start},end:{kind:'text',blockId:row.id,offset:reference.end}},text:value.proposal.new});
      if (result?.success===false || revision!==prior+1) throw Error('采用结果需要核对；保留当前草稿，未写磁盘');
      return {revision,prior_revision:prior};
    } finally { applying=null; }
  }
  function batchPositions(row) {
    const node=editor().state.doc.nodeAt(row.position);
    if(!node||node.type.name!=='paragraph')throw Error('批量操作目标不是普通段落');
    const positions=new Array(row.text.length+1);
    let offset=0;
    node.descendants((child,childOffset)=>{
      const start=row.position+1+childOffset;
      if(child.isText){
        for(let index=0;index<=child.text.length;index++)positions[offset+index]=start+index;
        offset+=child.text.length;
        return false;
      }
      if(child.type.name==='tab'){
        positions[offset]=start;positions[offset+1]=start+1;offset++;return false;
      }
      if(child.type.name==='hardBreak'||child.type.name==='lineBreak'){
        if(['page','column'].includes(child.attrs.lineBreakType))return false;
        positions[offset]=start;positions[offset+1]=start+1;offset++;return false;
      }
      if(child.type.name==='run')return true;
      if(!child.textContent)return false;
      throw Error('批量操作暂不支持带文字的未知内联对象：'+child.type.name);
    });
    if(offset!==row.text.length)throw Error('批量操作无法稳定映射段落文字');
    for(let index=0;index<=row.text.length;index++)if(positions[index]===undefined)throw Error('批量操作无法稳定映射段落文字');
    return positions;
  }
  function applyBatch(value) {
    if(!editable||value.revision!==revision||value.instruction!==instruction)throw Error('草稿或要求已改变，旧批量操作未采用');
    const operations=value.operations;
    if(!Array.isArray(operations)||!operations.length||operations.length>8)throw Error('批量操作数量必须为1—8项');
    const byId=new Map(rows().map(row=>[row.id,row])),groups=new Map();
    for(const operation of operations){
      const row=byId.get(operation.target_id);
      if(!row||!Number.isSafeInteger(operation.start)||!Number.isSafeInteger(operation.end)||operation.start<0||operation.end<=operation.start||operation.end>row.text.length)throw Error('批量操作目标或区间无效');
      if(row.text.slice(operation.start,operation.end)!==operation.expected_before)throw Error('批量操作原文不匹配，未采用');
      if(!groups.has(row.id))groups.set(row.id,[]);
      groups.get(row.id).push(operation);
    }
    const entries=[];
    for(const [identifier,spans] of groups){
      spans.sort((left,right)=>left.start-right.start);
      for(let index=1;index<spans.length;index++)if(spans[index].start<spans[index-1].end)throw Error('批量操作区间重叠');
      entries.push({row:byId.get(identifier),positions:batchPositions(byId.get(identifier)),operations:spans});
    }
    entries.sort((left,right)=>right.row.position-left.row.position);
    const prior=revision,transaction=editor().state.tr;
    applying=value.token;
    try {
      for(const entry of entries)for(const operation of [...entry.operations].reverse())transaction.insertText(operation.after,entry.positions[operation.start],entry.positions[operation.end]);
      editor().dispatch(transaction.scrollIntoView());
      if(revision!==prior+1)throw Error('批量采用结果需要核对；保留当前草稿，未写磁盘');
      return {revision,prior_revision:prior,operations:operations.length};
    } finally { applying=null; }
  }
  const actions={open,snapshot,selection,apply,applyBatch,tableCommand:value=>runTableCommand(value.command),history:historyState,undo:()=>historyStep('undo'),redo:()=>historyStep('redo'),locate:value=>{locate(value.id);return state();},
    enable:value=>{editable=false;engine.setDocumentMode(value.editable===true?'editing':'viewing');lastBodyDocument=editor().state.doc;editable=value.editable===true;$('paragraph-style').disabled=$('paragraph-after').disabled=$('paragraph-layout').disabled=!editable;tableActions();queueMicrotask(localizeToolbar);return {editable};},
    instruction:value=>{instruction=value.revision;return {instruction};},
    paper:value=>{document.body.classList.toggle('paper-dark',value.dark);return {};},
    theme:value=>{document.body.dataset.theme=['light','dark'].includes(value.theme)?value.theme:'dark';return {};},
    chatSidebar:({closed})=>{document.body.classList.toggle('chat-sidebar-closed',closed);return {};},
    zoom:value=>{value.fit?engine.setZoomMode('fit-width'):engine.setZoom(value.percent);return {};},
    find:value=>{if(!value.text||value.text.length>200)throw Error('请输入1—200字');const row=rows().find(item=>item.text.includes(value.text));if(!row)throw Error('当前正文未找到文字');locate(row.id);return state();},
    close:()=>{ready=false;engine?.destroy();engine=null;return {};}};
  window.addEventListener('message', event => {
    if (event.source!==parent || event.data?.type!=='office-editor-connect' || !event.ports[0] || port) return;
    port=event.ports[0];
    port.onmessage=async event=>{
      const {id,action,value}=event.data || {};
      if(!Number.isSafeInteger(id) || !Object.hasOwn(actions,action))return;
      try {const result=await actions[action](value||{});port.postMessage({id,ok:true,result});}
      catch(error){port.postMessage({id,ok:false,error:String(error.message||'编辑操作失败').slice(0,240)});}
    };port.start();send('connected',{});
  });
  $('paragraph-style').addEventListener('change',()=>{
    if(!editable)return;
    try{
      const segments=editor().doc.selection.current().target?.segments||[];
      if(!segments.length)throw Error('请先将光标放入正文段落');
      const ids=new Set(segments.map(segment=>segment.blockId)),e=editor(),transaction=e.state.tr;
      for(const row of rows().filter(row=>ids.has(row.id))){
        const node=transaction.doc.nodeAt(row.position);
        if(node?.type.name!=='paragraph'||node.attrs.sdBlockId!==row.id)throw Error('目标段落已改变，请重新选择');
        transaction.setNodeMarkup(row.position,undefined,{...node.attrs,paragraphProperties:{...(node.attrs.paragraphProperties||{}),styleId:$('paragraph-style').value}});
      }
      e.dispatch(transaction);
    }catch(error){$('frame-error').textContent=error.message;}
  });
  $('paragraph-after').addEventListener('change',()=>{
    try {layoutApply(layoutSelection(),{after:Number($('paragraph-after').value)});
    }catch(error){$('frame-error').textContent=error.message;}
  });
  $('paragraph-layout').addEventListener('mousedown',event=>{if(event.button===0)event.preventDefault();});
  $('paragraph-layout').addEventListener('click',()=>{try{layoutDialog();}catch(error){$('frame-error').textContent=error.message;}});
  for(const side of ['before','after']){
    const button=$('table-body-'+side);
    button.addEventListener('mousedown',event=>{if(event.button===0)event.preventDefault();});
    button.addEventListener('click',()=>{
      const notice=$('frame-error');
      try{
        continueBody(side,tableAction);
        if(tableFailure&&notice.textContent===tableFailure.message)notice.textContent=tableFailure.previous;
        tableFailure=null;
      }catch(error){
        const previous=tableFailure&&notice.textContent===tableFailure.message?tableFailure.previous:notice.textContent;
        tableFailure={message:error.message,previous};notice.textContent=error.message;
      }
    });
  }
  for(const button of document.querySelectorAll('[data-table-command]')){
    button.addEventListener('mousedown',event=>{if(event.button===0)event.preventDefault();});
    button.addEventListener('click',()=>{try{runTableCommand(button.dataset.tableCommand);}catch(error){$('frame-error').textContent=error.message;}});
  }
  $('table-column-width').addEventListener('mousedown',event=>{if(event.button===0)event.preventDefault();});
  $('table-column-width').addEventListener('click',()=>{try{columnWidthDialog();}catch(error){$('frame-error').textContent=error.message;}});
  document.addEventListener('compositionstart',()=>{composing=true;});
  document.addEventListener('compositionend',()=>{composing=false;});
  document.addEventListener('keydown',event=>{
    if(!(event.ctrlKey||event.metaKey)||event.altKey||event.isComposing||composing)return;
    if(event.key.toLowerCase()==='s'){event.preventDefault();send('save',{});return;}
    if(event.shiftKey||!editable||!ready||!['Home','End'].includes(event.key)||event.target.closest('input,textarea,select,button,dialog'))return;
    const e=editor();if((!event.target.closest('#editor')&&!e.view.dom.contains(event.target))||!tableContext())return;
    const end=event.key==='End',node=end?e.state.doc.lastChild:e.state.doc.firstChild;
    if(!node?.isTextblock)return;
    const position=end?e.state.doc.content.size-node.nodeSize+1+node.content.size:1;
    if(!e.state.doc.resolve(position).parent.inlineContent)return;
    event.preventDefault();event.stopImmediatePropagation();
    e.commands.setTextSelection({from:position,to:position});e.dispatch(e.state.tr.setStoredMarks([]).scrollIntoView());e.view.focus();
  },true);
  document.addEventListener('paste',event=>{
    if(!event.clipboardData?.getData('text/html'))return;
    event.preventDefault();if(!editable)return;
    const html=event.clipboardData.getData('text/html'),plain=event.clipboardData.getData('text/plain');
    if(html.length>100000||plain.length>20000){send('notice',{text:'粘贴超过本次2万字预算；未截断插入，请分段粘贴。'});return;}
    const documentValue=new DOMParser().parseFromString(html,'text/html'),output=[];
    const markNames={STRONG:'bold',B:'bold',EM:'italic',I:'italic',U:'underline'};
    function inline(node,marks=[]){
      if(node.nodeType===Node.TEXT_NODE)return node.textContent?[{type:'text',text:node.textContent,...(marks.length?{marks}: {})}]:[];
      if(node.nodeType!==Node.ELEMENT_NODE||['SCRIPT','STYLE','IFRAME','OBJECT','IMG','SVG','MATH'].includes(node.tagName))return [];
      const next=markNames[node.tagName]?[...marks,{type:markNames[node.tagName]}]:marks;
      if(node.tagName==='BR')return [{type:'lineBreak'}];
      return [...node.childNodes].flatMap(child=>inline(child,next));
    }
    const blocks=[...documentValue.body.children].filter(node=>/^(P|DIV|LI|H[1-6])$/.test(node.tagName));
    if(blocks.length)for(const node of blocks)output.push({type:'paragraph',content:inline(node)});
    else output.push({type:'paragraph',content:inline(documentValue.body)});
    if(output.some(node=>node.content.length))editor().commands.insertContent(output);
  },true);
  document.addEventListener('click',event=>{const link=event.target.closest('a');if(link){event.preventDefault();send('notice',{text:'链接保留在文档中；编辑区不会自动访问外部地址。'});}},true);
  let scrollTimer;
  $('page-scroll').addEventListener('scroll',()=>{clearTimeout(scrollTimer);scrollTimer=setTimeout(()=>{if(ready)send('state',state());},100);},{passive:true});
})();
