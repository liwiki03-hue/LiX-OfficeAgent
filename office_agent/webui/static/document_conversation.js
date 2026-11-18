(() => {
  'use strict';
  const rows=new Map(),nodes=new Map(),timings=[],paintTimings=new Map();
  let context=null,connected=true,handlers={},paged=null;
  const terminal=new Set(['complete','failed','stopped','uncertain','interrupted']);
  const make=(tag,text,parent)=>{const node=document.createElement(tag);node.textContent=text||'';parent?.append(node);return node;};
  function afterVisiblePaint(element, callback, onFrame){
    requestAnimationFrame(()=>requestAnimationFrame(()=>{
      onFrame?.();
      if(document.hidden||!element.isConnected||element.checkVisibility?.({opacityProperty:true,visibilityProperty:true})===false)return;
      const rect=element.getBoundingClientRect();
      const left=Math.max(0,rect.left),right=Math.min(innerWidth,rect.right),top=Math.max(0,rect.top),bottom=Math.min(innerHeight,rect.bottom);
      if(left>=right||top>=bottom)return;
      const x=(left+right)/2;
      if([top+2,(top+bottom)/2,bottom-2].some(y=>element.contains(document.elementFromPoint(x,y))))callback(performance.now());
    }));
  }
  function visibleAnswer(row,node){
    const point=paintTimings.get(row.id);
    const final=row.status==='complete'&&Boolean(row.answer?.trim()||row.intent==='edit'&&row.results?.length);
    const preview=node.previewBody&&!node.preview.hidden&&node.preview.dataset.kind==='current';
    if(!point||point.final_answer_ms!==undefined||node.measuring||!final&&(!preview||point.first_effective_ms!==undefined))return;
    const element=final?node.text:node.previewBody,signature=node.signature;
    node.measuring=true;
    afterVisiblePaint(element,at=>{
      if(node.signature!==signature)return;
      if(point.first_effective_ms===undefined){
        point.first_effective_ms=at;point.first_effective_source=final?'final':'proposal_preview';
        timings.push({id:row.id,phase:'first_effective_visible',elapsed_ms:at-point.request_ms});
      }
      if(final){
        point.final_answer_ms=at;
        timings.push({id:row.id,phase:'final_answer_visible',elapsed_ms:at-point.request_ms});
      }
      if(timings.length>256)timings.splice(0,timings.length-256);
      report(row.id);
    },()=>{
      node.measuring=false;
      if(node.signature!==signature&&nodes.get(row.id)===node)visibleAnswer(rows.get(row.id),node);
    });
  }
  function report(id){
    const api=window.OfficeAuth?.wrapApi(window.pywebview?.api);
    if(typeof api?.conversation_observe!=='function')return;
    for(const point of timings.filter(value=>value.id===id&&!value.reported)){
      point.reported=true;api.conversation_observe(point.id,point.phase,point.elapsed_ms).catch(()=>{});
    }
  }
  function follow(force=false,wasNear=false,priorTop=null){
    const scroll=document.getElementById('floating-scroll');if(!scroll)return;
    const top=priorTop??scroll.scrollTop,near=wasNear||scroll.scrollHeight-scroll.scrollTop-scroll.clientHeight<72;
    if(!force&&!near){const latest=document.getElementById('floating-latest');if(latest)latest.hidden=false;return;}
    requestAnimationFrame(()=>{
      if(force||Math.abs(scroll.scrollTop-top)<2){scroll.scrollTop=scroll.scrollHeight;const latest=document.getElementById('floating-latest');if(latest)latest.hidden=true;}
    });
  }
  function mount(row){
    const list=document.getElementById('floating-timeline');if(!list)return null;
    let node=nodes.get(row.id);
    if(node?.user.isConnected&&node.assistant.isConnected)return node;
    const user=make('article','',list);user.className='floating-message user';user.dataset.clientId=row.id;user.setAttribute('aria-label','用户的问题');
    make('p',row.question,user);
    const assistant=make('article','',list);assistant.className='floating-message assistant conversation-turn';assistant.dataset.turnId=row.id;assistant.setAttribute('aria-label','LiX 回复');
    const text=make('p','',assistant);text.className='conversation-body';text.setAttribute('role','status');text.setAttribute('aria-live','polite');text.setAttribute('aria-atomic','true');
    node={user,assistant,text};nodes.set(row.id,node);return node;
  }
  function render(row){
    const scroll=document.getElementById('floating-scroll'),priorTop=scroll?.scrollTop,stick=scroll&&scroll.scrollHeight-scroll.scrollTop-scroll.clientHeight<72;
    const node=mount(row);if(!node)return;
    const undo=row.undo||{},preview=row.preview,publication=row.publication;
    const editor=paged?.document_id===row.document_id?paged:null;
    const proposal=preview&&editor?.proposals?.find(value=>value.id===preview.id);
    const previewState=preview&&(editor||publication?.reason==='draft_advanced')&&(!terminal.has(row.status)||publication?.reason==='draft_advanced')
      ? proposal&&['pending','adopted'].includes(proposal.status)
        &&editor.draft_revision===(proposal.applied_revision??preview.draft_revision)
        &&editor.instruction_revision===preview.instruction_revision?'current':'previous'
      : '';
    const canPreview=publication?.preview_latest_draft&&editor?.publication_result?.id===publication.id;
    const editingDuringValidation=publication?.reason==='validating'&&editor?.publishing
      &&editor.draft_revision>publication.draft_revision?editor.draft_revision:null;
    const signature=JSON.stringify([row.sequence,row.status,row.text,row.local,undo,previewState,proposal?.status,canPreview,editingDuringValidation]);
    if(node.signature!==signature){
      node.signature=signature;
      if(row.local){node.receipt??=make('small','',node.user);node.receipt.textContent=row.status==='sending'?'正在发送':row.status==='failed'?'未受理':'等待核对回执';}
      else{node.receipt?.remove();node.receipt=null;}
      let text=row.text||row.answer||'消息还在等待应用接收';
      if(row.stage==='proposal'&&proposal?.status==='adopted'&&previewState==='current')text='修改已进入草稿，待验证副本。';
      if(editingDuringValidation)text+='\n正在处理第 '+publication.draft_revision+' 版，第 '+editingDuringValidation+' 版仍是草稿。';
      if(undo.message)text=['ready','blocked'].includes(undo.status)?text+'\n'+undo.message:undo.message;
      if(!connected&&!row.local&&!terminal.has(row.status))text+='\n暂时没有新的进度，可以重新连接查看。';
      if(node.text.textContent!==text)node.text.textContent=text;
      if(previewState){
        if(!node.preview){
          node.preview=make('div','',node.assistant);node.preview.className='conversation-preview';
          node.assistant.insertBefore(node.preview,node.text);
          node.preview.setAttribute('aria-live','polite');
          node.previewLabel=make('small','',node.preview);
          node.previewBody=make('p','',node.preview);node.previewBody.className='conversation-body';
        }
        node.preview.hidden=false;node.preview.dataset.kind=previewState;
        node.previewLabel.textContent=previewState==='current'?'预览':'上一版预览';
        node.previewBody.textContent=preview.text;
      }else if(node.preview)node.preview.hidden=true;
      node.assistant.dataset.state=row.status;
      if(terminal.has(row.status)){node.pending?.remove();node.pending=null;}
      const canRetry=publication?.reason!=='draft_advanced'&&row.failure_code!=='number_conflict'&&['failed','interrupted','uncertain'].includes(row.status);
      const action=undo.available?'undo':undo.recoverable?'reconcile':canRetry?'retry':'';
      const actionKey=JSON.stringify([action,Boolean(undo.busy),canPreview?publication.id:null]);
      if(node.actionKey!==actionKey){
        node.actionKey=actionKey;node.actions?.remove();node.actions=null;
        if(action||canPreview){
          const actions=make('div','',node.assistant);actions.className='conversation-actions';node.actions=actions;
          if(canPreview){
            const previewButton=make('button','预览最新草稿',actions);previewButton.type='button';
            previewButton.onclick=async()=>{previewButton.disabled=true;try{await handlers.previewDraft?.(publication,previewButton);}finally{if(previewButton.isConnected)previewButton.disabled=false;}};
          }
          if(action){
          const button=make('button','',actions);button.type='button';button.disabled=Boolean(undo.busy);
          if(action==='undo'){
            const icon=make('span','',button);icon.className='agent-icon agent-icon-undo';icon.setAttribute('aria-hidden','true');make('span','撤销本次修改',button);button.setAttribute('aria-label','撤销本次修改');button.dataset.undoTurn=row.id;
            button.onclick=async()=>{button.disabled=true;try{await handlers.undo?.(row.id);}finally{if(button.isConnected)button.disabled=false;}};
          }else if(action==='reconcile'){
            button.textContent='核对撤销结果';button.dataset.undoTurn=row.id;
            button.onclick=async()=>{button.disabled=true;try{await handlers.undo?.(row.id);}finally{if(button.isConnected)button.disabled=false;}};
          }else{
            button.textContent=row.status==='uncertain'?'核对已有结果':'从此处恢复';button.dataset.retryTurn=row.id;
            button.onclick=async()=>{button.disabled=true;try{await handlers.retry?.(row.id,row.local);}finally{if(button.isConnected)button.disabled=false;}};
          }
          }
        }
      }
    }
    if(!terminal.has(row.status)){
      const seconds=Math.floor(Date.now()/1000-row.created);
      if(seconds>=5){node.elapsed??=make('small','',node.assistant);node.elapsed.className='conversation-elapsed';node.elapsed.textContent='已用 '+seconds+' 秒';}
    }else{node.elapsed?.remove();node.elapsed=null;}
    if(stick)follow(false,true,priorTop);
    visibleAnswer(row,node);
  }
  function merge(row){
    if(!row||typeof row.id!=='string'||typeof row.question!=='string'||!Number.isInteger(row.sequence))return;
    const previous=rows.get(row.id);
    if(previous&&!previous.local&&previous.sequence>row.sequence)return;
    if(previous&&!previous.local&&previous.sequence===row.sequence&&JSON.stringify(previous.undo)===JSON.stringify(row.undo)){render(previous);return;}
    const pendingAt=previous?.sentAt;
    rows.set(row.id,{...row,local:false,sentAt:pendingAt,acknowledged:previous?.acknowledged});render(rows.get(row.id));
    const step=row.steps?.at(-1);if(step?.at&&step.at!==previous?.steps?.at(-1)?.at)timings.push({id:row.id,phase:'stage_visible:'+step.stage,elapsed_ms:Math.max(0,Date.now()-step.at*1000)});
    if(pendingAt&&!previous?.acknowledged){timings.push({id:row.id,phase:'accepted_visible',elapsed_ms:performance.now()-pendingAt});rows.get(row.id).acknowledged=true;}
    if(timings.length>256)timings.splice(0,timings.length-256);report(row.id);
  }
  function echo(message){
    if(rows.has(message.id)){const prior=rows.get(message.id);if(prior.local){prior.status='sending';prior.text='我会先核对当前文档和这条消息，再继续处理。';}render(prior);follow(true);return;}
    const row={id:message.id,question:message.text,document_id:message.document,request:{...message},status:'sending',stage:'sending',text:'我会先核对当前文档和这条消息，再继续处理。',created:Date.now()/1000,sequence:0,steps:[],local:true,sentAt:performance.now()};
    paintTimings.set(row.id,{id:row.id,request_ms:row.sentAt});
    if(paintTimings.size>160)paintTimings.delete(paintTimings.keys().next().value);
    rows.set(row.id,row);render(row);follow(true);
    requestAnimationFrame(()=>requestAnimationFrame(()=>{
      const point=paintTimings.get(row.id),node=nodes.get(row.id);if(!point||!node?.user.isConnected)return;
      const at=performance.now();point.question_echo_ms??=at;timings.push({id:row.id,phase:'local_echo',elapsed_ms:at-row.sentAt});
    }));
    setTimeout(()=>{const current=rows.get(row.id);if(current?.local&&current.status==='sending'){current.text='消息还在等待应用接收';render(current);}},1800);
  }
  function failed(id,text,uncertain=false){const row=rows.get(id);if(!row||!row.local&&terminal.has(row.status))return;row.status=uncertain?'interrupted':'failed';row.text=text;render(row);}
  function reset(){for(const node of nodes.values()){node.user.remove();node.assistant.remove();}rows.clear();nodes.clear();paintTimings.clear();context=null;paged=null;}
  function restore(state,editor){
    if(!state)return;
    if(context&&context!==state.context_id&&state.context_id)reset();
    context=state.context_id||context;
    paged=editor??null;
    for(const row of state.turns||[])merge(row);
    for(const row of rows.values())if(row.local)render(row);
  }
  function measureVisibleRows(){for(const [id,node] of nodes)visibleAnswer(rows.get(id),node);}
  document.addEventListener('scroll',measureVisibleRows,true);
  document.addEventListener('visibilitychange',measureVisibleRows);
  setInterval(()=>{for(const row of rows.values())if(!terminal.has(row.status))render(row);},1000);
  window.OfficeConversation={echo,merge,restore,failed,configure:value=>{handlers=value;},
    undoMessage:id=>rows.get(id)?.undo?.message,
    undoError:(id,message)=>{const row=rows.get(id);if(row){row.undo={...(row.undo||{}),status:'blocked',available:false,busy:false,message};render(row);}},
    painted:()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))),afterVisiblePaint,
    connection:value=>{connected=value;for(const row of rows.values()){const node=nodes.get(row.id);if(node)node.signature=null;render(row);}},
    proposalHost:id=>{const node=nodes.get(id),row=rows.get(id);if(!node||!row||terminal.has(row.status))return null;node.pending??=make('div','',node.assistant);node.pending.className='conversation-pending';return node.pending;},
    pending:()=>[...rows.values()].some(row=>!terminal.has(row.status)&&!['clarify','reference','scope','deferred','choices'].includes(row.stage)),
    has:id=>rows.has(id),message:id=>rows.get(id)?.local&&rows.get(id).request?{...rows.get(id).request}:null,timings:()=>timings.slice(-160),paintTimings:()=>Array.from(paintTimings.values(),value=>({...value})),reset};
})();
