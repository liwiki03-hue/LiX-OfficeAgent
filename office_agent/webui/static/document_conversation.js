(() => {
  'use strict';
  const rows=new Map(),nodes=new Map(),timings=[];
  let context=null,connected=true,handlers={};
  const terminal=new Set(['complete','failed','stopped','uncertain','interrupted']);
  const make=(tag,text,parent)=>{const node=document.createElement(tag);node.textContent=text||'';parent?.append(node);return node;};
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
    const undo=row.undo||{},signature=JSON.stringify([row.sequence,row.status,row.text,row.local,undo]);
    if(node.signature!==signature){
      node.signature=signature;
      if(row.local){node.receipt??=make('small','',node.user);node.receipt.textContent=row.status==='sending'?'正在发送':row.status==='failed'?'未受理':'等待核对回执';}
      else{node.receipt?.remove();node.receipt=null;}
      let text=row.text||row.answer||'消息还在等待应用接收';
      if(undo.message)text=['ready','blocked'].includes(undo.status)?text+'\n'+undo.message:undo.message;
      if(!connected&&!row.local&&!terminal.has(row.status))text+='\n暂时没有新的进度，可以重新连接查看。';
      if(node.text.textContent!==text)node.text.textContent=text;
      node.assistant.dataset.state=row.status;
      if(terminal.has(row.status)){node.pending?.remove();node.pending=null;}
      const action=undo.available?'undo':undo.recoverable?'reconcile':['failed','interrupted','uncertain'].includes(row.status)?'retry':'';
      const actionKey=action+':'+Boolean(undo.busy);
      if(node.actionKey!==actionKey){
        node.actionKey=actionKey;node.actions?.remove();node.actions=null;
        if(action){
          const actions=make('div','',node.assistant);actions.className='conversation-actions';node.actions=actions;
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
    if(!terminal.has(row.status)){
      const seconds=Math.floor(Date.now()/1000-row.created);
      if(seconds>=5){node.elapsed??=make('small','',node.assistant);node.elapsed.className='conversation-elapsed';node.elapsed.textContent='已用 '+seconds+' 秒';}
    }else{node.elapsed?.remove();node.elapsed=null;}
    if(stick)follow(false,true,priorTop);
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
    rows.set(row.id,row);render(row);follow(true);
    requestAnimationFrame(()=>requestAnimationFrame(()=>{timings.push({id:row.id,phase:'local_echo',elapsed_ms:performance.now()-row.sentAt});}));
    setTimeout(()=>{const current=rows.get(row.id);if(current?.local&&current.status==='sending'){current.text='消息还在等待应用接收';render(current);}},1800);
  }
  function failed(id,text,uncertain=false){const row=rows.get(id);if(!row||!row.local&&terminal.has(row.status))return;row.status=uncertain?'interrupted':'failed';row.text=text;render(row);}
  function reset(){for(const node of nodes.values()){node.user.remove();node.assistant.remove();}rows.clear();nodes.clear();context=null;}
  function restore(state){
    if(!state)return;
    if(context&&context!==state.context_id&&state.context_id)reset();
    context=state.context_id||context;
    for(const row of state.turns||[])merge(row);
    for(const row of rows.values())if(row.local)render(row);
  }
  setInterval(()=>{for(const row of rows.values())if(!terminal.has(row.status))render(row);},1000);
  window.OfficeConversation={echo,merge,restore,failed,configure:value=>{handlers=value;},
    undoMessage:id=>rows.get(id)?.undo?.message,
    undoError:(id,message)=>{const row=rows.get(id);if(row){row.undo={...(row.undo||{}),status:'blocked',available:false,busy:false,message};render(row);}},
    painted:()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))),
    connection:value=>{connected=value;for(const row of rows.values()){const node=nodes.get(row.id);if(node)node.signature=null;render(row);}},
    proposalHost:id=>{const node=nodes.get(id),row=rows.get(id);if(!node||!row||terminal.has(row.status))return null;node.pending??=make('div','',node.assistant);node.pending.className='conversation-pending';return node.pending;},
    pending:()=>[...rows.values()].some(row=>!terminal.has(row.status)&&!['clarify','reference','scope','deferred','choices'].includes(row.stage)),
    has:id=>rows.has(id),message:id=>rows.get(id)?.local&&rows.get(id).request?{...rows.get(id).request}:null,timings:()=>timings.slice(-160),reset};
})();
