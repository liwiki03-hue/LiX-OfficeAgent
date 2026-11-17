(() => {
  'use strict';
  const $=id=>document.getElementById(id);
  const ask=message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  const make=(tag,text,parent,id)=>{const node=document.createElement(tag);if(text)node.textContent=text;if(id)node.id=id;parent?.append(node);return node;};
  const button=(text,parent,action,id)=>{const node=make('button',text,parent,id);node.type='button';node.addEventListener('click',()=>Promise.resolve(action()).catch(error=>{if(!error.stale)notice(error.message);}));return node;};
  const tool=(text,icon,parent,action,id)=>window.OfficeAgentElements.iconButton(button(text,parent,action,id),icon);
  let adapter, identity='', host, opened=false, opening=false, generation=0, reference=null, lastState, changes=Promise.resolve(), saving=false, sending=false, applying=false, instruction=0, signature='', recoveryNotice='', autosave, errorGeneration=0;
  let announcedState='', announcementTimer, instructionEpoch=0,requirementsDialog=null,pendingFeedback=null;
  const retiredSessions=new Set();
  const retiredInboxes=new Set();
  const automaticAttempts=new Set();
  let automaticSaving=null, publicationProgress=null,lastExportRevision=-1,undoing=false;
  const api=()=>window.OfficeAuth.wrapApi(window.pywebview.api);
  const active=()=>!!adapter;
  function notice(text){pendingFeedback=null;const target=$('paged-feedback')||$('floating-status');if(target){target.textContent=text||'';errorGeneration++;}}
  async function call(name,...args){const mine=generation,result=await api()[name](...args);if(mine!==generation){const error=Error('文档已切换，旧结果未进入当前草稿');error.stale=true;throw error;}if(!result?.ok)throw Error(result?.error||'宿主未完成操作；请核对当前草稿');return result;}
  async function publish(name,args,id,text){
    const mine=generation;
    notice(text);
    publicationProgress={id,generation:mine,feedback:errorGeneration};
    try{const result=await call(name,...args);renderState(lastState);return result;}
    catch(error){if(mine===generation&&publicationProgress?.id===id)publicationProgress=null;throw error;}
  }
  function layout(){
    const root=$('floating-root');if(!root)return;
    const closed=root.classList.contains('paged-ai-closed'), outline=root.classList.contains('paged-outline-open');
    if($('paged-toggle-ai')){
      window.OfficeAgentElements.iconButton($('paged-toggle-ai'),closed?'panel-right-open':'panel-right-close','AI 对话',closed?'展开对话侧栏':'收起对话侧栏');
      $('paged-toggle-ai').setAttribute('aria-expanded',String(!closed));
    }
    $('paged-toggle-outline')?.setAttribute('aria-expanded',String(outline));
    if(adapter)adapter.frame.inert=matchMedia('(max-width:850px)').matches&&(!closed||outline);
  }
  function setSidebar(closed,returnFocus=false){
    const root=$('floating-root');if(!root?.classList.contains('paged-active'))return;
    root.classList.toggle('paged-ai-closed',closed);
    if(matchMedia('(max-width:850px)').matches&&!closed)root.classList.remove('paged-outline-open');
    localStorage.setItem('office-agent.paged-ai-closed',String(closed));
    layout();
    if(returnFocus)$('paged-toggle-ai')?.focus();
  }
  function dialog(title,build,returnFocus=document.activeElement){
    const previous=returnFocus,node=make('dialog','',document.body);node.className='paged-dialog';
    node.setAttribute('aria-label',title);make('h2',title,node);build(node);
    const buttons=[...node.children].filter(child=>child.tagName==='BUTTON');
    if(buttons.length){const actions=make('div','',node);actions.className='paged-dialog-actions';actions.append(...buttons);}
    node.addEventListener('close',()=>{node.remove();previous?.focus();});node.showModal();return node;
  }
  async function sync(bootstrap=false){
    const mine=generation,rev=adapter.revision;await changes;
    if(!bootstrap&&lastExportRevision===rev&&lastState?.snapshot_revision===rev&&adapter.revision===rev)return lastState;
    const result=await adapter.exportSnapshot(rev);
    if(mine!==generation||adapter.revision!==rev)throw Error('草稿已经改变，请重新核对');
    const response=await call('paged_sync',host.document.id,rev,result.data,result.manifest,bootstrap);
    if(mine===generation){lastExportRevision=rev;renderState(response);}
    return response;
  }
  function olderState(state){
    if(!lastState)return false;
    if(lastState.inbox_id!==state.inbox_id)return retiredInboxes.has(state.inbox_id);
    if(state.draft_revision<lastState.draft_revision)return true;
    if(lastState.requirement_session_id!==state.requirement_session_id)return retiredSessions.has(state.requirement_session_id||'');
    return lastState.instruction_revision>state.instruction_revision;
  }
  function acceptState(state){
    if(olderState(state))return lastState;
    if(lastState&&lastState.inbox_id!==state.inbox_id)retiredInboxes.add(lastState.inbox_id);
    if(lastState?.inbox_id===state.inbox_id&&lastState.requirement_session_id!==state.requirement_session_id)retiredSessions.add(lastState.requirement_session_id||'');
    lastState=state;return state;
  }
  function renderState(state){
    if(!state||!$('paged-save-state'))return;
    state=acceptState(state);
    if(publicationProgress&&state.publication_result?.id===publicationProgress.id&&!state.publishing){
      const progress=publicationProgress;publicationProgress=null;
      if(progress.generation===generation&&progress.feedback===errorGeneration)notice(state.publication_result.text);
    }
    if(host&&state.document_id===host.document?.id&&state.access?.serial>=(host.access?.serial??-1))host.access=state.access;
    if(pendingFeedback){
      const current=pendingFeedback.generation===errorGeneration&&pendingFeedback.inbox===state.inbox_id&&pendingFeedback.session===state.requirement_session_id;
      if(!current)pendingFeedback=null;
      else if(!state.pending_requirements&&state.instruction_revision>=pendingFeedback.revision)notice('');
    }
    if(requirementsDialog&&(state.inbox_id!==requirementsDialog.inbox||state.requirement_session_id!==requirementsDialog.session))requirementsDialog.invalidate();
    const unsynced=adapter&&adapter.revision>state.draft_revision;
    $('paged-save-state').textContent=state.uncertain?'结果待核实':state.publishing?'正在生成并验证副本':!state.editable?'只读':unsynced?'有未保存编辑 · 正在核对草稿 '+adapter.revision:state.dirty?'有未保存编辑 · 草稿 '+state.draft_revision:state.published_revision>0?'副本已验证 · 草稿 '+state.published_revision:'原稿 · 未修改';
    $('paged-save-state').title=!state.editable&&!state.uncertain&&!state.publishing?'只读'+((state.reasons||[]).length?'：'+state.reasons.join('、'):''):$('paged-save-state').textContent;
    if(state.publication_rolled_back&&!state.uncertain&&!state.publishing){
      $('paged-save-state').textContent='副本已撤销 · 草稿未交付';
      $('paged-save-state').title='原稿与草稿保留，可重新生成副本';
      if($('paged-feedback')?.textContent.startsWith('草稿 ')&&$('paged-feedback').textContent.includes('副本已通过验证'))notice('副本已撤销，草稿保留；可重新审阅后生成新副本。');
    }
    if(state.round_undo_restored?.revision===state.draft_revision){
      $('paged-save-state').textContent=state.round_undo_restored.before_dirty?'已撤销 · 此前人工编辑未保存':'已撤销 · 正文已恢复';
      $('paged-save-state').title='正文已回到本轮修改前，原件与此前内容保留';
      if($('paged-feedback')?.textContent==='副本已撤销，草稿保留；可重新审阅后生成新副本。')notice('');
    }
    const status=state.uncertain?'结果待核实':state.publishing?'正在生成并验证副本':!state.editable?'当前文档只读':unsynced||state.dirty?'有未保存编辑':state.published_revision>0?'副本已验证':'原稿未修改';
    $('paged-save-state').dataset.kind=state.uncertain?'uncertain':!state.editable?'readonly':unsynced||state.dirty?'dirty':'normal';
    if(status!==announcedState){announcedState=status;clearTimeout(announcementTimer);const mine=generation;announcementTimer=setTimeout(()=>{if(mine===generation&&$('paged-state-announcement'))$('paged-state-announcement').textContent=status;},700);}
    $('paged-save').disabled=!opened||saving||state.publishing||state.uncertain||!state.editable||!state.dirty||unsynced;
    $('paged-editable-copy').hidden=state.editable;
    $('paged-editable-copy').disabled=!opened||saving||host?.busy||state.publishing||state.uncertain;
    $('paged-restore').hidden=!state.recovery_available||state.dirty;
    $('paged-use-original').hidden=!state.recovery_available||state.dirty;
    $('paged-recovery-actions').hidden=!state.recovery_available||state.dirty;
    $('paged-reference').disabled=!opened||!state.editable||state.recovery_available;
    $('paged-table-reference').disabled=$('paged-reference').disabled;
    $('paged-pending').hidden=!state.pending_requirements;
    $('paged-pending').disabled=sending||applying||host?.busy||state.publishing||state.uncertain||!state.editable||!reference||reference.draft_revision!==adapter?.revision;
    $('paged-pending').title=$('paged-pending').disabled?'请等待当前步骤完成，并引用最新草稿后处理':'使用当前引用处理已保存要求';
    const received=state.requirements||[],waiting={reference:reference?.draft_revision===adapter?.revision?'待处理':'待引用',model:'待核对接口',budget:'调用预算已用尽',review:'待核对建议',next_step:'下一步处理',applying:'采用后待处理',publishing:'保存后待处理',verify_result:'待核对结果',readonly:'只读，暂不处理',recovery:'待核对恢复草稿',ready:'待处理'};
    waiting.readonly=state.copy_error||['failed','uncertain'].includes(state.publication_result?.status)?'副本未就绪':'准备可编辑副本';
    if(!state.editable&&state.publishing)waiting.publishing='正在创建可编辑副本';
    const cancelled=received.filter(row=>row.status==='cancelled').length;
    const receiptText=received.length?'已收到 '+received.length+' 条'+(cancelled?' · 已取消 '+cancelled+' 条':'')+(state.pending_requirements?' · '+(waiting[state.pending_reason]||'待处理'):''):'';
    if($('paged-requirement-state').textContent!==receiptText)$('paged-requirement-state').textContent=receiptText;
    $('paged-requirement-state').dataset.inbox=state.inbox_id||'';
    $('paged-requirement-details').hidden=!received.length;
    const recoveryState=state.recovery_state||(state.recovery_available?'available':null);
    if(state.recovery_status&&['available','invalid','failed'].includes(recoveryState)&&recoveryNotice!==state.recovery_status){recoveryNotice=state.recovery_status;notice(state.recovery_status);}
    if(recoveryState==='saved'||recoveryState==='none')recoveryNotice='';
    if(state.published_revision>0&&!state.publishing&&!state.uncertain&&['正在生成并验证草稿 '+state.published_revision+' 的副本；后续输入仍属于新草稿。','正在保存并验证草稿 '+state.published_revision+' 的副本；原文件不变。'].includes($('paged-feedback')?.textContent))notice('草稿 '+state.published_revision+' 的副本已通过验证；原文件未改动。'+(state.dirty?'后续编辑尚未交付。':''));
    if(state.editable_copy_id&&!state.publishing&&!state.uncertain&&$('paged-feedback')?.textContent==='正在创建并验证简化副本；原稿保持只读。')notice('简化副本已生成，可在结果中打开副本编辑；原稿未改动。');
    const proposals=state.proposals||[],key=JSON.stringify([proposals,host?.conversation?.context_id]);
    if(key===signature)return;signature=key;
    const list=$('paged-proposals');list.replaceChildren();
    document.querySelectorAll('.conversation-pending').forEach(node=>node.replaceChildren());
    for(const proposal of proposals){
      if(proposal.status!=='pending')continue;
      const owner=proposal.conversation_id&&window.OfficeConversation?.proposalHost(proposal.conversation_id);
      const parent=owner||(!host?.conversation?.turns?.length?list:null);if(!parent)continue;
      const item=make('article','',parent);item.className='paged-proposal';
      const labels={pending:'待审阅',stale:'已过期，请重新生成',adopted:'已采用到草稿',rejected:'已拒绝'};
      item.dataset.status=proposal.status;
      let actions;
      if(proposal.edits){
        make('h3','表格修改 · '+proposal.edits.length+'处',item);make('p',labels[proposal.status]||proposal.status,item);
        if(proposal.reason)make('p',proposal.reason,item);
        for(const edit of proposal.edits)window.OfficeAgentElements.edit(item,{title:'表格文字',before:edit.old,after:edit.new,status:labels[proposal.status]||proposal.status});
        actions=make('div','',item);actions.className='agent-tool-actions';
      }else actions=window.OfficeAgentElements.edit(item,{title:'文字修改',before:proposal.old,after:proposal.new,status:labels[proposal.status]||proposal.status,reason:proposal.reason});
      for(const change of proposal.allowed_changes||[])make('p','按明确要求修改年份：'+change.before+' → '+change.after+'；其余数字、单位、公式和引文保持。',item);
      tool('定位','locate-fixed',actions,()=>adapter.request('locate',{id:proposal.reference.node_id}));
      if(proposal.status==='pending'){
        tool('拒绝','x',actions,async()=>{const result=await call('paged_reject',proposal.id);renderState(result);});
        tool('采用到草稿','check',actions,()=>adopt(proposal)).classList.add('agent-adopt');
      }
    }
  }
  async function adopt(proposal,accessId=null){
    if(applying||sending)throw Error('当前操作尚未完成，请稍后采用');
    applying=true;
    try{
      await changes;await sync();const mine=generation,rev=adapter.revision;
      if(accessId)await adapter.request('instruction',{revision:lastState.instruction_revision});
      const grant=await call('paged_prepare_apply',proposal.id,host.document.id,rev,accessId);
      if(mine!==generation)throw Error('当前文档已改变');
      const result=await adapter.applyApprovedProposal(grant.proposal,rev,grant.proposal.instruction_revision,grant.token);
      await changes;
      const response=await call('paged_applied',grant.token,host.document.id,result.prior_revision,result.revision);
      reference=null;$('paged-reference-state').textContent='已修改草稿';renderState(response);
      await sync();
      if(accessId&&host.access?.id===accessId){
        try{
          const target=result.reference||(proposal.edits?{nodes:proposal.edits.map(edit=>{const cell=proposal.reference.cells.find(row=>row.node_id===edit.node_id);return {node_id:edit.node_id,start:cell.start,end:cell.start+edit.new.length};}),whole_table:false}:{node_id:proposal.reference.node_id,start:proposal.reference.start,end:proposal.reference.start+proposal.new.length});
          reference=(await call('paged_reference',host.document.id,adapter.revision,target)).reference;
          $('paged-reference-state').textContent='继续修改';
        }catch(error){if(error.stale)throw error;$('paged-reference-state').textContent='已修改草稿；下次发送时使用当前选区或光标所在段落';}
        if(host.access.level==='save_copy')automaticSaving={accessId,revision:adapter.revision};
      }
    }finally{applying=false;}
  }
  async function capture(allowParagraph=false,table=false){
    if(!opened)throw Error('编辑器仍在打开');
    const feedbackVersion=errorGeneration;
    await sync();const target=await adapter.getSelectionReference(allowParagraph,table);
    const result=await call('paged_reference',host.document.id,adapter.revision,target);
    reference=result.reference;$('paged-reference-state').textContent='已引用：'+reference.old.slice(0,100);
    if(feedbackVersion===errorGeneration)notice('');
  }
  async function submit(clientId,text){
    if(sending)throw Error('正在确认接收；请勿重复发送');
    const mine=generation,feedbackVersion=errorGeneration,receivedEpoch=instructionEpoch,inbox=lastState?.inbox_id;
    sending=true;
    try{
      const append=/末尾|文末|结尾|最后|\bend\b|\bappend\b/i.test(text)&&/添加|增加|补充|追加|插入|\badd\b|\bappend\b/i.test(text);
      if(!append&&!applying&&lastState?.editable&&['edit','save_copy'].includes(host.access?.level||'save_copy')&&reference?.draft_revision!==adapter.revision){
        try{await capture(!/所选|选中|选区/.test(text)&&(/光标|当前段|这一?段/.test(text)||adapter.revision>0&&/最新草稿/.test(text)));}catch(error){if(error.stale)throw error;reference=null;}
      }
      const target=!append&&!applying&&reference?.draft_revision===adapter.revision?reference.id:null;
      const result=await call('paged_submit',clientId,text,target,host.document.id,host.document.version,lastState?.requirement_session_id||'',lastState?.inbox_id);
      if(receivedEpoch!==instructionEpoch||lastState?.inbox_id!==inbox){const error=Error('本轮已停止或切换，旧接收结果不更新当前输入');error.stale=true;throw error;}
      const newer=olderState(result);
      lastState=acceptState(result);instruction=lastState.instruction_revision;renderState(lastState);
      if(!applying)await adapter.request('instruction',{revision:instruction}).catch(()=>{});
      if(mine===generation&&feedbackVersion===errorGeneration){
        notice(newer?'要求已保存；列表已有更新，请核对当前有效要求。':result.pending_requirements?'要求 '+result.record.revision+' 已保存；'+({reference:'正在准备处理范围',model:'请选择有效模型，要求已保留',budget:'本轮调用预算已用尽，核对后开始新一轮',review:'核对建议与保护要求后主动处理',next_step:'将在下一安全步骤处理',applying:'采用完成后继续处理',publishing:'副本验证完成后继续处理',verify_result:'先核对不确定的副本结果',readonly:'将自动创建可编辑副本并继续处理',recovery:'先核对恢复草稿'}[result.pending_reason]||'等待处理'):'');
        if(!newer&&result.pending_requirements)pendingFeedback={generation:errorGeneration,inbox:lastState.inbox_id,session:lastState.requirement_session_id,revision:result.record.revision};
      }
      return result;
    }catch(error){if(mine===generation&&!error.stale)notice(error.message);throw error;
    }finally{sending=false;}
  }
  async function startPending(){
    if(sending||applying)throw Error('当前步骤尚未完成');
    if(!window.OfficeModelSelector?.readyForRequest())throw Error('请选择有效模型；已保存要求保持不变');
    const expected=lastState,receivedEpoch=instructionEpoch,feedbackVersion=errorGeneration;
    sending=true;
    try{
      await sync();
      if(!reference||reference.draft_revision!==adapter.revision)throw Error('请引用最新草稿，已保存要求无需重输');
      const result=await call('paged_start_pending',host.document.id,adapter.revision,reference.id,expected.instruction_revision,expected.requirement_session_id);
      if(receivedEpoch!==instructionEpoch||lastState?.inbox_id!==expected.inbox_id){const error=Error('旧处理回执不更新当前要求');error.stale=true;throw error;}
      const newer=olderState(result);
      renderState(result);if(feedbackVersion===errorGeneration)notice(newer?'要求列表已更新，请核对后主动处理。':result.started?'正在处理已保存要求':'要求仍待处理，请核对当前模型与文档状态');
    }finally{sending=false;}
  }
  async function applyOperations(operations){
    if(!opened||applying)throw Error('当前编辑步骤尚未完成');
    applying=true;
    try{
      await sync();
      const instruction=lastState?.instruction_revision??0;
      const prepared=await call('paged_prepare_operations',host.document.id,adapter.revision,instruction,operations);
      const result=await adapter.applyValidatedOperations(operations,{token:prepared.token,instruction_revision:instruction});
      await changes;
      const response=await call('paged_applied_operations',prepared.token,host.document.id,result.prior_revision,result.revision);
      reference=null;renderState(response);await sync();return response;
    }finally{applying=false;}
  }
  async function idMap(){await sync();return call('paged_id_map',host.document.id);}
  async function referenceLocator(locatorId,start,end){
    const result=await call('paged_reference_locator',host.document.id,adapter.revision,locatorId,start,end);
    reference=result.reference;$('paged-reference-state').textContent='已引用：'+reference.old.slice(0,100);
    return result;
  }
  function manageRequirements(){
    if(requirementsDialog?.node.isConnected){requirementsDialog.node.focus();return;}
    const owner={document:host.document.id,source:host.document.version,inbox:lastState.inbox_id,session:lastState.requirement_session_id},mine=generation;
    let shown=lastState,busy=false,dirty=false,stale=false,editing=null;
    const node=dialog('已收到要求',()=>{},$('paged-requirement-details'));
    make('p',host.document.name,node);
    const message=make('p','',node);message.setAttribute('role','alert');
    const body=make('div','',node);body.className='paged-requirements';
    const footer=make('div','',node);footer.className='paged-dialog-actions';
    const controls=disabled=>{for(const item of node.querySelectorAll('button'))item.disabled=disabled;if(editing)editing.disabled=disabled;};
    const current=()=>mine===generation&&lastState?.inbox_id===owner.inbox&&lastState.requirement_session_id===owner.session&&!stale;
    const invalidate=()=>{stale=true;message.textContent='文档或要求会话已改变；本页不能提交。可复制未保存内容后关闭。';controls(true);closing.disabled=false;if(editing){editing.disabled=false;editing.readOnly=true;}};
    const close=async()=>{if(busy)return;if(dirty&&!await ask('放弃尚未保存的要求修改？'))return;node.close();};
    const refresh=button('重新读取要求',footer,async()=>{
      if(!current()){invalidate();return;}
      if(dirty&&!await ask('重新读取会放弃当前未保存的要求修改，是否继续？'))return;
      busy=true;controls(true);
      try{const result=await call('paged_requirements',owner.document,owner.source,owner.inbox);if(!current())return invalidate();shown=acceptState(result);dirty=false;editing=null;message.textContent='';renderState(shown);render();}
      catch(error){if(!error.stale)message.textContent=error.message;}
      finally{busy=false;controls(stale);if(stale)closing.disabled=false;}
    });
    const closing=button('关闭',footer,close);
    async function commit(row,action,text,pending){
      if(busy)return;
      if(!current()){invalidate();return;}
      const identity=JSON.stringify([row.id,row.version,shown.instruction_revision,action,text]);
      if(pending.identity!==identity){pending.identity=identity;pending.id=crypto.randomUUID();}
      busy=true;controls(true);const feedbackVersion=errorGeneration;let committed=false;
      try{
        const result=await call('paged_change_requirement',owner.document,owner.source,owner.inbox,shown.requirement_session_id,pending.id,row.id,row.version,shown.instruction_revision,action,text);
        if(!current())return invalidate();
        const newer=olderState(result);
        shown=acceptState(result);instructionEpoch++;instruction=shown.instruction_revision;dirty=false;editing=null;
        if(!applying)await adapter.request('instruction',{revision:instruction}).catch(()=>{});
        if(!current())return invalidate();
        message.textContent=newer?'操作已记录；要求列表随后有更新，请核对当前版本。':action==='cancel'?'已取消后续处理；已发请求不保证撤回，已采用草稿和副本保留。':'要求已修改；引用当前文字后主动处理，不会自动调用模型。';
        renderState(shown);render();committed=true;
        if(feedbackVersion===errorGeneration)notice(message.textContent);
      }catch(error){if(!error.stale)message.textContent=error.message;else invalidate();}
      finally{busy=false;controls(stale);if(stale)closing.disabled=false;else if(committed&&node.isConnected)(body.querySelector('[data-requirement="'+row.id+'"] button')||refresh).focus();}
    }
    function render(){
      body.replaceChildren();
      const labels={cancelled:'已取消后续处理',included:'已纳入模型请求，未代表修改完成',stopped:'已停止',reference:'待确认引用',model:'待核对接口',budget:'调用预算已用尽',review:'待核对建议',next_step:'下一安全步骤处理',applying:'采用后待处理',publishing:'保存后待处理',verify_result:'待核对结果',readonly:'只读，暂不处理',recovery:'待核对恢复草稿',ready:'待处理'};
      for(const row of shown.requirements||[]){
        const item=make('section','',body);item.dataset.requirement=row.id;
        make('h3','要求 '+row.revision+' · '+(labels[row.status]||'待处理'),item);make('p',row.text,item);
        if(row.history?.length){const history=make('details','',item);make('summary','原文与变更记录',history);for(const old of row.history){make('p',(old.status==='superseded'?'已被替代：':'取消前内容：')+old.text,history);}}
        if(['cancelled','stopped'].includes(row.status))continue;
        const actions=make('div','',item);actions.className='paged-requirement-actions';
        button('修改要求',actions,async()=>{
          if(dirty&&!await ask('放弃尚未保存的要求修改？'))return;
          dirty=false;editing=null;
          render();const target=body.querySelector('[data-requirement="'+row.id+'"]');
          editing=make('textarea','',target);editing.value=row.text;editing.maxLength=4000;editing.setAttribute('aria-label','修改要求 '+row.revision);
          editing.addEventListener('input',()=>{dirty=editing.value!==row.text;});const pending={};
          button('保存修改',target,()=>commit(row,'replace',editing.value,pending));
          button('放弃修改',target,()=>{dirty=false;editing=null;render();});editing.focus();
        });
        button('取消要求',actions,async()=>{
          if(dirty&&!await ask('放弃尚未保存的要求修改？'))return;
          dirty=false;editing=null;render();const target=body.querySelector('[data-requirement="'+row.id+'"]');
          make('p','仅取消后续处理；已发请求不保证撤回，已采用草稿和已生成副本保留。',target);const pending={};
          button('确认取消要求',target,()=>commit(row,'cancel','',pending));button('保留要求',target,render);
        });
      }
    }
    node.addEventListener('cancel',event=>{event.preventDefault();close();});
    node.addEventListener('close',()=>{if(requirementsDialog?.node===node)requirementsDialog=null;});
    requirementsDialog={node,inbox:owner.inbox,session:owner.session,get dirty(){return dirty;},get busy(){return busy;},invalidate};render();refresh.focus();
  }
  async function editableCopy(){
    if(saving||!opened||lastState?.editable)return;
    saving=true;renderState(lastState);
    const mine=generation,documentId=host.document.id,rev=lastState.draft_revision;
    let plan,finished=false;
    const node=dialog('创建可编辑副本',body=>{
      make('p','当前原稿只读。创建简化文字副本后，可单独编辑；不会覆盖原文件。',body);
      const reasons=make('details','',body);make('summary','为什么只读',reasons);make('p',(lastState.reasons||[]).join('、')||'当前编辑器未确认可以安全往返。',reasons);
      make('p','正在核对可复制的内容…',body,'paged-copy-message');
    },$('paged-editable-copy'));
    node.addEventListener('close',()=>{
      if(!finished&&plan&&mine===generation)call('paged_publish_editable_copy',plan.id,documentId,rev,false).catch(()=>{});
    },{once:true});
    try{
      const result=await call('paged_plan_editable_copy',documentId,rev);
      if(mine!==generation||!node.isConnected||!node.open){
        if(mine===generation)await call('paged_publish_editable_copy',result.plan.id,documentId,rev,false);
        return;
      }
      plan=result.plan;$('paged-copy-message').remove();
      make('p','保留 '+plan.summary.paragraphs+' 段、'+plan.summary.tables+' 张表的文字。',node);
      make('p',plan.retained,node);
      const limits=make('ul','',node);for(const text of plan.limitations)make('li',text,limits);
      const target=make('details','',node);make('summary','新文件位置',target);make('p',plan.output,target);
      const continuation=make('label','',node);continuation.className='document-access-continue';
      const continueEditing=make('input','',continuation);continueEditing.type='checkbox';continueEditing.checked=true;
      make('span','创建后直接打开，按当前方式修改副本',continuation);
      const actions=make('div','',node);actions.className='paged-dialog-actions';
      button('取消',actions,()=>node.close());
      const confirm=button('确认创建简化副本',actions,async()=>{
        if(mine!==generation){node.close();return;}
        confirm.disabled=true;
        try{
          await publish('paged_publish_editable_copy',[plan.id,documentId,rev,true,continueEditing.checked],plan.id,'正在创建并验证简化副本；原稿保持只读。');finished=true;node.close();
        }catch(error){if(!error.stale)notice(error.message);node.close();}
      });confirm.className='paged-primary';
    }catch(error){
      if(!error.stale&&node.isConnected&&node.open){$('paged-copy-message').textContent=error.message;button('关闭',node,()=>node.close());}
    }finally{saving=false;if(mine===generation)renderState(lastState);}
  }
  async function save(){
    if(saving||!opened)return;
    saving=true;
    try{
      await sync();const rev=adapter.revision,mine=generation;
      const result=await call('paged_plan',host.document.id,rev);
      if(mine!==generation)return;
      const plan=result.plan;
      if(host.access?.level==='save_copy'&&host.access.id){
        await publish('paged_publish',[plan.id,host.document.id,rev,false,host.access.id],plan.id,'正在保存并验证草稿 '+rev+' 的副本；原文件不变。');
        return;
      }
      dialog('生成验证副本',node=>{
        make('p','原文件保持不变。以下是当前草稿的实际变化；确认后备份、写候选、重开并通过渲染闸门。',node);
        make('p','草稿 '+rev+' · '+plan.changes.length+' 项变化 · 输出为新的 DOCX 净稿副本',node);
        if(plan.pending_requirement_count)make('p','仍有 '+plan.pending_requirement_count+' 条已保存要求未纳入模型请求；本次只发布当前草稿。',node);
        const list=make('div','',node);list.className='paged-change-list';
        for(const change of plan.changes){
          const action=change.operation==='insert'?'新增段落':change.operation==='delete'?'删除段落':change.before===change.after&&change.format_changed?'调整格式':change.format_changed?'文字与格式变化':'文字变化';
          make('h3',(change.kind==='paragraph'?'段落 ':'')+change.position+' · '+action,list);
          if(change.operation!=='insert')make('pre','原文：'+change.before,list);
          if(change.operation!=='delete')make('pre','草稿：'+change.after,list);
          if(change.format_before){make('p','原格式：'+change.format_before,list);make('p','草稿格式：'+change.format_after,list);}
        }
        button('取消',node,async()=>{await call('paged_publish',plan.id,host.document.id,rev,false);node.close();});
        const confirm=button('确认生成副本',node,async()=>{
          confirm.disabled=true;
          try{await publish('paged_publish',[plan.id,host.document.id,rev,true],plan.id,'正在生成并验证草稿 '+rev+' 的副本；后续输入仍属于新草稿。');node.close();}
          catch(error){notice(error.message);node.close();}
        });confirm.className='paged-primary';
      },$('paged-save'));
    }finally{saving=false;renderState(lastState);}
  }
  async function driveAutomatic(){
    const access=host?.access;
    if(!opened||!access?.id||sending||applying||saving||undoing||lastState?.round_undo_busy||host.busy||lastState?.deferred||lastState?.preparing||lastState?.publishing||lastState?.uncertain||host.stage==='stopped'||document.querySelector('dialog[open]'))return;
    if(!lastState?.editable){
      if(!lastState?.pending_requirements||lastState.copy_error)return;
      const key=access.id+':editable-copy:'+lastState.requirement_session_id+':'+lastState.instruction_revision;
      if(automaticAttempts.has(key))return;
      automaticAttempts.add(key);
      notice('正在创建可编辑副本，完成后继续处理已发送要求；原稿保留。');
      const feedbackVersion=errorGeneration;
      const result=await call('paged_automatic_editable_copy',host.document.id,adapter.revision,lastState.inbox_id,lastState.requirement_session_id,lastState.instruction_revision,access.id);
      publicationProgress={id:result.copy_plan_id,generation,feedback:feedbackVersion};renderState(lastState);
      return;
    }
    if(lastState.continuation_pending&&lastState.pending_requirements){
      if(!window.OfficeModelSelector?.readyForRequest())return;
      const key=access.id+':continue-copy:'+lastState.requirement_session_id+':'+lastState.instruction_revision;
      if(automaticAttempts.has(key))return;
      automaticAttempts.add(key);
      await sync();
      const target=reference?.draft_revision===adapter.revision?{reference}:await call('paged_automatic_reference',host.document.id,adapter.revision,lastState.requirement_session_id,lastState.instruction_revision);
      reference=target.reference;$('paged-reference-state').textContent=reference.kind==='append'?'追加到文末':'已自动引用正文';
      const result=await call('paged_start_pending',host.document.id,adapter.revision,reference.id,lastState.instruction_revision,lastState.requirement_session_id);
      renderState(result);notice(result.started?'正在处理已发送要求':'请选择有效模型，要求已保留');
      return;
    }
    const proposals=(lastState.proposals||[]).filter(row=>row.status==='pending'&&row.draft_revision===adapter.revision&&row.instruction_revision===lastState.instruction_revision);
    if(proposals.length>1){
      const key=access.id+':alternatives:'+adapter.revision;
      if(!automaticAttempts.has(key)){automaticAttempts.add(key);notice('有多条备选建议，请选择一条采用；未自动选择。');}
      return;
    }
    const proposal=proposals[0];
    if(proposal){
      const key=access.id+':apply:'+proposal.id;
      if(automaticAttempts.has(key))return;
      automaticAttempts.add(key);await adopt(proposal,access.id);return;
    }
    if(access.level==='save_copy'&&automaticSaving?.accessId===access.id&&automaticSaving.revision===adapter.revision&&!lastState.pending_requirements&&lastState.dirty){
      const key=access.id+':save:'+adapter.revision;
      if(automaticAttempts.has(key))return;
      automaticAttempts.add(key);automaticSaving=null;await save();
    }
  }
  async function beforeSwitch(){
    if(requirementsDialog?.busy){notice('正在保存要求，请等待接收结果后切换');return false;}
    if(requirementsDialog?.dirty&&!await ask('切换文档将放弃尚未保存的要求修改，是否继续？'))return false;
    requirementsDialog?.node.close();
    if(!active()||adapter.revision===(lastState?.published_revision||0))return true;
    return new Promise(resolve=>dialog('当前有未保存编辑',node=>{
      make('p','保存副本、保留恢复草稿或放弃后才能切换。恢复不会恢复任何执行批准。',node);
      button('先生成副本',node,()=>{node.close();resolve(false);return save();});
      button('保留恢复草稿并切换',node,async()=>{await sync();await call('paged_permit_switch',host.document.id,adapter.revision,'recover');node.close();resolve(true);});
      button('放弃草稿并切换',node,async()=>{await call('paged_permit_switch',host.document.id,adapter.revision,'discard');node.close();resolve(true);});
      button('取消',node,()=>{node.close();resolve(false);});
      node.addEventListener('cancel',()=>resolve(false),{once:true});
    }));
  }
  function mount(){
    const root=$('floating-root');root.classList.add('paged-active');
    const top=make('div','',root,'paged-top');top.append(root.querySelector('header'));
    const bar=make('div','',null,'paged-filebar');
    const header=top.querySelector('header');header.insertBefore(bar,$('floating-choose'));
    tool('大纲','list-tree',bar,()=>{root.classList.toggle('paged-outline-open');if(matchMedia('(max-width:850px)').matches&&root.classList.contains('paged-outline-open'))root.classList.add('paged-ai-closed');layout();},'paged-toggle-outline').setAttribute('aria-controls','paged-outline');
    const recovery=make('div','',root,'paged-recovery-actions');recovery.hidden=true;
    button('恢复本机草稿',recovery,restore,'paged-restore').hidden=true;
    button('使用原稿',recovery,async()=>{const result=await call('paged_use_original',host.document.id);lastState=result;renderState(result);await adapter.request('enable',{editable:result.editable});notice('使用已准入原稿开始编辑；旧恢复草稿不代表已交付副本。');},'paged-use-original').hidden=true;
    tool('保存副本','save',bar,save,'paged-save').classList.add('paged-primary');
    tool('创建可编辑副本','file',bar,editableCopy,'paged-editable-copy').hidden=true;
    tool('AI 对话','panel-right-close',bar,()=>setSidebar(!root.classList.contains('paged-ai-closed')),'paged-toggle-ai').setAttribute('aria-controls','floating-conversation');
    const work=make('div','',root,'paged-work');
    const outline=make('aside','',work,'paged-outline');outline.setAttribute('aria-label','文档大纲');make('h2','大纲',outline);make('nav','',outline,'paged-outline-items');
    const middle=make('section','',work,'paged-editor-main');middle.setAttribute('aria-label','分页文档编辑');
    work.append($('floating-conversation'));
    tool('收起对话侧栏','panel-right-close',$('floating-conversation-heading'),()=>setSidebar(true,true),'paged-close-ai').setAttribute('aria-controls','floating-conversation');
    $('floating-conversation').append($('floating-controls'));
    const referenceBar=make('div','',$('floating-controls'),'paged-reference-bar');
    tool('引用选区','quote',referenceBar,()=>capture(),'paged-reference');make('p','选择文字讨论修改',referenceBar,'paged-reference-state');
    tool('引用整表，最多8个段落','freeze',referenceBar,()=>capture(false,true),'paged-table-reference');
    make('span','',referenceBar,'paged-requirement-state').setAttribute('role','status');
    tool('查看已收到要求','chat',referenceBar,manageRequirements,'paged-requirement-details').hidden=true;
    button('处理已保存要求',referenceBar,startPending,'paged-pending').hidden=true;
    $('floating-controls').prepend(referenceBar);
    $('floating-scroll').append(make('div','',null,'paged-proposals'));
    const bottom=make('footer','',root,'paged-statusbar'),metadata=make('div','',bottom,'paged-metadata');
    make('span','分页中…',metadata,'paged-statistics');
    make('span','正在打开…',metadata,'paged-save-state').setAttribute('aria-live','off');
    const announcement=make('span','',root,'paged-state-announcement');announcement.className='agent-visually-hidden';announcement.setAttribute('role','status');announcement.setAttribute('aria-live','polite');announcement.setAttribute('aria-atomic','true');
    tool('查找','search',bottom,()=>dialog('查找正文',node=>{const input=make('input','',node);input.maxLength=200;input.setAttribute('aria-label','查找文字');button('定位第一个匹配',node,async()=>{await adapter.request('find',{text:input.value});node.close();});button('关闭',node,()=>node.close());}));
    const paper=tool('纸张深色显示','moon',bottom,()=>{root.classList.toggle('paged-paper-dark');paper.setAttribute('aria-pressed',String(root.classList.contains('paged-paper-dark')));return adapter.request('paper',{dark:root.classList.contains('paged-paper-dark')});});paper.setAttribute('aria-pressed','false');
    const zoom=make('select','',bottom,'paged-zoom');zoom.setAttribute('aria-label','页面缩放');
    for(const value of [50,75,100,125,150]){const option=make('option',value+'%',zoom);option.value=value;if(value===100)option.selected=true;}
    zoom.onchange=()=>adapter.request('zoom',{percent:Number(zoom.value)}).catch(error=>notice(error.message));
    tool('适合页宽','scan',bottom,()=>adapter.request('zoom',{fit:true}));
    make('p','',root,'paged-feedback').setAttribute('role','status');
    if(localStorage.getItem('office-agent.paged-ai-closed')==='true')root.classList.add('paged-ai-closed');
    layout();
  }
  function frameEvent(type,value){
    if(type==='changed'){
      reference=null;const mine=generation,documentId=host.document.id;
      changes=changes.then(()=>{if(mine!==generation)return;return call('paged_changed',documentId,value.revision,value.apply_token).then(result=>renderState(result));}).catch(error=>notice(error.message));
      $('paged-reference-state').textContent='草稿已改变，请重新引用当前文字';
      clearTimeout(autosave);autosave=setTimeout(()=>{if(opened)sync().catch(error=>notice('恢复快照未更新：'+error.message));},1800);
    }else if(type==='state'){
      $('paged-statistics').textContent='第 '+value.page+' / '+value.pages+' 页 · '+value.characters+' 字符 · 本引擎分页';
      $('paged-statistics').title=$('paged-statistics').textContent;
      const outline=$('paged-outline-items'),key=JSON.stringify(value.outline);
      if(outline.dataset.signature!==key){outline.dataset.signature=key;outline.replaceChildren();for(const row of value.outline)button(row.text||'空标题',outline,()=>adapter.request('locate',{id:row.id}));if(!value.outline.length)make('p','文档未使用标题样式',outline);}
    }else if(type==='save')save().catch(error=>notice(error.message));
    else if(type==='notice')notice(value.text);
  }
  async function restore(){
    const result=await call('paged_restore',host.document.id);adapter.closeDocument();opened=false;
    adapter=new DocumentEditorAdapter($('paged-editor-main'),frameEvent);await adapter.openDocument(result);opened=true;await sync();await adapter.request('enable',{editable:result.editable});notice('已恢复本机草稿；未恢复模型请求、写入或撤销批准。');
  }
  async function undoRound(turnId){
    if(undoing||applying||sending||saving||!opened)throw Error('当前操作尚未完成，请稍后核对撤销');
    const mine=generation,documentId=host.document.id,sourceVersion=host.document.version;
    const row=host.conversation?.turns.find(value=>value.id===turnId);
    if(!row)throw Error('此轮不属于当前对话');
    await sync();
    if(!await ask('撤销这轮修改：“'+row.question.slice(0,100)+(row.question.length>100?'…':'')+'”。将恢复本轮AI修改前的草稿，保留之前的人工编辑，并按审计撤回这一轮'+(row.undo?.outputs_count||'已登记的')+'个未改变的最终结果文件；原件和工作基线保留。若已有后续编辑或依赖，将保留现场并说明原因。确认撤销？'))return;
    if(mine!==generation||host.document.id!==documentId)throw Error('文档已切换，未撤销');
    undoing=true;applying=true;clearTimeout(autosave);
    let token;
    try{
      const result=await call('conversation_undo_begin',turnId,documentId,adapter.revision,true);
      if(result.complete)return;
      token=result.token;
      if(result.phase==='restoring'){
        adapter.closeDocument();opened=false;
        adapter=new DocumentEditorAdapter($('paged-editor-main'),frameEvent);adapter.frame.inert=true;
        await adapter.openDocument(result);
        if(mine!==generation)throw Error('编辑器已切换，恢复结果未进入其他文档');
        const restored=await adapter.exportSnapshot(adapter.revision);
        const accepted=await call('conversation_undo_restored',turnId,token,documentId,adapter.revision,restored.data,restored.manifest);
        lastExportRevision=adapter.revision;renderState(accepted);
      }
      const deadline=performance.now()+30000;
      let status;
      do{
        status=await call('conversation_undo_status',turnId,token,documentId);
        if(['files_done','complete'].includes(status.phase))break;
        if(status.phase==='uncertain')throw Error(status.message);
        if(performance.now()>deadline)throw Error('仍未取得完整撤销回执，请核对已有结果，不要重复删除');
        await new Promise(resolve=>setTimeout(resolve,150));
      }while(true);
      const checked=await adapter.exportSnapshot(adapter.revision);
      renderState(await call('paged_sync',documentId,adapter.revision,checked.data,checked.manifest));
      renderState(await call('conversation_undo_finish',turnId,token,documentId,adapter.revision));
      opened=true;adapter.frame.inert=false;await adapter.request('enable',{editable:lastState.editable});layout();
      window.OfficeFloating?.refresh();
    }catch(error){
      if(token&&mine===generation){
        await api().conversation_undo_abort(turnId,token,documentId).catch(()=>{});
        try{
          const current=await call('paged_open',documentId,sourceVersion);
          adapter?.closeDocument();adapter=new DocumentEditorAdapter($('paged-editor-main'),frameEvent);
          await adapter.openDocument(current);opened=true;lastExportRevision=-1;renderState(current);await adapter.request('enable',{editable:current.editable});layout();
        }catch(_){opened=false;}
      }
      await window.OfficeFloating?.refresh();throw error;
    }finally{
      undoing=false;applying=false;if(adapter&&opened)layout();
    }
  }
  async function update(state){
    const supported=window.isSecureContext&&state.document?.ready&&state.document.name.toLowerCase().endsWith('.docx')&&typeof api().paged_open==='function';
    if(!supported){if(adapter||$('floating-root')?.classList.contains('paged-active'))close();return;}
    host=state;const key=state.document.id+':'+state.document.version;
    if(identity===key){if(state.paged){if(olderState(state.paged))return;renderState(state.paged);const desired=state.stage==='stopped'?-1:lastState.instruction_revision;if(adapter&&!sending&&!applying&&instruction!==desired){instruction=desired;await adapter.request('instruction',{revision:instruction});}await driveAutomatic();}return;}
    if(opening){close();opening=false;}
    close();identity=key;host=state;opening=true;const mine=++generation;
    try{
      const result=await call('paged_open',state.document.id,state.document.version);
      if(mine!==generation)return;
      lastState=acceptState(result);host.pagedCapabilities=result.capabilities;mount();
      adapter=new DocumentEditorAdapter($('paged-editor-main'),frameEvent);
      await adapter.openDocument(result);if(mine!==generation)return;opened=true;
      if(result.editable){
        try{await sync(result.draft_revision===0);}
        catch(error){
          const current=await call('paged_open',state.document.id,state.document.version);
          if(current.editable)throw error;
          renderState(current);notice('原稿只读；发送要求后自动创建可编辑副本。');
        }
        await adapter.request('enable',{editable:lastState.editable&&!lastState.recovery_available});
        if(lastState.editable&&typeof api().paged_editor_ready==='function')await call('paged_editor_ready',state.document.id,adapter.revision);
      }
      layout();
      renderState(lastState);await call('floating_present','review');
    }catch(error){if(mine===generation){
      notice(error.message);opened=false;adapter?.closeDocument();adapter=null;
      const container=$('paged-editor-main');
      if(container){container.replaceChildren();make('p','分页编辑器未能完成打开或无修改验证。原稿和已校验恢复草稿保留；没有生成副本。',container);button('重新打开当前草稿',container,()=>{identity='';opening=false;return update(host);});}
      if($('paged-save'))$('paged-save').disabled=true;
    }}
    finally{if(mine===generation)opening=false;}
  }
  function close(){
    automaticAttempts.clear();automaticSaving=null;publicationProgress=null;lastExportRevision=-1;
    requirementsDialog?.invalidate();
    retiredSessions.clear();retiredInboxes.clear();lastState=null;pendingFeedback=null;
    generation++;opened=false;clearTimeout(autosave);clearTimeout(announcementTimer);announcedState='';adapter?.closeDocument();adapter=null;reference=null;identity='';signature='';recoveryNotice='';
    const root=$('floating-root');if(!root?.classList.contains('paged-active'))return;
    const top=$('paged-top'),header=top?.querySelector('header');if(header)root.prepend(header);
    $('floating-body')?.append($('floating-conversation'));
    (matchMedia('(max-width:899px)').matches?root:$('floating-conversation'))?.append($('floating-controls'));
    for(const id of ['paged-top','paged-filebar','paged-work','paged-statusbar','paged-feedback','paged-reference-bar','paged-proposals','paged-close-ai','paged-recovery-actions','paged-state-announcement'])$(id)?.remove();
    root.classList.remove('paged-active','paged-ai-closed','paged-outline-open');
  }
  function previewStructure(view,title){
    if(!view)throw Error('尚未取得结构内容');
    dialog(title,node=>{
      make('p','此处是宿主有界结构文本，不是实时分页预览。',node);
      for(const block of view.blocks||[]){
        if(block.kind==='paragraph')make('p',block.text||'（空段落）',node);
        else{const table=make('table','',node);for(const row of block.rows||[]){const tr=make('tr','',table);for(const value of row)make('td',value,tr);}}
      }
      if(view.has_more)make('p','此摘录尚未覆盖全文。',node);
      button('关闭',node,()=>node.close());
    });
  }
  const requireAdapter=()=>{if(!opened||!adapter)throw Error('编辑器尚未打开');return adapter;};
  async function messageContext(text){
    if(lastState&&!lastState.editable)return {revision:lastState.draft_revision,reference:null};
    const deadline=performance.now()+10000;
    while(adapter&&!opened&&performance.now()<deadline)await new Promise(resolve=>setTimeout(resolve,25));
    if(!opened||!adapter)throw Error('编辑器尚未就绪，原话保留。');
    if(lastState.editable&&!lastState.publishing&&!/^(?:进度|到哪|停止|停一下)/.test(text))await sync();
    return {revision:adapter.revision,reference:reference?.draft_revision===adapter.revision?reference.id:null};
  }
  window.OfficePaged={active,update,submit,beforeSwitch,close,previewStructure,messageContext,reportError:notice,presentationSource:async()=>{if(!opened||applying)throw Error('请等待当前草稿就绪');if(lastState.editable)await sync();return {document_id:host.document.id,draft_revision:adapter.revision};},invalidate:()=>{instructionEpoch++;instruction=-1;if(adapter)adapter.request('instruction',{revision:-1}).catch(()=>{});},capabilities:()=>requireAdapter().capabilities(),contract:()=>requireAdapter().contract(),selection:()=>requireAdapter().getSelection(),snapshot:()=>requireAdapter().getSnapshot(),subscribeChanges:handler=>requireAdapter().subscribeChanges(handler),tableCommand:command=>requireAdapter().applyTableCommand(command),undo:()=>requireAdapter().undo(),redo:()=>requireAdapter().redo(),history:()=>requireAdapter().history(),applyOperations,idMap,referenceLocator};
  window.OfficePaged.undoRound=undoRound;
  window.addEventListener('resize',layout);
  new MutationObserver(()=>{if(active()&&opened)adapter.request('theme',{theme:document.documentElement.dataset.theme||'dark'}).catch(()=>{});}).observe(document.documentElement,{attributes:true,attributeFilter:['data-theme']});
  document.addEventListener('keydown',event=>{
    if(event.key==='Escape'&&!event.isComposing&&!event.defaultPrevented&&!document.querySelector('dialog[open],#floating-root details[open]')&&active()&&matchMedia('(max-width:850px)').matches){
      $('floating-root').classList.remove('paged-outline-open');setSidebar(true,true);
    }
  });
  window.addEventListener('officeauth:unauthenticated',()=>{requirementsDialog?.node.close();close();});
})();
