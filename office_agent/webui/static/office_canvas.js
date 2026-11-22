(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const Grid = window.OfficeCanvasGrid;
  const make = (tag, text, parent, id) => { const node = document.createElement(tag); if (text) node.textContent = text; if (id) node.id = id; parent?.append(node); return node; };
  let host, state, view, identity = '', fallback = '', epoch = 0, busy = false, opening = false, selected = null, pendingField = false, previewPage = 1, lastError = '', preview = null, aiSignature = '';
  let rangeAnchor = null, rangeFocus = null;
  let pointerSelection=null, suppressSelectionClick=false;
  let paneObserver=null;
  let windowLoading=false, scrollSuppressedUntil=0;
  let viewRequest=0,scrollTimer=null,scrollPosition={top:0,left:0},prefetch=null;
  let wheelIntent=null;
  const automaticAttempts=new Set();
  let automaticSaving=null;
  const active = () => !!identity;
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const notice = text => { if ($('canvas-feedback')) $('canvas-feedback').textContent = text || ''; };
  const button = (text, parent, action, id, icon) => {
    const node = make('button', text, parent, id); node.type = 'button';
    node.onclick = event => Promise.resolve(action(event)).catch(error => { if (!error.stale) notice(error.message); });
    if (icon) window.OfficeAgentElements.iconButton(node, icon, text);
    return node;
  };
  async function call(name, ...args) {
    const mine = epoch, result = await api()[name](...args);
    if (mine !== epoch) { const error = Error('文档已切换，旧结果已作废'); error.stale = true; throw error; }
    if (!result?.ok) throw Error(result?.error || result?.reason || '宿主未完成操作');
    return result;
  }
  function dialog(title, build) {
    const previous = document.activeElement, node = make('dialog', '', document.body); node.className = 'paged-dialog'; node.dataset.canvasDocument=identity;node.setAttribute('aria-label', title); make('h2', title, node); build(node);
    node.addEventListener('close', () => { node.remove(); if (previous?.isConnected) previous.focus(); }); node.showModal(); return node;
  }
  function status(next) {
    if (state?.ai && next.ai && next.ai.document_id===state.ai.document_id && next.ai.draft_revision===state.ai.draft_revision && next.ai.sequence<state.ai.sequence) next={...next,ai:state.ai};
    if (next.ai) next = {...next, document_id:next.ai.document_id, draft_revision:next.ai.draft_revision};
    if (next.document_id && next.document_id !== host?.document?.id || state && next.draft_revision != null && next.draft_revision < state.draft_revision) return;
    const access=next.access||next.ai?.access;
    if(host&&access?.serial>=(host.access?.serial??-1))host.access=access;
    if(next.draft_revision!=null&&next.draft_revision!==state?.draft_revision)prefetch=null;
    state = {...state, ...next};
    if (!$('canvas-state')) return;
    $('canvas-state').textContent = state.uncertain ? '结果待核实' : state.publishing ? '正在写副本与验证' : state.previewing || busy ? '正在处理' : pendingField ? '有未保存编辑' : state.dirty ? state.checkpoint_revision===state.draft_revision ? '本机草稿已保存 · 未生成副本' : '有未保存编辑' : state.published_revision && state.published_revision===state.draft_revision ? '副本已验证' : '原稿 · 未修改';
    const recovery=state.recovery_available;
    const multiple=view?.kind==='xlsx'&&rangeAnchor&&Grid.rectangle(rangeAnchor,rangeFocus).count>1;
    $('canvas-save').disabled = busy || state.publishing || state.uncertain || recovery || !state.dirty || pendingField;
    $('canvas-draft-save').disabled = busy || state.publishing || state.uncertain || recovery || !state.operation_count || pendingField || state.checkpoint_revision===state.draft_revision;
    $('canvas-recovery-actions').hidden = !recovery;
    $('canvas-draft-status').textContent = state.recovery_status || '';
    $('canvas-draft-location').textContent = state.recovery_location || '';
    if(recovery) $('canvas-value').readOnly=true;
    if(['invalid','failed'].includes(state.recovery_state)&&state.recovery_status!==lastError){lastError=state.recovery_status;notice(state.recovery_status);}
    $('canvas-undo').disabled = busy || state.publishing || state.uncertain || !state.undo_count || pendingField;
    $('canvas-undo').dataset.tooltip=state.history_origin==='restored_operations'?'恢复记录按单项撤回；本次新增区域按编辑组撤回':'撤回草稿上一个编辑组';
    $('canvas-redo').hidden=view?.kind!=='xlsx';
    $('canvas-redo').disabled = busy || state.publishing || state.uncertain || !state.redo_count || pendingField;
    $('canvas-apply').disabled = busy || state.publishing || recovery || multiple || !selected || !pendingField || selected.editable === false && !(selected.formula_editable && $('canvas-type').value === 'formula');
    for (const id of ['canvas-bold', 'canvas-italic', 'canvas-size']) $(id).disabled = busy || state.publishing || recovery || multiple || !selected || selected.editable === false || selected.format_editable === false;
    $('canvas-range-actions').hidden=view?.kind!=='xlsx';
    for(const id of ['canvas-copy-range','canvas-paste-range'])$(id).disabled=busy||state.publishing||state.uncertain||recovery||pendingField||!selected;
    $('canvas-geometry').disabled=busy||state.publishing||state.uncertain||recovery||pendingField||!selected||multiple||!view?.geometry;
    $('canvas-panes').disabled=busy||state.publishing||state.uncertain||recovery||pendingField||view?.kind!=='xlsx'||!view?.panes;
    if(view?.panes)$('canvas-panes').dataset.tooltip=view.panes.supported?('冻结窗格 · 前'+view.panes.rows+'行 / 前'+view.panes.cols+'列'):('文件窗格未展示：'+view.panes.reason);
    $('canvas-pane-state').textContent=view?.panes?.supported===false?'文件窗格未展示':view?.panes&&(view.panes.rows||view.panes.cols)?'冻结 '+view.panes.rows+'行 / '+view.panes.cols+'列':'';
    $('canvas-pane-state').title=view?.panes?.reason||'冻结范围来自当前文件草稿';
    $('canvas-type').disabled=!!multiple;
    if(multiple)$('canvas-value').readOnly=true;
    $('canvas-preview').disabled = busy || state.publishing || view?.kind!=='pptx'&&pendingField;
    if (state.error && state.error !== lastError) { lastError = state.error; notice(state.error); }
    $('canvas-ai-note').hidden = true;
    if ($('canvas-reference')) $('canvas-reference').disabled = busy || state.publishing || recovery || pendingField || !selected || !multiple && selected.editable === false;
    renderAI();
    const output=state.last_output;
    $('canvas-continue').hidden=!output; $('canvas-output-details').hidden=!output;
    if(output){$('canvas-output-path').textContent=output.path;if(!state.publishing&&!state.uncertain&&$('canvas-feedback').textContent==='正在备份、重开和渲染验证；请等待结果，不要重复确认')notice('副本已生成：'+output.name);}
  }
  function renderAI() {
    const panel = $('canvas-ai-panel'), ai = state?.ai;
    const signature=String(busy)+':'+JSON.stringify(ai);
    if (!panel || !ai || aiSignature === signature) return;
    aiSignature = signature; panel.replaceChildren();
    const reference = ai.read_reference||ai.reference;
    if(reference)make('p','引用：'+reference.item_name+' · '+reference.location,panel).className='canvas-reference-status';
    if (ai.running&&!window.OfficeConversation?.pending()) make('p','正在处理…',panel).setAttribute('role','status');
    if (ai.error) make('p',ai.error,panel).setAttribute('role','alert');
    if(!window.OfficeConversation)for(const answer of ai.answers||[])make('p',answer.text,panel);
    if (!ai.running && reference && ai.requirements?.length) button('处理已接收要求',panel,()=>guarded(async()=>{
      status(await call('canvas_ai_resume',host.document.id,state.draft_revision,reference.id,ai.session_id,ai.instruction_revision));
    }),'canvas-ai-resume');
    for (const proposal of ai.proposals || []) {
      if (!['pending','adopted'].includes(proposal.status)) continue;
      const card = make('article','',panel); card.className='paged-proposal';
      make('strong',proposal.status==='adopted'?'已采用到草稿':'修改建议 · 尚未采用',card);make('p','原文：'+proposal.old,card);make('p','建议：'+proposal.new,card);make('p',proposal.reason,card);
      if(proposal.edits){card.querySelectorAll('p').forEach(node=>node.remove());for(const edit of proposal.edits){make('strong',edit.target,card);make('p','原文：'+edit.old,card);make('p','修改：'+edit.new,card);}make('p',proposal.reason,card);}
      make('small',proposal.model.model_id+' · '+proposal.model.connection_name,card);
      if(proposal.status==='pending'){
        button('采用到草稿',card,()=>guarded(()=>adoptAI(proposal)),'canvas-ai-adopt-'+proposal.id).disabled=ai.running || busy;
        button('拒绝',card,()=>guarded(async()=>{status(await call('canvas_ai_reject',host.document.id,state.draft_revision,proposal.id));}));
      }
    }
    if(ai.reference?.kind==='range'&&ai.reference.cells.some(cell=>cell.action==='number'&&!cell.protected&&!cell.formula&&!cell.merged))button('选择可修改的数值格…',panel,()=>numericRange(ai.reference));
  }
  async function adoptAI(proposal,accessId=null){
    checkField();const item=view.item_id,reference=state.ai.reference;
    status(await call('canvas_ai_adopt',host.document.id,state.draft_revision,proposal.id,accessId));
    preview=null;await loadItem(item);
    if(accessId&&host.access?.id===accessId){
      try{status(reference.kind==='range'?await call('canvas_ai_reference_range',host.document.id,state.draft_revision,reference.item_id,reference.start,reference.end,reference.numeric_targets):await call('canvas_ai_reference',host.document.id,state.draft_revision,reference.item_id,reference.target,reference.allow_numeric_change));}
      catch(error){if(error.stale)throw error;notice('草稿已修改；请重新选择下一次修改的位置。');}
      if(host.access.level==='save_copy')automaticSaving={accessId,revision:state.draft_revision};
    }else notice('已修改草稿');
  }
  async function driveAutomatic(){
    const access=host?.access;
    if(!access?.id||busy||pendingField||state?.publishing||state?.previewing||state?.uncertain||state?.ai?.running||host.busy||host.stage==='stopped'||document.querySelector('dialog[open]'))return;
    const proposals=(state.ai?.proposals||[]).filter(row=>row.status==='pending'&&row.draft_revision===state.draft_revision);
    if(proposals.length>1){
      const key=access.id+':alternatives:'+state.draft_revision;
      if(!automaticAttempts.has(key)){automaticAttempts.add(key);notice('有多条备选建议，请选择一条采用；未自动选择。');}
      return;
    }
    const proposal=proposals[0];
    if(proposal){
      const key=access.id+':apply:'+proposal.id;
      if(automaticAttempts.has(key))return;
      automaticAttempts.add(key);await guarded(()=>adoptAI(proposal,access.id));return;
    }
    if(access.level==='save_copy'&&automaticSaving?.accessId===access.id&&automaticSaving.revision===state.draft_revision&&state.dirty){
      const key=access.id+':save:'+state.draft_revision;
      if(automaticAttempts.has(key))return;
      automaticAttempts.add(key);automaticSaving=null;await guarded(save);
    }
  }
  async function referenceAI(editing=false) {
    checkField();if(!selected)return;
    if(view.kind==='xlsx'&&selectedRange().count>1){
      const range=selectedRange();status(await call(editing?'canvas_ai_reference_range':'canvas_ai_read_range',host.document.id,state.draft_revision,view.item_id,range.start,range.end,...(editing?[[]]:[])));
      $('floating-root').classList.add('office-ai-open');$('canvas-toggle-ai').setAttribute('aria-expanded','true');$('floating-input').focus();notice('');return;
    }
    const numeric=view.kind==='xlsx'&&selected.type==='n';
    const select=async()=>{
      status(await call(editing?'canvas_ai_reference':'canvas_ai_read_reference',host.document.id,state.draft_revision,view.item_id,selected.id,...(editing?[numeric]:[])));
      $('floating-root').classList.add('office-ai-open');$('canvas-toggle-ai').setAttribute('aria-expanded','true');
      $('floating-input').focus();notice('');
    };
    if(numeric&&editing)return new Promise(resolve=>{let confirmed=false;const node=dialog('确认数值修改范围',node=>{make('p','仅允许修改 '+selected.id+' 的数值，按当前编辑方式处理；原文件保留。',node);button('取消',node,()=>node.close());button('允许本格数值建议',node,async()=>{confirmed=true;await select();node.close();resolve(true);});});node.addEventListener('close',()=>{if(!confirmed)resolve(false);});});
    await select();return true;
  }
  function numericRange(reference){
    const doc=host.document.id,revision=state.draft_revision;
    dialog('选择数值修改范围',node=>{
      make('p','仅勾选的普通数值格允许修改；公式、合并和保护格保持只读。',node);
      const inputs=[];
      for(const cell of reference.cells.filter(cell=>cell.action==='number'&&!cell.protected&&!cell.formula&&!cell.merged)){
        const label=make('label',cell.target+' · '+cell.old,node),input=make('input','',label);input.type='checkbox';input.value=cell.target;input.checked=reference.numeric_targets.includes(cell.target);inputs.push(input);
      }
      button('取消',node,()=>node.close());button('应用选择',node,async()=>{status(await call('canvas_ai_reference_range',doc,revision,reference.item_id,reference.start,reference.end,inputs.filter(input=>input.checked).map(input=>input.value)));node.close();});
    });
  }
  async function submit(clientId,text) {
    checkField();
    if(['edit','save_copy'].includes(host.access?.level||'save_copy')&&selected&&(!state?.ai?.reference||state.ai.reference.draft_revision!==state.draft_revision||state.ai.reference.read_only))if(!await referenceAI(true))throw Error('未授权修改数值；原话保留，可直接提问解释。');
    const reference=state?.ai?.reference;
    if(!reference||reference.draft_revision!==state.draft_revision)throw Error('请先选择当前内容并点击“引用给AI”；输入已保留。');
    const result=await call('canvas_ai_submit',host.document.id,state.draft_revision,reference.id,state.ai.session_id,clientId,text);
    status(result);return result;
  }
  async function guarded(action) {
    if (busy) return;
    const mine = epoch; busy = true; status({});
    try { await action(); } finally { if (mine === epoch) { busy = false; status({}); } }
  }
  function checkField() {
    if (pendingField) throw Error('编辑栏还有未采用文字；请按Enter采用，或按Esc取消');
  }
  async function loadItem(itemId, row = view?.row || 1, col = view?.col || 1, preserve=false) {
    checkField();
    if(!preserve){wheelIntent=null;clearTimeout(scrollTimer);scrollTimer=null;}
    const request=++viewRequest,revision=state.draft_revision;
    const result = await call('canvas_view', host.document.id, revision, itemId, row, col);
    if(request!==viewRequest||revision!==state?.draft_revision||pendingField){const error=Error('当前草稿或输入已变化，窗口结果已作废');error.stale=true;throw error;}
    const anchor=preserve?captureAnchor(result.view):null,previous=preserve?selected?.id:null;
    const oldRange=preserve?[rangeAnchor,rangeFocus]:null,focused=document.activeElement?.dataset?.target;
    status(result); view = result.view; selected = null; preview = null; render();
    if(preserve){
      if(previous&&$('canvas-surface').querySelector('button[data-target="'+previous+'"]')){
        setSelected(cellAt(previous));
        if(oldRange?.[0]&&oldRange[1]!==oldRange[0]){rangeAnchor=oldRange[0];setSelected(cellAt(oldRange[1]),true);}
      }
      restoreAnchor(anchor);
      if(focused)$('canvas-surface').querySelector('button[data-target="'+focused+'"]')?.focus({preventScroll:true});
    }else{
      const surface=$('canvas-surface');surface.scrollTop=0;surface.scrollLeft=0;prefetch=null;
    }
    scrollPosition={top:$('canvas-surface').scrollTop||0,left:$('canvas-surface').scrollLeft||0};
    if (view.kind === 'pptx') {slideMessage='';renderPage();}
  }
  function visibleCells(direction=null){
    const surface=$('canvas-surface'),bounds=surface?.getBoundingClientRect?.();
    if(!bounds)return [];
    const frozenRows=view.panes?.rows||0,frozenCols=view.panes?.cols||0;
    let top=bounds.top,left=bounds.left;
    for(const node of surface.querySelectorAll('thead th,[data-frozen-row="true"]'))top=Math.max(top,node.getBoundingClientRect().bottom);
    for(const node of surface.querySelectorAll('tbody th,[data-frozen-col="true"]'))left=Math.max(left,node.getBoundingClientRect().right);
    const dy=direction==='down'?surface.scrollHeight-surface.clientHeight-surface.scrollTop:direction==='up'?-surface.scrollTop:0;
    const dx=direction==='right'?surface.scrollWidth-surface.clientWidth-surface.scrollLeft:direction==='left'?-surface.scrollLeft:0;
    const cells=[];
    for(const node of surface.querySelectorAll('.canvas-grid button[data-target]')){
      const [row,col]=Grid.position(node.dataset.target),box=node.getBoundingClientRect();
      if(row<=frozenRows||col<=frozenCols)continue;
      if(box.bottom-dy>top+1&&box.top-dy<bounds.bottom-1&&box.right-dx>left+1&&box.left-dx<bounds.right-1)cells.push({row,col,target:node.dataset.target,top:box.top-bounds.top,left:box.left-bounds.left});
    }
    return cells;
  }
  function captureAnchor(next){
    const indices=Grid.indices(next);
    return visibleCells().find(cell=>indices.rows.includes(cell.row)&&indices.cols.includes(cell.col))||null;
  }
  function restoreAnchor(anchor){
    if(!anchor)return;
    const surface=$('canvas-surface'),node=surface.querySelector('button[data-target="'+anchor.target+'"]');
    if(!node)return;
    const bounds=surface.getBoundingClientRect(),box=node.getBoundingClientRect();
    surface.scrollTop+=box.top-bounds.top-anchor.top;
    surface.scrollLeft+=box.left-bounds.left-anchor.left;
  }
  async function loadWindow(row, col) {
    if(windowLoading||!view||view.kind!=='xlsx'||pendingField||Date.now()<scrollSuppressedUntil)return;
    if(row===view.row&&col===view.col)return;
    const mine=epoch;
    windowLoading=true;
    try{
      await loadItem(view.item_id,row,col,true);
    }finally{if(mine===epoch){windowLoading=false;scrollSuppressedUntil=Date.now()+80;}}
  }
  function prefetchWindow(target){
    if(!target||prefetch?.pending||typeof api().canvas_prefetch!=='function')return;
    const key=[epoch,identity,state.draft_revision,view.item_id,target.row,target.col].join(':');
    if(prefetch?.key===key)return;
    const current={key,pending:true};prefetch=current;
    call('canvas_prefetch',host.document.id,state.draft_revision,view.item_id,target.row,target.col)
      .catch(()=>{}).finally(()=>{if(prefetch===current)current.pending=false;});
  }
  function scrollWindow() {
    if(!view||view.kind!=='xlsx'||pendingField||state?.publishing||state?.previewing||state?.uncertain||document.querySelector('dialog[open]')){wheelIntent=null;return;}
    const surface=$('canvas-surface');if(!surface)return;
    if(surface.clientHeight<=0||surface.clientWidth<=0)return;
    if(windowLoading||busy||Date.now()<scrollSuppressedUntil){clearTimeout(scrollTimer);scrollTimer=setTimeout(scrollWindow,100);return;}
    const dy=surface.scrollTop-scrollPosition.top,dx=surface.scrollLeft-scrollPosition.left;
    scrollPosition={top:surface.scrollTop,left:surface.scrollLeft};
    const intent=wheelIntent;wheelIntent=null;
    if(intent&&(intent.epoch!==epoch||intent.item!==view.item_id||intent.revision!==state.draft_revision))return;
    const direction=intent?.direction||(Math.abs(dy)>=Math.abs(dx)&&Math.abs(dy)>.5?(dy>0?'down':'up'):Math.abs(dx)>.5?(dx>0?'right':'left'):null);
    if(!direction)return;
    const distance={down:surface.scrollHeight-surface.clientHeight-surface.scrollTop,up:surface.scrollTop,right:surface.scrollWidth-surface.clientWidth-surface.scrollLeft,left:surface.scrollLeft}[direction];
    const cells=visibleCells(direction),visible=cells.length?{top:Math.min(...cells.map(cell=>cell.row)),bottom:Math.max(...cells.map(cell=>cell.row)),left:Math.min(...cells.map(cell=>cell.col)),right:Math.max(...cells.map(cell=>cell.col))}:null;
    const target=Grid.adjacent(view,direction,visible);
    if(distance<=160)prefetchWindow(target);
    if(distance<=32&&target)guarded(()=>loadWindow(target.row,target.col)).catch(error=>{if(!error.stale)notice(error.message);});
  }
  function wheelWindow(event){
    if(event.defaultPrevented||event.ctrlKey||event.metaKey||view?.kind!=='xlsx'||pendingField||state?.publishing||state?.previewing||state?.uncertain)return;
    const surface=$('canvas-surface');
    if(!surface||surface.clientHeight<=0||surface.clientWidth<=0||document.querySelector('dialog[open]')||event.target?.closest?.('input,textarea,select,[contenteditable="true"]'))return;
    const scale=event.deltaMode===1?16:1,horizontal=event.deltaMode===2?surface.clientWidth:scale,vertical=event.deltaMode===2?surface.clientHeight:scale;
    let dx=event.deltaX*horizontal,dy=event.deltaY*vertical;
    if(event.shiftKey&&!dx){dx=event.deltaY*horizontal;dy=0;}
    if(!Number.isFinite(dx)||!Number.isFinite(dy)||!dx&&!dy)return;
    const direction=Math.abs(dy)>=Math.abs(dx)?(dy>0?'down':'up'):(dx>0?'right':'left');
    if(!Grid.adjacent(view,direction)){wheelIntent=null;return;}
    const distance={down:surface.scrollHeight-surface.clientHeight-surface.scrollTop,up:surface.scrollTop,right:surface.scrollWidth-surface.clientWidth-surface.scrollLeft,left:surface.scrollLeft}[direction];
    if(distance>1&&!windowLoading&&!busy&&Date.now()>=scrollSuppressedUntil){wheelIntent=null;return;}
    wheelIntent={epoch,item:view.item_id,revision:state.draft_revision,direction};
    if(distance<=1)event.preventDefault();
    clearTimeout(scrollTimer);scrollTimer=setTimeout(scrollWindow,0);
  }
  function setSelected(target,extend=false) {
    if (pendingField) { notice('请先采用或取消编辑栏文字'); return false; }
    let range;
    if(view.kind==='xlsx'){
      range=Grid.rectangle(extend&&rangeAnchor?rangeAnchor:target.id,target.id);
      rangeAnchor=extend&&rangeAnchor?rangeAnchor:target.id;rangeFocus=target.id;
    }
    selected = target; $('canvas-position').textContent = target.location || target.id;
    if(range)$('canvas-position').textContent=range.label;
    $('canvas-value').value = target.formula != null ? '='+target.formula : target.text ?? target.value ?? '';
    $('canvas-value').readOnly = target.editable === false && !target.formula_editable;
    $('canvas-type').hidden = view.kind !== 'xlsx'; $('canvas-type').value = target.formula != null ? 'formula' : target.type === 'n' ? 'number' : 'text';
    $('canvas-type').querySelector('option[value="formula"]').disabled = target.formula_editable === false;
    $('canvas-size').value = target.style?.size || (target.size ? Number(target.size)/100 : 12);
    $('canvas-bold').setAttribute('aria-pressed', String(target.style?.bold ?? target.bold ?? false));
    $('canvas-italic').setAttribute('aria-pressed', String(target.style?.italic ?? target.italic ?? false));
    for (const node of $('canvas-surface').querySelectorAll('[data-target]')) {
      const included=range?Grid.includes(range,node.dataset.target):node.dataset.targets?JSON.parse(node.dataset.targets).includes(target.id):node.dataset.target===target.id;
      node.setAttribute('aria-selected', String(included));node.dataset.range=String(included&&range?.count>1);
      if(node.matches('.canvas-grid button'))node.tabIndex=node.dataset.target===target.id?0:-1;
    }
    pendingField = false; status({});
    if(range?.count>1){notice('已选 '+range.label+' · '+range.count+'格，可复制、粘贴或引用给 AI。');return true;}
    const display=target.display?.text;
    notice(target.formula != null ? (target.calculation?.status === 'calculated' ? '本机受限计算：'+target.calculation.value+' · 28位十进制预览，保存后由Office重算' : '此公式只读：'+(target.calculation?.reason || '超出受限计算范围'))+(display?' · 显示 '+display:'') : display ? '显示：'+display+'；编辑使用原始值：'+target.value : target.editable === false ? '富文本或受保护内容只读；原内容保留' : target.format_editable === false ? '此格依赖继承样式；可改文字/数值，格式调整暂未开放' : '');
    return true;
  }
  function colName(number) { let text = ''; while (number) { const rem = (number-1)%26; text = String.fromCharCode(65+rem)+text; number = Math.floor((number-1)/26); } return text; }
  function pos(text) { const match = /^([A-Z]+)(\d+)$/.exec(text); if (!match) return null; let col = 0; for (const char of match[1]) col = col*26+char.charCodeAt(0)-64; return [Number(match[2]), col]; }
  function selectedRange(){if(!rangeAnchor)throw Error('请先选择单元格');return Grid.rectangle(rangeAnchor,rangeFocus);}
  async function selectScope(axis,index){
    checkField();if(busy||view.kind!=='xlsx')return;
    const result=await call('canvas_selection_scope',host.document.id,state.draft_revision,view.item_id,axis,index);
    rangeAnchor=result.start;setSelected(cellAt(result.end),true);notice(result.scope+' · '+selectedRange().label+'（最多100格）');
  }
  function wireSelection(surface){
    const finish=cancel=>{
      const drag=pointerSelection;if(!drag)return;
      pointerSelection=null;
      if(surface.hasPointerCapture?.(drag.id))surface.releasePointerCapture(drag.id);
      if(cancel&&drag.selected&&drag.epoch===epoch&&drag.item===view?.item_id&&drag.revision===state?.draft_revision){rangeAnchor=drag.before[0];rangeFocus=drag.before[1];setSelected(drag.selected,true);}
      suppressSelectionClick=drag.moved;setTimeout(()=>{suppressSelectionClick=false;},0);
    };
    surface.addEventListener('pointerdown',event=>{
      const node=event.target.closest('.canvas-grid button[data-target]');
      if(!node||event.button!==0||busy||pendingField)return;
      if(event.ctrlKey||event.metaKey){event.preventDefault();notice('当前支持连续矩形选择；使用 Shift 扩选。');return;}
      const before=[rangeAnchor,rangeFocus],previous=selected;
      try{if(!setSelected(cellAt(node.dataset.target),event.shiftKey))return;}
      catch(error){notice(error.message);return;}
      pointerSelection={id:event.pointerId,anchor:rangeAnchor,before,selected:previous,moved:false,epoch,item:view.item_id,revision:state.draft_revision};
      surface.setPointerCapture?.(event.pointerId);
    });
    surface.addEventListener('pointermove',event=>{
      const drag=pointerSelection;if(!drag||drag.id!==event.pointerId)return;
      if(drag.epoch!==epoch||drag.item!==view.item_id||drag.revision!==state.draft_revision){finish(true);return;}
      const target=document.elementFromPoint(event.clientX,event.clientY)?.closest('.canvas-grid button[data-target]');
      try{if(target&&target.dataset.target!==rangeFocus){rangeAnchor=drag.anchor;setSelected(cellAt(target.dataset.target),true);drag.moved=true;}}
      catch(error){notice(error.message);}
      const box=surface.getBoundingClientRect();
      const dy=event.clientY<box.top+20?-24:event.clientY>box.bottom-20?24:0,dx=event.clientX<box.left+20?-24:event.clientX>box.right-20?24:0;
      if(dx||dy)surface.scrollBy(dx,dy);
    });
    surface.addEventListener('pointerup',()=>finish(false));surface.addEventListener('pointercancel',()=>finish(true));
    surface.addEventListener('dblclick',event=>{
      if(event.target.closest('.canvas-grid button[data-target]'))return;
      const target=document.elementFromPoint(event.clientX,event.clientY)?.closest('.canvas-grid button[data-target]');
      if(target&&surface.contains(target))target.ondblclick?.(event);
    });
    surface.addEventListener('keydown',event=>{if(event.key==='Escape'&&pointerSelection){event.preventDefault();finish(true);}});
    surface.addEventListener('contextmenu',event=>{
      const target=event.target.closest('.canvas-grid button[data-target]');if(!target||busy||pendingField)return;
      event.preventDefault();
      if(!rangeAnchor||!Grid.includes(selectedRange(),target.dataset.target))setSelected(cellAt(target.dataset.target));
      const range=selectedRange(),node=dialog('选区 '+range.label,()=>{});
      node.style.margin='0';node.style.inset='auto';node.style.position='fixed';node.style.width='220px';
      node.style.left=Math.max(8,Math.min(event.clientX,innerWidth-236))+'px';node.style.top=Math.max(8,Math.min(event.clientY,innerHeight-260))+'px';
      button('复制',node,()=>{node.close();return copyRange(true);});
      button('引用给 AI',node,()=>{node.close();return referenceAI(true);});
      button('清除选择',node,()=>{rangeAnchor=rangeFocus=null;selected=null;surface.querySelectorAll('[aria-selected]').forEach(item=>item.setAttribute('aria-selected','false'));status({});node.close();});
      button('关闭',node,()=>node.close());
    });
  }
  function cellAt(id){
    const cell=view.cells.find(cell=>cell.id===id)||{id,value:'',type:'inlineStr',editable:!view.protected,formula_editable:view.formula_creation_allowed};
    const covered=Grid.covered(view.merges,id);
    return {...cell,editable:cell.editable&&!covered,formula_editable:cell.formula_editable&&!covered};
  }
  async function moveGrid(row,col,extend){
    checkField();if(row<1||row>10000||col<1||col>256)return;
    const id=Grid.column(col)+row,anchor=rangeAnchor;
    if(extend&&anchor)Grid.rectangle(anchor,id);
    let node=$('canvas-surface').querySelector('button[data-target="'+id+'"]');
    if(!node){await loadItem(view.item_id,row,col);node=$('canvas-surface').querySelector('button[data-target="'+id+'"]');}
    if(!node)throw Error('目标是合并覆盖位置，请选择左上承载格');
    if(extend&&anchor)rangeAnchor=anchor;
    if(setSelected(cellAt(id),extend)){node.focus();node.scrollIntoView({block:'nearest',inline:'nearest'});}
  }
  async function historyStep(action){
    checkField();const item=view.item_id,target=selected?.id,row=view.row||1,col=view.col||1;
    status(await call(action==='redo'?'canvas_redo':'canvas_undo',host.document.id,state.draft_revision));
    await loadItem(item,row,col);
    if(target&&view.kind==='xlsx')setSelected(cellAt(target));
    notice(action==='redo'?'已重做编辑组；未写入原文件。':'已撤回编辑组；保存过的副本不受影响。');
  }
  async function copyRange(immediate=false){
    checkField();const range=selectedRange();
    const result=await call('canvas_copy_range',host.document.id,state.draft_revision,view.item_id,range.start,range.end);
    if(immediate){try{await navigator.clipboard.writeText(result.tsv);notice('已复制 '+range.label);return;}catch{}}
    dialog('复制区域 '+range.label,node=>{
      const text=make('textarea','',node);text.value=result.tsv;text.readOnly=true;text.rows=8;text.setAttribute('aria-label','区域原始内容');
      const feedback=make('p',result.copy_scope,node);feedback.setAttribute('role','status');
      button('复制',node,async()=>{try{if(!navigator.clipboard?.writeText)throw Error();await navigator.clipboard.writeText(result.tsv);feedback.textContent='已复制区域原始内容。';}catch{feedback.textContent='请在内容框中按Ctrl+C复制。';text.focus();text.select();}});
      button('关闭',node,()=>node.close());
    });
  }
  function pasteRange(initial=''){
    checkField();const range=selectedRange();
    if(initial.length>65536)throw Error('粘贴内容超过64 Ki字符，请缩小区域');
    const doc=host.document.id,revision=state.draft_revision,item=view.item_id;
    dialog('粘贴到 '+range.start,node=>{
      const input=make('textarea','',node);input.rows=8;input.maxLength=65536;input.value=initial;input.setAttribute('aria-label','制表符分隔的粘贴内容');
      const mode=make('select','',node);mode.setAttribute('aria-label','区域粘贴类型');
      for(const [value,label] of [['text','全部作为文本'],['numbers','识别明确十进制数值']]){const option=make('option',label,mode);option.value=value;}
      mode.value='text';
      const feedback=make('p','以文本粘贴，不执行公式。最多100格。',node);feedback.setAttribute('role','status');
      const changes=make('section','',node);changes.className='canvas-range-changes';
      let plan=null,closing=false,committing=false;
      const cancel=async()=>{
        if(closing||committing)return;closing=true;
        if(plan)await call('canvas_commit_range',doc,revision,plan.id,false).catch(()=>{});
        node.close();
      };
      const previewButton=button('预览区域',node,async()=>{
        if(busy)return;
        input.disabled=mode.disabled=previewButton.disabled=true;
        try{await guarded(async()=>{
          const result=await call('canvas_prepare_range',doc,revision,item,range.start,input.value,mode.value);
          if(!node.isConnected||!node.open){await call('canvas_commit_range',doc,revision,result.plan.id,false).catch(()=>{});return;}
          plan=result.plan;changes.replaceChildren();
          feedback.textContent=plan.start+':'+plan.end+' · 修改'+result.changed_cells+'格，未变化'+result.unchanged_cells+'格。仅采用到草稿，保存副本仍需独立确认。';
          for(const change of plan.changes){const entry=make('article','',changes);make('strong',change.target+' · '+(change.action==='number'?'数值':'文本'),entry);make('p','原内容：'+change.before,entry);make('p','改为：'+change.after,entry);}
          accept.hidden=false;back.hidden=false;previewButton.hidden=true;
        });}catch(error){if(!error.stale)feedback.textContent=error.message;input.disabled=mode.disabled=previewButton.disabled=false;}
      });
      const accept=button('采用到草稿',node,async()=>{
        if(busy||!plan||committing)return;committing=true;accept.disabled=true;let adopted=false;
        try{await guarded(async()=>{
          const result=await call('canvas_commit_range',doc,revision,plan.id,true);
          adopted=true;
          status(result);pendingField=false;node.close();await loadItem(item,view.row,view.col);
          setSelected(cellAt(range.start));notice('已采用'+result.group_operations+'项区域修改；可整组撤销，原文件未改动。');
        });}catch(error){if(!error.stale){if(adopted)notice('区域已采用到草稿，视图刷新失败；请重新打开当前工作表核对，不要重复粘贴。');else feedback.textContent=error.message;}committing=false;accept.disabled=adopted;}
      });accept.hidden=true;
      const back=button('返回修改',node,async()=>{
        if(busy||committing)return;
        if(plan)await call('canvas_commit_range',doc,revision,plan.id,false).catch(()=>{});
        plan=null;changes.replaceChildren();accept.hidden=back.hidden=true;previewButton.hidden=false;input.disabled=mode.disabled=previewButton.disabled=false;
        feedback.textContent='以文本粘贴，不执行公式。最多100格。';input.focus();
      });back.hidden=true;
      button('取消',node,cancel);
      node.addEventListener('cancel',event=>{event.preventDefault();cancel().catch(()=>node.close());});
    });
  }
  function geometryDialog(){
    checkField();if(!selected||selectedRange().count!==1)throw Error('请先选择一个单元格作为行列位置');
    const doc=host.document.id,revision=state.draft_revision,item=view.item_id,target=selected.id,[row,col]=pos(target);
    dialog('行列尺寸',node=>{
      const axis=make('select','',node);axis.setAttribute('aria-label','尺寸作用范围');
      for(const [value,label] of [['row_height','整行 '+row+' · 磅'],['column_width','整列 '+colName(col)+' · 字符宽度']]){const option=make('option',label,axis);option.value=value;}
      axis.value='row_height';
      const input=make('input','',node);input.type='number';input.step='0.05';input.setAttribute('aria-label','行列尺寸数值');
      const label=make('label','',node),reset=make('input','',label);reset.type='checkbox';label.append(document.createTextNode('恢复自动/默认尺寸'));
      const feedback=make('p','',node);feedback.setAttribute('role','status');
      let plan=null,committing=false;
      const refresh=()=>{
        const dimension=axis.value==='row_height'?view.geometry.rows[String(row)]:view.geometry.columns[String(col)];
        input.min=axis.value==='row_height'?6:2;input.max=axis.value==='row_height'?200:80;input.value=dimension?.value??'';
        reset.checked=false;input.disabled=dimension?.editable!==true;
        feedback.textContent=dimension?.editable?('包含未显示的单元格；当前'+(dimension.value==null?'自动/默认':dimension.value+(axis.value==='row_height'?'磅':'字符宽度'))+'。仅修改草稿，保存副本仍需独立确认。'):dimension?.reason||'当前尺寸无法读取';
        previewButton.disabled=dimension?.editable!==true;
      };
      const cancel=async()=>{if(committing)return;if(plan)await call('canvas_commit_geometry',doc,revision,plan.id,false).catch(()=>{});node.close();};
      const previewButton=button('预览影响',node,async()=>{
        if(busy)return;
        if(!reset.checked&&input.value===''){feedback.textContent='请输入尺寸，或选择恢复自动/默认';return;}
        axis.disabled=input.disabled=reset.disabled=previewButton.disabled=true;
        try{await guarded(async()=>{
          const result=await call('canvas_prepare_geometry',doc,revision,item,target,axis.value,reset.checked?null:Number(input.value));
          if(!node.isConnected||!node.open){await call('canvas_commit_geometry',doc,revision,result.plan.id,false).catch(()=>{});return;}
          plan=result.plan;const change=plan.changes[0];feedback.textContent=change.scope+'：'+change.before+' → '+change.after+'（'+change.unit+'）。采用后可撤销；原文件不变。';
          axis.disabled=input.disabled=reset.disabled=true;previewButton.hidden=true;adopt.hidden=back.hidden=false;
        });}catch(error){if(!error.stale){feedback.textContent=error.message;axis.disabled=reset.disabled=previewButton.disabled=false;input.disabled=reset.checked;}}
      });
      const adopt=button('采用到草稿',node,async()=>{
        if(busy||!plan||committing)return;committing=true;adopt.disabled=true;let adopted=false;
        try{await guarded(async()=>{
          const result=await call('canvas_commit_geometry',doc,revision,plan.id,true);adopted=true;preview=null;status(result);node.close();await loadItem(item,view.row,view.col);setSelected(cellAt(target));notice('尺寸已采用到草稿；原文件未改动。');
        });}catch(error){if(!error.stale){if(adopted)notice('尺寸已采用，视图刷新失败；请重新打开工作表核对，勿重复提交。');else feedback.textContent=error.message;}committing=false;adopt.disabled=adopted;}
      });adopt.hidden=true;
      const back=button('返回修改',node,async()=>{
        if(busy||committing)return;
        if(plan)await call('canvas_commit_geometry',doc,revision,plan.id,false).catch(()=>{});
        plan=null;axis.disabled=reset.disabled=false;adopt.hidden=back.hidden=true;previewButton.hidden=false;refresh();
      });back.hidden=true;
      axis.onchange=refresh;reset.onchange=()=>{input.disabled=reset.checked;};
      button('取消',node,cancel);node.addEventListener('cancel',event=>{event.preventDefault();cancel().catch(()=>node.close());});refresh();
    });
  }
  function panesDialog(){
    checkField();const doc=host.document.id,revision=state.draft_revision,item=view.item_id,info=view.panes;
    dialog('冻结窗格',node=>{
      const fields={};
      for(const [key,label,max] of [['rows','冻结前几行',4],['cols','冻结前几列',3]]){
        const line=make('label',label,node),input=make('input','',line);input.type='number';input.min=0;input.max=max;input.step=1;input.value=info[key];input.setAttribute('aria-label',label);fields[key]=input;
      }
      const feedback=make('p',info.editable?'固定文件开头的行列，计入30×16视窗预算；两项填0可取消冻结。':info.reason,node);feedback.setAttribute('role','status');
      let plan=null,committing=false;
      const cancel=async()=>{if(committing)return;if(plan)await call('canvas_commit_panes',doc,revision,plan.id,false).catch(()=>{});node.close();};
      const previewButton=button('预览影响',node,async()=>{
        if(busy)return;
        if(Object.values(fields).some(input=>input.value==='')){feedback.textContent='请填写两个整数；不冻结的方向填0。';return;}
        for(const input of Object.values(fields))input.disabled=true;previewButton.disabled=true;
        try{await guarded(async()=>{
          const result=await call('canvas_prepare_panes',doc,revision,item,Number(fields.rows.value),Number(fields.cols.value));
          if(!node.isConnected||!node.open){await call('canvas_commit_panes',doc,revision,result.plan.id,false).catch(()=>{});return;}
          plan=result.plan;const change=plan.changes[0];feedback.textContent=change.before+' → '+change.after+'。只修改窗格与活动选区；采用后可撤销，保存副本仍需独立确认。';
          previewButton.hidden=true;adopt.hidden=back.hidden=false;
        });}catch(error){if(!error.stale){feedback.textContent=error.message;for(const input of Object.values(fields))input.disabled=false;previewButton.disabled=false;}}
      });previewButton.disabled=!info.editable;
      const adopt=button('采用到草稿',node,async()=>{
        if(busy||!plan||committing)return;committing=true;adopt.disabled=true;let adopted=false;
        try{await guarded(async()=>{
          const result=await call('canvas_commit_panes',doc,revision,plan.id,true);adopted=true;preview=null;status(result);node.close();await loadItem(item,view.row,view.col);notice('窗格已采用到草稿；原文件未改动。');
        });}catch(error){if(!error.stale){if(adopted)notice('窗格已采用，视图刷新失败；请重新读取核对，不要重复提交。');else feedback.textContent=error.message;}committing=false;adopt.disabled=adopted;}
      });adopt.hidden=true;
      const back=button('返回修改',node,async()=>{
        if(busy||committing)return;
        if(plan)await call('canvas_commit_panes',doc,revision,plan.id,false).catch(()=>{});
        plan=null;for(const input of Object.values(fields))input.disabled=false;previewButton.hidden=false;previewButton.disabled=false;adopt.hidden=back.hidden=true;feedback.textContent='固定文件开头的行列；保存副本仍需独立确认。';
      });back.hidden=true;
      button('取消',node,cancel);node.addEventListener('cancel',event=>{event.preventDefault();cancel().catch(()=>node.close());});
    });
  }
  function pinGrid(table,indices){
    if(!view.panes?.supported||!view.panes.rows&&!view.panes.cols)return;
    table.classList.add('canvas-frozen-grid');
    const head=table.querySelector('thead'),body=table.querySelector('tbody');
    const update=()=>{
      if(!table.isConnected)return;
      const first= head?.children[0],widths=first?[...first.children].map(cell=>cell.getBoundingClientRect?.().width||0):[];
      let left=widths[0]||48;
      for(let index=0;index<indices.cols.length;index++){
        const col=indices.cols[index];
        if(col>view.panes.cols)break;
        const header=first?.children[index+1];if(header){header.dataset.frozenCol='true';header.style.left=left+'px';}
        for(const line of body?.children||[])for(const cell of line.children)if(cell.dataset.column===String(col)){cell.dataset.frozenCol='true';cell.style.left=left+'px';}
        left+=widths[index+1]||104;
      }
      let top=first?.getBoundingClientRect?.().height||28;
      for(const line of body?.children||[]){
        if(Number(line.dataset.row)>view.panes.rows)break;
        for(const cell of line.children){cell.dataset.frozenRow='true';cell.style.top=top+'px';}
        top+=line.getBoundingClientRect?.().height||32;
      }
    };
    update();
    if(typeof ResizeObserver==='function'){paneObserver=new ResizeObserver(update);paneObserver.observe(table);}
  }
  function grid(parent) {
    const indices=Grid.indices(view);
    const table = make('table', '', parent); table.className = 'canvas-grid'; table.setAttribute('role','grid');table.setAttribute('aria-multiselectable','true');table.setAttribute('aria-label', '工作表 '+view.items.find(item => item.id === view.item_id)?.name);
    if(view.geometry){
      table.classList.add('canvas-sized-grid');
      table.title=view.geometry.display||'行列尺寸近似展示，实际排版请查看打印预览';
      const columnWidth=col=>{const dimension=view.geometry.columns[String(col)],value=dimension?.value??dimension?.default;return value!=null?Math.max(24,Math.min(2050,value*8+5)):104;};
      const group=make('colgroup','',table);make('col','',group).style.width='48px';
      for(const col of indices.cols){
        const node=make('col','',group);node.style.width=columnWidth(col)+'px';
      }
      table.style.width=(48+indices.cols.map(columnWidth).reduce((a,b)=>a+b,0))+'px';
    }
    const head = make('thead', '', table), tr = make('tr', '', head); make('th', '', tr);
    for (const col of indices.cols) {const header=make('th', '', tr);header.scope='col';header.dataset.column=String(col);button(colName(col),header,()=>selectScope('column',col)).setAttribute('aria-label','选择已用范围内整列 '+colName(col));}
    const body = make('tbody', '', table), cells = new Map(view.cells.map(cell => [cell.id, cell]));
    const merges = view.merges.map(ref => { const [start, end = start] = ref.split(':'); return [pos(start), pos(end)]; }).filter(pair => pair[0] && pair[1]);
    for (const row of indices.rows) {
      const line = make('tr', '', body),header=make('th', '', line);header.scope='row';header.dataset.row=String(row);button(String(row),header,()=>selectScope('row',row)).setAttribute('aria-label','选择已用范围内整行 '+row);
      line.dataset.row=String(row);
      const dimension=view.geometry?.rows[String(row)];if(dimension?.value!=null){line.style.height=dimension.value+'pt';line.title='文件行高 '+dimension.value+'磅；网格为近似展示';}
      for (const col of indices.cols) {
        const id = colName(col)+row, merge = merges.find(([[r1,c1],[r2,c2]]) => row >= r1 && row <= r2 && col >= c1 && col <= c2);
        if (merge && (row !== merge[0][0] || col !== merge[0][1]) && indices.rows.includes(merge[0][0]) && indices.cols.includes(merge[0][1])) continue;
        const cell = cells.get(id) || {id, value:'', type:'inlineStr', editable:!view.protected, formula_editable:view.formula_creation_allowed}, td = make('td', '', line); td.dataset.target = id;
        td.dataset.row=String(row);td.dataset.column=String(col);
        if (merge && row === merge[0][0] && col === merge[0][1]) { td.rowSpan = indices.rows.filter(value=>value>=row&&value<=merge[1][0]).length; td.colSpan = indices.cols.filter(value=>value>=col&&value<=merge[1][1]).length; }
        const covered = merge && (row !== merge[0][0] || col !== merge[0][1]);
        const fallback = cell.formula != null ? cell.calculation?.status === 'calculated' ? cell.calculation.value : cell.value ? cell.value+'（旧缓存）' : '未计算' : cell.value;
        const display = cell.display?.text ?? fallback;
        const select = event => {if(suppressSelectionClick&&event?.type==='click')return false;if(event?.ctrlKey||event?.metaKey){notice('当前支持连续矩形选择；使用 Shift 扩选。');return false;}return !busy&&setSelected({...cell, editable:cell.editable && !covered, formula_editable:cell.formula_editable && !covered},!!event?.shiftKey);};
        const node = button(display || '\u00a0', td, select); node.dataset.target = id; node.title = id+(cell.display?.format ? ' · 显示格式 '+cell.display.format : '')+(cell.formula != null ? ' · ='+cell.formula+' · '+(cell.calculation?.status === 'calculated' ? '本机受限计算；28位精度' : cell.calculation?.reason || '未计算') : covered ? ' · 合并覆盖坐标' : ' · 双击编辑');
        node.tabIndex=row===view.row&&col===view.col?0:-1;
        node.onkeydown=event=>{
          if(event.isComposing||event.keyCode===229)return;
          if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='c'){event.preventDefault();copyRange(true).catch(error=>notice(error.message));return;}
          if((event.ctrlKey||event.metaKey)&&['z','y'].includes(event.key.toLowerCase())){event.preventDefault();guarded(()=>historyStep(event.key.toLowerCase()==='y'||event.shiftKey?'redo':'undo')).catch(error=>notice(error.message));return;}
          if(event.key==='Enter'){event.preventDefault();if(select(event))$('canvas-value').focus();return;}
          const move={ArrowUp:[-1,0],ArrowDown:[1,0],ArrowLeft:[0,-1],ArrowRight:[0,1]}[event.key];
          if(move){event.preventDefault();guarded(()=>moveGrid(row+move[0],col+move[1],event.shiftKey)).catch(error=>notice(error.message));}
        };
        node.ondblclick = () => { if(select()){ $('canvas-value').focus(); $('canvas-value').select(); } };
        const style = cell.style || {}; node.style.fontWeight = style.bold ? '700' : '400'; node.style.fontStyle = style.italic ? 'italic' : 'normal';
        if (Number(style.size) >= 6 && Number(style.size) <= 72) node.style.fontSize = Number(style.size)+'pt';
        if (/^#[0-9a-f]{6}$/i.test(style.background || '')) td.style.backgroundColor = style.background;
        if (/^#[0-9a-f]{6}$/i.test(style.color || '')) node.style.color = style.color;
        if (['left','center','right'].includes(style.align)) node.style.textAlign = style.align;
      }
    }
    pinGrid(table,indices);
  }
  function slide(parent) {
    const stage = make('div', '', parent, 'canvas-slide'); stage.style.aspectRatio = String(view.width/view.height);
    const image = make('img', '', stage, 'canvas-slide-image'); image.hidden = true;
    make('p', '页图准备中；可先编辑文字片段', stage, 'canvas-slide-placeholder');
    const boxes = new Map();
    for (const run of view.runs) {
      if (!run.box) continue;
      const key=JSON.stringify(run.box);if(!boxes.has(key))boxes.set(key,[]);boxes.get(key).push(run);
    }
    for (const runs of boxes.values()) {
      const run=runs[0],[x,y,w,h] = run.box;
      if (![x,y,w,h].every(Number.isFinite) || x < 0 || y < 0 || w <= 0 || h <= 0 || x+w > 1.01 || y+h > 1.01) continue;
      const hit = button('编辑文字', stage, () => openSlideEditor(runs)); hit.className = 'canvas-hit'; hit.dataset.target = run.id;hit.dataset.targets=JSON.stringify(runs.map(item=>item.id)); hit.style.left = x*100+'%'; hit.style.top = y*100+'%'; hit.style.width = w*100+'%'; hit.style.height = h*100+'%'; hit.setAttribute('aria-label','编辑此处文字'+(runs.length>1?' · '+runs.length+'个片段':'：'+run.text.slice(0,40)));hit.hidden=true;
    }
    make('p','',parent,'canvas-slide-status').setAttribute('role','status');
    const list = make('section', '', parent); list.className = 'canvas-runs'; make('h3','本页文字片段 · 点击编辑',list);
    for (const run of view.runs) { const node = button((run.location?.startsWith('表格') ? run.location+' · ' : '')+(run.text || '（空文字）'),list,()=>openSlideEditor([run])); node.dataset.target = run.id; node.setAttribute('aria-label','编辑片段 '+run.id+' '+run.text);node.title=run.location || '普通文字';
      button('引用片段 '+run.id,list,async()=>{setSelected({...run,editable:true});await referenceAI(false);},null,'quote');
    }
    if (!view.runs.length) make('p','本页没有可编辑的普通文字片段',list);
    paintSlidePreview();
  }
  function openSlideEditor(runs){
    checkField();if(document.querySelector('dialog[open]'))return;
    const bar=$('canvas-editbar'),home=bar.parentNode,next=bar.nextSibling;
    dialog('编辑本页文字',node=>{
      node.dataset.slideEditor='true';
      if(runs.length>1){
        make('p','此位置包含多个真实文字片段，请选择要修改的一项。',node);
        const choices=make('div','',node);choices.className='canvas-runs';
        for(const run of runs)button(run.text||'（空文字）',choices,()=>setSelected({...run,editable:true}));
        selected=null;$('canvas-position').textContent='请选择片段';$('canvas-value').value='';$('canvas-value').readOnly=true;status({});
      }else setSelected({...runs[0],editable:true});
      node.append(bar);button('返回页图',node,()=>node.close());
      node.addEventListener('close',()=>{if(home.isConnected)home.insertBefore(bar,next?.parentNode===home?next:null);});
    });
    if(runs.length===1)$('canvas-value').focus({preventScroll:true});
  }
  let slideRequest=0,slideRendering=null,slideDesired=null,slideMessage='';
  function slideKey(){return view?.kind==='pptx'?[epoch,host.document.id,view.item_id,state.draft_revision].join(':'):null;}
  function paintSlidePreview(){
    const stage=$('canvas-slide'),image=$('canvas-slide-image');if(!stage||!image)return;
    const current=preview?.document_id===host.document.id&&preview?.item_id===view.item_id&&preview?.draft_revision===state.draft_revision;
    image.hidden=!current;
    if(current){image.src=preview.image;image.alt=view.items.find(item=>item.id===view.item_id)?.name+' · 实际渲染页图';stage.dataset.revision=String(preview.draft_revision);}
    else{image.removeAttribute('src');delete stage.dataset.revision;}
    $('canvas-slide-placeholder').hidden=!!current;
    for(const hit of stage.querySelectorAll('.canvas-hit'))hit.hidden=!current;
    const text=$('canvas-slide-status');
    if(text)text.textContent=slideMessage||(current?'页图对应草稿 '+preview.draft_revision+' · 静态预览':'草稿 '+state.draft_revision+' · 页图尚未更新，可先编辑文字');
  }
  function render() {
    const nav = $('canvas-items'); nav.replaceChildren();
    for (const item of view.items.filter(item => !item.hidden)) { const node = button(item.name,nav,()=>guarded(()=>loadItem(item.id,1,1))); node.setAttribute('aria-current',String(item.id===view.item_id)); }
    $('canvas-window').hidden = view.kind !== 'xlsx'; $('canvas-preview-page').hidden = view.kind !== 'xlsx';
    $('canvas-type').hidden = view.kind !== 'xlsx';
    $('canvas-main').dataset.kind=view.kind;
    const surface = $('canvas-surface');paneObserver?.disconnect();paneObserver=null;surface.replaceChildren();
    if (view.kind === 'xlsx') { $('canvas-window').value = colName(view.col)+view.row; grid(surface); } else slide(surface);
    $('canvas-position').textContent = '未选中'; $('canvas-value').value = ''; $('canvas-value').readOnly = true; selected = null;
    rangeAnchor=rangeFocus=null;
    $('canvas-scope').textContent = (view.kind === 'xlsx' ? view.scope : view.readonly_objects)+(state.history_origin==='restored_operations'?'；恢复记录按单项撤销，本次新增区域按编辑组撤销。':'');
    status({});
  }
  async function apply(action, value) {
    if (!selected || selected.editable === false && !(action === 'formula' && selected.formula_editable)) return;
    if(view.kind==='xlsx'&&selectedRange().count>1)throw Error('区域当前支持复制和粘贴；编辑文字或格式请先选单格');
    const target = selected, item = view.item_id, row = view.row || 1, col = view.col || 1;
    const result = await call('canvas_apply',host.document.id,state.draft_revision,{item_id:item,target:target.id,action,value,expected:action==='formula'&&target.formula!=null?'='+target.formula:target.text ?? target.value ?? ''});
    const scroll={top:$('canvas-surface').scrollTop,left:$('canvas-surface').scrollLeft};
    pendingField = false; status(result); preview = null;slideMessage='';
    const refreshed = await call('canvas_view',host.document.id,state.draft_revision,item,row,col); view = refreshed.view; render();
    const current = view.kind === 'xlsx' ? view.cells.find(cell => cell.id === target.id) : view.runs.find(run => run.id === target.id);
    if (current) setSelected({...current,editable:current.editable !== false});
    if(view.kind==='pptx'){$('canvas-surface').scrollTop=scroll.top;$('canvas-surface').scrollLeft=scroll.left;}
    if(view.kind==='pptx')document.querySelector('dialog[data-slide-editor]')?.close();
    notice(view.kind === 'pptx' ? '草稿已更新；点击“刷新页图”查看实际排版' : result.change.derived || '草稿已更新；原文件未改动');
  }
  async function renderPage() {
    if(view.kind==='pptx')return renderSlidePage();
    checkField(); notice('正在渲染当前草稿…');
    const result = await call('canvas_preview',host.document.id,state.draft_revision,view.item_id,previewPage); preview = result;
    dialog('实际打印页预览',node=>{ const img = make('img','',node); img.src = result.image; img.alt = '第 '+result.page+' / '+result.pages+' 页'; img.style.maxWidth='100%'; make('p',result.renderer+' · '+result.page+'/'+result.pages,node); button('关闭',node,()=>node.close()); });
    notice('页图对应草稿 '+state.draft_revision+' · '+result.renderer);
  }
  async function renderSlidePage(){
    const key=slideKey();if(!key)return;
    slideDesired=key;
    if(slideRendering)return;
    const ticket={key,epoch,document:host.document.id,item:view.item_id,revision:state.draft_revision,request:++slideRequest};
    slideRendering=ticket;slideMessage='草稿 '+ticket.revision+' · 页图更新中，文字编辑仍可使用';paintSlidePreview();
    try{
      const result=await call('canvas_preview',ticket.document,ticket.revision,ticket.item,1);
      if(slideRendering!==ticket||slideKey()!==key||ticket.request!==slideRequest)return;
      if(result.document_id!==ticket.document||result.item_id!==ticket.item||result.draft_revision!==ticket.revision)throw Error('页图版本与当前草稿不符');
      preview=result;slideMessage='';paintSlidePreview();
      if($('canvas-feedback')?.textContent==='草稿已更新；点击“刷新页图”查看实际排版')notice('页图对应草稿 '+ticket.revision+' · '+result.renderer);
    }catch(error){
      if(!error.stale&&slideRendering===ticket&&slideKey()===key){slideMessage='页图尚未更新：'+error.message+'；可点击刷新页图重试，编辑内容保留';paintSlidePreview();}
    }finally{
      if(slideRendering===ticket){slideRendering=null;if(ticket.epoch===epoch&&slideDesired!==key&&slideDesired===slideKey())renderSlidePage();}
    }
  }
  async function save() {
    checkField();
    const result = await call('canvas_plan',host.document.id,state.draft_revision);
    if(host.access?.level==='save_copy'&&host.access.id){
      await call('canvas_publish',host.document.id,state.draft_revision,result.plan.id,false,host.access.id);
      notice('正在保存并验证副本；原文件不变。');return;
    }
    dialog('确认保存副本',node=>{
      make('p',result.plan.mode,node); make('p','输出：'+result.plan.output.split(/[\\/]/).pop(),node);
      for (const change of result.plan.changes) { const entry = make('article','',node); make('strong',change.item+' · '+(change.scope || change.location || change.target)+' · '+({row_height:'行高',column_width:'列宽'}[change.action]||change.action),entry); make('p','原内容：'+change.before+(change.unit?'（'+change.unit+'）':''),entry); make('p','修改为：'+(typeof change.after==='object'?JSON.stringify(change.after):change.after)+(change.unit?'（'+change.unit+'）':''),entry);if(change.calculation)make('p','最终草稿 '+change.final_formula+' → 本机受限计算：'+change.calculation.value+' · 28位十进制；保存后由Office重算',entry); if(change.derived)make('p',change.derived,entry); }
      button('取消',node,async()=>{await call('canvas_publish',host.document.id,state.draft_revision,result.plan.id,false);node.close();});
      const confirm = button('确认生成副本',node,async()=>{confirm.disabled=true;try{await call('canvas_publish',host.document.id,state.draft_revision,result.plan.id,true);node.close();notice('正在备份、重开和渲染验证；请等待结果，不要重复确认');}catch(error){notice(error.message);node.close();}}); confirm.className='floating-primary-button';
    });
  }
  async function checkpoint(action) {
    checkField();
    const doc=host.document.id, revision=state.draft_revision;
    const result=await call('canvas_draft_plan',doc,revision,action), plan=result.plan;
    dialog(action==='save'?'保存本机草稿':'恢复本机草稿',node=>{
      make('p',plan.source_name+' · '+plan.count+'项编辑',node);
      make('p',action==='save'?'本机记录含原文片段和修改内容，未加密；再次保存会更新此前记录。':'将从已重新授权的原文件回放编辑；不恢复旧模型请求或批准。',node);
      make('p',plan.notice,node);
      const details=make('details','',node);make('summary','编辑内容与保存位置',details);make('p',plan.location,details);make('p',plan.saved_at,details);
      for(const change of plan.changes){const entry=make('article','',details);make('strong',change.item+' · '+(change.location||change.target),entry);make('p','原内容：'+change.before,entry);make('p','修改为：'+(typeof change.after==='object'?JSON.stringify(change.after):change.after),entry);}
      let submitted=false;
      const cancel=async()=>{if(submitted)return;submitted=true;try{await call('canvas_draft_confirm',doc,revision,plan.id,false);}finally{node.close();}};
      button('取消',node,cancel);
      const confirm=button(action==='save'?'确认保存本机草稿':'确认恢复内容',node,async()=>{
        if(submitted)return;submitted=true;confirm.disabled=true;
        try{await guarded(async()=>{status(await call('canvas_draft_confirm',doc,revision,plan.id,true));node.close();if(action==='restore')await loadItem(view.item_id,1,1);notice(state.recovery_status);});}
        catch(error){node.close();if(!error.stale)notice(error.message);}
      });
      node.addEventListener('cancel',()=>{cancel().catch(error=>{if(!error.stale)notice(error.message);});});
    });
  }
  function mount() {
    const root = $('floating-root'); root.classList.add('office-active');
    const toolbar = make('div','',root,'canvas-toolbar');
    button('保存副本',toolbar,()=>guarded(save),'canvas-save','save'); button('撤回草稿上一步',toolbar,()=>guarded(()=>historyStep('undo')),'canvas-undo','undo');
    button('重做草稿编辑组',toolbar,()=>guarded(()=>historyStep('redo')),'canvas-redo','redo');
    button('刷新页图',toolbar,()=>view.kind==='pptx'?renderPage():guarded(renderPage),'canvas-preview','scan');
    const page = make('input','',toolbar,'canvas-preview-page'); page.type='number';page.min=1;page.max=100;page.value=1;page.setAttribute('aria-label','打印预览页码');page.onchange=()=>{previewPage=Number(page.value);};
    button('对话侧栏',toolbar,()=>{root.classList.toggle('office-ai-open');$('canvas-toggle-ai').setAttribute('aria-expanded',String(root.classList.contains('office-ai-open')));},'canvas-toggle-ai','panel-right-open').setAttribute('aria-expanded','false');
    const work = make('div','',root,'canvas-work'), main = make('section','',work,'canvas-main');main.setAttribute('aria-label','原生文档部件编辑');
    const bar=make('div','',main,'canvas-editbar');make('span','未选中',bar,'canvas-position');
    const kind=make('select','',bar,'canvas-type');kind.setAttribute('aria-label','单元格类型');for(const [value,label] of [['text','文本'],['number','数值'],['formula','受限公式']]){const opt=make('option',label,kind);opt.value=value;}
    const input=make('textarea','',bar,'canvas-value');input.rows=2;input.setAttribute('aria-label','编辑选中内容'); input.maxLength=4000;
    input.oninput=()=>{pendingField=true;status({});};kind.onchange=()=>{pendingField=true;input.readOnly=selected?.editable===false&&!(selected?.formula_editable&&kind.value==='formula');status({});if(kind.value==='formula')notice('支持同表受限数值/文本函数；区域小数用NUMBERVALUE明确分隔符，空白不按0计算。');};
    input.onkeydown=event=>{if(event.isComposing||event.keyCode===229)return;if(event.key==='Escape'){event.preventDefault();pendingField=false;if(selected)setSelected(selected);if(view.kind==='pptx')document.querySelector('dialog[data-slide-editor]')?.close();}if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();guarded(()=>apply(view.kind==='xlsx'?kind.value:'text',input.value)).catch(error=>notice(error.message));}};
    button('采用编辑',bar,()=>guarded(()=>apply(view.kind==='xlsx'?kind.value:'text',input.value)),'canvas-apply','check');
    button('引用给AI',bar,()=>guarded(()=>referenceAI(true)),'canvas-reference','quote');
    button('加粗',bar,()=>guarded(async()=>{checkField();await apply('format',{bold:$('canvas-bold').getAttribute('aria-pressed')!=='true'});}),'canvas-bold','bold');
    button('斜体',bar,()=>guarded(async()=>{checkField();await apply('format',{italic:$('canvas-italic').getAttribute('aria-pressed')!=='true'});}),'canvas-italic','italic');
    const size=make('select','',bar,'canvas-size');size.setAttribute('aria-label','字号');for(const value of [8,10,11,12,14,16,18,20,24,28,32,36,48,60,72]){const opt=make('option',String(value),size);opt.value=value;}size.onchange=()=>guarded(async()=>{checkField();await apply('format',{size:Number(size.value)});}).catch(error=>notice(error.message));
    const windowInput=make('input','',bar,'canvas-window');windowInput.setAttribute('aria-label','跳转单元格');windowInput.placeholder='A1';windowInput.onkeydown=event=>{if(event.key==='Enter'&&!event.isComposing){event.preventDefault();const position=pos(windowInput.value.toUpperCase());if(!position){notice('请输入A1形式的位置');return;}guarded(()=>loadItem(view.item_id,...position)).catch(error=>notice(error.message));}};
    const rangeActions=make('span','',bar,'canvas-range-actions');rangeActions.title='Shift+单击或Shift+方向键选择矩形区域';
    button('复制区域',rangeActions,()=>guarded(copyRange),'canvas-copy-range','copy');
    button('粘贴区域',rangeActions,()=>pasteRange(),'canvas-paste-range','paste');
    button('行列尺寸',rangeActions,geometryDialog,'canvas-geometry','sliders');
    button('冻结窗格',rangeActions,panesDialog,'canvas-panes','freeze');
    make('span','',rangeActions,'canvas-pane-state').setAttribute('role','status');
    const surface=make('div','',main,'canvas-surface');
    wireSelection(surface);
    surface.addEventListener('scroll',scrollWindow,{passive:true});
    surface.addEventListener('wheel',wheelWindow,{passive:false});
    surface.onpaste=event=>{if(view?.kind!=='xlsx'||!selected||event.isComposing||busy)return;event.preventDefault();try{pasteRange(event.clipboardData?.getData('text/plain')||'');}catch(error){notice(error.message);}};
    surface.oncopy=event=>{if(view?.kind!=='xlsx'||!selected||busy)return;event.preventDefault();guarded(copyRange).catch(error=>notice(error.message));};
    make('nav','',main,'canvas-items').setAttribute('aria-label','工作表或幻灯片');
    work.append($('floating-conversation'));$('floating-conversation').append($('floating-controls'));
    const collapse=button('收起对话侧栏',$('floating-conversation-heading'),()=>{
      root.classList.remove('office-ai-open');
      const toggle=$('canvas-toggle-ai');toggle.setAttribute('aria-expanded','false');toggle.focus();
    },'canvas-close-ai','panel-right-close');
    collapse.setAttribute('aria-controls','floating-conversation');
    const aiPanel=make('section','',$('floating-scroll'),'canvas-ai-panel');aiPanel.setAttribute('aria-label','当前草稿修改建议');
    make('p','交互编辑后请保存副本，再选择该副本交给Agent继续。当前版本不把手工草稿自动发送给模型。',$('floating-controls'),'canvas-ai-note').hidden=true;
    const footer=make('footer','',root,'canvas-footer');make('span','正在打开',footer,'canvas-state').setAttribute('role','status');
    const recovery=make('span','',footer,'canvas-recovery-actions');recovery.hidden=true;
    button('恢复本机草稿',recovery,()=>guarded(()=>checkpoint('restore')),'canvas-draft-restore');
    button('本次使用原稿',recovery,()=>guarded(async()=>{status(await call('canvas_use_original',host.document.id,state.draft_revision));await loadItem(view.item_id);notice(state.recovery_status);}),'canvas-use-original');
    const draftDetails=make('details','',footer,'canvas-draft-details');const draftSummary=make('summary','本机草稿',draftDetails);const draftPanel=make('div','',draftDetails,'canvas-draft-panel');make('span','',draftPanel,'canvas-draft-status').setAttribute('role','status');button('保存本机草稿',draftPanel,()=>guarded(()=>checkpoint('save')),'canvas-draft-save');make('p','',draftPanel,'canvas-draft-location');draftPanel.onkeydown=event=>{if(event.key==='Escape'&&!event.isComposing){event.preventDefault();draftDetails.open=false;draftSummary.focus();}};
    button('继续编辑副本',footer,async()=>{if(!await beforeSwitch())return;await call('floating_select_output',state.last_output.id);await window.OfficeFloating.refresh();},'canvas-continue').hidden=true;
    const outputDetails=make('details','',footer,'canvas-output-details');outputDetails.hidden=true;make('summary','副本位置',outputDetails);make('p','',outputDetails,'canvas-output-path');
    const details=make('details','',footer);make('summary','能力范围',details);make('p','',details,'canvas-scope');
    make('p','',root,'canvas-feedback').setAttribute('role','status');
  }
  function close() {
    automaticAttempts.clear();automaticSaving=null;
    slideRequest++;slideRendering=null;slideDesired=null;slideMessage='';
    paneObserver?.disconnect();paneObserver=null;
    $('canvas-surface')?.removeEventListener?.('scroll',scrollWindow);$('canvas-surface')?.removeEventListener?.('wheel',wheelWindow);wheelIntent=null;
    clearTimeout(scrollTimer);scrollTimer=null;prefetch=null;viewRequest++;windowLoading=false;scrollSuppressedUntil=0;scrollPosition={top:0,left:0};
    epoch++;identity='';opening=false;busy=false;selected=null;rangeAnchor=rangeFocus=null;pointerSelection=null;pendingField=false;preview=null;lastError='';state=null;view=null;previewPage=1;aiSignature='';
    for(const node of document.querySelectorAll('dialog[data-canvas-document]'))node.close();
    const root=$('floating-root');if(!root?.classList.contains('office-active'))return;
    $('floating-body')?.append($('floating-conversation'));(matchMedia('(max-width:899px)').matches?root:$('floating-conversation'))?.append($('floating-controls'));
    for(const id of ['canvas-toolbar','canvas-work','canvas-footer','canvas-feedback','canvas-ai-note','canvas-ai-panel','canvas-close-ai'])$(id)?.remove();
    root.classList.remove('office-active','office-ai-open');
  }
  async function update(next) {
    const supported=next.document?.ready&&/\.(xlsx|pptx)$/i.test(next.document.name);
    if(!supported){if(active())close();return;}
    host=next;const key=next.document.id+':'+next.document.version;
    if(fallback===key)return;
    if(identity===key){if(next.canvas)status(next.canvas);await driveAutomatic();return;}
    if(opening)return;close();identity=key;opening=true;const mine=epoch;mount();
    try{const result=await call('canvas_open',next.document.id,next.document.version);if(mine!==epoch)return;state=result;view=result.view;render();if(view.kind==='pptx')renderPage();}
    catch(error){if(mine===epoch){if(!view){close();fallback=key;if($('floating-status'))$('floating-status').textContent='当前文档使用原有结构视图：'+error.message;}else{notice('页图未完成：'+error.message+'；可查看文字片段，原稿未改动');}}}
    finally{if(mine===epoch)opening=false;}
  }
  async function beforeSwitch(){
    if(!active()||!state?.dirty&&!pendingField)return true;
    if(state.publishing||state.previewing||busy){notice('请等待当前步骤结束');return false;}
    return new Promise(resolve=>dialog('处理未保存编辑',node=>{make('p','继续切换会放弃当前内存编辑；主动保存过的本机草稿保留，可重新打开准确原文件后恢复。',node);let answered=false;button('返回保存副本',node,()=>{answered=true;node.close();resolve(false);});button('放弃并继续',node,async()=>{await call('canvas_discard',host.document.id,state.draft_revision,true);pendingField=false;answered=true;node.close();resolve(true);});node.addEventListener('close',()=>{if(!answered)resolve(false);});}));
  }
  async function messageContext(text){
    checkField();
    if(!/^(?:进度|到哪|停止|停一下)/.test(text)&&selected)await referenceAI(false);
    return {revision:state.draft_revision,reference:(state?.ai?.read_reference||state?.ai?.reference)?.id||null};
  }
  window.OfficeNative={active,update,beforeSwitch,close,submit,messageContext,referenceSelection:()=>{
    if(!selected)throw Error('请先从主文档选中单元格、范围或幻灯片文字。');
    return referenceAI();
  }};
  window.addEventListener('officeauth:unauthenticated',close);
  window.addEventListener('beforeunload',event=>{if(active()&&(state?.dirty||pendingField)){event.preventDefault();event.returnValue='';}});
})();
