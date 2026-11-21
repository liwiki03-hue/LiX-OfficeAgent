(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  let snapshot = null, cursor = 0, polling = false, epoch = 0, submitting = false, pending = null;
  let operation = 0, feedbackOperation = 0;
  let accessDialog = null, autoAction = false;
  const automaticAttempts = new Set();
  const accessLabels = {save_copy:'自动存副本',edit:'仅修改草稿'};
  let planSignature = '', outputSignature = '';
  let adopted = new Set(), seenMessages = new Set(), seenSkills = new Set(), viewVersion = '', modelHome = null, timer = null, lastPoll = 0;
  const make = (tag, text, parent, id) => {
    const node = document.createElement(tag);
    if (text) node.textContent = text;
    if (id) node.id = id;
    if (parent) parent.append(node);
    return node;
  };
  const button = (text, parent, action, id) => {
    const node = make('button', text, parent, id); node.type = 'button';
    node.addEventListener('click', action); return node;
  };
  const api = () => window.OfficeAuth?.wrapApi(window.pywebview?.api);
  const access = () => snapshot?.access || {id:'',level:'save_copy',serial:0};
  function accessSettings() {
    if (!snapshot?.document?.ready || accessDialog?.open) return;
    const owner = {id:snapshot.document.id,version:snapshot.document.version,epoch:snapshot.epoch}, previous = document.activeElement;
    const node = make('dialog','',document.body); node.className='paged-dialog document-access-dialog'; node.setAttribute('aria-label','编辑方式'); accessDialog=node;
    make('h2','编辑方式',node); make('p',snapshot.document.name,node);
    const choices = make('div','',node); choices.className='document-access-choices';
    for (const [level,description] of [['save_copy','发送要求后直接修改，自动生成验证副本。'],['edit','直接修改草稿，由你选择何时保存。']]) {
      const label=make('label','',choices),input=make('input','',label); input.type='radio';input.name='document-access';input.value=level;input.checked=access().level===level;
      const text=make('span','',label);make('strong',accessLabels[level],text);make('small',description,text);
    }
    make('p','原文件保留。停止后不再处理，重新发送要求可继续。',node).className='document-access-scope';
    const message=make('p','',node);message.setAttribute('role','alert');
    const actions=make('div','',node);actions.className='paged-dialog-actions';
    button('取消',actions,()=>node.close());
    if(access().id)button('停止本次修改',actions,async()=>{
      if(snapshot?.document?.id!==owner.id||snapshot?.epoch!==owner.epoch){message.textContent='文档已切换，请关闭后重新选择。';return;}
      const result=await call('floating_stop');if(result?.ok)node.close();else message.textContent=result?.error||'未取得停止结果，请核对当前状态。';
    });
    const apply=button('应用',actions,async()=>{
      if(snapshot?.document?.id!==owner.id||snapshot?.epoch!==owner.epoch){message.textContent='文档已切换，请关闭后重新选择。';return;}
      const level=node.querySelector('input:checked').value;apply.disabled=true;
      try{
        const result=await api().floating_set_access(owner.id,owner.version,owner.epoch,level,true);
        if(!result?.ok)throw Error(result?.error||'编辑方式未更新');
        if(snapshot?.document?.id===owner.id&&snapshot.epoch===owner.epoch){snapshot.access=result.access;render(snapshot);}
        node.close();await refresh();
      }catch(error){message.textContent=error.message;apply.disabled=false;}
    });apply.className='paged-primary';
    node.addEventListener('close',()=>{if(accessDialog===node)accessDialog=null;node.remove();if(previous?.isConnected)previous.focus();});node.showModal();
  }
  async function openOutput(row) {
    if(window.OfficePaged?.active()&&!await window.OfficePaged.beforeSwitch())return;
    if(window.OfficeNative?.active()&&!await window.OfficeNative.beforeSwitch())return;
    const result=await call('floating_select_output',row.id,true,access().id||null);
    if(result?.ok)$('floating-document').replaceChildren();
  }
  async function driveAccess(state) {
    if(autoAction||state.busy||!state.document?.ready||state.stage==='uncertain'||state.stage==='stopped'||document.querySelector('dialog[open]'))return;
    const created=state.outputs.find(row=>!row.undone&&row.continue_editing?.serial===state.access?.serial&&row.source_version===state.document.version);
    const plan=state.access?.id&&state.access.level==='save_copy'&&!state.paged&&!state.canvas&&state.plans.length===1?state.plans[0]:null;
    const key=created?'open:'+created.id:plan?state.access.id+':plan:'+plan.plan.id:null;
    if(!key||automaticAttempts.has(key))return;
    automaticAttempts.add(key);autoAction=true;
    try{
      if(created)await openOutput(created);
      else await call('floating_confirm',state.session_id,plan.plan.id,plan.revision,state.document.version,false,state.access.id);
    }finally{autoAction=false;}
  }
  function feedback(text) { if ($('floating-status')) { $('floating-status').textContent = text; $('floating-status').dataset.feedbackKind = ''; } }
  async function call(name, ...args) {
    const mine = epoch, documentId = snapshot?.document?.id, ownOperation = ++operation;
    try {
      const result = await api()[name](...args);
      if (mine !== epoch) return null;
      if (['floating_read','floating_preview_output','floating_protect','floating_restore_drafts','floating_submit','floating_reference'].includes(name) && documentId !== snapshot?.document?.id) return null;
      if (!result?.ok && ownOperation >= feedbackOperation) { feedbackOperation = ownOperation; feedback(result?.reason==='cancelled'?'已取消选择文件':result?.error || result?.reason || '此步骤未完成；输入已保留。'); }
      else if (result?.ok && !result.cancelled && ownOperation >= feedbackOperation && ['floating_submit','floating_read','floating_choose','floating_admit_drop','floating_start_again','floating_protect','floating_resolve_protection'].includes(name)) { feedbackOperation = ownOperation; feedback(''); }
      await refresh();
      return result;
    } catch (_) { if (mine === epoch && ownOperation >= feedbackOperation) { feedbackOperation = ownOperation; feedback('未取得宿主结果；请先核对当前状态，不要重复确认。'); } return null; }
  }
  function positionBack() {
    const back = $('floating-return');
    if (!back || back.hidden) return;
    const composer = $('simple-compose-area') || $('prompt')?.closest('.composer');
    const rect = composer?.getBoundingClientRect();
    const bottom = rect && rect.height ? Math.max(24, innerHeight - rect.top + 12) : 24;
    back.style.bottom = Math.round(bottom) + 'px';
  }
  function secondary(on, history = false) {
    const returningToHistory = history && !$('floating-root').hidden && typeof currentView !== 'undefined' && currentView === 'history';
    window.OfficeWindowChrome?.secondary(on);
    $('workspace').classList.toggle('floating-primary', !on);
    $('floating-root').hidden = on;
    $('floating-return').hidden = !on;
    for (const id of ['simple-model-menu', 'simple-model-state']) {
      const node = $(id); if (node) (on ? modelHome : $('floating-models')).append(node);
    }
    if (history) { window.OfficeWorkspace?.open('history'); $('simple-history-documents')?.click(); ensureOutputRecords(); if(snapshot)renderOutputs(snapshot); }
    if (returningToHistory && typeof refreshHistory === 'function' && !historyLoading && !historySelectingFile) refreshHistory();
    if (!on && $('model-mode').value !== 'agent') {
      $('model-mode').value = 'agent'; $('model-mode').dispatchEvent(new Event('change'));
    }
    if (on) positionBack();
  }
  function addEntry(text, kind = 'assistant') {
    const list = $('floating-timeline'), scroll = $('floating-scroll');
    const near = scroll.scrollHeight - scroll.scrollTop - scroll.clientHeight < 90;
    const entry = make('article', '', list); entry.className = 'floating-message ' + kind;
    make('strong', kind === 'user' ? '你的要求' : '处理进展', entry);
    make('p', text, entry);
    while (list.children.length > 160) list.firstElementChild.remove();
    if (near) scroll.scrollTop = scroll.scrollHeight;
    else $('floating-latest').hidden = false;
    return entry;
  }
  function showSkillRead(receipt,state){
    if(!receipt||receipt.document_id!==state.document?.id||receipt.source_version!==state.document?.version||receipt.session_id!==state.session_id||receipt.epoch!==state.epoch||receipt.instruction_revision!==state.revision||receipt.draft_revision!==state.paged?.draft_revision)return;
    if(typeof receipt.id!=='string'||!/^[a-f0-9]{64}$/.test(receipt.body_sha256)||!Number.isInteger(receipt.characters))return;
    const key=[receipt.id,receipt.body_sha256,receipt.session_id,receipt.draft_revision,receipt.instruction_revision].join(':');
    if(seenSkills.has(key))return;seenSkills.add(key);if(seenSkills.size>64)seenSkills.delete(seenSkills.values().next().value);
    if(state.conversation?.turns?.length)return;
    const entry=addEntry('已读取参考：'+(receipt.display_name||receipt.id));entry.dataset.skillRead=receipt.id;
    const detail=make('details','',entry);make('summary','参考读取详情',detail);
    make('p','仅代表宿主准备参考；不代表已发送、已执行或质量通过。',detail);
    make('p',receipt.id+' · '+receipt.characters+' 字符 · SHA256 '+receipt.body_sha256,detail);
  }
  const sourceViewLabel = () => snapshot?.document?.ready ? '已读取原稿 · 结构文本视图' : '上次读取版本 · 当前源文件需重新读取';
  function showView(view, label = null, outputId = null) {
    const pane = $('floating-document'); pane.replaceChildren();
    pane.dataset.outputId = outputId || '';
    make('p', (label || sourceViewLabel()) + ' · 不是实时 Word 分页画布', pane).className = 'floating-note';
    if (snapshot?.view) button(snapshot.document?.ready ? '返回已读取原稿' : '查看上次读取版本', pane, () => showView(snapshot.view));
    if (view.blocks) {
      for (const block of view.blocks) {
        const entry = make('section', '', pane); entry.id = 'floating-block-' + block.index;
        if (block.kind === 'paragraph') {
          make('p', block.text || '（空段落）', entry);
          if (block.locator_id && (outputId || snapshot?.document?.ready)) {
            entry.dataset.locator = block.locator_id;
            if (!outputId) button('引用到对话', entry, () => reference(block.locator_id));
            const label = make('label', '', entry); label.className = 'floating-protection';
            const check = make('input', '', label); check.type = 'checkbox';
            check.checked = snapshot?.protected.includes(block.paragraph_index) || false;
            make('span', '保护此段，不作修改', label);
            check.onchange = async () => { const result = await call('floating_protect', block.locator_id, check.checked); if (!result?.ok) check.checked = !check.checked; };
          }
        } else {
          const table = make('table', '', entry);
          for (const row of block.rows || []) { const tr = make('tr', '', table); row.forEach(text => make('td', text, tr)); }
          if (block.rows_truncated) make('p', '本表仍有未读取部分。请在更多工具中按区域读取。', entry);
        }
      }
      if (view.has_more && (outputId || snapshot?.document?.ready)) button('继续读取下一段范围', pane, async () => {
        const result = outputId ? await call('floating_preview_output', outputId, true, view.next_offset) : await call('floating_read', snapshot.read_token, true, view.next_offset);
        if (result?.ok) showView(result.view, label, outputId);
      });
    } else if (view.slides) {
      make('p','PPTX 有界大纲；此视图不还原复杂图形，也不是可编辑画布。',pane);
      for (const slide of view.slides) make('h3','第 '+(slide.slide_index+1)+' 页 · '+(slide.title||'未命名页面'),pane);
      if(view.has_more&&(outputId||snapshot?.document?.ready))button('继续读取后续幻灯片',pane,async()=>{const result=outputId?await call('floating_preview_output',outputId,true,view.next_offset):await call('floating_read',snapshot.read_token,true,view.next_offset);if(result?.ok)showView(result.view,label,outputId);});
    } else {
      for (const sheet of view.sheets || []) {
        make('h3', sheet.title, pane);
        const table = make('table', '', pane);
        for (const cell of sheet.cells || []) {
          const row = make('tr', '', table);
          make('th', cell.cell || cell.coordinate || '', row);
          make('td', String(cell.formula_text ?? cell.value ?? cell.formula ?? ''), row);
          if (cell.locator_id && (outputId || snapshot?.document?.ready)) {
            row.dataset.locator = cell.locator_id;
            const td = make('td', '', row), label = make('label', '', td);
            if (!outputId) button('引用到对话', td, () => reference(cell.locator_id));
            const check = make('input', '', label); check.type = 'checkbox';
            check.checked = snapshot?.protected_cells.some(pair => pair[0] === sheet.title && pair[1] === cell.coordinate) || false;
            make('span', '保护此格', label);
            check.onchange = async () => { const result = await call('floating_protect', cell.locator_id, check.checked); if (!result?.ok) check.checked = !check.checked; };
          }
        }
        if (sheet.cells_truncated) make('p', '本工作表仅展示有界摘录，仍有未读取单元格。', pane);
      }
      make('p', '有界工作表摘录；公式与缓存不是已核算结果，未显示范围不能解释为空。', pane);
    }
  }
  async function reference(identifier) {
    const result = await call('floating_reference', snapshot.document.id, snapshot.document.version, identifier);
    if (!result?.ok) return;
    const input = $('floating-input');
    input.value = (input.value ? input.value + '\n' : '') + '针对' + result.label + '：';
    $('floating-root').dataset.tab = 'conversation'; input.focus();
  }
  function renderPlans(state) {
    const signature = JSON.stringify([state.plans, state.busy, [...adopted]]);
    if (signature === planSignature) return;
    planSignature = signature;
    const pane = $('floating-plans'); pane.replaceChildren();
    const ids = new Set(state.plans.map(row => row.plan.id));
    adopted = new Set([...adopted].filter(id => ids.has(id)));
    for (const row of state.plans) {
      const plan = row.plan, card = make('article', '', pane); card.className = 'floating-diff';
      const actions = window.OfficeAgentElements.edit(card, {title:plan.target || '文字修改', before:plan.before, after:plan.after, risks:plan.risks || [], status:'待审阅 · 尚未写入 · 要求版本 ' + row.revision});
      if (plan.location?.document_version === state.document.version) button('定位原文', actions, () => {
        showView(snapshot.view); $('floating-root').dataset.tab = 'document';
        const target = [...$('floating-document').querySelectorAll('[data-locator]')].find(node => node.dataset.locator === plan.location.locator_id);
        if (target) { target.tabIndex = -1; target.focus(); target.scrollIntoView({block:'center'}); }
        else feedback('此位置不在当前有界摘录中，请先读取对应范围。');
      });
      const accept = button(adopted.has(plan.id) ? '已采用建议' : '采用建议', actions, () => { adopted.add(plan.id); renderPlans(snapshot); });
      accept.disabled = adopted.has(plan.id);
      button('不采用', actions, () => call('floating_confirm', state.session_id, plan.id, row.revision, state.document.version, false));
      const confirm = button('确认生成副本…', actions, async () => {
        if (!await ask('请核对当前差异。将按要求版本 ' + row.revision + ' 备份、生成新副本、重开校验并渲染。不会覆盖原件。确认执行这一组？')) return;
        confirm.disabled = true;
        await call('floating_confirm', state.session_id, plan.id, row.revision, state.document.version, true);
      });
      confirm.className = 'floating-primary-button'; confirm.disabled = state.busy || !adopted.has(plan.id);
    }
  }
  function renderOutputs(state) {
    const pane = $('floating-outputs'); if(!pane)return;
    const signature = JSON.stringify([state.outputs, state.revision, state.document?.version, state.paged?.editable,state.paged?.proposals,state.conversation]);
    if (signature === outputSignature) return;
    outputSignature = signature;
    pane.replaceChildren();
    const records=$('floating-turn-records');
    if(records){
      records.replaceChildren();
      for(const turn of state.conversation?.turns||[]){
        const entry=make('article','',records);entry.className='floating-history-turn';entry.dataset.recordTurn=turn.id;
        const title=make('h4',turn.question,entry);title.title=turn.question;
        const status={failed:'失败',stopped:'已停止',interrupted:'未完成',uncertain:'待核对',waiting:'待处理',running:'处理中'}[turn.status];
        if(status)make('small',status,entry);
        const details=make('details','',entry);make('summary','过程与依据',details);
        make('p',turn.undo?.message||turn.text||turn.answer,details);
        if(turn.model)make('p','模型：'+turn.model.model_id+' · '+(turn.model.connection_name||turn.model.connection_id),details);
        for(const step of turn.steps||[])make('p',step.text,details);
        for(const change of turn.changes||[]){make('p','原文：'+change.before,details);make('p','修改：'+change.after,details);}
        for(const reference of turn.references||[])make('p',[reference.scope,reference.location].filter(Boolean).join(' · '),details);
        if(turn.undo?.detail)make('p',turn.undo.detail,details);
        for(const receipt of turn.skills||[]){
          const skill=make('details','',entry);skill.dataset.skillRead=receipt.id;
          make('summary','参考：'+receipt.display_name,skill);
          make('p','仅代表宿主准备参考；不代表已发送、已执行或质量通过。',skill);
          make('p',receipt.id+' · '+receipt.characters+' 字符 · SHA256 '+receipt.body_sha256,skill);
        }
      }
    }
    const proposals=$('floating-proposal-records');
    if(proposals){
      proposals.replaceChildren();
      for(const proposal of state.paged?.proposals||[]){
        const entry=make('article','',proposals);entry.className='floating-history-turn';
        make('h4',proposal.status==='adopted'?'已采用的修改建议':'修改建议记录',entry);
        if(proposal.model)make('p','模型：'+proposal.model.model_id+' · '+(proposal.model.connection_name||proposal.model.connection_id),entry);
        for(const edit of proposal.edits||[proposal]){make('p','原文：'+edit.old,entry);make('p','修改：'+edit.new,entry);}
        for(const change of proposal.allowed_changes||[])make('p',change.before+' → '+change.after,entry);
      }
    }
    for (const row of state.outputs) {
      const card = make('article', '', pane); card.className = 'floating-result';
      const simplified=row.kind==='paged_editable_copy';
      const current=state.document?.ready&&row.output_sha256===state.document.version;
      make('h3', row.undone ? '副本已撤销' : current ? state.paged?.editable?'当前编辑副本':'已打开副本' : simplified ? '简化文字副本已生成' : '副本已生成并通过验证', card);
      make('p', String(row.output || '').split(/[\\/]/).pop(), card);
      if(simplified&&!row.undone)make('p',current?'格式已简化；原文件保留。':'正文文字已核对；原样式与分页已重建，打开后再核对编辑能力。',card);
      const details = make('details', '', card); make('summary', '路径与验证详情', details);
      make('p', row.output, details);
      make('p', row.render_diff ? '指定渲染器闸门已通过；不代表 Word 人工质量验收。' : '未提供渲染结果，请核对审计。', details);
      if (!row.undone) {
        if (row.source_version === state.document?.version && row.instruction_revision < state.revision) make('p', '这是上一版要求的结果，有新要求待处理；尚未完成最新要求。', card);
        button('查看副本内容…', card, async () => {
          const response = await call('floating_preview_output', row.id, true);
          if (response?.ok) {
            if(window.OfficePaged?.active())window.OfficePaged.previewStructure(response.view,'已验证副本 · 只读结构文本');
            else {showView(response.view, '已验证副本 · 只读', row.id); $('floating-root').dataset.tab = 'document';}
          }
        });
        if(!current)button(simplified?'打开副本编辑…':'基于此副本继续…', card, () => openOutput(row));
        button('撤销本次副本…', card, async () => {
          if (await ask('仅删除本次审计登记且内容未改变的副本，保留原稿与备份。确认撤销？')) call('floating_rollback', row.id, true);
        });
      }
    }
  }
  function render(state) {
    const doc = state.document;
    const permission=$('floating-access');
    if(permission){permission.textContent=accessLabels[state.access?.level]||accessLabels.save_copy;permission.disabled=!doc?.ready;permission.dataset.level=state.access?.level||'save_copy';permission.title='编辑方式：'+permission.textContent;permission.setAttribute('aria-label',permission.title);}
    $('floating-root').dataset.document = doc ? doc.ready ? 'ready' : 'awaiting' : 'none';
    $('floating-title').textContent = doc?.name || 'LiX · 文档助手';
    const phase = state.open_request?.phase;
    const opening = ['receiving','admission_pending','parsing'].includes(phase);
    $('floating-identity').textContent = doc ? (state.stage === 'changed' ? '源文件已改变，请重新打开' : doc.ready ? doc.editable ? '' : '只读' : doc.source_origin === 'native_drop' ? '正在打开…' : '等待读取同意') : '本机打开，发送才调用模型';
    $('floating-read').hidden = !doc || doc.ready || doc.source_origin === 'native_drop';
    $('floating-composer').hidden = false;
    $('floating-drop').hidden = !state.draft_drop_pending&&(!state.drop_pending || !['failed','blocked','cancelled'].includes(phase));
    $('floating-drop').textContent=state.draft_drop_pending?'先处理草稿，再打开拖入的文件…':'重试本地打开此文件';
    $('floating-send').disabled = submitting || !doc || (!doc.ready && phase !== 'parsing');
    $('floating-send').textContent = state.busy ? '追加要求' : phase === 'parsing' ? '保存要求，解析后处理' : '发送要求';
    $('floating-send').setAttribute('aria-label', $('floating-send').textContent);
    $('floating-send').title = $('floating-send').textContent;
    $('floating-stop').disabled = !state.busy && !opening;
    $('floating-resolve-protection').hidden = !state.needs_clarification;
    const stages = {rolled_back:'副本已撤销，原稿保留', verified:state.paged||state.canvas?'副本已验证，继续输入即可修改':'副本已验证；可明确选择版本继续', uncertain:'结果不确定；请核对产物和审计，不要重复执行', failed:'本轮未完成；要求与记录保留', changed:'来源已改变，请重新选择并读取', stopped:'已停止后续步骤；远端请求不保证已取消', needs_clarification:'请在文档视图勾选需要保护的位置，再继续生成方案'};
    const currentOutputs = state.outputs.filter(row => !row.undone && row.source_version === doc?.version);
    const pendingResult = state.stage === 'verified' && currentOutputs.length > 0 && currentOutputs.every(row => row.instruction_revision < state.revision);
    const back = $('floating-return');
    if (back) {
      const awaitingConfirmation = Boolean(state.needs_clarification || state.plans.length);
      const awaitingReview = Boolean(pendingResult || state.outputs.some(row => !row.undone));
      const running = Boolean(state.busy || opening);
      const attention = awaitingConfirmation || awaitingReview;
      back.dataset.state = attention ? 'attention' : running ? 'running' : 'idle';
      back.dataset.attention = String(attention);
      const detail = awaitingConfirmation && awaitingReview ? '有待确认方案和待验收副本'
        : awaitingConfirmation ? '有待确认方案' : awaitingReview ? '有待验收副本' : running ? '任务运行中' : '';
      const label = '返回悬浮文档助手' + (detail ? '，' + detail : '');
      back.setAttribute('aria-label', label);
      back.title = label;
      positionBack();
    }
    const openingText = {receiving:'正在打开…',admission_pending:'正在检查文件…',parsing:'正在打开…',failed:state.open_request?.error || '打开失败，可重试',cancelled:'已取消打开，输入保留'};
    $('floating-progress').textContent = openingText[phase] || (state.needs_clarification ? stages.needs_clarification : state.busy ? (state.stage === 'executing' || state.paged?.publishing || state.canvas?.publishing ? '正在写副本与验证；补充要求将排队' : '正在生成方案；可以继续追加要求') : state.plans.length ? '方案待审阅，尚未写入' : pendingResult ? '上一版副本已验证；新要求待明确选择版本后处理' : stages[state.stage] || '就绪');
    const publication = state.paged?.publication_result;
    const manualVersion = publication?.preview_latest_draft && !publication.conversation_id
      && !state.paged.uncertain && !state.paged.publication_rolled_back;
    const turnInView=state.conversation?.turns?.some(turn=>turn.document_id===doc?.id&&['running','waiting','complete'].includes(turn.status));
    $('floating-progress').hidden = Boolean(!openingText[phase]&&!state.plans.length&&!pendingResult&&!state.needs_clarification
      && (manualVersion&&state.stage==='verified'&&!state.busy || turnInView&&(state.busy||state.stage==='verified')));
    $('floating-progress').classList.add('state-badge');
    if(state.paged?.requirements?.length&&state.paged.requirements.every(row=>row.cancelled)){
      $('floating-progress').hidden=false;
      $('floating-progress').textContent=state.paged.publishing?'副本事务已开始，将按原快照完成；取消仅作用于后续处理':state.busy?'要求已取消，等待已发请求结束；不再派发后续步骤':'本轮要求已全部取消；草稿和副本保留';
    }
    $('floating-progress').dataset.state = state.stage === 'uncertain' ? 'unknown' : state.stage === 'failed' ? 'error' : state.busy || opening ? 'running' : state.plans.length ? 'waiting' : state.stage === 'verified' ? 'done' : 'ready';
    $('floating-revisions').textContent = state.revision ? '要求已收到 ' + state.revision + ' · 已纳入模型请求 ' + state.applied_revision : '';
    $('floating-actual-model').textContent = state.actual_model ? '本轮实际使用：' + state.actual_model.model_id + ' · ' + state.actual_model.connection_name : '';
    $('floating-open-trace').textContent = (state.open_trace || []).filter(row => row.open_request_id === state.open_request?.id).map(row => row.stage + ' · ' + row.elapsed_ms + ' ms · ' + row.code).join('\n');
    $('floating-new-conversation').hidden = !doc?.ready || !state.conversation?.at_limit;
    $('floating-new-conversation').disabled = submitting;
    if($('floating-old-conversations'))$('floating-old-conversations').hidden = !doc?.ready;
    const identity = doc ? doc.id + ':' + doc.version : '';
    if (viewVersion !== identity) { viewVersion = identity; $('floating-document').replaceChildren(); }
    renderPlans(state); if($('floating-root').hidden)renderOutputs(state);
  }
  function ensureOutputRecords(){
    if($('floating-outputs'))return $('floating-outputs');
    const parent=$('simple-document-records')||$('history-view');if(!parent)return null;
    const section=make('section','',parent,'floating-output-records');section.setAttribute('aria-label','当前文档结果记录');
    make('h3','本文件记录',section);
    button('查看旧对话',section,event=>showArchivedConversations(event.currentTarget),'floating-old-conversations').hidden=!snapshot?.document?.ready;
    make('div','',section,'floating-turn-records');make('div','',section,'floating-proposal-records');return make('div','',section,'floating-outputs');
  }
  async function refresh(automatic=false) {
    if (!window.OfficeFloating.active || polling || !window.OfficeAuth?.isAuthenticated()) return;
    const interval=document.hidden||['orb','hidden'].includes(snapshot?.presentation)?5000:snapshot?.busy||window.OfficeConversation?.pending()?350:1500;
    if (automatic && Date.now() - lastPoll < interval) return;
    lastPoll = Date.now();
    polling = true; const mine = epoch;
    try {
      const state = await api().floating_snapshot(cursor);
      if (mine !== epoch || !state?.ok) return;
      window.OfficeConversation?.connection(true);
      if(snapshot?.conversation?.context_id&&state.conversation?.context_id&&snapshot.conversation.context_id!==state.conversation.context_id){
        $('floating-timeline').replaceChildren();window.OfficeConversation?.reset();window.OfficePlugins?.clear();pending=null;epoch++;
      }
      if(snapshot?.epoch===state.epoch&&snapshot.access?.serial>state.access?.serial)state.access=snapshot.access;
      if (snapshot && snapshot.epoch !== state.epoch) { if(snapshot.conversation?.context_id!==state.conversation?.context_id){$('floating-timeline').replaceChildren();window.OfficeConversation?.reset();} adopted.clear(); seenMessages.clear(); seenSkills.clear(); automaticAttempts.clear();accessDialog?.close();feedback(''); }
      const switched = !snapshot || snapshot.document?.id !== state.document?.id || snapshot.document?.version !== state.document?.version;
      const availabilityChanged = snapshot?.document?.ready !== state.document?.ready;
      snapshot = state;
      window.OfficeConversation?.restore(state.conversation,state.paged);
      if(switched){seenSkills.clear();window.OfficePlugins?.clear();}
      if((switched||availabilityChanged)&&state.document?.ready)window.OfficePlugins?.reload();
      if (state.gap) {
        feedback('进展记录已压缩，当前要求、计划与产物已从宿主快照恢复。');
        for (const message of state.messages) if (!seenMessages.has(message.id)) { if(!window.OfficeConversation?.has(message.client_id))addEntry(message.text, 'user'); seenMessages.add(message.id); }
      }
      for (const envelope of state.events) {
        if (envelope.event_seq <= cursor) continue;
        const event = envelope.event;
        if (typeof handleEvent === 'function') handleEvent(event);
        if (event.type === 'instruction_received' && !seenMessages.has(event.payload.id)) { if ($('floating-status').dataset.feedbackKind === 'stop_requested') feedback(''); if(!window.OfficeConversation?.has(event.payload.client_id))addEntry(event.payload.text, 'user'); seenMessages.add(event.payload.id); }
        else if (event.type === 'instruction_applied') $('floating-revisions').textContent = '已将要求版本 ' + event.payload.revision + ' 纳入模型请求';
        else if (event.type === 'requirement_changed'&&!state.conversation?.turns?.length) addEntry('要求 '+event.payload.target_revision+(event.payload.action==='cancel'?'已取消后续处理；已采用草稿和已生成副本保留。':'已修改；请核对当前引用后主动处理。'));
        else if (event.type === 'stale_action_blocked'&&!state.conversation?.turns?.length) addEntry('要求或草稿已更新，旧响应或工具结果已作废；未沿用旧计划。');
        else if (event.type === 'document_ready') showView(event.payload.view);
        else if (event.type === 'answer'&&!state.conversation?.turns?.length) addEntry(event.text);
        else if (['draft_answer','canvas_ai_answer'].includes(event.type)&&!state.conversation?.turns?.length) addEntry(event.payload.text);
        else if (event.type === 'draft_copy_continued'&&!state.conversation?.turns?.length) feedback(event.payload.text);
        else if (event.type === 'skill_loaded') showSkillRead(event.payload,state);
        else if (['draft_model_error','draft_switch_requested'].includes(event.type)) feedback(event.payload.text);
        else if (event.type === 'draft_publish_failed'&&!event.payload.publication?.preview_latest_draft) feedback(event.payload.text);
        else if (event.type === 'draft_published'&&!event.payload.publication?.preview_latest_draft&&!state.conversation?.turns?.length) addEntry((event.payload.receipt.kind==='paged_editable_copy'?'简化文字副本已生成：':'草稿 '+event.payload.draft_revision+' 的副本已验证：')+String(event.payload.receipt.output).split(/[\\/]/).pop());
        else if (event.type === 'orb_position') { try { localStorage.setItem('office-agent.orb-position', JSON.stringify(event.payload)); } catch (_) {} }
        else if (['native_warning', 'admission_failed', 'stop_requested', 'read_denied', 'needs_clarification', 'new_round','instruction_waiting'].includes(event.type)) { feedback(event.payload.text); $('floating-status').dataset.feedbackKind = event.type; }
        else if (event.type === 'log' && ['safety', 'error', 'warn'].includes(event.tag)&&!state.conversation?.turns?.length) addEntry(event.text);
      }
      cursor = state.cursor; render(state);
      window.OfficePaged?.update(state).catch(error => feedback(error.message));
      window.OfficeNative?.update(state).catch(error => feedback(error.message));
      driveAccess(state).catch(error=>feedback(error.message));
      if ((switched || state.gap || availabilityChanged && !$('floating-document').dataset.outputId) && state.view) showView(state.view);
      if (switched && state.queued_for_next_version.length && !$('floating-input').value.trim()) {
        $('floating-input').value = state.queued_for_next_version.join('\n');
        feedback('写副本期间收到的补充已保留到输入框；重新读取后核对并主动发送。');
      }
    } catch(error) { window.OfficeConversation?.connection(false); throw error; } finally { polling = false; }
  }
  async function submit(recovery=null) {
    const restored=recovery&&typeof recovery.id==='string'&&typeof recovery.text==='string'?recovery:null;
    const input = $('floating-input'), text = restored?restored.text:input.value.trim();
    if (submitting || !text || !snapshot?.document) return;
    if(restored&&(restored.document!==snapshot.document.id||restored.version!==snapshot.document.version)){feedback('文档已改变，旧消息保留；没有发送当前草稿或重复执行。');return;}
    const atLimit=snapshot.conversation?.at_limit;
    const localOnly=snapshot.document.ready&&!window.OfficePaged?.active()&&!window.OfficeNative?.active()&&!window.OfficeModelSelector?.readyForRequest(false);
    if (!restored&&(!pending || pending.text !== text || pending.session !== snapshot.session_id || pending.document !== snapshot.document.id)) {
      const bytes = crypto.getRandomValues(new Uint8Array(16));
      pending = {id: Array.from(bytes, value => value.toString(16).padStart(2, '0')).join(''), text, session: snapshot.session_id, document:snapshot.document.id, version:snapshot.document.version, open:snapshot.document.open_request_id || null, context:snapshot.conversation?.context_id||null, replyTo:snapshot.conversation?.pending_scope||null};
    }
    if(!restored && !Object.hasOwn(pending,'plugin'))pending.plugin=window.OfficePlugins?.choice()||null;
    const sending = restored||pending, ownEpoch=epoch; submitting = true; render(snapshot);
    if(!localOnly)window.OfficeConversation?.echo(sending);
    let result,receipt;
    try {
      if(!localOnly)await window.OfficeConversation?.painted();
      const readContext=sending.readContext||(atLimit&&!sending.replyTo?{revision:0,reference:null}:window.OfficeNative?.active()?await window.OfficeNative.messageContext(sending.text):window.OfficePaged?.active()?await window.OfficePaged.messageContext(sending.text):{revision:0,reference:null});
      sending.readContext=readContext;window.OfficeConversation?.messageContext(sending.id,readContext);
      const options=sending.plugin?[localOnly,sending.replyTo||null,sending.plugin]:sending.replyTo?[localOnly,sending.replyTo]:localOnly?[true]:[];
      receipt=typeof api().conversation_begin==='function'?await api().conversation_begin(sending.id,sending.text,sending.document,sending.version,readContext.revision,readContext.reference,sending.context,...options):null;
      if(ownEpoch!==epoch||snapshot?.document?.id!==sending.document)throw Error('对话已切换，旧消息未继续处理。');
      if(localOnly&&!receipt)throw Error('请先在输入区选择有效的文档 Agent 模型；要求已保留。');
      if(receipt&&!receipt.ok)throw Error(receipt.error||'消息未受理；原话保留。');
      if(receipt?.ok && sending.plugin)window.OfficePlugins?.clear();
      if($('floating-status').dataset.feedbackKind==='model_required')feedback('');
      if(localOnly)window.OfficeConversation?.echo(sending);
      if(receipt?.turn)window.OfficeConversation?.acknowledge?.(sending.id,receipt.turn);
      if(receipt&&!receipt.handled)await window.OfficeConversation?.painted();
      const editText=receipt?.edit_text||sending.text;
      result=receipt?.handled?receipt:window.OfficeNative?.active() ? await window.OfficeNative.submit(sending.id,editText) : window.OfficePaged?.active() ? await window.OfficePaged.submit(sending.id,editText) : await call('floating_submit', sending.session, sending.id, editText, sending.open, sending.document, sending.version);
      if(result&&!result.ok)throw Error(result.error||'消息未受理；原话保留。');
      await refresh();
    } catch (error) {
      if(ownEpoch===epoch&&snapshot?.document?.id===sending.document){
        feedback(error.message);
        if(receipt?.reason==='model_required')$('floating-status').dataset.feedbackKind=receipt.reason;
        if(window.OfficePaged?.active())window.OfficePaged.reportError(error.message);
        window.OfficeConversation?.failed(sending.id,error.message,!result);
      }
    }
    if (result?.ok && ownEpoch===epoch && snapshot?.document?.id===sending.document) { if (input.value.trim() === text) input.value = ''; if(pending?.id===sending.id)pending = null; }
    submitting = false; if (snapshot) render(snapshot);
  }
  async function showArchivedConversations(trigger) {
    const owner=snapshot?.document;
    if(!owner?.ready)return;
    const result=await call('conversation_archive',owner.id,owner.version);
    if(!result?.ok||snapshot?.document?.id!==owner.id)return;
    const dialog=make('dialog','',document.body);dialog.className='floating-history-dialog';dialog.setAttribute('aria-label','旧对话');
    make('h2','旧对话',dialog);
    const close=button('关闭',dialog,()=>dialog.close());
    const list=make('div','',dialog), messages=make('div','',dialog);
    if(!result.archives.length)make('p','这个文件还没有归档对话。',list);
    let selection=0;
    for(const [index,archive] of result.archives.entries()){
      button('旧对话 '+(index+1)+' · '+new Date(archive.modified*1000).toLocaleString(),list,async()=>{
        const mine=++selection;
        const reply=await call('conversation_archive',owner.id,owner.version,archive.id);
        if(!dialog.open||mine!==selection||snapshot?.document?.id!==owner.id)return;
        messages.replaceChildren();
        if(!reply?.ok){make('p',reply?.error||'旧对话无法读取。',messages);return;}
        for(const turn of reply.turns){
          const entry=make('article','',messages);entry.className='floating-history-turn';
          make('h4',turn.question,entry);make('p',turn.answer||turn.text,entry);
          for(const output of turn.results||[])if(output.verified){
            button('查看副本 · '+output.name,entry,async()=>{
              const verified=await call('conversation_output',output.id,owner.id,owner.version,archive.id);
              if(!dialog.open||snapshot?.document?.id!==owner.id)return;
              if(!verified?.ok){make('p',verified?.error||'旧副本无法核对。',entry);return;}
              dialog.close();await openHistoryOutput(output.id);
            });
          }
        }
      });
    }
    dialog.addEventListener('close',()=>{dialog.remove();if(trigger?.isConnected)trigger.focus();});
    dialog.showModal();close.focus();
  }
  async function restoreDrafts() {
    const result = await call('floating_restore_drafts');
    if (!result?.ok) return;
    if (!result.tasks.length) { feedback('本次有界检查未找到可恢复要求；旧候选未保存原文的记录不能恢复。'); return; }
    const previous = document.activeElement;
    const dialog = make('dialog', '', document.body); dialog.className = 'floating-history-dialog';
    make('h2', '复用历史要求', dialog);
    make('p', '只填入草稿，不发送或恢复批准；请核对是否已执行。', dialog);
    if (result.limited) make('p', '仅显示有界扫描中的最近5组，历史可能不完整。', dialog);
    const fill = async text => {
      if ($('floating-input').value.trim() && !await ask('替换当前未发送草稿？不会发送请求。')) return;
      $('floating-input').value = text; dialog.close();
      feedback('历史要求已填入草稿；请核对已有产物，避免重复修改，再主动发送。');
    };
    for (const [index, task] of result.tasks.entries()) {
      make('h3', '历史要求 ' + (index + 1) + ' · 需核对执行状态', dialog);
      const combined = task.messages.join('\n');
      if (task.cancelled_count) make('p','已排除 '+task.cancelled_count+' 条取消要求。',dialog);
      if (!task.messages.length) make('p','该组没有可复用的有效要求。',dialog);
      if (task.messages.length && combined.length <= 4000) button('复用这一组要求', dialog, () => fill(combined));
      for (const text of task.messages) { make('pre', text, dialog); button('复用这条要求', dialog, () => fill(text)); }
    }
    button('关闭', dialog, () => dialog.close());
    dialog.addEventListener('close', () => { dialog.remove(); previous?.focus(); }); dialog.showModal();
  }
  function mount() {
    if ($('floating-root')) return;
    const root = make('section', '', $('workspace'), 'floating-root'); root.setAttribute('aria-label', '悬浮文档工作区');
    const header = make('header', '', root);
    const logo = make('img', '', header); logo.src = 'brand/lix.png'; logo.alt = 'LiX'; logo.className = 'lix-brand-icon'; logo.width = logo.height = 24;
    make('strong', 'LiX · 文档助手', header, 'floating-title');
    window.OfficeAgentElements.iconButton(button('选择文档', header, async () => { if (window.OfficePaged?.active() && !await window.OfficePaged.beforeSwitch()) return; if(window.OfficeNative?.active()&&!await window.OfficeNative.beforeSwitch())return; const selected=await call('floating_choose'); if(selected?.ok&&!selected.cancelled&&selected.document_id===snapshot?.document?.id&&selected.epoch===snapshot?.epoch)await call('floating_present', 'review'); }, 'floating-choose'), 'folder-open');
    const collapse = button('收起', header, () => call('floating_present', 'orb'));
    const menuDetails = make('details', '', header); make('summary', '更多', menuDetails);
    window.OfficeAgentElements.iconButton(menuDetails.querySelector('summary'), 'ellipsis', '更多');
    window.OfficeWindowChrome?.attach(header, collapse);
    const menu = make('div', '', menuDetails); menu.className = 'floating-menu';
    const group = name => { const folded=['当前任务','工具与记录'].includes(name),section = make(folded?'details':'section', '', menu); section.setAttribute('aria-label', name); if(folded)make('summary',name,section);return section; };
    const documentActions = group('文档');
    button('文档审阅', documentActions, () => {menuDetails.open=false;call('floating_present', 'review');});
    button('结构文本',documentActions,()=>{menuDetails.open=false;if(window.OfficePaged?.active())window.OfficePaged.previewStructure(snapshot.view,'原稿 · 只读有界结构文本');else{showView(snapshot.view);root.dataset.tab='document';}}).title='查看当前有界结构文本';
    button('紧凑对话', documentActions, () => {menuDetails.open=false;call('floating_present', 'compact');});
    const taskActions = group('当前任务');
    button('重新生成方案', taskActions, () => { menuDetails.open = false; call('floating_start_again', false); });
    button('继续排队要求', taskActions, () => { menuDetails.open = false; call('floating_start_again', false); });
    button('新一轮要求…', taskActions, async () => { menuDetails.open = false; if (await ask('开始新一轮会使旧计划和批准失效；不删除已生成副本。是否继续？')) {window.OfficePaged?.invalidate();call('floating_start_again', true);} });
    const tools = group('工具与记录');
    button('历史记录', tools, () => { menuDetails.open = false; secondary(true, true); });
    button('复用历史要求', tools, () => { menuDetails.open = false; restoreDrafts(); }).title='恢复本文件历史要求到草稿，不自动发送';
    button('生成演示稿', tools, () => { menuDetails.open=false; window.OfficePresentations.open().catch(error=>feedback(error.message)); });
    button('问答与工具', tools, () => { menuDetails.open = false; secondary(true); }).title='问答、图片与专业工具';
    const management = group('管理');
    button('设置', management, () => { menuDetails.open = false; window.OfficeSettings?.show(); });
    button('打开详情', management, () => { $('floating-open-trace').hidden = !$('floating-open-trace').hidden; }).title='查看本次打开过程';
    button('关闭文档…',management,async()=>{menuDetails.open=false;if(window.OfficePaged?.active()&&!await window.OfficePaged.beforeSwitch())return;if(window.OfficeNative?.active()&&!await window.OfficeNative.beforeSwitch())return;await call('floating_close_document');});
    make('pre', '', menu, 'floating-open-trace').hidden = true;
    menuDetails.addEventListener('keydown', event => {
      if (event.isComposing) return;
      if (event.key === 'Escape') {event.preventDefault();event.stopPropagation();menuDetails.open=false;menuDetails.querySelector('summary').focus();return;}
      if (!['ArrowDown','ArrowUp','Home','End'].includes(event.key)) return;
      event.preventDefault();menuDetails.open=true;
      const items=Array.from(menu.querySelectorAll('button:not(:disabled),details>summary')).filter(node=>node.checkVisibility());
      const index=items.indexOf(document.activeElement);
      const next=event.key==='Home'?0:event.key==='End'?items.length-1:event.key==='ArrowDown'?(index+1)%items.length:(index+items.length-1)%items.length;
      items[next]?.focus();
    });
    document.addEventListener('pointerdown', event => {if(menuDetails.open&&!menuDetails.contains(event.target))menuDetails.open=false;});
    const fitMenu=()=>{if(!menuDetails.open)return;const top=menuDetails.getBoundingClientRect().bottom,composer=$('floating-controls')?.getBoundingClientRect().top||innerHeight;menu.style.maxHeight=Math.max(88,Math.min(innerHeight-12,composer-8)-top)+'px';};
    menuDetails.addEventListener('toggle',fitMenu);window.addEventListener('resize',fitMenu);
    make('p', '本机打开，发送才调用模型', root, 'floating-identity');
    button('重试本地打开此文件', root, async () => {if(window.OfficePaged?.active()&&!await window.OfficePaged.beforeSwitch())return;if(window.OfficeNative?.active()&&!await window.OfficeNative.beforeSwitch())return;await call('floating_admit_drop');}, 'floating-drop').hidden = true;
    const read = make('section', '', root, 'floating-read'); read.hidden = true;
    make('p', '允许读取当前文件并在软件内查看？发送要求时，模型会收到任务所需的有界片段。读取不批准写入。', read);
    button('同意读取并打开', read, () => call('floating_read', snapshot.read_token, true));
    button('暂不读取', read, () => call('floating_read', snapshot.read_token, false));
    const tabs = make('nav', '', root, 'floating-tabs'); tabs.setAttribute('aria-label', '工作区视图');
    for (const [id, label] of [['document', '文档'], ['conversation', '对话']]) button(label, tabs, () => { root.dataset.tab = id; });
    root.dataset.tab = 'conversation';
    const body = make('div', '', root, 'floating-body');
    const welcome = make('section', '', body, 'floating-welcome');
    make('span', '', welcome).className = 'agent-icon agent-icon-document';
    make('p', 'LiX', welcome).className = 'agent-eyebrow';
    make('h1', '打开文档', welcome);
    make('p', '拖入文档，或选择文件。', welcome).className = 'agent-welcome-note';
    button('选择文档', welcome, () => $('floating-choose').click()).className = 'agent-welcome-choose';
    const examples = make('div', '', welcome); examples.className = 'agent-suggestions';
    for (const text of ['润色文字，保留原意', '统一术语与表达', '核对数字和单位']) button(text, examples, async () => {
      const input = $('floating-input');
      if (input.value && !await ask('替换未发送的要求？不会调用模型。')) return;
      input.value = text + '。'; input.focus();
    });
    make('p', '本机打开，发送时才调用模型。', welcome).className = 'agent-welcome-safety';
    (async () => {
      let seen = false;
      try { seen = localStorage.getItem('lix.shell-intro-seen') === 'true'; } catch (_) {}
      try { const saved = await api().floating_intro(); seen ||= saved?.ok && saved.intro_seen; } catch (_) {}
      if(!seen && welcome.isConnected){
        const guide=make('aside','',welcome);guide.className='lix-first-guide';guide.setAttribute('aria-label','首次使用');
        make('p','拖文档到球上打开；点击球展开，× 收起；在托盘菜单退出。',guide);
        const label=make('label','',guide),remember=make('input','',label);remember.type='checkbox';remember.checked=true;label.append(document.createTextNode('以后不再提示'));
        button('知道了',guide,async()=>{
          if(remember.checked){
            try { localStorage.setItem('lix.shell-intro-seen','true'); } catch (_) {}
            try { await api().floating_intro(true); } catch (_) {}
          }
          guide.remove();
        });
      }
    })();
    make('section', '', body, 'floating-document').setAttribute('aria-label', '文档结构文本');
    const conversation = make('section', '', body, 'floating-conversation');
    const conversationHeader = make('div', '', conversation, 'floating-conversation-heading');
    make('span', '', conversationHeader).className = 'agent-icon agent-icon-chat';
    make('strong', '对话', conversationHeader);
    const nextConversation=button('开始新对话',conversationHeader,async()=>{
      const owner=snapshot?.document, context=snapshot?.conversation?.context_id;
      if(submitting||!owner?.ready||!context)return;
      nextConversation.disabled=true;
      try{const result=await call('conversation_new',context,owner.id,owner.version);
        if(result?.ok&&snapshot?.document?.id===owner.id)$('floating-input').focus();
      }finally{nextConversation.disabled=submitting;}
    },'floating-new-conversation');nextConversation.hidden=true;
    const scroll = make('div', '', conversation, 'floating-scroll');
    const timeline = make('div', '', scroll, 'floating-timeline'); timeline.setAttribute('aria-label', '要求与处理进展');
    const latest = button('回到最新', conversation, () => { scroll.scrollTop = scroll.scrollHeight; latest.hidden = true; }, 'floating-latest');
    const down = make('span', '', latest); down.className='agent-icon agent-icon-chevron-down'; down.setAttribute('aria-hidden','true');
    latest.hidden = true;
    make('div', '', scroll, 'floating-plans');
    make('p', '就绪', root, 'floating-progress').setAttribute('role', 'status');
    const taskDetails = make('details', '', root, 'floating-task-details'); make('summary', '保护与任务版本', taskDetails);
    make('p', '', taskDetails, 'floating-revisions');
    make('p', '', taskDetails, 'floating-actual-model');
    make('p', '', root, 'floating-status').setAttribute('role', 'status');
    const resolveProtection = button('按已勾选保护位置继续', root, () => call('floating_resolve_protection'), 'floating-resolve-protection'); resolveProtection.hidden = true;
    make('p', '正文局部文字修改默认保护数字、单位、公式和引文的字面值；不代表事实核验。', taskDetails).className = 'floating-note';
    make('p', '本机任务要求保存在当前文档父目录的 .office_agent 中。历史恢复只填草稿，不恢复批准。', taskDetails).className = 'floating-note';
    const composer = make('form', '', root, 'floating-composer'); composer.onsubmit = event => { event.preventDefault(); submit(); };
    const label = make('label', '提问或说明修改要求（运行中可追加）', composer); label.htmlFor = 'floating-input'; label.className = 'agent-visually-hidden';
    const input = make('textarea', '', composer, 'floating-input'); input.rows = 2; input.maxLength = 4000; input.placeholder = '提问或说明修改要求…';
    input.onkeydown = event => { if (event.key === 'Enter' && !event.shiftKey && !event.isComposing && event.keyCode !== 229) { event.preventDefault(); submit(); } };
    const footer = make('div', '', composer); footer.className = 'agent-composer-footer';
    make('div', '', footer, 'floating-models');
    window.OfficePlugins?.attach(footer);
    button('自动存副本',footer,accessSettings,'floating-access');
    const actions = make('div', '', footer); actions.className = 'floating-actions';
    button('停止后续步骤', actions, () => {window.OfficePaged?.invalidate();return call('floating_stop');}, 'floating-stop');
    button('发送要求', actions, submit, 'floating-send').className = 'floating-primary-button';
    const controls = make('div', '', conversation, 'floating-controls');
    for (const id of ['floating-progress','floating-resolve-protection','floating-task-details','floating-status','floating-composer']) controls.append($(id));
    const contextMenu = make('div', '', taskDetails); contextMenu.className = 'agent-context-body';
    for (const child of [...taskDetails.children]) if (child.tagName !== 'SUMMARY' && child !== contextMenu) contextMenu.append(child);
    taskDetails.querySelector('summary').setAttribute('aria-label', '保护与任务版本');
    taskDetails.querySelector('summary').title = '保护与任务版本';
    footer.prepend(taskDetails);
    const fit = () => (root.classList.contains('paged-active') || root.classList.contains('office-active') ? conversation : matchMedia('(max-width:899px)').matches ? root : conversation).append(controls);
    window.addEventListener('resize', fit); fit();
    const back = button('', $('workspace'), () => secondary(false), 'floating-return');
    back.hidden = true;
    back.setAttribute('aria-label', '返回悬浮文档助手');
    back.title = '返回悬浮文档助手';
    back.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 4h6v16H6V4h8z"/><path d="M10 12h6M10 16h4"/>
      <path d="M11 4 8 8l3 4"/>
    </svg>`;
    window.addEventListener('resize', positionBack);
    if (window.ResizeObserver) {
      const composerArea = $('simple-compose-area');
      if (composerArea) new ResizeObserver(positionBack).observe(composerArea);
    }
    modelHome = $('simple-model-menu')?.parentElement;
    for (const id of ['simple-model-menu', 'simple-model-state']) if ($(id)) $('floating-models').append($(id));
    $('model-mode').value = 'agent'; $('model-mode').dispatchEvent(new Event('change'));
    $('workspace').classList.add('floating-primary');
    root.addEventListener('keydown', event => {
      if (event.key !== 'Escape' || event.isComposing) return;
      const opened = [...root.querySelectorAll('details[open]')].reverse()[0];
      if (opened) { event.preventDefault(); event.stopPropagation(); opened.open = false; opened.querySelector('summary')?.focus(); }
    });
    const selector = '#floating-task-details, #simple-model-menu, #floating-root > header > details, #paged-top > header > details';
    root.addEventListener('click', event => {
      const node = event.target.closest('summary')?.parentElement;
      if (!node?.matches(selector) || node.open) return;
      for (const other of root.querySelectorAll(selector)) if (other !== node) other.open = false;
    }, true);
    root.addEventListener('toggle', event => {
      if (!event.target.open || !event.target.matches(selector)) return;
      for (const node of root.querySelectorAll(selector)) if (node !== event.target) node.open = false;
    }, true);
  }
  async function connect() {
    const mine=epoch;
    if (typeof window.pywebview?.api?.floating_snapshot !== 'function' || !await window.OfficeAuth.ensure()) return;
    if(mine!==epoch)return;
    window.OfficeFloating.active = true; mount();
    if(window.isSecureContext&&typeof api().paged_enable_workspace==='function')await api().paged_enable_workspace();
    if(mine!==epoch)return;
    await refresh();
    try {
      const position = JSON.parse(localStorage.getItem('office-agent.orb-position') || 'null');
      if (position) await api().floating_restore_position(position);
    } catch (_) {}
    if (timer === null) timer = setInterval(() => refresh(true).catch(() => {}), 250);
  }
  async function openHistoryOutput(identifier){
    if(window.OfficePaged?.beforeSwitch&&!await window.OfficePaged.beforeSwitch())return false;
    if(window.OfficeNative?.beforeSwitch&&!await window.OfficeNative.beforeSwitch())return false;
    const result=await call('floating_select_history_output',identifier);
    if(!result?.ok)throw Error(result?.error||'副本未打开，请核对当前文件');
    secondary(false);await refresh();return true;
  }
  window.OfficeFloating = {active: false, history: () => secondary(true, true), refresh, openHistoryOutput};
  window.OfficeConversation?.configure({pluginConsent:async (id,token,approved)=>{
    const doc=snapshot?.document;
    if(!doc)return;
    const result=await call('conversation_plugin_consent',id,doc.id,doc.version,snapshot.conversation?.context_id,token,approved);
    if(!result?.ok)feedback(result?.error||'请刷新后重试。');
    await refresh();
  },pluginSource:async(id,index)=>{await call('conversation_plugin_source',id,index);},open:async id=>{
    let row=snapshot?.outputs.find(value=>value.id===id&&!value.undone);
    if(!row){const result=await api().conversation_output(id,snapshot.document.id,snapshot.document.version);if(!result?.ok){feedback(result?.error||'结果记录已改变，请核对当前文件与历史。');return;}row=result.output;}
    await openOutput(row);
  },previewDraft:async (publication,trigger)=>{
    try{await window.OfficePaged.previewLatestDraft(publication,trigger);}
    catch(error){if(!error.stale)feedback(error.message);}
  },undo:async id=>{
    try{
      if(!window.OfficePaged?.active()||typeof window.OfficePaged.undoRound!=='function')throw Error('当前编辑器没有可靠的回合恢复点');
      await window.OfficePaged.undoRound(id);await refresh();
    }catch(error){await refresh();window.OfficeConversation?.undoError(id,window.OfficeConversation?.undoMessage?.(id)||'撤销尚未完成，当前内容保留。请从历史记录核对结果。');}
  },reselect:async id=>{
    if(submitting||snapshot?.conversation?.pending_scope!==id)return;
    const input=$('floating-input');
    if(input.value){feedback('输入框已有草稿。请自行发送原问题或“用新选区回答”。');return;}
    input.value='用新选区回答';
    await submit();
  },retry:async(id,local)=>{
    if(local){const message=window.OfficeConversation?.message(id);if(!message){feedback('原消息回执正在核对，没有发送当前草稿。');await refresh();return;}await submit(message);return;}
    const result=await api().conversation_retry(id,snapshot.document.id,snapshot.document.version);
    if(!result?.ok){feedback(result?.error||'恢复未完成，已有内容保留。');return;}
    if(result.output_id)await call('floating_select_output',result.output_id,true,result.access_id);
    await refresh();
  }});
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)refresh().catch(()=>{});});
  window.addEventListener('pywebviewready', connect);
  window.addEventListener('officeworkspace:view',event=>{if(event.detail?.screen==='history'){ensureOutputRecords();if(snapshot)renderOutputs(snapshot);}});
  window.addEventListener('officeauth:authenticated', connect);
  window.addEventListener('officeauth:unauthenticated', () => { epoch++; snapshot = null; cursor = 0; pending = null; adopted.clear(); seenMessages.clear(); seenSkills.clear(); if (timer) clearInterval(timer); timer = null; });
  if (window.pywebview?.api) connect();
})();
