(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  if (!$('daily-conversations') || !window.OfficeWorkspace || $('simple-reading')) return;
  const make = (tag, id, text, parent) => {
    const node = document.createElement(tag);
    if (id) node.id = id;
    if (text) node.textContent = text;
    if (parent) parent.append(node);
    return node;
  };
  const button = (id, text, parent, action) => {
    const node = make('button', id, text, parent);
    node.type = 'button'; node.className = 'btn secondary'; node.onclick = action;
    return node;
  };
  const move = (id, parent) => { const node = $(id); if (node) parent.append(node); };
  const field = (id, parent) => { const node = $(id); if (node) parent.append(node.closest('label') || node); };
  const menu = (id, title, parent) => {
    const node = make('details', id, '', parent); node.className = 'simple-menu';
    make('summary', '', title, node);
    const body = make('div', '', '', node); body.className = 'simple-menu-body';
    return [node, body];
  };
  const app = $('workspace'), shell = $('document-workspace'), task = $('workspace-task');
  app.classList.add('simple-workbench');
  let mode = window.OfficeModelSelector.mode(), state = null, historyKind = 'documents';
  let screen = window.OfficeWorkspace.screen(), queued = false, inspector = null, fileSelected = false;
  let sessionsKey = '';
  const menus = [];
  const body = make('div', 'simple-body', '', shell);
  const read = make('section', 'simple-reading', '', body);
  read.setAttribute('aria-label', '当前任务内容'); read.tabIndex = 0;
  const welcome = make('div', 'simple-image-welcome', '', read);
  make('h2', '', '为报告生成一张配图', welcome);
  make('p', '', '描述画面，查看方案后再确认生成。', welcome);
  read.append($('daily-conversations'), $('content-panes'));
  body.append($('workspace-stage'));
  const start = make('section', 'simple-start', '', read);
  const startTitle = make('h2', 'simple-start-title', '', start);
  const startNote = make('p', 'simple-start-note', '', start);
  const starts = make('div', 'simple-start-actions', '', start);
  button('simple-start-document', '选择文档…', starts, () => $('btn-select2').click());
  button('simple-start-draft', '润色一段文字', starts, () => {
    const input = $('prompt');
    if (input.value && !window.confirm('替换未发送的草稿？不会调用模型。')) return;
    input.value = '请润色以下文字，保留数字、单位、引文和限定条件：\n';
    input.dispatchEvent(new Event('input', {bubbles:true})); input.focus();
  });
  button('simple-start-question', '概括报告要点', starts, () => {
    if ($('prompt').value && !window.confirm('替换未发送的草稿？不会调用模型。')) return;
    $('prompt').value = '请概括以下技术报告的目的、方法、主要结果与适用条件：\n';
    $('prompt').dispatchEvent(new Event('input', {bubbles:true})); $('prompt').focus();
  });
  const compose = make('section', 'simple-compose-area', '', shell);
  compose.setAttribute('aria-label', '当前任务输入');
  const composer = document.querySelector('.composer');
  compose.append($('pending-notice'), composer);
  const dock = make('div', 'simple-input-tools', '', composer);
  dock.append($('model-mode').closest('label'));
  const [models, modelBody] = menu('simple-model-menu', '选择模型', dock); menus.push(models);
  const modelState = make('span', 'simple-model-state', '', dock); modelState.setAttribute('role', 'status');
  modelBody.append($('model-selector-bar'));
  $('model-selector-bar').querySelectorAll('label').forEach(node => { if (node.contains($('model-search')) || node.contains($('model-selector'))) modelBody.prepend(node); });
  const [sendOptions, sendBody] = menu('simple-send-options', '发送选项', dock); menus.push(sendOptions);
  field('chat-stream', sendBody); move('chat-preview-context', sendBody); move('chat-cancel-continuation', sendBody);
  const [budget, budgetBody] = menu('simple-budget-menu', '回答长度', dock); menus.push(budget);
  sendBody.append(budget);
  field('chat-reply-length', budgetBody); move('chat-reply-save', budgetBody);
  make('p', '', '保存后用于后续问答；预算不保证字数。', budgetBody);
  const [imageOptions, imageBody] = menu('simple-image-options', '图片参数', dock); menus.push(imageOptions);
  move('daily-image-parameters', imageBody);
  const [assistTools, assistBody] = menu('simple-assist-menu', '辅助工具', dock); menus.push(assistTools);
  const vision = composer.querySelector('.vision-actions'); if (vision) assistBody.append(vision);
  move('tools-menu', assistBody); move('input-help', assistBody);
  dock.prepend(assistTools);
  dock.append($('btn-stop'), $('btn-ask'));
  button('simple-attach-document', '选择文档并切换用途…', assistBody, async () => {
    if (await switchMode('agent')) $('btn-select2').click();
  });
  const [sessionMenu, sessionBody] = menu('simple-session-menu', '会话操作', document.querySelector('.topbar')); menus.push(sessionMenu);
  make('h3', '', '当前会话', sessionBody);
  field('chat-title', sessionBody); move('chat-rename', sessionBody); move('chat-title-hint', sessionBody);
  move('chat-save-now', sessionBody); move('chat-delete', sessionBody);
  make('h3', '', '导出', sessionBody); move('chat-export-controls', sessionBody);
  make('h3', '', '保存容量', sessionBody); move('chat-capacity', sessionBody);
  const account = $('settings-pane-account');
  move('technical-details', $('settings-pane-machine'));
  move('more-menu', $('settings-pane-machine'));
  const preferences = make('section', 'simple-history-preferences', '', account);
  make('h3', '', '问答与图片历史', preferences);
  field('chat-save-history', preferences);
  const storage = $('chat-storage')?.closest('details'); if (storage) preferences.append(storage);
  field('image-reuse-preference', preferences);
  const rail = $('nav-work').parentElement;
  const [newMenu, newBody] = menu('simple-new-task', '＋ 新建任务', rail); menus.push(newMenu); rail.prepend(newMenu);
  move('chat-new', newBody);
  button('simple-new-document', '选择文档…', newBody, async () => {
    if (!await switchMode('agent')) return;
    newMenu.open = false; window.OfficeWorkspace.open('work');
    $('btn-select2').click();
  });
  button('simple-new-image', '准备图片提示词', newBody, async () => {
    if (!await switchMode('image')) return;
    newMenu.open = false;
    window.OfficeWorkspace.open('work');
    $('prompt').focus();
  });
  const createConversation = $('chat-new').onclick;
  $('chat-new').onclick = async () => {
    if (!await switchMode('chat')) return;
    await createConversation();
    newMenu.open = false; window.OfficeWorkspace.open('work');
  };
  $('nav-work').querySelector('span').textContent = '工作区';
  $('nav-work').title = '工作区'; $('nav-work').setAttribute('aria-label', '工作区');
  $('nav-resources').textContent = '资料库'; $('nav-resources').title = '资料库';
  $('nav-history').querySelector('span').textContent = '历史记录';
  $('nav-history').setAttribute('aria-label', '历史记录');
  $('workspace-view-switch').append($('nav-review'));
  $('nav-review').className = 'btn secondary'; $('nav-review').textContent = '修改建议';
  $('nav-review').setAttribute('aria-label', '当前文档修改建议');
  $('workspace-view-document').textContent = '原文';
  $('workspace-view-task').textContent = '修改要求';
  $('workspace-view-review').textContent = '修改建议';
  $('workspace-view-review').hidden = true;
  const [documentsMenu,documentsBody] = menu('simple-document-menu','文档',document.querySelector('.topbar')); menus.push(documentsMenu);
  documentsBody.append($('workspace-view-switch'));
  const panelHeading = make('div', 'simple-panel-heading', '', $('workspace-stage'));
  const panelTitle = make('h2', 'simple-panel-title', '文档', panelHeading);
  $('workspace-stage').prepend(panelHeading);
  button('simple-panel-close', '回到对话', panelHeading, () => { inspector=null; window.OfficeWorkspace.open('work'); update(); $('prompt').focus(); });
  const showInspector = kind => {
    inspector = kind;
    if (kind === 'task' && shell.dataset.mode === 'dialogue') shell.dataset.mode='task';
    if (kind !== 'review') window.OfficeWorkspace.open('work');
    update();
  };
  $('workspace-view-document').addEventListener('click', () => showInspector('document'));
  $('workspace-view-task').addEventListener('click', () => showInspector('task'));
  $('nav-review').addEventListener('click', () => showInspector('review'));
  button('simple-view-preview', '预览', $('workspace-view-switch'), () => {
    window.OfficeWorkspace.open('review'); $('workspace-review-kind-pages').click(); showInspector('review');
  });
  const recent = make('section', 'simple-recent', '', rail);
  make('h2', '', '最近会话', recent);
  const recentList = make('div', 'simple-recent-list', '', recent);
  const [sessions, sessionsBody] = menu('simple-session-list', '查找全部会话', recent); menus.push(sessions);
  const sessionSearch = $('chat-search');
  field('chat-search', sessionsBody); field('chat-session', sessionsBody); move('chat-refresh', sessionsBody);
  const history = make('section', 'simple-history', '', $('history-reader')); history.hidden = true;
  const historyNav = make('nav', 'simple-history-kinds', '', history); historyNav.setAttribute('aria-label', '历史记录分类');
  const historyContent = make('div', 'simple-history-content', '', history);
  const historyChats = make('section', 'simple-history-chats', '', historyContent);
  make('h2', '', '对话', historyChats);
  const historyChatList = make('div', 'simple-history-chat-list', '', historyChats);
  historyChats.insertBefore(sessionSearch.closest('label'), historyChatList);
  const sessionSearchHandler = sessionSearch.oninput;
  sessionSearch.oninput = () => { sessionSearchHandler?.(); renderSessions(); };
  $('simple-session-list').querySelector('summary').textContent = '查找会话';
  sessions.addEventListener('toggle', () => {
    const searchLabel = sessionSearch.closest('label');
    if (sessions.open) {
      if (window.OfficeSettings?.isOpen() || typeof modalOpen !== 'undefined' && modalOpen) {
        sessions.open = false;
        return;
      }
      sessionsBody.prepend(searchLabel);
      showHistory('chat');
      queueMicrotask(() => { if (sessions.open) sessionSearch.focus(); });
    } else {
      const restoreFocus = sessionsBody.contains(document.activeElement);
      historyChats.insertBefore(searchLabel, historyChatList);
      if (restoreFocus) sessions.querySelector('summary').focus();
    }
  });
  const documentHistory = make('section', 'simple-document-records', '', historyContent);
  documentHistory.append($('history-view'), $('history-detail'));
  for (const [key, title] of [['chat','对话'], ['documents','文档任务'], ['image','图片']]) {
    button('simple-history-' + key, title, historyNav, () => showHistory(key));
  }
  const imageHistory = $('daily-image-history');
  imageHistory.querySelector('h3').textContent = '图片记录';
  $('nav-history').addEventListener('click', () => showHistory(historyKind));
  const status = $('status-card');
  compose.prepend(status);
  const docDialog = button('workspace-mode-dialogue', 'AI 对话', $('workspace-task-modes'), () => {
    shell.dataset.mode = 'dialogue'; inspector=null; window.OfficeWorkspace.open('work'); update(); $('prompt').focus();
  });
  $('workspace-mode-task').addEventListener('click', () => showInspector('task'));
  $('workspace-mode-range').addEventListener('click', () => showInspector('task'));
  $('range-compose').addEventListener('click', () => {
    if ($('prompt').value) window.OfficeWorkbench.compose();
  });
  const fileDetails = make('details', 'simple-file-details', '', $('workspace-document-bar'));
  make('summary', '', '文件详情', fileDetails);
  fileDetails.append($('workspace-document-identity'));
  const outputDetails = make('details', 'simple-output-details');
  make('summary', '', '副本详情', outputDetails);
  $('range-output').before(outputDetails);
  outputDetails.append($('range-output'));
  const outputSummary = make('p', 'simple-output-summary'); outputSummary.setAttribute('role','status');
  outputDetails.before(outputSummary);
  const summarizeOutput = () => {
    const first = $('range-output').textContent.split('\n')[0];
    outputSummary.textContent = first.startsWith('已保存副本：') ? '已生成副本 ' + first.slice(6).split(/[\\/]/).pop() + ' · 可预览并继续审阅' : first;
    outputDetails.hidden = !first;
  };
  new MutationObserver(summarizeOutput).observe($('range-output'), {childList:true,subtree:true,characterData:true});
  summarizeOutput();
  $('workspace-current-status').hidden = true;
  $('workspace-document').querySelector('p').textContent = '只读结构文本 · 选择文段后可定位修改。';
  const resources = $('workspace-resources');
  resources.querySelector('h2').textContent = '资料与研究';
  resources.querySelector('p').textContent = '选择需要的工具；读取、联网和写入仍分别确认。';
  const projectPanel = $('workbench-panel');
  if (projectPanel) {
    const management = make('details', 'simple-project-management', '', projectPanel);
    make('summary', '', '项目记录', management);
    const columns = projectPanel.querySelector('.workbench-columns');
    const nodes = [];
    for (let node=$('wb-name').previousElementSibling; node && node!==columns; node=node.nextElementSibling) nodes.push(node);
    nodes.forEach(node => { if (node!==management) management.append(node); });
    projectPanel.insertBefore(management,columns);
    const actions = $('wb-load').parentElement;
    management.prepend($('wb-save'),$('wb-save-confirm'));
    const projectStep = make('p', 'simple-project-step', '先读取当前项目与文段，或新建项目。', projectPanel);
    actions.after(projectStep);
    const stepNav = make('nav','simple-research-steps','',projectPanel); stepNav.setAttribute('aria-label','当前研究步骤');
    projectStep.after(stepNav);
    [...columns.children].forEach((section,index) => {
      const titles=['目标与问题','材料与证据','修改队列'];
      section.dataset.researchStep=String(index);
      button('',titles[index] || '研究详情',stepNav,() => {
        columns.dataset.step=String(index);
        [...stepNav.children].forEach((node,i)=>node.setAttribute('aria-pressed',String(i===index)));
      }).setAttribute('aria-pressed',String(index===0));
    });
    columns.dataset.step='0';
  }
  const settingsForm = $('settings-pane-model').querySelector('.settings-form');
  if (settingsForm) {
    settingsForm.prepend($('connection-manager'));
    const note = settingsForm.querySelector(':scope > p:not([id])');
    note.textContent='本机工作区配置；保存后生效方式见对应连接及配置说明。';
    const diagnostics = make('details','simple-settings-diagnostics','',settingsForm);
    make('summary','','运行配置与存储详情',diagnostics);
    diagnostics.append($('settings-runtime'));
    const footer = make('div','simple-settings-actions','',$('settings-pane-model'));
    footer.append($('settings-save'),$('settings-discard'),$('settings-rollback'));
    const retrieval = make('details','simple-retrieval-settings','',settingsForm);
    make('summary','','文档检索配置',retrieval);
    const first = $('settings-embedding-channel').previousElementSibling.previousElementSibling;
    const nodes=[];
    for (let node=first;node && node!==diagnostics;node=node.nextElementSibling) {
      if (node===footer || node===retrieval) break;
      nodes.push(node);
    }
    nodes.forEach(node=>retrieval.append(node));
    settingsForm.insertBefore(retrieval,diagnostics);
    const guide = $('settings-setup');
    if (guide) diagnostics.append(guide);
  }
  const authBoundary = $('auth-screen')?.querySelector('.auth-boundary');
  if (authBoundary) {
    const details = make('details','simple-auth-details','',authBoundary.parentElement);
    make('summary','','本地账户共用数据；文件修改仍需逐次确认',details); details.append(authBoundary);
  }
  const historyClose = imageHistory.querySelector(':scope > button:last-child');
  if (historyClose) historyClose.onclick = () => window.OfficeWorkspace.open('work');
  move('daily-open-images', assistBody); move('daily-open-chats', assistBody);
  $('daily-open-images').addEventListener('click', () => showHistory('image', false));
  $('daily-open-chats').addEventListener('click', () => window.OfficeWorkspace.open('work'));
  $('model-open-input').hidden = true;
  $('daily-next-action').hidden = true;
  const lengthInfo = Array.from($('daily-conversations').children).find(node => node.tagName === 'SMALL');
  if (lengthInfo) budgetBody.append(lengthInfo);
  move('chat-latest', compose);
  const latest = $('chat-latest');
  let following = true;
  const scrollState = () => {
    following = read.scrollHeight - read.clientHeight - read.scrollTop < 60;
    latest.hidden = following || mode !== 'chat';
  };
  read.addEventListener('scroll', scrollState, {passive:true});
  const settleReading = () => requestAnimationFrame(() => {
    if (following && mode === 'chat') read.scrollTop = read.scrollHeight;
    scrollState();
  });
  function decorateTurns() {
    for (let index=menus.length-1; index>=0; index--) {
      if (!menus[index].isConnected) menus.splice(index,1);
    }
    for (const article of $('chat-transcript').children) {
      if (!article.classList.contains('daily-turn') || article.dataset.uiDecorated) continue;
      article.dataset.uiDecorated = 'true';
      const actions = make('div', '', '', article); actions.className = 'simple-message-actions';
      const [more, body] = menu('', '更多', actions); more.classList.add('simple-message-menu');
      menus.push(more); initializeMenu(more);
      more.addEventListener('keydown', event => {
        if (event.key === 'Escape' && !event.isComposing) { event.preventDefault(); event.stopPropagation(); closeMenu(more); }
      });
      for (const child of [...article.children]) {
        if (child.tagName === 'LABEL' || child.tagName === 'DETAILS') body.append(child);
        if (child.tagName === 'BUTTON') {
          if (['复制回答', '继续回答', '填回原问题'].includes(child.textContent)) actions.prepend(child);
          else body.append(child);
        }
      }
    }
    settleReading();
    schedule();
  }
  new MutationObserver(decorateTurns).observe($('chat-transcript'), {childList:true});
  new MutationObserver(settleReading).observe($('chat-live-answer'), {childList:true});
  async function switchMode(next) {
    if ($('model-mode').disabled) return false;
    if (mode === next) return true;
    $('model-mode').value = next; await $('model-mode').onchange();
    return window.OfficeModelSelector.mode() === next;
  }
  function showHistory(kind, load = true) {
    if (window.OfficeSettings?.isOpen() || typeof modalOpen !== 'undefined' && modalOpen) return;
    historyKind = kind;
    window.OfficeWorkspace.open('history');
    renderHistoryKind();
    if (kind === 'image' && load) $('daily-open-images').click();
    screen = 'history'; update();
  }
  function renderHistoryKind() {
    const kind = historyKind;
    if (imageHistory.parentElement !== historyContent) historyContent.append(imageHistory);
    imageHistory.hidden = kind !== 'image';
    historyChats.hidden = kind !== 'chat'; documentHistory.hidden = kind !== 'documents';
    historyNav.querySelectorAll('button').forEach(node => node.setAttribute('aria-pressed', String(node.id === 'simple-history-' + kind)));
  }
  function modelLabel() {
    const selection = window.OfficeModelSelector.selection();
    const text = selection.valid ? selection.model + ' · ' + selection.connection : '选择模型';
    models.querySelector('summary').textContent = text;
    models.querySelector('summary').title = selection.valid ? selection.model + ' · ' + selection.connection + ' · ' + selection.state : '选择已保存模型或配置接口';
    modelState.textContent = selection.valid ? selection.state : '选择失效或未配置';
    modelState.dataset.invalid = String(!selection.valid);
    modelState.title = $('model-selection-status').textContent;
  }
  function update() {
    screen = window.OfficeWorkspace.screen();
    app.dataset.uiMode = mode; app.dataset.uiScreen = screen;
    const inWork = screen === 'work' || screen === 'review';
    const activePanel = screen === 'resources' ? 'resources' : mode === 'agent' ? (screen === 'review' ? 'review' : inspector) : null;
    body.dataset.panel = activePanel || '';
    read.hidden = !inWork;
    history.hidden = screen !== 'history';
    if (screen === 'history') renderHistoryKind();
    $('workspace-stage').hidden = !activePanel || screen === 'history';
    $('workspace-stage').dataset.inspector = activePanel || '';
    $('workspace-document').hidden = activePanel !== 'document';
    task.hidden = activePanel !== 'task';
    $('workspace-review').hidden = activePanel !== 'review';
    $('workspace-resources').hidden = activePanel !== 'resources';
    panelHeading.hidden = activePanel === 'resources';
    panelTitle.textContent = {document:'原文 · 只读结构',task:'修改要求',review:'修改建议与预览'}[activePanel] || '文档';
    documentsMenu.hidden = mode !== 'agent' || !inWork;
    $('workspace-document-bar').hidden = mode !== 'agent' || !inWork || !fileSelected;
    sessionMenu.hidden = mode !== 'chat' || screen === 'history' || screen === 'resources';
    compose.hidden = !inWork;
    composer.hidden = false;
    const structured = activePanel === 'task';
    composer.querySelector('.prompt-bar').hidden = structured;
    $('btn-ask').hidden = structured;
    $('content-panes').hidden = mode === 'chat' || (mode === 'agent' && !fileSelected);
    welcome.hidden = true;
    const chatEmpty = mode === 'chat' && !$('chat-transcript').querySelector('.daily-turn') && $('chat-live-answer').hidden;
    const documentEmpty = mode === 'agent' && !fileSelected;
    const imageEmpty = mode === 'image' && !$('log-pane').querySelector('.imagegen-item');
    start.hidden = !inWork || Boolean(activePanel) || !(chatEmpty || documentEmpty || imageEmpty);
    startTitle.textContent = mode === 'agent' ? '从一份文档开始' : mode === 'image' ? '为报告生成一张配图' : '今天想处理什么？';
    startNote.textContent = mode === 'agent' ? '选择文件后授权读取；修改先预览，确认后另存副本。' : mode === 'image' ? '描述画面，查看方案后再确认生成。' : '直接提问，也可以继续最近的会话。';
    $('simple-start-document').hidden = mode !== 'agent';
    $('simple-start-draft').hidden = $('simple-start-question').hidden = mode === 'image';
    $('chat-transcript').classList.toggle('empty',chatEmpty);
    sendOptions.hidden = mode !== 'chat';
    budget.hidden = mode !== 'chat';
    imageOptions.hidden = mode !== 'image';
    assistTools.hidden = false;
    docDialog.setAttribute('aria-pressed', String(shell.dataset.mode === 'dialogue'));
    const title = screen === 'resources' ? '资料库' : screen === 'history' ? '历史记录' : mode === 'chat' ? state?.title || '新会话' : mode === 'image' ? '图片生成' : screen === 'review' ? '修改审阅' : '文档工作区';
    if ($('page-title').textContent !== title) $('page-title').textContent = title;
    $('prompt').placeholder = mode === 'chat' ? '输入问题…' : mode === 'image' ? '描述要生成的报告配图…' : '描述文档修改要求…';
    $('chat-save-status').hidden = !['failed','uncertain','unsaved'].includes(state?.save_state);
    settleFeedback();
    latest.hidden = following || mode !== 'chat';
    $('chat-cancel-continuation').hidden = !($('chat-context-content').textContent.includes('关联续写'));
    if (screen === 'history') {
      $('nav-work').removeAttribute('aria-current'); $('nav-resources').removeAttribute('aria-current');
      $('nav-history').setAttribute('aria-current', 'page');
    }
    modelLabel();
    sizeInput();
  }
  function settleFeedback() {
    const hostFeedback = window.OfficeRequestFeedback?.snapshot()?.source;
    const dailyFeedback = window.OfficeDaily.feedbackSnapshot()?.source;
    const sameNotice = hostFeedback?.channel === 'chat' && dailyFeedback?.terminal &&
      dailyFeedback.session_id === state?.current && hostFeedback.session_id === dailyFeedback.session_id && hostFeedback.turn_id === dailyFeedback.turn_id;
    status.hidden = Boolean(hostFeedback?.channel === 'chat' && (sameNotice || mode !== 'chat' || hostFeedback.session_id !== state?.current));
    for (const node of $('chat-transcript').querySelectorAll('[data-turn-notice]')) {
      node.hidden = Boolean(mode === 'chat' && dailyFeedback?.terminal && dailyFeedback.session_id === state?.current && node.dataset.turnNotice === dailyFeedback.turn_id);
    }
  }
  function sizeInput() {
    const input = $('prompt');
    if (input.tagName !== 'TEXTAREA' || input.hidden || input.offsetParent === null) return;
    const top = input.scrollTop;
    input.style.height='auto'; input.style.height=Math.min(216,Math.max(64,input.scrollHeight))+'px';
    input.scrollTop=top;
  }
  $('prompt').addEventListener('input',sizeInput);
  const schedule = () => {
    if (queued) return; queued = true;
    requestAnimationFrame(() => { queued = false; update(); });
  };
  window.addEventListener('officeworkspace:view', event => {
    screen = event.detail.screen;
    if (screen === 'review') inspector='review';
    update();
  });
  window.addEventListener('officeapp:state', event => { fileSelected=event.detail.hasFile; schedule(); });
  window.addEventListener('officeworkspace:locate',()=>{inspector='document';schedule();});
  window.addEventListener('officeworkspace:task',event=>{if(event.detail?.current_name) fileSelected=true; schedule();});
  $('nav-work').addEventListener('click',()=>{inspector=null;schedule();});
  window.addEventListener('officemodels:mode', event => { mode = event.detail.mode; update(); });
  function showSessions(value) {
    state = value;
    const key=JSON.stringify([value.current,value.sessions]);
    if(key!==sessionsKey){sessionsKey=key;renderSessions();}
    settleFeedback();
    schedule();
  }
  function renderSessions() {
    if (!state) return;
    recentList.replaceChildren();
    historyChatList.replaceChildren();
    const rows = [...state.sessions].sort((a,b) => b.created-a.created || a.id.localeCompare(b.id));
    const addSession = (item, parent) => {
      const node = button('', item.title, parent, async () => {
        if ($('chat-session').disabled || $('model-mode').disabled) return;
        if (!await switchMode('chat')) return;
        $('chat-session').value = item.id; $('chat-session').dispatchEvent(new Event('change'));
        window.OfficeWorkspace.open('work');
      });
      node.title = item.title; node.dataset.sessionId=item.id; node.setAttribute('aria-current', String(item.id === state.current));
      if (parent === recentList) { node.textContent=''; make('span','',''+item.title,node).className='simple-session-title'; }
    };
    rows.slice(0,6).forEach(item => addSession(item,recentList));
    const query = sessionSearch.value.trim().toLocaleLowerCase();
    const matches=rows.filter(item=>item.title.toLocaleLowerCase().includes(query));
    matches.forEach(item => addSession(item,historyChatList));
    if (!matches.length) make('p','','没有匹配的会话。',historyChatList);
  }
  window.addEventListener('officedaily:state', event => showSessions(event.detail));
  new MutationObserver(schedule).observe($('model-selection-status'), {childList:true,subtree:true,characterData:true});
  new MutationObserver(schedule).observe($('chat-notice'), {childList:true,subtree:true,characterData:true});
  new MutationObserver(schedule).observe($('error-line'), {childList:true,subtree:true,characterData:true,attributes:true,attributeFilter:['style']});
  new MutationObserver(schedule).observe($('log-pane'), {childList:true,subtree:true});
  new MutationObserver(schedule).observe($('chat-transcript'), {childList:true});
  const closeMenu = node => {
    const focus = node.contains(document.activeElement); node.open = false;
    const body=node.querySelector(':scope > .simple-menu-body');
    if (body?.matches(':popover-open')) body.hidePopover();
    node.querySelector('summary').setAttribute('aria-expanded','false');
    if (focus) node.querySelector('summary').focus();
  };
  const placeMenu = node => {
    if (!node.open) return;
    const popup=node.querySelector(':scope > .simple-menu-body');
    const anchor=node.querySelector('summary').getBoundingClientRect();
    const height=popup.getBoundingClientRect().height;
    const width=popup.getBoundingClientRect().width;
    const below=innerHeight-anchor.bottom-16;
    popup.style.left=Math.max(12,Math.min(anchor.left,innerWidth-width-12))+'px';
    popup.style.top=Math.max(12,below>=Math.min(height,240) ? Math.min(anchor.bottom+8,innerHeight-height-12) : anchor.top-height-8)+'px';
  };
  function initializeMenu(node) {
    const popup=node.querySelector(':scope > .simple-menu-body');
    if (typeof popup.showPopover==='function') popup.setAttribute('popover','manual');
    node.querySelector('summary').setAttribute('aria-expanded','false');
    node.addEventListener('toggle', () => {
      if (!node.open) { closeMenu(node); return; }
      menus.filter(other => other!==node && !other.contains(node) && !node.contains(other)).forEach(closeMenu);
      node.querySelector('summary').setAttribute('aria-expanded','true');
      if (popup.hasAttribute('popover') && !popup.matches(':popover-open')) popup.showPopover();
      placeMenu(node);
    });
  }
  menus.forEach(initializeMenu);
  window.addEventListener('resize',()=>menus.forEach(placeMenu));
  document.addEventListener('keydown', event => {
    if (event.key !== 'Escape' || event.isComposing || document.querySelector('dialog[open]') || window.OfficeSettings?.isOpen()) return;
    const current = [...menus].reverse().find(node => node.open && node.contains(document.activeElement)) || [...menus].reverse().find(node => node.open);
    if (!current) return;
    event.preventDefault(); event.stopImmediatePropagation(); closeMenu(current);
  }, true);
  document.addEventListener('click', event => {
    menus.filter(node => node.open && !node.contains(event.target)).forEach(closeMenu);
  });
  window.addEventListener('officeauth:unauthenticated', () => { state=null; recentList.replaceChildren(); menus.forEach(closeMenu); schedule(); });
  document.addEventListener('officesettings:opened', () => {menus.forEach(closeMenu);preferences.append($('chat-notice'));});
  document.addEventListener('officesettings:closed', () => $('daily-conversations').prepend($('chat-notice')));
  window.OfficeWorkbench = {compose: () => {
    window.OfficeWorkspace.open('work');
    inspector=null;
    update(); $('prompt').focus();
  }, state:()=>({screen:window.OfficeWorkspace.screen(),mode,inspector,historyKind}), hasFile:()=>fileSelected};
  if ($('workspace-dialogue').open) window.OfficeWorkspace.open('work');
  showSessions(window.OfficeDaily.viewSnapshot()); decorateTurns(); update();
})();
