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
  let screen = window.OfficeWorkspace.screen(), queued = false;
  const menus = [];
  const read = make('section', 'simple-reading', '', shell);
  read.setAttribute('aria-label', '当前任务内容'); read.tabIndex = 0;
  const welcome = make('div', 'simple-image-welcome', '', read);
  make('h2', '', '为报告生成一张配图', welcome);
  make('p', '', '描述画面，查看方案后再确认生成。', welcome);
  read.append($('daily-conversations'), $('content-panes'));
  const compose = make('section', 'simple-compose-area', '', shell);
  compose.setAttribute('aria-label', '当前任务输入');
  const composer = document.querySelector('.composer');
  compose.append($('pending-notice'), composer);
  const dock = make('div', 'simple-input-tools', '', compose);
  dock.append($('model-mode').closest('label'));
  const [models, modelBody] = menu('simple-model-menu', '选择模型', dock); menus.push(models);
  const modelState = make('span', 'simple-model-state', '', dock); modelState.setAttribute('role', 'status');
  modelBody.append($('model-selector-bar'));
  $('model-selector-bar').querySelectorAll('label').forEach(node => { if (node.contains($('model-search')) || node.contains($('model-selector'))) modelBody.prepend(node); });
  const [sendOptions, sendBody] = menu('simple-send-options', '发送选项', dock); menus.push(sendOptions);
  field('chat-stream', sendBody); move('chat-preview-context', sendBody); move('chat-cancel-continuation', sendBody);
  const [budget, budgetBody] = menu('simple-budget-menu', '回答长度', dock); menus.push(budget);
  field('chat-reply-length', budgetBody); move('chat-reply-save', budgetBody);
  make('p', '', '保存后用于后续问答；预算不保证字数。', budgetBody);
  const [imageOptions, imageBody] = menu('simple-image-options', '图片参数', dock); menus.push(imageOptions);
  move('daily-image-parameters', imageBody);
  const [assistTools, assistBody] = menu('simple-assist-menu', '辅助工具', dock); menus.push(assistTools);
  const vision = composer.querySelector('.vision-actions'); if (vision) assistBody.append(vision);
  move('tools-menu', assistBody); move('input-help', assistBody);
  const [sessionMenu, sessionBody] = menu('simple-session-menu', '会话操作', document.querySelector('.topbar')); menus.push(sessionMenu);
  make('h3', '', '当前会话', sessionBody);
  field('chat-title', sessionBody); move('chat-rename', sessionBody); move('chat-title-hint', sessionBody);
  move('chat-save-now', sessionBody); move('chat-delete', sessionBody);
  make('h3', '', '导出', sessionBody); move('chat-export-controls', sessionBody);
  make('h3', '', '保存容量', sessionBody); move('chat-capacity', sessionBody);
  const account = $('settings-pane-account');
  move('technical-details', $('settings-pane-machine'));
  move('more-menu', sessionBody);
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
  button('simple-view-preview', '预览', $('workspace-view-switch'), () => {
    window.OfficeWorkspace.open('review'); $('workspace-review-kind-pages').click();
  });
  const recent = make('section', 'simple-recent', '', rail);
  make('h2', '', '最近会话', recent);
  const recentList = make('div', 'simple-recent-list', '', recent);
  const [sessions, sessionsBody] = menu('simple-session-list', '查找全部会话', recent); menus.push(sessions);
  field('chat-search', sessionsBody); field('chat-session', sessionsBody); move('chat-refresh', sessionsBody);
  const history = make('section', 'simple-history', '', shell); history.hidden = true;
  const historyNav = make('nav', 'simple-history-kinds', '', history); historyNav.setAttribute('aria-label', '历史记录分类');
  const historyContent = make('div', 'simple-history-content', '', history);
  for (const [key, title] of [['chat','对话'], ['documents','文档任务'], ['image','图片']]) {
    button('simple-history-' + key, title, historyNav, () => showHistory(key));
  }
  const imageHistory = $('daily-image-history');
  imageHistory.querySelector('h3').textContent = '图片记录';
  $('nav-history').addEventListener('click', () => showHistory(historyKind));
  const status = $('status-card');
  compose.prepend(status);
  const docDialog = button('workspace-mode-dialogue', 'AI 对话', $('workspace-task-modes'), () => {
    shell.dataset.mode = 'dialogue'; shell.dataset.pane = 'task'; update(); $('prompt').focus();
  });
  $('workspace-mode-task').addEventListener('click', update);
  $('workspace-mode-range').addEventListener('click', update);
  $('range-compose').addEventListener('click', () => {
    if ($('prompt').value) window.OfficeWorkbench.compose();
  });
  const fileDetails = make('details', 'simple-file-details', '', $('workspace-document-bar'));
  make('summary', '', '文件详情', fileDetails);
  fileDetails.append($('workspace-document-identity'));
  const outputDetails = make('details', 'simple-output-details');
  make('summary', '', '副本路径与验证详情', outputDetails);
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
    for (const article of $('chat-transcript').children) {
      if (!article.classList.contains('daily-turn') || article.dataset.uiDecorated) continue;
      article.dataset.uiDecorated = 'true';
      const actions = make('div', '', '', article); actions.className = 'simple-message-actions';
      const [more, body] = menu('', '更多', actions); more.classList.add('simple-message-menu');
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
    if (kind === 'documents') {
      history.hidden = true;
      if (typeof navigate === 'function') navigate('history');
      $('history-reader').prepend(historyNav);
      shell.hidden = true;
    } else {
      if (typeof navigate === 'function') navigate('work');
      shell.hidden = false; history.hidden = false;
      history.prepend(historyNav);
      historyContent.replaceChildren();
      if (kind === 'image') {
        historyContent.append(imageHistory); imageHistory.hidden = false;
        if (load) $('daily-open-images').click();
      } else {
        make('h2', '', '普通问答历史', historyContent);
        make('p', '', '选择左侧会话继续查看。恢复不会发送请求或恢复批准。', historyContent);
        sessions.open = true;
      }
      screen = 'history'; update();
    }
  }
  function modelLabel() {
    const selection = window.OfficeModelSelector.selection();
    const text = selection.valid ? selection.model + ' · ' + selection.connection : '选择模型 / 配置接口';
    models.querySelector('summary').textContent = text;
    models.querySelector('summary').title = text;
    modelState.textContent = selection.valid ? selection.state : '选择失效或未配置';
    modelState.dataset.invalid = String(!selection.valid);
    modelState.title = $('model-selection-status').textContent;
  }
  function update() {
    app.dataset.uiMode = mode; app.dataset.uiScreen = screen;
    read.hidden = screen === 'resources' || screen === 'history' || mode === 'agent';
    history.hidden = screen !== 'history';
    $('workspace-stage').hidden = screen === 'history' || screen !== 'resources' && mode !== 'agent';
    $('workspace-view-switch').hidden = mode !== 'agent' || screen === 'resources' || screen === 'history';
    $('workspace-document-bar').hidden = mode !== 'agent' || screen === 'history';
    sessionMenu.hidden = mode !== 'chat' || screen === 'history' || screen === 'resources';
    if (mode === 'agent' && screen !== 'resources') {
      if (compose.parentElement !== task) task.append(compose);
      if ($('content-panes').parentElement !== read) read.append($('content-panes'));
      if (shell.dataset.mode === 'dialogue') {
        task.insertBefore(read, compose);
        read.hidden = false;
      }
    } else {
      if (read.parentElement !== shell) shell.append(read);
      if (compose.parentElement !== shell) shell.append(compose);
    }
    compose.hidden = screen === 'history' || screen === 'resources' || mode === 'agent' && screen === 'review';
    composer.hidden = mode === 'agent' && shell.dataset.mode !== 'dialogue';
    if (mode === 'chat' && $('daily-conversations').parentElement !== read) read.prepend($('daily-conversations'));
    $('content-panes').hidden = mode === 'chat';
    welcome.hidden = mode !== 'image' || Boolean($('log-pane').children.length);
    sendOptions.hidden = budget.hidden = mode !== 'chat';
    imageOptions.hidden = mode !== 'image';
    assistTools.hidden = mode === 'chat';
    docDialog.setAttribute('aria-pressed', String(shell.dataset.mode === 'dialogue'));
    const title = screen === 'resources' ? '资料库' : screen === 'history' ? '历史记录' : mode === 'chat' ? state?.title || '新会话' : mode === 'image' ? '图片生成' : screen === 'review' ? '修改审阅' : '文档工作区';
    if ($('page-title').textContent !== title) $('page-title').textContent = title;
    $('prompt').placeholder = mode === 'chat' ? '输入问题…' : mode === 'image' ? '描述要生成的报告配图…' : '描述文档修改要求…';
    $('chat-save-status').hidden = !['failed','uncertain','unsaved'].includes(state?.save_state);
    const hostFeedback = window.OfficeRequestFeedback?.snapshot()?.source;
    const dailyFeedback = window.OfficeDaily.feedbackSnapshot()?.source;
    const sameNotice = hostFeedback?.channel === 'chat' && dailyFeedback?.terminal &&
      dailyFeedback.session_id === state?.current && hostFeedback.session_id === dailyFeedback.session_id && hostFeedback.turn_id === dailyFeedback.turn_id;
    status.hidden = Boolean(hostFeedback?.channel === 'chat' && (sameNotice || mode !== 'chat' || hostFeedback.session_id !== state?.current));
    for (const node of $('chat-transcript').querySelectorAll('[data-turn-notice]')) {
      node.hidden = Boolean(mode === 'chat' && dailyFeedback?.terminal && dailyFeedback.session_id === state?.current && node.dataset.turnNotice === dailyFeedback.turn_id);
    }
    latest.hidden = following || mode !== 'chat';
    $('chat-cancel-continuation').hidden = !($('chat-context-content').textContent.includes('关联续写'));
    if (screen === 'history') {
      $('nav-work').removeAttribute('aria-current'); $('nav-resources').removeAttribute('aria-current');
      $('nav-history').setAttribute('aria-current', 'page');
    }
    modelLabel();
  }
  const schedule = () => {
    if (queued) return; queued = true;
    requestAnimationFrame(() => { queued = false; update(); });
  };
  window.addEventListener('officeworkspace:view', event => {
    screen = event.detail.screen; schedule();
  });
  window.addEventListener('officemodels:mode', event => { mode = event.detail.mode; schedule(); });
  function showSessions(value) {
    state = value; recentList.replaceChildren();
    for (const item of [...state.sessions].sort((a,b) => b.created-a.created || a.id.localeCompare(b.id)).slice(0,6)) {
      const node = button('', item.title, recentList, async () => {
        if ($('chat-session').disabled || $('model-mode').disabled) return;
        if (!await switchMode('chat')) return;
        $('chat-session').value = item.id; $('chat-session').dispatchEvent(new Event('change'));
        window.OfficeWorkspace.open('work');
      });
      node.title = item.title; node.setAttribute('aria-current', String(item.id === state.current));
    }
    schedule();
  }
  window.addEventListener('officedaily:state', event => showSessions(event.detail));
  new MutationObserver(schedule).observe($('model-selection-status'), {childList:true,subtree:true,characterData:true});
  new MutationObserver(schedule).observe($('chat-notice'), {childList:true,subtree:true,characterData:true});
  new MutationObserver(schedule).observe($('error-line'), {childList:true,subtree:true,characterData:true,attributes:true,attributeFilter:['style']});
  new MutationObserver(schedule).observe($('log-pane'), {childList:true});
  const closeMenu = node => {
    const focus = node.contains(document.activeElement); node.open = false;
    if (focus) node.querySelector('summary').focus();
  };
  for (const node of menus) node.addEventListener('toggle', () => {
    if (node.open) menus.filter(other => other !== node && !other.contains(node)).forEach(closeMenu);
  });
  document.addEventListener('keydown', event => {
    if (event.key !== 'Escape' || event.isComposing || document.querySelector('dialog[open]') || window.OfficeSettings?.isOpen()) return;
    const current = menus.find(node => node.open && node.contains(document.activeElement)) || menus.find(node => node.open);
    if (!current) return;
    event.preventDefault(); event.stopImmediatePropagation(); closeMenu(current);
  }, true);
  document.addEventListener('click', event => {
    menus.filter(node => node.open && !node.contains(event.target)).forEach(closeMenu);
  });
  window.addEventListener('officeauth:unauthenticated', () => { state=null; recentList.replaceChildren(); menus.forEach(closeMenu); schedule(); });
  document.addEventListener('officesettings:opened', () => preferences.append($('chat-notice')));
  document.addEventListener('officesettings:closed', () => $('daily-conversations').prepend($('chat-notice')));
  window.OfficeWorkbench = {compose: () => {
    window.OfficeWorkspace.open('work');
    if (mode === 'agent') { shell.dataset.mode = 'dialogue'; shell.dataset.pane = 'task'; }
    update(); $('prompt').focus();
  }};
  if ($('workspace-dialogue').open) window.OfficeWorkspace.open('work');
  showSessions(window.OfficeDaily.viewSnapshot()); decorateTurns(); update();
})();
