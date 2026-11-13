(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  if (!$('workspace') || $('document-workspace')) return;
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
  const move = (id, parent) => { if ($(id)) parent.append($(id)); };
  const label = id => document.querySelector(`label[for="${id}"]`);
  const field = (id, parent) => { if (label(id)) parent.append(label(id)); move(id, parent); };
  const span = (first, last, parent) => {
    const nodes = [];
    for (let node = first; node; node = node.nextSibling) { nodes.push(node); if (node === last) break; }
    nodes.forEach(node => parent.append(node));
  };
  const disclosure = (id, title, parent) => {
    const node = make('details', id, '', parent); make('summary', '', title, node); return node;
  };
  const shell = make('section', 'document-workspace');
  shell.setAttribute('aria-label', '文档编辑工作区');
  shell.dataset.screen = 'work'; shell.dataset.pane = 'document'; shell.dataset.mode = 'task';
  $('work-view').prepend(shell);
  $('workspace').classList.add('document-layout');
  const top = make('div', 'workspace-document-bar');
  const strip = document.querySelector('.file-strip');
  top.append(strip);
  const identity = make('span', 'workspace-document-identity', '尚未取得版本身份', top);
  const currentStatus = make('p', 'workspace-current-status', '', top);
  currentStatus.setAttribute('role', 'status');
  document.querySelector('.topbar').append(top);
  const views = make('nav', 'workspace-view-switch', '', shell);
  views.setAttribute('aria-label', '当前工作区视图');
  const stage = make('div', 'workspace-stage', '', shell);
  const documentPane = make('section', 'workspace-document', '', stage);
  documentPane.setAttribute('aria-label', '文档结构文本');
  make('h2', '', '文档结构', documentPane);
  make('p', '', '只读结构文本，可选择文字。这里不是 Word 分页编辑画布；读取前仍需同意范围。', documentPane);
  const taskPane = make('aside', 'workspace-task', '', stage);
  taskPane.setAttribute('aria-label', '修改要求与保护');
  make('h2', '', '修改要求', taskPane);
  const modes = make('div', 'workspace-task-modes', '', taskPane);
  modes.setAttribute('role', 'group'); modes.setAttribute('aria-label', '任务方式');
  const reviewPane = make('section', 'workspace-review', '', stage);
  reviewPane.setAttribute('aria-label', '集中修改审阅');
  make('h2', '', '修改审阅', reviewPane);
  make('p', '', '采用建议只加入待修改项。预览实际影响后，独立确认才会生成新副本。每组为独立事务。', reviewPane);
  reviewPane.dataset.kind = 'task';
  const reviewKinds = make('div', 'workspace-review-kinds', '', reviewPane);
  reviewKinds.setAttribute('role', 'group'); reviewKinds.setAttribute('aria-label', '审阅内容');
  for (const [kind, title] of [['task', '完整任务'], ['range', '选区修改'], ['pages', '静态页图']]) button('workspace-review-kind-' + kind, title, reviewKinds, () => setReviewKind(kind));
  const resources = make('section', 'workspace-resources', '', stage);
  resources.setAttribute('aria-label', '资料与检索');
  make('h2', '', '资料与检索', resources);
  make('p', '', '切换入口保留文档和任务草稿，不自动读取、联网或调用模型。', resources);
  const rangeGrid = $('range-panel').querySelector('.workbench-grid');
  const [source, editor, queue] = Array.from(rangeGrid.children);
  source.id = 'workspace-range-source'; editor.id = 'workspace-range-editor'; queue.id = 'workspace-range-review';
  const readActions = $('range-load').parentElement;
  documentPane.append(readActions, source);
  taskPane.append(editor);
  reviewPane.append(queue);
  const rangeExtra = disclosure('workspace-range-advanced', '范围边界与直接格式继承', editor);
  span(label('range-inheritance'), $('range-inheritance').nextElementSibling, rangeExtra);
  const sourceExtra = disclosure('workspace-source-details', '标记边界与只读域详情', source);
  field('range-island', sourceExtra); move('range-structures', sourceExtra);
  $('range-original').rows = 16;
  const taskReview = make('section', 'workspace-task-review', '', reviewPane);
  make('h3', '', '完整任务 · 范围与完成账本', taskReview);
  const taskPanel = $('doc-task-panel');
  taskPane.append(taskPanel);
  const request = make('div', 'workspace-request');
  field('doc-task-request', request);
  request.append($('doc-task-create').parentElement);
  taskPanel.prepend(request);
  const modelActions = disclosure('workspace-model-actions', '模型润色与对话方案（主动调用）', request);
  move('doc-task-polish', modelActions); move('doc-task-import', modelActions);
  $('doc-task-create').className = 'btn primary';
  $('doc-task-create').textContent = '生成精确术语建议';
  const protection = disclosure('workspace-protection', '保护要求与适用限制', taskPanel);
  span(label('doc-task-profile'), $('doc-task-protection'), protection);
  const scope = disclosure('workspace-scope', '大纲、精确搜索与章节范围', documentPane);
  span(label('doc-task-query'), $('doc-task-section'), scope);
  const advanced = disclosure('workspace-task-advanced', '高级任务与结构操作', taskPanel);
  Array.from(taskPanel.children).filter(node => node.tagName === 'DETAILS' && ![protection, advanced, $('doc-task-return-panel')].includes(node)).forEach(node => advanced.append(node));
  const summaryDetails = disclosure('workspace-task-checks', '覆盖、保护和确定性检查详情', taskReview);
  move('doc-task-summary', summaryDetails);
  move('doc-task-ledger', taskReview); move('doc-task-next', taskReview);
  const taskActions = $('doc-task-route').parentElement;
  taskReview.append(taskActions);
  const readablePlan = make('div', 'workspace-readable-plan', '', taskReview);
  readablePlan.setAttribute('aria-label', '本组实际变化');
  ['doc-task-preview', 'doc-task-visual'].forEach(id => move(id, taskReview));
  const deliveries = make('section', 'workspace-deliveries', '', reviewPane);
  make('h3', '', '已生成的副本与继续编辑', deliveries);
  ['doc-task-deliveries', 'doc-task-prepare-return'].forEach(id => move(id, deliveries));
  const technicalPlan = disclosure('workspace-plan-details', '执行路线、定位覆盖和完整确认明细', taskReview);
  move('doc-task-preview', technicalPlan);
  const continuation = disclosure('workspace-continuation', '原稿 → 系统副本 → 用户当前稿', reviewPane);
  move('doc-task-return-panel', continuation);
  const archive = disclosure('workspace-task-archive', '保存或恢复任务（不恢复批准）', taskReview);
  field('doc-task-receipt', archive); move('doc-task-restore', archive);
  const rangeReview = $('workspace-range-review');
  ['range-review', 'range-assessment', 'range-status'].forEach(id => move(id, rangeReview));
  move('range-delivery', reviewPane);
  const deliveryDetails = disclosure('workspace-review-archive', '回执与审阅存档详情', $('range-delivery'));
  field('range-delivery-id', deliveryDetails); move('range-delivery-help', deliveryDetails);
  field('range-archive-id', deliveryDetails); move('range-restore-review', deliveryDetails); move('range-remove-review', deliveryDetails);
  const comparisons = make('div', 'workspace-page-controls');
  comparisons.setAttribute('role', 'group'); comparisons.setAttribute('aria-label', '静态页面显示');
  const pair = document.querySelector('.range-page-pair');
  pair.before(comparisons); pair.dataset.side = 'before';
  for (const [side, name] of [['before', '修改前'], ['after', '修改后']]) button('workspace-page-' + side, name, comparisons, () => {
    pair.dataset.side = side;
    Array.from(comparisons.children).forEach(node => node.setAttribute('aria-pressed', String(node.id === 'workspace-page-' + side)));
  }).setAttribute('aria-pressed', String(side === 'before'));
  const previous = button('workspace-review-previous', '上一审阅项', comparisons, () => {
    const select = $('range-review-item');
    if (!select.disabled && select.selectedIndex > 0) { select.selectedIndex--; select.dispatchEvent(new Event('change')); }
  });
  const updatePrevious = () => { previous.disabled = $('range-review-item').disabled || $('range-review-item').selectedIndex <= 0; };
  new MutationObserver(updatePrevious).observe($('range-review-item'), {attributes: true, childList: true});
  $('range-review-item').addEventListener('change', updatePrevious); updatePrevious();
  const legacy = $('professional-panel');
  legacy.querySelector('summary').textContent = '兼容：整段修改与专业对话授权';
  const tools = disclosure('workspace-dialogue', '对话与辅助工具', resources);
  tools.append(document.querySelector('.composer'), $('status-card'), $('pending-notice'), $('content-panes'));
  ['workbench-panel', 'continuous-panel', 'literature-panel', 'assist-panel', 'grounded-panel', 'citations-panel'].forEach(id => {
    const node = $(id);
    if (node.tagName === 'DETAILS') resources.append(node);
    else { const parent = disclosure('workspace-tool-' + id, node.getAttribute('aria-label'), resources); parent.append(node); }
  });
  resources.append(legacy);
  const resourceSettings = make('div', 'workspace-resource-settings', '', resources);
  button('workspace-rag-settings', '文档索引与 RAG 检索', resourceSettings, () => {
    window.OfficeSettings?.show(); window.OfficeSettings?.selectTab('tab-rag');
  });
  const account = button('nav-account', '账户', document.querySelector('.nav-group-tail'), () => {
    window.OfficeSettings?.show(); window.OfficeSettings?.selectTab('tab-account');
  });
  account.className = 'nav-item'; account.title = '账户与会话'; account.setAttribute('aria-haspopup', 'dialog');
  const filterLabel = make('label', '', '筛选修改项');
  filterLabel.htmlFor = 'workspace-review-filter';
  const filter = make('select', 'workspace-review-filter');
  for (const [value, title] of [['all', '全部'], ['pending', '待审阅'], ['adopted', '已采用（尚未写入）'], ['needs_review', '待复核或失败'], ['unsupported', '未支持'], ['completed', '已执行并验证']]) {
    const option = make('option', '', title, filter); option.value = value;
  }
  reviewPane.querySelector('p').after(filterLabel, filter);
  const filterStatus = make('p', 'workspace-filter-status'); filter.after(filterStatus);
  filterStatus.setAttribute('role', 'status');
  const applyFilter = () => {
    const rows = document.querySelectorAll('#range-queue > [data-review-state], #doc-task-ledger > [data-review-state]');
    rows.forEach(row => { row.hidden = filter.value !== 'all' && row.dataset.reviewState !== filter.value; });
    const activeRows = Array.from(rows).filter(row => reviewPane.dataset.kind === 'range' ? row.closest('#range-queue') : row.closest('#doc-task-ledger'));
    filterStatus.textContent = activeRows.length ? `当前筛选 ${activeRows.filter(row => !row.hidden).length} / ${activeRows.length} 项；仅当前已读取的清单，后续账本需主动翻页。` : '尚无修改建议。先在工作台读取文档并准备任务。';
  };
  filter.onchange = applyFilter;
  ['range-queue', 'doc-task-ledger'].forEach(id => new MutationObserver(applyFilter).observe($(id), {childList: true, subtree: true, attributes: true, attributeFilter: ['data-review-state']}));
  $('range-panel').hidden = true;
  const rail = $('nav-work').parentElement;
  const navButtons = new Map();
  $('nav-work').querySelector('span').textContent = '工作台';
  $('nav-work').setAttribute('aria-label', '工作台'); $('nav-work').title = '工作台';
  const nav = (key, title, id) => {
    const node = button(id, title, rail, () => open(key));
    node.className = 'nav-item'; node.title = title; node.setAttribute('aria-controls', 'document-workspace');
    $('nav-history').before(node); navButtons.set(key, node); return node;
  };
  navButtons.set('work', $('nav-work'));
  nav('review', '修改审阅', 'nav-review'); nav('resources', '资料与检索', 'nav-resources');
  $('nav-work').addEventListener('click', () => open('work'));
  $('nav-history').addEventListener('click', () => { shell.hidden = true; navButtons.forEach(node => node.removeAttribute('aria-current')); });
  const paneButtons = new Map();
  for (const [key, title] of [['document', '文档'], ['task', '任务'], ['review', '审阅']]) paneButtons.set(key, button('workspace-view-' + key, title, views, () => {
    shell.dataset.pane = key; if (key === 'review') open('review'); else if (shell.dataset.screen !== 'work') open('work');
    update();
  }));
  for (const [key, title] of [['task', '完整任务'], ['range', '选区修改']]) button('workspace-mode-' + key, title, modes, () => {
    shell.dataset.mode = key; update();
  });
  const quick = make('div', 'workspace-quick-actions', '', request);
  quick.setAttribute('role', 'group'); quick.setAttribute('aria-label', '准备任务，不自动执行');
  for (const [title, text] of [['统一术语', '将全文所有“旧术语”替换为“规范术语”，保留数字、单位、公式和引文'], ['润色', '润色所选章节，保留事实、数字、单位、公式、引文和限定条件']]) button('', title, quick, () => {
    const input = $('doc-task-request'); if (input.value && !window.confirm('替换尚未提交的修改要求？不读取、不调用模型、不写入。')) return;
    input.value = text; input.dispatchEvent(new Event('input', {bubbles: true})); input.focus();
    if (title === '润色') { modelActions.open = true; scope.open = true; }
  });
  function open(screen) {
    if (typeof modalOpen !== 'undefined' && modalOpen || window.OfficeSettings?.isOpen()) return;
    if (typeof navigate === 'function') navigate('work');
    shell.hidden = false; shell.dataset.screen = screen;
    if (screen === 'review') shell.dataset.pane = 'review';
    else if (screen === 'work' && shell.dataset.pane === 'review') shell.dataset.pane = 'document';
    update();
  }
  function update() {
    navButtons.forEach((node, key) => {
      const active = !$('work-view').hidden && shell.dataset.screen === key;
      node.classList.toggle('active', active);
      if (active) node.setAttribute('aria-current', 'page'); else node.removeAttribute('aria-current');
    });
    paneButtons.forEach((node, key) => node.setAttribute('aria-pressed', String(shell.dataset.pane === key)));
    Array.from(modes.children).forEach(node => node.setAttribute('aria-pressed', String(node.id === 'workspace-mode-' + shell.dataset.mode)));
    if (!$('work-view').hidden) $('page-title').textContent = {work: '文档工作台', review: '修改审阅', resources: '资料与检索'}[shell.dataset.screen];
    const taskActive = shell.dataset.mode === 'task';
    currentStatus.textContent = shell.dataset.screen === 'resources' ? $('status-line').textContent : taskActive ? $('doc-task-workflow').textContent : $('range-status').textContent;
    window.dispatchEvent(new CustomEvent('officeworkspace:view', {detail: {screen: shell.dataset.screen, pane: shell.dataset.pane}}));
  }
  document.addEventListener('focusin', () => { if (!$('work-view').hidden) update(); });
  ['doc-task-workflow', 'range-status', 'status-line'].forEach(id => new MutationObserver(update).observe($(id), {childList: true, subtree: true, characterData: true}));
  new MutationObserver(() => {
    identity.textContent = '当前选择 · 身份以宿主读取和计划为准';
    update();
  }).observe($('file-name'), {childList: true, subtree: true, characterData: true});
  window.addEventListener('officeworkspace:task', event => {
    identity.textContent = event.detail?.current_name ? `原稿：${event.detail.source_name} → 当前：${event.detail.current_name} · ${event.detail.receipts?.length || 0}次独立事务` : '尚未取得版本身份';
    readablePlan.replaceChildren();
  });
  const actionNames = {replace: '替换文字', insert: '插入文字', delete: '删除文字', format: '字符格式', paragraph_format: '段落格式', paragraphs_replace: '跨段替换', chapter_move: '章节移动', comment: '添加批注'};
  window.addEventListener('officeworkspace:plan', event => {
    readablePlan.replaceChildren();
    const {before, after, group, groups, input, mode} = event.detail;
    make('h3', '', `本组 ${after.length} 项 · 第 ${group}/${groups} 组`, readablePlan);
    make('p', '', `实际输入：${input}；${mode === 'tracked' ? '原生文字修订' : '净稿'}；输出为新副本。分页可能改变，引用与格式影响以完整确认明细和宿主校验为准。`, readablePlan);
    after.forEach((item, index) => {
      const row = make('article', '', '', readablePlan);
      make('h4', '', `${index + 1}. ${actionNames[item.action] || item.action} · ${before[index]?.location || '见宿主定位'}`, row);
      make('p', '', '修改前：' + (before[index]?.text ?? before[index]?.texts?.join('\n') ?? '见完整明细'), row);
      make('p', '', '修改后：' + (item.text ?? item.texts?.join('\n') ?? '格式或位置调整，见完整明细'), row);
      make('p', '', '原因：' + item.reason, row);
    });
  });
  window.addEventListener('officeworkspace:clear', () => { readablePlan.replaceChildren(); });
  function setReviewKind(kind) {
    reviewPane.dataset.kind = kind;
    Array.from(reviewKinds.children).forEach(node => node.setAttribute('aria-pressed', String(node.id === 'workspace-review-kind-' + kind)));
    filter.hidden = filterLabel.hidden = filterStatus.hidden = kind === 'pages';
    applyFilter();
  }
  window.addEventListener('officeworkspace:review', () => { open('review'); setReviewKind('pages'); });
  window.addEventListener('officeworkspace:locate', () => { open('work'); shell.dataset.pane = 'document'; update(); });
  window.addEventListener('officeworkspace:escape', () => {
    if (shell.dataset.screen !== 'work') open('work');
    const target = shell.dataset.pane === 'document' && matchMedia('(max-width:750px)').matches ? $('range-original') : shell.dataset.mode === 'task' ? $('doc-task-request') : $('range-text');
    target.focus();
  });
  $('range-compose').addEventListener('click', () => { if ($('prompt').value) { open('resources'); tools.open = true; $('prompt').focus(); } });
  window.addEventListener('officeworkspace:compose', () => {
    if (window.OfficeWorkbench) window.OfficeWorkbench.compose();
    else { open('resources'); tools.open = true; $('prompt').focus(); }
  });
  button('workspace-see-task-review', '查看建议与整组影响', taskPanel, () => { open('review'); setReviewKind('task'); });
  button('workspace-see-range-review', '查看待修改清单', editor, () => { open('review'); setReviewKind('range'); });
  window.addEventListener('officeauth:unauthenticated', () => {
    identity.textContent = '会话已清理，需重新读取'; shell.dataset.screen = 'work'; shell.dataset.pane = 'document'; update();
  });
  new MutationObserver(() => {
    const nodes = Array.from($('range-queue').children).filter(node => node.id !== 'workspace-queue-empty');
    if (nodes.length) $('workspace-queue-empty')?.remove();
    else if (!$('workspace-queue-empty')) make('p', 'workspace-queue-empty', '尚无选区修改建议。在工作台读取文段、选中文字后加入清单。', $('range-queue'));
  }).observe($('range-queue'), {childList: true});
  $('doc-task-prepare-return').addEventListener('click', () => { continuation.open = true; });
  window.OfficeWorkspace = {open, screen: () => shell.dataset.screen};
  setReviewKind('task'); update();
})();
