(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const dialog = $('settings-dialog');
  if (!dialog || dialog.classList.contains('settings-workspace')) return;
  dialog.classList.add('settings-workspace');
  const pages = [
    ['machine', '通用与外观', '调整显示偏好，查看本机环境。'],
    ['model', '模型与服务', '管理连接和模型；保存与实际调用验证分别进行。'],
    ['skills', '技能', '为文档任务添加可复用的工作方法。'],
    ['rag', '文档检索', '查看索引，在明确授权的范围内查找内容。'],
    ['account', '数据与隐私', '管理本机历史、记忆和当前会话。'],
    ['updates', '关于与更新', '查看版本与发布状态，按需检查更新。']
  ];
  const make = (tag, text, parent, className) => {
    const element = document.createElement(tag);
    if (text) element.textContent = text;
    if (className) element.className = className;
    if (parent) parent.append(element);
    return element;
  };
  const group = (parent, title, elements) => {
    const section = make('section', '', null, 'settings-group');
    make('h4', title, section);
    for (const element of elements.filter(Boolean)) section.append(element);
    parent.append(section);
    return section;
  };
  const nav = dialog.querySelector('.settings-nav');
  for (const [key, title, description] of pages) {
    nav.append($('tab-' + key));
    const pane = $('settings-pane-' + key);
    const heading = make('header', '', null, 'settings-page-heading');
    make('h3', title, heading); make('p', description, heading); pane.prepend(heading);
  }
  const compactNav = make('label', '', null, 'settings-compact-nav');
  compactNav.htmlFor = 'settings-section';
  make('span', '设置分区', compactNav, 'visually-hidden');
  const sectionSelect = make('select', '', compactNav);
  sectionSelect.id = 'settings-section';
  for (const [key, title] of pages) {
    const option = make('option', title, sectionSelect); option.value = 'tab-' + key;
  }
  nav.after(compactNav);
  sectionSelect.value = dialog.dataset.tab || 'tab-model';
  sectionSelect.addEventListener('change', () => {
    window.OfficeSettings.selectTab(sectionSelect.value);
    sectionSelect.value = dialog.dataset.tab;
  });
  const machine = $('settings-pane-machine').querySelector('.settings-form');
  group(machine, '显示', [machine.querySelector('label[for="settings-theme-choice"]'), $('settings-theme-choice'), $('settings-theme-state')]);
  group(machine, '本机能力', [$('settings-probe'), $('settings-probe-status')]);
  const model = $('settings-pane-model').querySelector('.settings-form');
  const manager = $('connection-manager');
  if (manager) { manager.open = true; manager.classList.add('settings-group'); }
  const legacy = [...model.children].filter(element => ![manager, $('simple-settings-diagnostics'), $('simple-retrieval-settings')].includes(element));
  const legacyGroup = group(model, '启动默认连接', legacy);
  legacyGroup.querySelector('h4').after($('settings-status'), $('settings-load'));
  if (manager) { model.prepend(manager); manager.after(legacyGroup); }
  const rag = $('rag-panel').querySelector('.settings-form');
  const range = [], query = [];
  let querying = false;
  for (const element of [...rag.children]) {
    if (element.matches('label[for="rag-question"]')) querying = true;
    (querying ? query : range).push(element);
  }
  group(rag, '索引与读取范围', range);
  group(rag, '查找与结果', query);
  group(rag, '撤销本次操作', [$('rag-rollback-index'), $('rag-rollback-evidence')]);
  const account = $('settings-pane-account').querySelector('.settings-form');
  if ($('simple-history-preferences')) account.firstElementChild.after($('simple-history-preferences'));
  group(account, '当前工作区会话', [$('auth-logout'), $('auth-workspace-status')]).classList.add('settings-danger-zone');
  const about = $('settings-pane-updates');
  if ($('nav-about')) about.append($('nav-about'));
  const preferenceNodes = [...about.children].filter(element => element.tagName === 'LABEL' || element.tagName === 'P' && !element.id);
  const updatePrefs = group(about, '更新偏好', preferenceNodes);
  $('updates-changelog').parentElement.insertBefore(updatePrefs, $('updates-changelog').previousElementSibling);
  for (const label of dialog.querySelectorAll('label[for]')) {
    const control = $(label.htmlFor);
    if (!control || control.parentElement !== label.parentElement || !control.matches('input,select,textarea')) continue;
    const row = make('div', '', null, 'settings-row');
    label.before(row); row.append(label, control);
    if (control.matches('textarea,input[type="password"]') || /url|key|models|body|question|memory|multi/.test(control.id)) row.classList.add('settings-row-wide');
  }
  for (const label of dialog.querySelectorAll('label')) {
    if (label.querySelector('input[type="checkbox"]')) label.classList.add('settings-switch-row');
  }
  for (const parent of [legacyGroup, ...rag.querySelectorAll('.settings-group'), machine, account, updatePrefs, $('connection-editor')].filter(Boolean)) {
    const notes = [...parent.children].filter(element => element.tagName === 'P' && !element.id && element.textContent.length > 110);
    if (!notes.length) continue;
    const details = make('details', '', parent, 'settings-explanation');
    make('summary', '使用范围与详细说明', details); notes.forEach(note => details.append(note));
  }
  const search = make('input', '', null);
  search.id = 'settings-search'; search.type = 'search'; search.placeholder = '查找设置'; search.maxLength = 80;
  search.setAttribute('aria-label', '查找设置名称');
  dialog.querySelector('.settings-head').insertBefore(search, $('settings-expand'));
  const expand = $('settings-expand');
  const expandStatus = make('p', '', null, 'settings-window-status');
  expandStatus.id = 'settings-window-status'; expandStatus.setAttribute('role', 'status');
  dialog.querySelector('.settings-head').after(expandStatus);
  let expanding = false;
  const updateExpand = () => {
    const native = document.body.dataset.desktopWindow === 'true';
    const expanded = native ? document.body.dataset.windowMaximized === 'true' : dialog.dataset.expanded === 'true';
    const label = native ? expanded ? '还原窗口' : '放大窗口' : expanded ? '还原设置大小' : '展开设置';
    expand.title = label; expand.setAttribute('aria-label', label); expand.setAttribute('aria-pressed', String(expanded));
    expand.dataset.expanded = String(expanded);
    expand.querySelector('path').setAttribute('d', expanded ? 'M3 8h5V3M16 3v5h5M21 16h-5v5M8 21v-5H3' : 'M8 3H3v5M16 3h5v5M3 16v5h5M21 16v5h-5');
  };
  expand.addEventListener('click', async () => {
    if (expanding) return;
    if (document.body.dataset.desktopWindow !== 'true') {
      dialog.dataset.expanded = String(dialog.dataset.expanded !== 'true'); updateExpand(); return;
    }
    expanding = true; expand.disabled = true; expandStatus.textContent = '';
    try {
      const result = await window.OfficeWindowChrome.maximize();
      if (!result?.ok) expandStatus.textContent = '窗口未能放大，请使用窗口边缘调整大小。';
    } catch (_) { expandStatus.textContent = '窗口操作未完成，请使用窗口边缘调整大小。'; }
    finally { expanding = false; expand.disabled = false; updateExpand(); }
  });
  new MutationObserver(updateExpand).observe(document.body, {attributes:true, attributeFilter:['data-desktop-window','data-window-maximized']});
  updateExpand();
  const results = make('section', '', null, 'settings-search-results');
  results.hidden = true; results.setAttribute('aria-label', '设置查找结果');
  dialog.querySelector('.settings-panes').prepend(results);
  const entries = [];
  for (const [key, title] of pages) {
    const pane = $('settings-pane-' + key);
    entries.push({key, title, text: title, target: pane.querySelector('.settings-page-heading h3')});
    for (const label of pane.querySelectorAll('label,summary')) {
      const text = [...label.childNodes].filter(child => child.nodeType === Node.TEXT_NODE).map(child => child.textContent).join(' ').trim();
      if (!text) continue;
      entries.push({key, title, text, target: label.htmlFor ? $(label.htmlFor) : label.querySelector('input,select,textarea') || label});
    }
  }
  const lookup = () => {
    results.replaceChildren();
    const query = search.value.trim().toLocaleLowerCase();
    results.hidden = !query;
    if (!query) return;
    make('h3', '查找设置', results);
    const matches = entries.filter(entry => (entry.title + ' ' + entry.text).toLocaleLowerCase().includes(query)).slice(0, 20);
    if (!matches.length) make('p', '没有匹配的设置名称。试试“主题”“模型”或“技能”。', results);
    for (const entry of matches) {
      const button = make('button', entry.text + ' · ' + entry.title, results, 'settings-search-result'); button.type = 'button';
      button.addEventListener('click', () => {
        window.OfficeSettings.selectTab('tab-' + entry.key);
        if ($('settings-pane-' + entry.key).hidden) return;
        search.value = ''; results.hidden = true;
        for (let parent = entry.target?.parentElement; parent && parent !== dialog; parent = parent.parentElement) if (parent.tagName === 'DETAILS') parent.open = true;
        entry.target?.scrollIntoView({block: 'center'});
        entry.target?.focus();
      });
    }
  };
  search.addEventListener('input', lookup);
  search.addEventListener('keydown', event => { if (event.key === 'Enter' && !event.isComposing) { event.preventDefault(); results.querySelector('button')?.focus(); } });
  document.addEventListener('officesettings:tab', () => { sectionSelect.value = dialog.dataset.tab; search.value = ''; results.hidden = true; dialog.querySelector('.settings-panes').scrollTop = 0; });
  document.addEventListener('officesettings:closed', () => { search.value = ''; results.hidden = true; });
  const updateSave = () => {
    const footer = $('simple-settings-actions');
    if (footer) footer.hidden = $('settings-save').disabled && $('settings-discard').disabled && $('settings-rollback').disabled;
    for (const id of ['connection-save', 'connection-discard', 'connection-undo']) if ($(id)) $(id).hidden = $(id).disabled;
  };
  for (const id of ['settings-save', 'settings-discard', 'settings-rollback', 'connection-save', 'connection-discard', 'connection-undo']) {
    if ($(id)) new MutationObserver(updateSave).observe($(id), {attributes: true, attributeFilter: ['disabled']});
  }
  updateSave();
})();
