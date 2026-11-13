(() => {
  'use strict';
  let header, controls, bar, maximum, locked = false, secondaryOpen = false, sequence = 0, resizeTimer;
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  async function action(name) {
    if (locked) return;
    locked = true;
    const mine = ++sequence;
    try {
      const result = await api().floating_window_action(name);
      if (!result?.ok) throw Error(result?.error || '窗口操作未完成，请使用悬浮球恢复窗口');
      if (mine === sequence) update(result);
      return result;
    } catch (error) {
      const status = document.getElementById('floating-status');
      if (status) status.textContent = error.message;
      return {ok:false};
    } finally { locked = false; }
  }
  function update(state) {
    if (!maximum) return;
    document.body.dataset.windowMaximized = String(!!state.maximized);
    window.OfficeAgentElements.iconButton(maximum, state.maximized ? 'minimize-2' : 'maximize-2', state.maximized ? '还原窗口' : '最大化窗口');
    maximum.setAttribute('aria-pressed', String(!!state.maximized));
  }
  function attach(parent, collapse) {
    if (header) return;
    header = parent;
    const title = document.getElementById('floating-title');
    title.classList.add('desktop-drag-region');
    controls = document.createElement('div'); controls.className = 'desktop-window-controls';
    for (const [name, label, icon] of [['minimize', '最小化窗口', 'minus'], ['maximize', '最大化窗口', 'maximize-2']]) {
      const node = document.createElement('button'); node.type = 'button'; node.hidden = true;
      node.dataset.windowAction = name;
      window.OfficeAgentElements.iconButton(node, icon, label);
      node.addEventListener('click', () => action(name));
      controls.append(node);
      if (name === 'maximize') maximum = node;
    }
    window.OfficeAgentElements.iconButton(collapse, 'x', '收起', '收起到悬浮球');
    collapse.classList.add('desktop-window-close');
    controls.append(collapse); parent.append(controls);
    bar = document.createElement('div'); bar.id = 'desktop-window-bar'; bar.hidden = true;
    const name = document.createElement('span'); name.textContent = 'LiX'; name.className = 'desktop-drag-region';
    const logo = document.createElement('img'); logo.src = 'brand/lix.png'; logo.alt = ''; logo.className = 'lix-brand-icon'; logo.width = logo.height = 24;
    name.prepend(logo); bar.append(name); document.body.prepend(bar);
    let lastDown;
    for (const node of [title, name]) node.addEventListener('pointerdown', event => {
      if (event.button !== 0 || document.body.dataset.desktopWindow !== 'true') return;
      event.preventDefault();
      const now = performance.now(), prior = lastDown;
      lastDown = {time:now,x:event.screenX,y:event.screenY};
      const doubleClick = prior && now-prior.time < 360 && Math.hypot(event.screenX-prior.x,event.screenY-prior.y) < 4;
      if (doubleClick) action('maximize');
      else if (document.body.dataset.windowMaximized !== 'true') action('drag');
    });
    for (const edge of ['left','right','top','bottom','top-left','top-right','bottom-left','bottom-right']) {
      const handle = document.createElement('div'); handle.className = 'desktop-resize-handle'; handle.dataset.edge = edge;
      handle.setAttribute('aria-hidden','true');
      handle.addEventListener('pointerdown', event => { if(event.button===0){event.preventDefault();action('resize-'+edge);} });
      document.body.append(handle);
    }
    refresh();
  }
  async function refresh() {
    const mine = ++sequence;
    try {
      const state = await api().floating_window_state();
      if (mine !== sequence || !state?.ok || !state.available || !controls) return;
      document.body.dataset.desktopWindow = 'true';
      controls.querySelectorAll('[data-window-action]').forEach(node => { node.hidden = false; });
      update(state);
      secondary(secondaryOpen);
    } catch (_) {}
  }
  function secondary(open) {
    secondaryOpen = open;
    if (!controls) return;
    const native = document.body.dataset.desktopWindow === 'true';
    document.body.classList.toggle('desktop-secondary', open && native);
    bar.hidden = !open || !native;
    (open && native ? bar : header).append(controls);
  }
  window.addEventListener('focus', () => { if (header) refresh(); });
  window.addEventListener('resize', () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (header) refresh(); }, 160); });
  window.OfficeWindowChrome = {attach, secondary, maximize: () => action('maximize')};
})();
