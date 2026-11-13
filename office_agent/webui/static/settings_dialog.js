/* 设置窗口：把原本挤在任务栏里的配置内容收进独立对话框，只做显示与焦点管理。 */
(() => {
  'use strict';
  const node = id => document.getElementById(id);
  const TABS = [
    ['tab-model', 'settings-pane-model'],
    ['tab-skills', 'settings-pane-skills'],
    ['tab-rag', 'settings-pane-rag'],
    ['tab-machine', 'settings-pane-machine'],
    ['tab-updates', 'settings-pane-updates'],
    ['tab-account', 'settings-pane-account']
  ];
  let open = false, previousFocus = null, initialized = false, backgroundWasInert = false;
  const canFocus = element => Boolean(element) && !element.hidden && !element.disabled &&
    typeof element.focus === 'function' && (element.offsetParent !== null || element.offsetParent === undefined);

  function selectTab(tabId) {
    const known = TABS.some(([id]) => id === tabId) ? tabId : TABS[0][0];
    const dialog = node('settings-dialog');
    if (dialog.dataset.tab === 'tab-skills' && known !== 'tab-skills' && window.OfficeSkills?.requestClose() === false) return;
    dialog.dataset.tab = known;
    TABS.forEach(([id, paneId]) => {
      const tab = node(id), pane = node(paneId);
      if (!tab || !pane) return;
      const active = id === known;
      tab.classList.toggle('active', active);
      if (active) tab.setAttribute('aria-current', 'true'); else tab.removeAttribute('aria-current');
      pane.hidden = !active;
    });
    document.dispatchEvent(new CustomEvent('officesettings:tab', { detail: { tab: known } }));
  }
  function visibleTabs() {
    return TABS.map(([id]) => node(id)).filter(tab => tab && !tab.hidden && !tab.disabled);
  }
  function onKey(event) {
    if (!open) return;
    if (event.isComposing) return;
    const modal = document.querySelector?.('dialog[open]');
    if (modal && !node('settings-dialog').contains(modal)) return;
    if (event.key === 'Escape') {
      event.preventDefault(); event.stopImmediatePropagation(); close();
      return;
    }
    if (event.key !== 'Tab') return;
    // 焦点圈在设置窗口内；窗口本身没有可聚焦项时把焦点交回入口按钮。
    const dialog = node('settings-dialog');
    const targets = Array.from(dialog.querySelectorAll(
      'button, input, select, textarea, summary, a[href], [tabindex]:not([tabindex="-1"])')).filter(canFocus);
    if (!targets.length) { event.preventDefault(); return; }
    const first = targets[0], last = targets[targets.length - 1];
    const inside = dialog.contains(document.activeElement);
    if (!inside || (event.shiftKey && document.activeElement === first) || (!event.shiftKey && document.activeElement === last)) {
      event.preventDefault(); (event.shiftKey ? last : first).focus();
    }
  }
  function onBackdrop(event) {
    if (open && event.target === node('settings-backdrop')) close();
  }
  function show() {
    if (open) return;
    const dialog = node('settings-dialog');
    if (!dialog) return;
    open = true;
    previousFocus = document.activeElement;
    if (node('workspace')) { backgroundWasInert = node('workspace').inert; node('workspace').inert = true; }
    node('settings-backdrop').hidden = false;
    dialog.hidden = false;
    selectTab(node('settings-dialog').dataset.tab || TABS[0][0]);
    const target = visibleTabs()[0] || node('settings-close');
    if (target) target.focus();
    document.dispatchEvent(new CustomEvent('officesettings:opened', { detail: { dialog } }));
  }
  function close() {
    if (!open) return;
    if (window.OfficeSettingsEditor?.requestClose() === false) return;
    if (window.OfficeModelSelector?.requestClose() === false) return;
    if (window.OfficeSkills?.requestClose() === false) return;
    const dialog = node('settings-dialog');
    open = false;
    dialog.hidden = true;
    node('settings-backdrop').hidden = true;
    if (node('workspace')) node('workspace').inert = backgroundWasInert;
    document.dispatchEvent(new CustomEvent('officesettings:closed', { detail: { dialog } }));
    const back = previousFocus;
    previousFocus = null;
    if (canFocus(back)) back.focus();
    else { const entry = node('nav-settings'); if (canFocus(entry)) entry.focus(); }
  }
  function init() {
    if (initialized) return;
    if (!node('settings-dialog') || !node('nav-settings')) return;
    initialized = true;
    TABS.forEach(([id]) => node(id)?.addEventListener('click', () => selectTab(id)));
    node('nav-settings').addEventListener('click', show);
    node('settings-close').addEventListener('click', close);
    node('settings-backdrop').addEventListener('click', onBackdrop);
    document.addEventListener('keydown', onKey, true);
    selectTab(TABS[0][0]);
  }
  window.OfficeSettings = { show, close, isOpen: () => open, selectTab: tabId => { if (open) selectTab(tabId); } };
  document.addEventListener('DOMContentLoaded', init);
  if (document.readyState !== 'loading') init();
})();
