(() => {
  'use strict';
  const make = (tag, text, parent, className) => {
    const node = document.createElement(tag);
    if (text) node.textContent = text;
    if (className) node.className = className;
    parent.append(node);
    return node;
  };
  function iconButton(node, icon, label = node.textContent, hint = label) {
    node.classList.add('agent-icon-button');
    node.setAttribute('aria-label', label);
    node.dataset.tooltip = hint;
    let image = node.querySelector(':scope > .agent-icon'), text = node.querySelector(':scope > .agent-visually-hidden');
    if (!image || !text) {
      node.replaceChildren();
      image = make('span', '', node); image.setAttribute('aria-hidden', 'true');
      text = make('span', '', node, 'agent-visually-hidden');
    }
    image.className = 'agent-icon agent-icon-' + icon;
    text.textContent = label;
    if (target === node && tooltip && !tooltip.hidden) showTooltip(node, true);
    return node;
  }
  let tooltip, target, timer;
  function hideTooltip() {
    clearTimeout(timer);
    if (target?.getAttribute('aria-describedby') === 'agent-button-tooltip') target.removeAttribute('aria-describedby');
    target = null;
    if (tooltip) tooltip.hidden = true;
  }
  function showTooltip(node, immediate = false) {
    hideTooltip();
    if (!node?.dataset.tooltip) return;
    target = node;
    const show = () => {
      if (target !== node || !node.isConnected || !node.getClientRects().length) return;
      if (!tooltip) {
        tooltip = make('div', '', document.body, 'agent-button-tooltip');
        tooltip.id = 'agent-button-tooltip';
        tooltip.setAttribute('role', 'tooltip');
      }
      tooltip.textContent = node.dataset.tooltip;
      tooltip.hidden = false;
      const box = node.getBoundingClientRect(), size = tooltip.getBoundingClientRect();
      tooltip.style.left = Math.max(8, Math.min(innerWidth - size.width - 8, box.left + (box.width - size.width) / 2)) + 'px';
      tooltip.style.top = (box.bottom + size.height + 14 < innerHeight ? box.bottom + 8 : Math.max(8, box.top - size.height - 8)) + 'px';
      node.setAttribute('aria-describedby', tooltip.id);
    };
    if (immediate) show(); else timer = setTimeout(show, 240);
  }
  document.addEventListener('pointerover', event => {
    const node = event.target.closest?.('.agent-icon-button');
    if (node && !node.contains(event.relatedTarget)) showTooltip(node);
  });
  document.addEventListener('pointerout', event => { if (target && !target.contains(event.relatedTarget)) hideTooltip(); });
  document.addEventListener('focusin', event => showTooltip(event.target.closest?.('.agent-icon-button'), true));
  document.addEventListener('focusout', hideTooltip);
  document.addEventListener('pointerdown', hideTooltip, true);
  document.addEventListener('keydown', event => { if (event.key === 'Escape') hideTooltip(); });
  document.addEventListener('scroll', hideTooltip, true);
  window.addEventListener('resize', () => {
    const focused = document.activeElement?.closest?.('.agent-icon-button');
    if (focused) showTooltip(focused, true); else hideTooltip();
  });
  function edit(parent, {before, after, title, status, reason, risks = []}) {
    const first = Array.from(before), second = Array.from(after);
    let start = 0, end = 0;
    while (start < Math.min(first.length, second.length) && first[start] === second[start]) start++;
    while (end < Math.min(first.length, second.length) - start && first[first.length - end - 1] === second[second.length - end - 1]) end++;
    const removed = first.length - start - end, added = second.length - start - end;
    const heading = make('header', '', parent, 'agent-tool-heading');
    make('span', '', heading, 'agent-icon agent-icon-file').setAttribute('aria-hidden', 'true');
    make('strong', title || '文字修改', heading, 'agent-tool-title');
    const counts = make('span', '', heading, 'agent-edit-counts');
    counts.setAttribute('aria-label', '删除 ' + removed + ' 字，新增 ' + added + ' 字');
    make('span', '−' + removed, counts, 'agent-removed-count');
    make('span', '+' + added, counts, 'agent-added-count');
    make('span', status, parent, 'agent-tool-status');
    if (reason) make('p', reason, parent, 'agent-tool-reason');
    const diff = make('div', '', parent, 'agent-edit-body');
    for (const [name, label, text] of [['before', '− 原文', before], ['after', '+ 建议', after]]) {
      const block = make('section', '', diff, 'agent-edit-' + name);
      make('span', label, block, 'agent-edit-label');
      make('pre', text, block);
    }
    for (const risk of risks) make('p', risk, parent, 'agent-tool-warning');
    const actions = make('div', '', parent, 'agent-tool-actions');
    return actions;
  }
  window.OfficeAgentElements = {edit, iconButton};
})();
