(() => {
  'use strict';
  if (window.OfficeConfirm) return;
  let active = null;
  const ask = (message, options = {}) => {
    if (active) return Promise.resolve(false);
    const previous = document.activeElement;
    const dialog = document.createElement('dialog');
    dialog.className = 'paged-dialog lix-confirm-dialog'; dialog.setAttribute('aria-label', 'LiX');
    const title = document.createElement('h2'); title.textContent = 'LiX';
    const text = document.createElement('p'); text.textContent = String(message); text.style.whiteSpace = 'pre-wrap';
    const actions = document.createElement('div'); actions.className = 'paged-dialog-actions';
    const cancel = document.createElement('button'); cancel.type = 'button'; cancel.textContent = options.cancel || '取消';
    const approve = document.createElement('button'); approve.type = 'button'; approve.textContent = options.approve || '确认'; approve.className = 'paged-primary';
    actions.append(cancel, approve); dialog.append(title, text, actions); document.body.append(dialog);
    let settled = false, resolve;
    const promise = new Promise(done => { resolve = done; });
    active = {message, promise, cancel: () => finish(false)};
    const finish = accepted => {
      if (settled) return;
      settled = true; active = null;
      if (dialog.open) dialog.close();
      dialog.remove(); if (previous?.isConnected) previous.focus(); resolve(accepted === true);
    };
    cancel.onclick = () => finish(false); approve.onclick = () => finish(true);
    dialog.addEventListener('cancel', event => { event.preventDefault(); finish(false); });
    dialog.addEventListener('close', () => finish(false));
    try { dialog.showModal(); cancel.focus(); }
    catch (_) { finish(false); }
    return promise;
  };
  const dismiss = () => active?.cancel();
  for (const name of ['officeauth:authenticated', 'officeauth:unauthenticated', 'officeworkspace:clear', 'pagehide']) window.addEventListener(name, dismiss);
  const observe = () => {
    const nodes = ['file-name', 'file-dir'].map(id => document.getElementById(id)).filter(Boolean);
    let identity = nodes.map(node => node.textContent).join('\u0000');
    const changed = () => {
      const next = nodes.map(node => node.textContent).join('\u0000');
      if (next !== identity) { identity = next; dismiss(); }
    };
    for (const node of nodes) new MutationObserver(changed).observe(node, {childList:true,characterData:true,subtree:true});
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', observe, {once:true});
  else observe();
  window.OfficeConfirm = {ask, dismiss};
})();
