"use strict";
(() => {
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  let ready = false, epoch = 0, pending = null;
  const wrappers = new WeakMap();
  const raw = () => window.pywebview?.api;
  function gate(value) {
    ready = value;
    const workspace = document.getElementById('workspace');
    if (workspace) { workspace.hidden = !value; workspace.inert = !value; workspace.setAttribute('aria-hidden', String(!value)); }
    const status = document.getElementById('auth-screen');
    if (status) { status.hidden = value; status.textContent = value ? '' : '本机工作区未就绪，请重新打开应用。'; }
  }
  function wrapApi(api) {
    if (!api) return null;
    if (wrappers.has(api)) return wrappers.get(api);
    const wrapped = new Proxy(api, {get(target, key) {
      const method = target[key];
      if (typeof method !== 'function') return method;
      return async (...args) => {
        const mine = epoch;
        if (!ready && key !== 'auth_status') return key === 'poll' ? [] : {ok:false, reason:'workspace_not_ready'};
        const result = await method.apply(target, args);
        if (mine !== epoch) return key === 'poll' ? [] : {ok:false, reason:'stale_workspace'};
        return result;
      };
    }});
    wrappers.set(api, wrapped); return wrapped;
  }
  async function ensure() {
    if (ready) return true;
    if (pending) return pending;
    if (!raw()) return false;
    pending = (async () => {
      try { const result = await raw().auth_status(); gate(result?.ok === true && result.authenticated === true); }
      catch (_) { gate(false); }
      return ready;
    })().finally(() => { pending = null; });
    return pending;
  }
  window.OfficeAuth = Object.freeze({ensure, wrapApi, isAuthenticated:() => ready});
  document.addEventListener('DOMContentLoaded', () => {
    const clear = document.getElementById('auth-logout');
    if (clear) clear.onclick = async () => {
      if (!await ask('清空当前界面会话、文件读取许可和未执行计划？已保存文档及历史记录保留。')) return;
      const result = await raw().logout();
      if (!result?.ok) { document.getElementById('auth-workspace-status').textContent = '当前任务未结束，暂不能清空会话。'; return; }
      epoch++; ready = false;
      window.dispatchEvent(new Event('officeauth:unauthenticated'));
      window.location.reload();
    };
  });
  window.addEventListener('pywebviewready', ensure);
})();
