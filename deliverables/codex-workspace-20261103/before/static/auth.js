/* Local authentication UI. Credentials stay in memory only; the bridge is the security boundary. */
"use strict";
(() => {
  const node = id => document.getElementById(id);
  let authenticated = false, ready = false, submitting = false, mode = "login";
  let statusPromise = null, sessionEpoch = 0;
  const rememberedKey = "office-agent.remembered-username";
  const rawApi = () => window.pywebview && window.pywebview.api;
  const wrappers = new WeakMap();
  const authMethods = new Set(["auth_status", "register", "login", "logout", "reset_accounts"]);
  const denied = value => value && typeof value === "object" &&
    [value.reason, value.error, value.code, value.status, value.error?.code].includes("not_authenticated");
  const success = value => Boolean(value && value.ok === true);
  const text = (id, value) => { node(id).textContent = value; };
  function clearPasswords() {
    ["auth-password", "auth-confirm"].forEach(id => { node(id).value = ""; node(id).type = "password"; });
    ["auth-show-password", "auth-show-confirm"].forEach(id => {
      node(id).setAttribute("aria-pressed", "false");
      node(id).setAttribute("aria-label", id === "auth-show-password" ? "显示密码" : "显示确认密码");
      node(id).textContent = "显示";
    });
  }
  function gate(value) {
    authenticated = value;
    const workspace = node("workspace");
    workspace.hidden = !value; workspace.inert = !value;
    workspace.setAttribute("aria-hidden", String(!value));
    node("auth-screen").hidden = value;
  }
  function setBusy(value) {
    submitting = value;
    node("auth-form").setAttribute("aria-busy", String(value));
    node("auth-form").querySelectorAll("input, button").forEach(input => { input.disabled = value || !ready; });
    ["auth-switch", "auth-reset", "auth-reset-confirm"].forEach(id => { node(id).disabled = value || !ready; });
    text("auth-submit", value ? "正在处理…" : mode === "login" ? "登录" : "注册");
  }
  function rememberUsername() {
    try {
      if (node("auth-remember").checked) localStorage.setItem(rememberedKey, node("auth-username").value.trim());
      else localStorage.removeItem(rememberedKey);
    } catch (_) { /* Remembering the username is optional. */ }
  }
  function validate(show = true) {
    const username = node("auth-username").value.trim(), password = node("auth-password").value;
    const errors = {
      username: [...username].length >= 4 && [...username].length <= 32 && !/[\/\\:\p{C}]/u.test(username) ? "" : "! 用户名需为 4–32 个字符，不能含路径符号或控制字符。",
      password: mode === "register"
        ? (password.length >= 8 && password.length <= 16 && /[A-Za-z]/.test(password) && /[0-9]/.test(password) ? "" : "! 密码需为 8–16 位，且包含英文字母和数字。")
        : (password.length > 0 && Array.from(password).length <= 1024 ? "" : "! 请输入账户密码。"),
      confirm: mode !== "register" || (password && node("auth-confirm").value === password) ? "" : "! 两次输入的密码需一致。"
    };
    Object.entries(errors).forEach(([field, error]) => {
      if (show) { text("auth-" + field + "-error", error); node("auth-" + field).setAttribute("aria-invalid", String(Boolean(error))); }
    });
    const score = Number(password.length >= 8 && password.length <= 16) + Number(/[A-Za-z]/.test(password)) + Number(/[0-9]/.test(password));
    node("auth-strength-meter").value = score;
    text("auth-strength-label", score === 3 ? "已满足密码规则；建议使用不重复的密码" : "8–16 位，包含英文字母和数字");
    return Object.keys(errors).find(field => errors[field]);
  }
  function switchMode(next, message = "") {
    mode = next; clearPasswords();
    node("auth-confirm-field").hidden = next !== "register";
    node("auth-confirm").required = next === "register";
    node("auth-strength").hidden = next !== "register";
    node("auth-password").autocomplete = next === "register" ? "new-password" : "current-password";
    node("auth-password").maxLength = next === "register" ? 16 : 2048;
    node("auth-confirm").maxLength = 16;
    text("auth-title", next === "register" ? "注册本地账户" : "登录本地账户");
    text("auth-switch", next === "register" ? "已有账户？登录" : "没有账户？注册");
    ["username", "password", "confirm"].forEach(field => { text("auth-" + field + "-error", ""); node("auth-" + field).removeAttribute("aria-invalid"); });
    text("auth-status", message); setBusy(false);
  }
  function expire() {
    if (!authenticated) return;
    sessionEpoch++; gate(false); clearPasswords();
    // Remove all business DOM (including archived results and file paths), not just its visible text.
    window.dispatchEvent(new Event("officeauth:unauthenticated"));
    node("workspace").replaceChildren();
    node("modal").replaceChildren(); node("modal").style.display = "none";
    // Reload discards app.js lexical state and reconnects through auth_status, never an old snapshot.
    window.location.reload();
  }
  function wrapApi(raw) {
    if (!raw || typeof raw !== "object") return null;
    if (wrappers.has(raw)) return wrappers.get(raw);
    const wrapped = new Proxy(raw, { get(target, key) {
      const method = target[key];
      if (typeof method !== "function") return method;
      return async (...args) => {
        if (!authMethods.has(key) && !authenticated) return key === "poll" ? [] : {ok: false, error: "not_authenticated"};
        const epoch = sessionEpoch;
        const result = await method.apply(target, args);
        if (denied(result)) { expire(); return key === "poll" ? [] : {ok: false, error: "not_authenticated"}; }
        if (!authMethods.has(key) && (epoch !== sessionEpoch || !authenticated)) return key === "poll" ? [] : {ok: false, error: "not_authenticated"};
        return result;
      };
    }});
    wrappers.set(raw, wrapped); return wrapped;
  }
  function ensure() {
    if (authenticated) return Promise.resolve(true);
    if (statusPromise) return statusPromise;
    const raw = rawApi();
    if (!raw) return Promise.resolve(false);
    statusPromise = (async () => {
      try {
        const result = await raw.auth_status();
        ready = success(result); setBusy(false);
        if (ready && result.authenticated === true) { gate(true); return true; }
        gate(false); text("auth-status", ready ? "请登录或注册本地账户。" : "本地认证服务暂不可用，请重新打开应用。");
      } catch (_) { gate(false); text("auth-status", "本地认证服务暂不可用，请重新打开应用。"); }
      return false;
    })().finally(() => { statusPromise = null; });
    return statusPromise;
  }
  window.OfficeAuth = Object.freeze({ensure, wrapApi, isAuthenticated: () => authenticated});
  // Register before app.js document handlers. Native dialog Escape still cancels safely.
  document.addEventListener("keydown", event => { if (!authenticated && event.key === "Escape") event.stopImmediatePropagation(); }, true);
  document.addEventListener("DOMContentLoaded", () => {
    try { const saved = localStorage.getItem(rememberedKey); if (saved && saved.length >= 4 && saved.length <= 32) { node("auth-username").value = saved; node("auth-remember").checked = true; } } catch (_) { /* Optional storage. */ }
    node("auth-theme").onclick = () => { if (typeof window.cycleTheme === "function") window.cycleTheme(); };
    const updateThemeLabel = () => {
      const names = {dark: "深色", light: "浅色", highcontrast: "高对比"};
      text("auth-theme", "主题：" + (names[document.documentElement.dataset.theme] || "深色") + " · 切换");
    };
    new MutationObserver(updateThemeLabel).observe(document.documentElement, {attributes: true, attributeFilter: ["data-theme"]}); updateThemeLabel();
    node("auth-switch").onclick = () => { if (!submitting) { switchMode(mode === "login" ? "register" : "login"); node("auth-username").focus(); } };
    ["username", "password", "confirm"].forEach(field => node("auth-" + field).addEventListener("input", () => validate(true)));
    [["auth-show-password", "auth-password"], ["auth-show-confirm", "auth-confirm"]].forEach(([button, input]) => {
      node(button).onclick = () => { const visible = node(input).type === "password"; node(input).type = visible ? "text" : "password"; node(button).setAttribute("aria-pressed", String(visible)); text(button, visible ? "隐藏" : "显示"); node(button).setAttribute("aria-label", (visible ? "隐藏" : "显示") + (input === "auth-confirm" ? "确认密码" : "密码")); };
    });
    node("auth-form").addEventListener("submit", async event => {
      event.preventDefault(); if (submitting || !ready) return;
      const invalid = validate(); if (invalid) { node("auth-" + invalid).focus(); return; }
      const action = mode; rememberUsername(); setBusy(true); text("auth-status", "正在处理…");
      try {
        const result = await rawApi()[action](node("auth-username").value.trim(), node("auth-password").value);
        clearPasswords();
        if (!success(result)) {
          const code = result?.error || result?.reason || result?.code;
          if (code === "locked") {
            const seconds = Number(result.retry_after);
            const wait = Number.isFinite(seconds) && seconds > 0 ? Math.ceil(seconds / 60) + " 分钟" : "5 分钟";
            text("auth-status", "尝试次数过多，暂时锁定。请约 " + wait + "后重试。锁定仅在当前进程有效，重启后清除。");
          } else if (code === "storage_unavailable") {
            text("auth-status", "本地账户存储不可用或已损坏。请检查本机权限；如需重建，可独立确认重置本地账户。文档与历史不会因此删除。");
          } else text("auth-status", action === "login" ? "用户名或密码不正确，或暂时无法登录。请稍后再试。" : "无法注册，请检查输入或稍后再试。");
          return;
        }
        if (action === "register") { switchMode("login", "注册成功，请使用刚设置的密码登录。"); node("auth-password").focus(); }
        else { gate(true); window.dispatchEvent(new Event("officeauth:authenticated")); }
      } catch (_) { clearPasswords(); text("auth-status", "本地认证服务暂不可用，请稍后再试。"); }
      finally { setBusy(false); }
    });
    const dialog = node("auth-reset-dialog");
    node("auth-reset").onclick = () => { if (!submitting && ready) { dialog.showModal(); node("auth-reset-cancel").focus(); } };
    node("auth-reset-cancel").onclick = () => dialog.close();
    node("auth-reset-confirm").onclick = async () => {
      if (submitting || !ready) return; setBusy(true);
      try {
        const result = await rawApi().reset_accounts(true);
        clearPasswords(); dialog.close();
        if (success(result)) { node("auth-username").value = ""; node("auth-remember").checked = false; rememberUsername(); switchMode("register", "本地账户已清除，文档和工作区记录保留。请重新注册。"); }
        else text("auth-status", "无法重置本地账户，请稍后再试。");
      } catch (_) { dialog.close(); text("auth-status", "无法重置本地账户，请稍后再试。"); }
      finally { setBusy(false); }
    };
    node("auth-logout").onclick = async () => {
      if (submitting || !authenticated) return; setBusy(true); node("auth-logout").disabled = true;
      try {
        const result = await rawApi().logout();
        if (success(result)) expire();
        else text("auth-workspace-status", [result?.reason, result?.error, result?.code].some(code => ["busy", "task_busy"].includes(code)) ? "任务正在处理，暂时不能退出。请等待任务结束后再试。" : "暂时无法退出登录，请稍后再试。");
      } catch (_) { text("auth-workspace-status", "暂时无法退出登录，请稍后再试。"); }
      finally { setBusy(false); if (node("auth-logout")) node("auth-logout").disabled = false; }
    };
    // Escape must not reach app-level focus/close handlers while the gate is visible.
    document.addEventListener("keydown", event => { if (!authenticated && event.key === "Escape" && !dialog.open) event.stopImmediatePropagation(); }, true);
    ensure();
  });
  window.addEventListener("pywebviewready", ensure);
})();
