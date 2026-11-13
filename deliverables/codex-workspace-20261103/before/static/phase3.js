"use strict";
(() => {
  const get = id => document.getElementById(id);
  const controls = ["general-load","general-save","general-delete","general-clear","multi-select","multi-clear","multi-load"].map(get);
  if (controls.some(control => !control)) return;
  let epoch = 0, busy = false, selected = null, items = [];
  const auth = () => window.OfficeAuth?.isAuthenticated() === true;
  const api = () => window.pywebview?.api;
  const id = value => typeof value === "string" && /^[a-f0-9]{32}$/.test(value);
  function buttons() { controls.forEach(control => { control.disabled = busy || !auth(); }); }
  function reset() {
    epoch++; busy = false; selected = null; items = [];
    ["general-items","multi-plans","general-status","multi-status"].forEach(name => get(name).textContent = "");
    get("general-content").value = ""; get("multi-names").value = ""; buttons();
  }
  async function run(status, operation) {
    if (busy || !auth()) return;
    const current = epoch; busy = true; buttons();
    const valid = () => current === epoch && auth();
    try { await operation(api(), valid); }
    catch (_) { if (valid()) get(status).textContent = "结果无法确认，请核查；不会自动重试。"; }
    finally { if (current === epoch) { busy = false; buttons(); } }
  }
  async function loadGeneral(service, valid) {
    const reply = await service.general_memory_snapshot();
    if (!valid()) return;
    if (!reply?.ok || !Array.isArray(reply.result?.items) || reply.result.items.length > 20) {
      get("general-status").textContent = "通用记忆不可用。"; return;
    }
    items = reply.result.items; selected = null; get("general-items").textContent = "";
    for (const item of items) {
      if (!id(item.memory_id) || typeof item.content !== "string" || item.content.length > 1000) continue;
      const label = document.createElement("label"), radio = document.createElement("input");
      radio.type = "radio"; radio.name = "general-memory-item";
      radio.addEventListener("change", () => { selected = item.memory_id; });
      label.append(radio, document.createTextNode(" " + item.category + "：" + item.content));
      get("general-items").append(label, document.createElement("br"));
    }
    get("general-status").textContent = "已保存 " + items.length + " 条；召回使用本地词项哈希，不代表语义模型理解。";
  }
  async function changeGeneral(service, valid, operation, value) {
    const reply = await service.plan_general_memory(operation, value);
    if (!valid()) return;
    const plan = reply?.result;
    if (!reply?.ok || !id(plan?.plan_id) || plan.operation !== operation || plan.requires_confirmation !== true ||
        (operation === "remember" && (plan.content !== value.content || plan.category !== value.category))) {
      get("general-status").textContent = "未产生可确认计划，内容可能不符合隐私或容量限制。"; return;
    }
    const text = operation === "remember" ? "确认保存并允许跨任务召回以下内容？召回可能发送给当前模型。\n\n" + plan.content :
      "确认停止召回" + (operation === "clear" ? "本账户全部通用记忆" : "所选通用记忆") + "？事务备份仍保留在本机。";
    const confirmed = window.confirm(text);
    if (!valid()) return;
    const outcome = await service.confirm_general_memory(plan.plan_id, confirmed === true);
    if (!valid()) return;
    if (!confirmed) { get("general-status").textContent = "已取消。"; return; }
    if (!outcome?.ok || outcome.result?.applied !== true) {
      get("general-status").textContent = "未确认成功，请核查；不会自动重试。"; return;
    }
    await loadGeneral(service, valid);
  }
  get("general-load").addEventListener("click", () => run("general-status", loadGeneral));
  get("general-save").addEventListener("click", () => {
    const value = {content: get("general-content").value.trim(), category: get("general-category").value};
    if (!value.content || value.content.length > 1000 || !["task","summary","preference"].includes(value.category)) return;
    run("general-status", (service, valid) => changeGeneral(service, valid, "remember", value));
  });
  get("general-delete").addEventListener("click", () => {
    if (selected) { const value = selected; run("general-status", (service, valid) => changeGeneral(service, valid, "delete", value)); }
  });
  get("general-clear").addEventListener("click", () => run("general-status", (service, valid) => changeGeneral(service, valid, "clear", null)));
  async function loadMulti(service, valid) {
    const response = await service.task_documents_snapshot();
    if (!valid()) return;
    get("multi-plans").textContent = "";
    if (!response?.ok || !Array.isArray(response.result?.files) || !Array.isArray(response.result?.plans)) {
      get("multi-status").textContent = "范围不可用，请重新选择。"; return;
    }
    get("multi-status").textContent = response.result.files.length ? "当前范围：" + response.result.files.join("、") : "当前为单文件任务。";
    for (const plan of response.result.plans.slice(0,20)) {
      if (!id(plan.id) || typeof plan.source !== "string") continue;
      const button = document.createElement("button"); button.className = "btn secondary"; button.type = "button";
      button.textContent = "查看单文件计划：" + plan.source.split(/[\\/]/).pop() + " · " + plan.kind;
      button.addEventListener("click", () => run("multi-status", async (bridge, live) => {
        const selected = await bridge.choose_document_plan(plan.id);
        if (live()) get("multi-status").textContent = selected?.ok ? "已打开此文件计划；请在修改卡中单独确认。" : "计划已失效。";
      }));
      get("multi-plans").append(button);
    }
  }
  get("multi-select").addEventListener("click", () => {
    const names = get("multi-names").value.split(/\r?\n/).map(value => value.trim()).filter(Boolean);
    if (names.length < 2 || names.length > 5 || new Set(names.map(name => name.toLowerCase())).size !== names.length) {
      get("multi-status").textContent = "请逐行填写2–5个不同文件名，包含当前文件。"; return;
    }
    run("multi-status", async (service, valid) => {
      const response = await service.plan_task_documents(names);
      if (!valid()) return;
      const plan = response?.result;
      if (!response?.ok || !id(plan?.plan_id) || plan.requires_confirmation !== true || JSON.stringify(plan.files) !== JSON.stringify(names)) {
        get("multi-status").textContent = "范围拒绝：只支持当前目录的安全DOCX/XLSX文件。"; return;
      }
      const confirmed = window.confirm("确认读取以下文件的有界摘要，并用于后续任务？内容可能发送给当前模型；每个写计划仍单独确认。\n\n" + plan.files.join("\n"));
      if (!valid()) return;
      const result = await service.confirm_task_documents(plan.plan_id, confirmed === true);
      if (!valid()) return;
      if (!confirmed) { get("multi-status").textContent = "已取消，原有范围未改变。"; return; }
      if (!result?.ok) { get("multi-status").textContent = "确认后范围已失效，请重新选择。"; return; }
      await loadMulti(service, valid);
    });
  });
  get("multi-load").addEventListener("click", () => run("multi-status", loadMulti));
  get("multi-clear").addEventListener("click", () => run("multi-status", async (service, valid) => {
    const response = await service.clear_task_documents();
    if (valid() && response?.ok) await loadMulti(service, valid);
  }));
  ["officeauth:authenticated","officeauth:unauthenticated","pagehide"].forEach(name => window.addEventListener(name, reset));
  reset();
})();
