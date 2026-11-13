/* Host-scoped preference settings. Browser supplies no account, recurrence or content authority. */
"use strict";
(() => {
  const byId = id => document.getElementById(id);
  const load = byId("memory-load"), toggle = byId("memory-toggle"), save = byId("memory-save");
  const remove = byId("memory-delete"), clear = byId("memory-clear");
  const manualKind = byId("memory-manual-kind"), manualText = byId("memory-manual-text");
  const manualSave = byId("memory-manual-save");
  const list = byId("memory-items"), status = byId("memory-status");
  if (!load || !toggle || !save || !remove || !clear || !list || !status ||
      !manualKind || !manualText || !manualSave) return;
  const labels = {persona: "人设设定", work: "工作背景", personal: "个人背景",
    writing: "长期写作约定", skill: "工作习惯与步骤", format: "格式偏好", role: "角色偏好"};
  let state = null, selected = null, busy = false, epoch = 0;
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const bridge = () => authenticated() ? window.OfficeAuth.wrapApi(window.pywebview?.api) : null;
  const id = value => typeof value === "string" && /^[0-9a-f]{32}$/.test(value);
  function buttons() {
    load.disabled = busy || !authenticated();
    toggle.disabled = busy || !authenticated() || !state;
    save.disabled = busy || !authenticated() || !state?.enabled || !state?.suggestion;
    manualKind.disabled = manualText.disabled = busy || !authenticated() || !state?.enabled;
    manualSave.disabled = manualText.disabled || !manualText.value.trim();
    remove.disabled = busy || !authenticated() || !selected;
    clear.disabled = busy || !authenticated() || !state?.items?.length;
    [load, toggle, save, manualSave, remove, clear].forEach(button => {
      button.title = !button.disabled ? '' : busy ? '正在处理，请等待' : !authenticated() ? '请先登录本机账户' :
        !state ? '请先查看本账户记忆' : button === save ? '需开启记忆并存在待确认候选' :
        button === manualSave ? '需开启记忆并填写内容' : '需存在或选中本账户记忆';
    });
  }
  function reset() {
    ++epoch; state = selected = null; list.replaceChildren();
    manualKind.value = "persona"; manualText.value = "";
    status.textContent = "记忆默认关闭；请先查看。"; buttons();
  }
  function populateManual() {
    const item = state?.items?.find(entry => entry.source === "manual" && entry.kind === manualKind.value);
    manualText.value = item?.text || ""; buttons();
  }
  async function refresh() {
    const service = bridge();
    if (busy || !service?.memory_snapshot) return;
    const current = epoch; busy = true; status.textContent = '正在读取本账户记忆…'; buttons();
    try {
      const response = await service.memory_snapshot();
      if (current !== epoch || !authenticated()) return;
      if (response?.ok !== true || !Array.isArray(response.result?.items)) {
        state = null; status.textContent = "记忆不可用（" + String(response?.reason || "memory_unavailable") + "）。"; return;
      }
      state = response.result; selected = null; list.replaceChildren();
      for (const item of state.items) {
        if (!id(item?.id) || typeof item.text !== "string") continue;
        const label = document.createElement("label"), radio = document.createElement("input");
        radio.type = "radio"; radio.name = "memory-selection";
        radio.addEventListener("change", () => {
          selected = item.id;
          if (item.source === "manual" && ["persona", "work", "personal", "writing", "skill"].includes(item.kind)) {
            manualKind.value = item.kind; populateManual();
          }
          buttons();
        });
        const lifespan = item.source === "manual" ? "手动长期／删除前有效" : "权重 " + String(item.weight);
        label.append(radio, document.createTextNode(" " + (labels[item.kind] || "记忆") + "：" + item.text + "（" + lifespan + "）"));
        list.append(label, document.createElement("br"));
      }
      populateManual();
      toggle.textContent = state.enabled ? "关闭记忆…" : "开启记忆…";
      status.textContent = (state.enabled ? "已启用" : "已关闭") + "；已保存 " + state.items.length + " 条。" +
        (state.suggestion ? "待确认的重复候选：" + state.suggestion.text : "");
    } catch (_) {
      if (current === epoch) { state = null; status.textContent = "记忆服务无响应，未自动重试。"; }
    } finally { busy = false; buttons(); }
  }
  async function change(action, value, explanation) {
    const service = bridge();
    if (busy || !state || !service?.plan_memory_change || !service?.confirm_memory_change) return;
    const current = epoch; busy = true; buttons(); let submitted = false;
    try {
      const response = await service.plan_memory_change(action, value);
      if (current !== epoch || !authenticated()) return;
      const planId = response?.plan?.plan_id;
      if (response?.ok !== true || !id(planId) || response.plan.action !== action ||
          (action === "set_manual" && (response.plan.value?.kind !== value.kind || response.plan.value?.text !== value.text))) {
        status.textContent = "方案不可用（" + String(response?.reason || "memory_plan_invalid") + "）。"; return;
      }
      const approved = window.confirm(explanation);
      if (current !== epoch || !authenticated()) return;
      submitted = approved === true;
      const result = await service.confirm_memory_change(planId, approved === true);
      if (current !== epoch || !authenticated()) return;
      if (!approved && result?.reason === "memory_confirmation_required") {
        status.textContent = "已取消，没有保存更改。"; return;
      }
      if (result?.ok !== true || result?.result?.saved !== true) {
        state = null; status.textContent = "结果无法确认（" + String(result?.reason || "memory_uncertain") + "），请核查设置，不会自动重试。"; return;
      }
      status.textContent = "已提交并复开验证。";
      state = null;
    } catch (_) {
      state = null;
      if (current === epoch) status.textContent = submitted ? "提交后结果未知；请核查，不会自动重试。" : "未提交：方案服务无响应。";
    } finally { busy = false; buttons(); }
    await refresh();
  }
  load.addEventListener("click", refresh);
  toggle.addEventListener("click", () => change("enable", !state.enabled,
    state.enabled ? "确认关闭记忆？历史项仍保留但不再引用；如需删除，请另行操作。" :
      "确认开启本账户记忆？只从直接输入提议短句，逐条确认才持久化；启用后引用的内容可能随任务传给模型服务。"));
  save.addEventListener("click", () => { if (state?.suggestion?.repeated) change("save", state.suggestion.text,
    "确认保存这条重复提出的候选？\n" + state.suggestion.text); });
  manualKind.addEventListener("change", populateManual);
  manualText.addEventListener("input", buttons);
  manualSave.addEventListener("click", () => {
    if (!state?.enabled || manualSave.disabled) return;
    const kind = manualKind.value, text = manualText.value.trim();
    if (!["persona", "work", "personal", "writing", "skill"].includes(kind) || text.length > 300 || !text) return;
    change("set_manual", {kind, text}, "确认保存" + labels[kind] + "？已有同类内容将被替换；此项不会自动衰减。\n\n" + text);
  });
  remove.addEventListener("click", () => { if (selected) change("delete", selected, "确认删除选中的一条记忆？"); });
  clear.addEventListener("click", () => change("clear", null, "确认删除本账户全部记忆？文档和其他账户不变。"));
  ["officeauth:authenticated", "officeauth:unauthenticated", "pagehide"].forEach(name => window.addEventListener(name, reset));
  reset();
})();
