/* Human-only approval controls; no SVG payload, provider calls or model history.
   按钮不再固定在卡片里，而是挂到「最近一次成功渲染」的图表消息内。 */
"use strict";
(() => {
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  let revision = 0, metadata = null, saved = null, inflight = false;
  let save = null, savePng = null, undo = null, status = null;
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const say = text => { if (status) status.textContent = text; };
  const hash = value => typeof value === "string" && /^[0-9a-f]{64}$/.test(value);
  const id = value => typeof value === "string" && /^[0-9a-f]{32}$/.test(value);
  function buttons() {
    if (save) save.disabled = inflight || !authenticated() || !metadata || Boolean(saved);
    if (savePng) savePng.disabled = inflight || !authenticated() || !metadata || Boolean(saved);
    if (undo) undo.disabled = inflight || !authenticated() || !saved;
  }
  function detach() {
    // 旧消息里的按钮失去权限：新建图表时旧图不再可保存。
    // 必须显式禁用其 DOM，否则用户仍能点到它。
    if (save) save.disabled = true;
    if (savePng) savePng.disabled = true;
    if (undo) undo.disabled = true;
    save = savePng = undo = status = null;
  }
  function retire() {
    ++revision; metadata = saved = null; detach(); say("");
  }
  window.OfficeDiagramExport = {
    retire,
    // 在图表消息内挂载保存/回滚按钮；同一时刻只有最新一张图可导出。
    mount(host) {
      detach();
      if (!host) return;
      status = document.createElement("p");
      status.className = "diagram-export-status";
      status.setAttribute("role", "status");
      save = document.createElement("button");
      save.type = "button"; save.className = "btn secondary";
      save.textContent = "保存这张图";
      savePng = document.createElement("button");
      savePng.type = "button"; savePng.className = "btn secondary";
      savePng.textContent = "保存 PNG（宿主栅格化）";
      undo = document.createElement("button");
      undo.type = "button"; undo.className = "btn secondary";
      undo.textContent = "撤销已保存的图";
      save.addEventListener("click", () => operate(false, false));
      savePng.addEventListener("click", () => operate(false, true));
      undo.addEventListener("click", () => operate(true, false));
      host.append(save, undo, status, savePng);
      buttons();
    },
    ready(value) {
      // 新预览到达：清空旧的保存/回滚身份，保留按钮挂载点。
      const keepSave = save, keepPng = savePng, keepUndo = undo, keepStatus = status;
      retire();
      save = keepSave; savePng = keepPng; undo = keepUndo; status = keepStatus;
      if (hash(value?.source_sha256) && hash(value?.svg_sha256)) {
        metadata = {source_sha256: value.source_sha256, svg_sha256: value.svg_sha256};
        say("可以把这张图存成新文件；原文档不受影响。请先确认图的内容对不对。");
      }
      buttons();
    }
  };
  ["officeauth:authenticated", "officeauth:unauthenticated", "officediagram:invalidate", "pagehide"].forEach(name => window.addEventListener(name, retire));
  function current(token) { return token === revision && authenticated(); }
  function failure(response) {
    const code = typeof response?.reason === "string" && /^[a-z_]{1,50}$/.test(response.reason) ? response.reason : "export_failed";
    say("没能完成（" + code + "）。不会自动重试；如果已经提交过，请核对文件和记录。");
  }
  async function operate(rollback, png) {
    if (inflight || !authenticated() || !api || (rollback ? !saved : !metadata || Boolean(saved))) return;
    const token = revision, expected = metadata, previous = saved;
    let submitted = false;
    inflight = true; buttons();
    try {
      const response = rollback ? await api.plan_diagram_rollback(previous.save_id) :
        png ? await api.plan_diagram_png_save(expected.source_sha256, expected.svg_sha256) :
          await api.plan_diagram_save(expected.source_sha256, expected.svg_sha256);
      if (!current(token)) return;
      if (response?.ok !== true) { failure(response); return; }
      const plan = response.plan, planId = rollback ? plan?.rollback_id : plan?.plan_id;
      if (!id(planId) || (!rollback && (typeof plan.selected_document !== "string" || plan.selected_document.length > 32768 ||
          plan.destination_basename !== "office-diagram-" + planId + (png ? ".png" : ".svg") ||
          (png && (!hash(plan.png_sha256) || plan.renderer !== "host-qt-svg-v1"))))) {
        say("保存方案格式不对，没有提交。"); return;
      }
      const approved = await ask(rollback ?
        "确认撤销？只会删除没被改动过的这张图：\n" + previous.path + "\n记录保留，历史不会恢复。" :
        "确认保存？请先自己看一眼图对不对。\n所在文档：" + plan.selected_document +
        "\n在同一目录新建：" + plan.destination_basename + "\n原文档不会被覆盖；提交后无法中途撤销。");
      if (!current(token)) return;
      submitted = approved === true;
      const result = rollback ? await api.confirm_diagram_rollback(planId, approved === true) :
        await api.confirm_diagram_save(planId, approved === true);
      if (!current(token)) return;
      if (!approved && result?.reason === "export_confirmation_required") {
        say(rollback ? "已取消，文件没有删除。" : "已取消，没有新建文件。"); return;
      }
      if (result?.ok !== true) { failure(result); return; }
      const outcome = result.result;
      if (rollback) {
        if (outcome?.removed === true && outcome.audit_updated === true && !outcome.code) {
          saved = null; say("已删除；记录已更新并保留。");
        } else {
          metadata = saved = null;
          say("撤销只完成了一部分，或结果没法确认（删除=" + String(outcome?.removed) + "，记录更新=" + String(outcome?.audit_updated) + "）。保存功能已关闭，请核对文件和记录。");
        }
      } else if (outcome?.saved === true && id(outcome.save_id) && typeof outcome.path === "string") {
        saved = {save_id: outcome.save_id, path: outcome.path};
        say("已保存并重新打开验证：" + outcome.path + "；原文档不变，已记录。要撤销需再确认一次。");
      } else {
        metadata = saved = null; say("保存结果没法确认，请核对文件和记录；不会自动重试。");
      }
    } catch (_) {
      if (current(token)) say(submitted ? "已提交但结果没法确认，请核对文件和记录；不会自动重试。" : "方案服务没有响应，没有提交。");
    } finally {
      inflight = false; buttons();
    }
  }
  retire();
})();
