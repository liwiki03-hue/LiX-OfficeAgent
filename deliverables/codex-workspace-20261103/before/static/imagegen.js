/* 生图改为对话驱动：没有独立表单卡片，结果内联在对话轮次里。
   生成 / 保存 / 撤销仍是三个互不代用的单独确认：每个按钮都挂在自己的消息上，
   只有用户明确点击才会提交；新建任务后旧消息的按钮立即失效。
   本模块只渲染 bridge 已校验的有界事件，不发起网络请求，也不持有跨会话授权。 */
"use strict";
(() => {
  const MAX_TEXT = 4000;
  let taskEpoch = 0;
  const openConfirmations = new Set();
  const cardScopes = new WeakMap();
  const text = value => typeof value === "string" ? value.slice(0, MAX_TEXT) : "";
  const tag = (name, className) => {
    const node = document.createElement(name);
    if (className) node.className = className;
    return node;
  };
  function line(parent, label, value) {
    const row = tag("p", "imagegen-line");
    const strong = tag("strong");
    strong.textContent = label + "：";
    row.append(strong, document.createTextNode(text(value)));
    parent.append(row);
    return row;
  }
  function actions(parent) {
    const box = tag("div", "actions imagegen-actions");
    parent.append(box);
    return box;
  }
  function add(parent, label, className, onClick) {
    const button = tag("button", "btn " + className);
    button.type = "button";
    button.textContent = label;
    button.onclick = onClick;
    parent.append(button);
    return button;
  }
  function showResult(card, reply) {
    const status = tag("p", "imagegen-status");
    const ok = Boolean(reply && reply.ok);
    status.className = "imagegen-status " + (ok ? "ok" : "warn");
    status.setAttribute("role", "status");
    if (ok) status.textContent = reply.status === "cancelled" ? "已取消，没有执行。" : "完成了。";
    else if (reply && reply.reason === "cancelled") status.textContent = "已取消选择目录，未写入图片；预览保留，可以再次选择保存位置。";
    else if (reply && reply.reason === "confirmation_required") status.textContent = "没有收到明确的同意，操作未执行。";
    else if (reply && reply.reason === "not_authenticated") status.textContent = "还没登录，操作被拒绝。";
    else if (reply && reply.reason === "no_session") status.textContent = "当前没有可用的生图任务；请重新说一次要什么图。";
    else if (reply && reply.reason === "stale") status.textContent = "任务已经变了，这个迟到的结果被丢弃。" +
      (reply.provider_request_state === "attempt_committed" ? " 生图请求可能计费，请先核对账单，不要直接重复确认。" : "请重新提交需求并核对新计划。");
    else if (reply && ["image_configuration_missing", "image_configuration_invalid"].includes(reply.code))
      status.textContent = "生图服务配置不完整或无效。" +
        (reply.provider_request_state === "attempt_committed" ? " 提供方尝试可能已发出并计费，请先核对账单，不要直接重复确认。" : "本次没有请求生图服务。") +
        "请检查项目配置文件中的提供方、密钥、地址和模型，重启后重新规划；不要在任务对话中粘贴密钥。";
    else if (reply && ["image_configuration_changed", "image_source_changed", "image_approval_expired"].includes(reply.code))
      status.textContent = "文件、配置或确认时限已改变；旧计划不能继续。" + (reply.provider_request_state === "attempt_committed" ? " 提供方尝试可能已发出并计费，请先核对账单。" : "") + "请重新提交需求并核对新计划。";
    else if (reply && reply.code === "image_generation_failed" && reply.failure_stage === "preflight" && reply.provider_request_state === "not_started")
      status.textContent = "生图在调用提供方前的本地准备阶段失败；此次未调用生图服务。请检查审计目录与文件权限，重新提交需求后核对新计划。";
    else if (reply && ["image_generation_failed", "image_failed"].includes(reply.code)) {
      const details = {
        image_response_timeout: "等待服务响应超时",
        image_response_encoding: "服务响应编码不可用",
        image_response_too_large: "服务响应超过大小限制",
        image_response_invalid: "服务响应不符合图片格式",
        image_decoder_unavailable: "图片解码不可用",
        image_raster_too_large: "返回图片超过尺寸限制",
        image_raster_invalid: "返回图片未通过安全校验",
        image_provider_failed: "提供方请求失败",
        image_authentication_failed: "鉴权失败，请修复当前接口密钥",
        image_model_not_found: "模型或接口路径不存在",
        image_rate_limited: "接口限流",
        image_service_failed: "服务异常",
        image_address_unreachable: "地址不可达"
      };
      const detail = Object.hasOwn(details, reply.provider_error) ? details[reply.provider_error] : Object.hasOwn(details, reply.failure_detail) ? details[reply.failure_detail] : "没有返回可用结果";
      status.textContent = "生图失败（" + detail + "）；提供方请求可能计费。请先核对服务账单与处理记录，不要直接重复确认。";
    }
    else if (reply && reply.code) status.textContent = "操作被拒绝（" + text(reply.code) + "），请核对当前状态与处理记录后再决定下一步。" +
      (["attempt_committed", "attempt_unknown"].includes(reply.provider_request_state) ? " 生图请求可能计费，请先核对账单，不要直接重复确认。" : "");
    else status.textContent = reply && reply.provider_request_state === "attempt_unknown" ?
      "生图操作未收到可靠结果；请求可能计费，请先核对账单，不要直接重复确认。" : "操作没有完成。";
    card.append(status);
    return status;
  }
  function confirmDialog(options) {
    return new Promise(resolve => {
      const dialog = tag("dialog", "confirm-dialog imagegen-confirm");
      dialog.setAttribute("aria-label", options.title);
      const title = tag("h2"); title.textContent = options.title;
      const list = tag("ul");
      options.lines.forEach(item => { const li = tag("li"); li.textContent = text(item); list.append(li); });
      const hint = tag("p", "confirm-hint");
      hint.textContent = "不会自动执行；只有你点下面这个按钮才会提交。";
      const row = tag("div", "actions");
      let settled = false;
      const finish = value => {
        if (settled) return;
        settled = true;
        openConfirmations.delete(finish);
        dialog.close();
        dialog.remove();
        resolve(value);
      };
      openConfirmations.add(finish);
      const cancel = add(row, "取消", "secondary", () => finish(false));
      add(row, options.confirmLabel, options.danger ? "danger" : "confirm", () => finish(true));
      dialog.append(title, list, hint, row);
      dialog.addEventListener("keydown", event => {
        event.stopPropagation(); // 不要在对话框下提交任务或切换历史
        if (event.key === "Tab") {
          const buttons = [...row.querySelectorAll("button")].filter(button => !button.disabled);
          const first = buttons[0], last = buttons[buttons.length - 1];
          if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
          else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
        }
        if (event.key === "Escape") { event.preventDefault(); finish(false); }
      });
      dialog.addEventListener("cancel", event => { event.preventDefault(); finish(false); });
      dialog.addEventListener("close", () => dialog.remove(), {once: true});
      document.body.append(dialog);
      dialog.showModal();
      cancel.focus(); // 默认焦点在取消上，避免回车误确认
    });
  }
  function prompt(call, card, operation, id, options) {
    const box = actions(card), born = taskEpoch;
    const submit = async confirmed => {
      if (born !== taskEpoch) return; // A newer task or logout invalidated this plan.
      box.querySelectorAll("button").forEach(button => { button.disabled = true; });
      let answer;
      try {
        answer = await call(operation, id, confirmed);
      } catch (_) {
        answer = {ok: false, error: "call_failed"};
        if (operation === "confirm_image_generation" && confirmed)
          answer.provider_request_state = "attempt_unknown";
      }
      if (answer && (answer.reason === "busy" || answer.code === "image_operation_busy")) {
        box.querySelectorAll("button").forEach(button => { button.disabled = false; });
        const hint = tag("p", "imagegen-status warn");
        hint.textContent = "另一张图正在处理，请等它完成后再逐张确认；本次没有提交请求。";
        card.append(hint);
        return answer;
      }
      box.remove();
      showResult(card, answer);
      if (born === taskEpoch && answer?.ok && !confirmed) {
        const kind = operation === 'confirm_image_save' ? 'image_result' : operation === 'confirm_image_rollback' ? 'image_saved' : null;
        if (kind) document.querySelectorAll('.imagegen-item').forEach(node => {
          const scope = cardScopes.get(node), target = cardScopes.get(card);
          if (scope?.kind !== kind || scope.epoch !== born || !target) return;
          if (kind === 'image_result' && scope.planId !== target.planId) return;
          if (kind === 'image_saved' && scope.saveId !== target.saveId) return;
          node.querySelectorAll('.imagegen-actions button').forEach(button => {button.disabled = false;});
        });
      }
      return answer;
    };
    add(box, options.cancelLabel, "secondary", () => submit(false));
    add(box, options.confirmLabel, options.danger ? "danger" : "confirm", async () => {
      if (born !== taskEpoch) return;
      if (await confirmDialog(options)) await submit(true);
      else if (born === taskEpoch) box.querySelectorAll("button").forEach(button => { button.disabled = false; });
    });
  }
  async function followUp(call, card, box, operation, id, born) {
    if (born !== taskEpoch) return;
    box.querySelectorAll("button").forEach(button => { button.disabled = true; });
    let answer;
    try {
      answer = await call(operation, id);
    } catch (_) {
      answer = {ok: false, error: "call_failed"};
    }
    if (!answer || !answer.ok) {
      if (born === taskEpoch) box.querySelectorAll('button').forEach(button => {button.disabled = false;});
      showResult(card, answer);
    }
  }
  function title(card, label) {
    const head = tag("div", "imagegen-title");
    head.textContent = label;
    card.append(head);
  }
  function renderPlan(call, card, event) {
    const plan = event.plan && typeof event.plan === "object" ? event.plan : {};
    const index = Number(plan.batch_index), total = Number(plan.batch_total);
    const position = Number.isInteger(index) && Number.isInteger(total) && total > 1 &&
      index >= 1 && index <= total && total <= 8 ? " " + index + "/" + total : "";
    title(card, "生图方案" + position + " · 还没调用、还没存文件");
    line(card, "发给模型的提示词（实际发送）", plan.effective_prompt);
    line(card, "尺寸 / 风格 / 张数", [plan.size, plan.style, plan.n].join(" · "));
    const configuration = plan.configuration && typeof plan.configuration === "object" ? plan.configuration : {};
    if (plan.route) line(card, "已选择的接口", plan.route.connection_name + " · " + plan.route.model_id);
    line(card, "使用的服务", [configuration.provider, configuration.model, configuration.base_url].filter(Boolean).join(" / "));
    const ttl = Number(plan.approval_ttl_seconds);
    line(card, "费用", "未知，且发出后无法收回；" + (Number.isInteger(ttl) && ttl > 0
      ? "计划创建后 " + ttl + " 秒内有效" : "确认时限从计划创建时开始") + "，过期请重新规划");
    prompt(call, card, "confirm_image_generation", text(plan.plan_id), {
      title: "确认联网生成图片？",
      lines: ["仅为本张图按上面的提示词和配置发一次请求；其他图仍需单独确认。",
              "费用未知且不可退回：请求一旦发出，就算结果不用也可能计费。",
              "这次只同意生成；要把图存成新文件还要再确认一次。"],
      confirmLabel: "确认生成（可能产生费用）", cancelLabel: "取消，不要这张图", danger: true,
    });
  }
  function renderResult(call, card, event) {
    title(card, "生图结果 · 还没存文件");
    line(card, "结果", event.text || "图片已生成");
    const host = tag("div", "imagegen-images");
    card.append(host);
    const planId = text(event.plan_id);
    window.OfficeImages?.appendImages(host, event.images, {transactionalSave: Boolean(planId)});
    if (!planId) return;
    const box = actions(card), born = taskEpoch;
    add(box, "保存成新的图片文件…", "secondary", () => followUp(call, card, box, "plan_image_save", planId, born));
  }
  function renderSavePlan(call, card, event) {
    const plan = event.plan && typeof event.plan === "object" ? event.plan : {};
    title(card, "保存图片 · 等你确认");
    line(card, "保存到", plan.path);
    line(card, "大小", text(String(plan.bytes ?? "")) + " 字节");
    line(card, "说明", "只新建文件，不会覆盖原文档，也不会覆盖已经存在的同名文件");
    prompt(call, card, "confirm_image_save", text(plan.save_id), {
      title: "确认另存为一个新图片文件？",
      lines: [text(plan.path), "只会新增文件；原文档和本次已生成的其他内容都不受影响。",
              "保存本身不会退回已经产生的云端费用。"],
      confirmLabel: "确认保存新文件", cancelLabel: "不保存", danger: false,
    });
  }
  function renderSaved(call, card, event) {
    title(card, "图片已保存");
    line(card, "位置", event.text || "新文件已写入");
    const saveId = text(event.save_id);
    if (!saveId) return;
    const box = actions(card), born = taskEpoch;
    add(box, "打开保存位置", "secondary", async () => {
      if (born !== taskEpoch) return;
      try { const reply = await call('open_saved_image_folder', saveId); if (born === taskEpoch) showResult(card, reply); }
      catch (_) { if (born === taskEpoch) showResult(card, {ok: false}); }
    });
    add(box, "撤销这个文件…", "secondary", () => followUp(call, card, box, "plan_image_rollback", saveId, born));
  }
  function renderRollbackPlan(call, card, event) {
    const plan = event.plan && typeof event.plan === "object" ? event.plan : {};
    title(card, "撤销已保存的图片 · 等你确认");
    line(card, "要删除", plan.path);
    line(card, "说明", "只有在文件仍是本次生成、内容没被改动过时才删除；改动过就保留。不退费用，也不恢复历史。");
    prompt(call, card, "confirm_image_rollback", text(plan.rollback_id), {
      title: "确认删除这个文件？",
      lines: [text(plan.path), "只在它仍是本次生成、且内容没变的情况下删除。",
              "云端费用不会因此退回，历史也不会恢复。"],
      confirmLabel: "确认删除这个文件", cancelLabel: "不删了", danger: true,
    });
  }
  const KINDS = new Set(["image_plan", "image_result", "image_save_plan", "image_saved",
                         "image_rollback_plan", "image_rolled_back"]);
  /* 结果不再堆在侧栏工具面板里，而是内联到对话流当前轮次。
     app.js 提供 mount，把消息节点接到对话流上；这样导出按钮、复制按钮都跟着消息走。 */
  function handleEvent(event, call, mount) {
    if (!event || typeof event !== "object" || typeof call !== "function") return;
    if (!KINDS.has(event.type)) return; // 未知事件类型不创建空卡片
    if (typeof mount !== "function") return;
    const card = tag("article", "imagegen-item");
    card.setAttribute("data-kind", text(event.type));
    cardScopes.set(card, {epoch:taskEpoch, kind:event.type, planId:text(event.plan_id || event.plan?.plan_id), saveId:text(event.save_id || event.plan?.save_id)});
    const label = { image_plan: "生图方案", image_result: "生图结果", image_save_plan: "保存图片",
                    image_saved: "图片已保存", image_rollback_plan: "撤销图片", image_rolled_back: "图片已撤销" }[event.type];
    card.setAttribute("aria-label", label);
    mount(card);
    if (event.type === "image_plan") renderPlan(call, card, event);
    else if (event.type === "image_result") renderResult(call, card, event);
    else if (event.type === "image_save_plan") renderSavePlan(call, card, event);
    else if (event.type === "image_saved") renderSaved(call, card, event);
    else if (event.type === "image_rollback_plan") renderRollbackPlan(call, card, event);
    else if (event.type === "image_rolled_back") {
      title(card, "图片已撤销");
      line(card, "结果", event.text || "文件已删除");
      line(card, "费用", "云端费用没有退回；历史也不会恢复");
    }
  }
  function retirePending(reason = "任务已改变；旧操作不能继续，请重新规划。") {
    ++taskEpoch;
    // An already open second-confirmation dialog must not outlive its task.
    [...openConfirmations].forEach(cancel => cancel(false));
    document.querySelectorAll(".imagegen-item .imagegen-actions button").forEach(button => {
      button.disabled = true;
      button.title = reason;
    });
  }
  window.addEventListener("officeauth:unauthenticated", () => {
    retirePending("已退出登录；这个确认已经失效，请重新登录并重新生成。");
  });
  window.OfficeImageGen = Object.freeze({handleEvent, retirePending});
})();
