/* Office Agent Web UI 前端逻辑：轮询后端事件并渲染，转发用户动作。 */
"use strict";

let api = null; // 认证握手成功之前，业务桥接不可达。
let pending = false;   // 是否有待确认的计划
let verified = false;  // 是否有已验证的输出副本
let busy = false;
let activeClarification = null;
let connected = false;
let hasFile = false;
let requesting = false;
let requestLabel = "";
let modalOpen = false;
let actionFocus = null;
let deferredModalFocus = null;
let fileWarning = "";
let taskWarning = "";
let taskSafety = "";
let planVersion = 0;
let decisionLimit = null;
let stopRequested = false;
let statusFeedback = null;

function el(id) { return document.getElementById(id); }
function setText(node, text) {
  const value = String(text ?? "");
  if (node.textContent !== value) node.textContent = value;
}
function canFocus(node) {
  if (!node || !node.isConnected || node.disabled ||
      node.getAttribute("aria-disabled") === "true" || node.closest("[hidden], [inert]")) return false;
  // Chromium 关闭 details 后仍可能给子控件返回几何尺寸，不能只依赖 getClientRects。
  for (let ancestor = node.parentElement; ancestor; ancestor = ancestor.parentElement) {
    if (ancestor.tagName === "DETAILS" && !ancestor.open &&
        !ancestor.querySelector(":scope > summary")?.contains(node)) return false;
  }
  return Boolean(node.getClientRects().length && getComputedStyle(node).visibility === "visible");
}
function focusVisible(node) {
  if (modalOpen || !canFocus(node)) return;
  node.focus({ preventScroll: true });
  node.scrollIntoView({ block: "nearest", inline: "nearest" });
}
function focusFallback(source) {
  if (currentView === "history") return el("btn-history-current");
  if (source && source.closest("#plan-card") && canFocus(el("plan-pane"))) return el("plan-pane");
  if (source && source.closest("#result-card") && canFocus(el("result-pane"))) return el("result-pane");
  if ((busy || requesting) && canFocus(el("log-pane"))) return el("log-pane");
  return el("prompt");
}
function repairFocus(previous) {
  // 只修复本次更新失效的焦点，不抢走用户已移动到其它控件的焦点。
  if (modalOpen || !previous || previous === document.body || canFocus(previous)) return;
  if (document.activeElement !== previous && document.activeElement !== document.body) return;
  focusVisible(focusFallback(previous));
}
function renderWarnings() {
  const parts = [];
  if (fileWarning) parts.push("本文件曾收到不落盘提示（尚未确认已恢复）：" + fileWarning);
  if (taskSafety) parts.push("当前任务安全提示（最近一条）：" + taskSafety);
  if (taskWarning) parts.push("当前任务提示（最近一条）：" + taskWarning);
  setText(el("banner-text"), parts.join("\n\n"));
  el("banner").style.display = parts.length ? "flex" : "none";
}

/* 主题只改根属性和入口标签，保留会话 DOM、面板状态与滚动。 */
const themes = ["dark", "light"];
let officeProbing = false;
const themeNames = { dark: "深色", light: "浅色" };
const themeStorageKey = "office-agent.theme";
let themeRevision = 0, themePending = null, themeWriting = false, themePersistence = '';
function applyTheme(theme) {
  if (!themes.includes(theme)) theme = "dark";
  document.documentElement.dataset.theme = theme;
  const next = themes[(themes.indexOf(theme) + 1) % themes.length];
  const label = "当前主题：" + themeNames[theme] + "；切换为" + themeNames[next];
  ["btn-theme", "settings-theme"].forEach((id) => {
    const button = el(id);
    if (!button) return;
    button.title = label;
    button.setAttribute("aria-label", label);
    if (id === "settings-theme") button.textContent = "切换主题（当前：" + themeNames[theme] + "）";
  });
  const state = el("settings-theme-state");
  if (state) state.textContent = "当前为" + themeNames[theme] + "。" + themePersistence;
  const choice = el("settings-theme-choice");
  if (choice) choice.value = theme;
}
function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem(themeStorageKey); } catch (e) { /* 存储受限时仍可切换 */ }
  // Fresh installs always open dark; an explicit saved choice still takes precedence.
  applyTheme(themes.includes(saved) ? saved : "dark");
  if (saved !== null && !themes.includes(saved)) {
    try { localStorage.setItem(themeStorageKey, "dark"); } catch (_) {}
  }
}
function cycleTheme() {
  const current = document.documentElement.dataset.theme;
  const next = themes[(themes.indexOf(current) + 1) % themes.length];
  saveTheme(next);
}
function saveTheme(next) {
  themePending = { theme: next, revision: ++themeRevision };
  themePersistence = "已应用，正在保存。";
  applyTheme(next);
  try { localStorage.setItem(themeStorageKey, next); } catch (e) { /* 保留本次窗口内的选择 */ }
  void flushTheme();
}
async function flushTheme() {
  if (themeWriting || !themePending) return;
  if (!api || typeof window.pywebview?.api?.floating_theme !== 'function') {
    themePersistence = "仅本窗口生效，尚未连接本机偏好。";
    applyTheme(document.documentElement.dataset.theme);
    return;
  }
  const owner = api;
  themeWriting = true;
  try {
    while (themePending && owner === api) {
      const pending = themePending;
      const result = await owner.floating_theme(pending.theme);
      if (owner !== api) break;
      if (!result?.ok || result.theme !== pending.theme) throw Error('Theme preference was not saved');
      if (themePending.revision === pending.revision) {
        themePending = null;
        themePersistence = "已保存在本机。";
        applyTheme(pending.theme);
      }
    }
  } catch (_) {
    if (owner === api) {
      themePersistence = "仅本窗口生效，偏好未保存。请重新选择后重试。";
      applyTheme(document.documentElement.dataset.theme);
    }
  } finally {
    themeWriting = false;
    if (owner !== api && themePending && api) void flushTheme();
  }
}
async function restoreTheme() {
  const owner = api, revision = themeRevision;
  if (!owner || typeof window.pywebview?.api?.floating_theme !== 'function') return;
  if (themePending) { void flushTheme(); return; }
  try {
    const result = await owner.floating_theme();
    if (owner !== api || revision !== themeRevision) return;
    if (!result?.ok) throw Error('Theme preference was not read');
    if (themes.includes(result.theme)) {
      themePersistence = "已保存在本机。";
      applyTheme(result.theme);
      try { localStorage.setItem(themeStorageKey, result.theme); } catch (_) {}
    } else saveTheme(document.documentElement.dataset.theme);
  } catch (_) {
    if (owner === api && revision === themeRevision) {
      themePersistence = "本机偏好未读取，当前选择仅本窗口生效。";
      applyTheme(document.documentElement.dataset.theme);
    }
  }
}
// 静态资源加载即启用主题，不等待 bridge；主题不调用任何后端 API。
initTheme();
el("btn-theme").addEventListener("click", cycleTheme);
el("settings-theme").addEventListener("click", cycleTheme);
el("settings-theme-choice")?.addEventListener("change", event => {
  const theme = event.target.value;
  if (!themes.includes(theme)) return;
  saveTheme(theme);
});

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/* 结构化表格 -> HTML <table>（数据来自后端，仍做转义防注入） */
const resultPageSize = 50;
let resultTable = null;
let resultText = "";
let resultPage = 0;
let resultEvent = null;
let resultSummary = "";

// 只使用 bridge 的真实 data；不尝试从模型/YAML 文本推断风险或总量。
function dataObject(value) { return value && typeof value === "object" && !Array.isArray(value) ? value : {}; }
function humanValue(value) {
  if (value == null) return "未提供";
  if (typeof value === "object") return "详情见技术信息";
  return String(value);
}
function toolSummary(ev) {
  if (!ev.tool) return String(ev.text ?? ""); // 模型 answer/state 原样保留。
  const d = dataObject(ev.data), c = dataObject(d.classification), m = dataObject(d.metadata || c.metadata);
  const lines = [];
  const add = text => lines.push(String(text));
  const risk = text => add("⚠ " + text);
  const field = (key, fallback) => d[key] !== undefined ? d[key] : c[fallback || key];
  if (ev.tool === "inspect" || ev.tool === "inspect_document") {
    const path = field("path");
    const kind = field("kind");
    const title = { spreadsheet: "Excel 工作簿", word: "Word 文档", legacy: "旧版 Office 文档" }[kind] || "文件报告";
    add(title + (typeof path === "string" ? " · " + path.split(/[\\/]/).pop() : ""));
    const size = d.size_bytes ?? m.size;
    if (typeof size === "number" && Number.isFinite(size)) add("文件大小：约 " + Math.round(size / 1024 * 10) / 10 + " KB");
    const readable = field("readable");
    if (readable === true) add("文件结构可读取（不代表全部功能或内容均已验证）");
    else if (readable === false) risk("当前文件不能正常读取");
    else add("可读状态未提供，不能据此判断可正常读取");
    [["protected", "加密或保护", "protected"], ["macros", "宏", "macro_enabled"],
      ["external_links", "外部链接", "external_relationships"]].forEach(([key, label, fallback]) => {
      const value = d[key] !== undefined ? d[key] : c.format === "ooxml" ? c[fallback] : undefined;
      if (value === true) risk("检测到" + label);
      else if (value === false) add("未检测到" + label);
      else add(label + "：检测状态未提供");
    });
  }
  // 清单字段的条目数只称“已接收”；仅后端明确 count 才声称总量。
  const inventories = {
    sheets: ["工作表", "sheet_count"], sections: ["节 / 页眉页脚", "section_count"],
    comments: ["批注", "comment_count"], images: ["图片", "image_count"],
    defined_names: ["命名区域", "defined_name_count"], charts: ["图表", "chart_count"],
    hidden_rows: ["隐藏行", "hidden_row_count"], hidden_columns: ["隐藏列", "hidden_column_count"],
    merged_ranges: ["合并区域", "merged_range_count"], conditional_formatting: ["条件格式区域", "conditional_formatting_range_count"],
    insertions: ["插入修订", "insertion_count"], deletions: ["删除修订", "deletion_count"],
    content_controls: ["内容控件", "content_control_count"], tables: ["表格", "table_count"],
    paragraphs: ["段落", "paragraph_count"], blocks: ["正文块", "total_blocks"], cells: ["单元格", "cell_count"]
  };
  const itemText = item => {
    if (!item || typeof item !== "object") return humanValue(item);
    const fields = ["title", "name", "location", "sheet", "cell", "coordinate", "range", "refers_to", "index", "id", "author", "placement", "position", "control_type", "tag", "text", "match", "formula_text", "value"];
    const parts = fields.filter(key => item[key] !== undefined).map(key => humanValue(item[key]));
    if (item.state !== undefined) parts.push(({visible: "可见", hidden: "隐藏", veryHidden: "深度隐藏"})[item.state] || humanValue(item.state));
    if (item.refers_to_truncated === true || item.text_truncated === true) parts.push("此条文本被截断");
    return parts.join(" · ") || "条目详情见技术信息";
  };
  function inventory(obj, prefix = "") {
    Object.entries(inventories).forEach(([key, [label, countKey]]) => {
      if (!Array.isArray(obj[key])) return;
      const items = obj[key], total = obj[countKey], clipped = obj[key + "_truncated"];
      add(prefix + label + "：已接收 " + items.length + " 条" + (typeof total === "number" ? " / 总量 " + total : "；总量未提供") +
        (clipped === true ? "；后端已截断，仅显示返回的前 " + items.length + " 条" : clipped === false ? "；后端未截断" : "；截断状态未提供"));
      items.forEach(item => {
        add("  • " + itemText(item));
        if (key === "sheets" && item && typeof item === "object") inventory(item, "    ");
        if (key === "sections" && item && typeof item === "object") {
          ["header", "footer"].forEach(part => { if (Array.isArray(item[part])) item[part].forEach(entry => add("    " + (part === "header" ? "页眉" : "页脚") + "：" + itemText(entry))); });
        }
      });
    });
  }
  if (ev.tool === "search") {
    const hits = Array.isArray(ev.data) ? ev.data : Array.isArray(d.matches) ? d.matches : null;
    if (hits) {
      add("搜索命中：已接收 " + hits.length + " 条" + (Array.isArray(ev.data) ? "（本次返回列表全部展示；文档命中总量及后端截断状态未提供）" :
        (typeof d.total === "number" ? " / 总量 " + d.total : "；总量未提供") + (d.truncated === true ? "；后端已截断" : d.truncated === false ? "；后端未截断" : "；截断状态未提供")));
      hits.forEach(hit => add("  • " + itemText(hit)));
    }
  } else inventory(d);
  inventory(dataObject(d.tracked_changes));
  [d, c].forEach(obj => {
    ["risks", "limitations", "suspicious_features", "parts_unscanned", "errors", "warnings", "merge_issues"].forEach(key => {
      if (Array.isArray(obj[key]) && obj[key].length) risk((key === "parts_unscanned" ? "未扫描部件：" : "文件限制 / 风险：") + obj[key].map(humanValue).join("；"));
    });
  });
  const changes = dataObject(d.tracked_changes);
  if (Array.isArray(changes.parts_unscanned) && changes.parts_unscanned.length) risk("修订存在未扫描部件：" + changes.parts_unscanned.join("；") + "。不能据此判断没有修订");
  if (d.merge_consistent === false) risk("表格合并结构不自洽，不能据此安全定位写入坐标");
  if (ev.tool === "read_excel_region" && Array.isArray(d.cells) && d.cells.some(cell => cell && cell.formula_text !== undefined)) add("公式仅显示存储文本；缓存值未经验证，不会重新计算公式。");
  if (d.drawings_readable === false) risk("绘图部件未能读取，不能判断没有图片");
  if (d.counts_cover_detailed_sheets_only === true) add("图表 / 图片总数仅覆盖已详列的工作表，不代表整个工作簿");
  if (d.has_more === true) add("正文尚有未返回内容；本次结束位置：" + humanValue(d.next_offset));
  if (d.truncated === true) add("后端已截断，本次结果不是完整内容");
  if (d.note) add(humanValue(d.note));
  if (d.error) risk("读取失败：" + humanValue(d.error));
  if (d.readable === false && ev.tool !== "inspect" && ev.tool !== "inspect_document") risk("内容不可读取；请核对限制信息");
  if (ev.table && Array.isArray(ev.table.rows)) add("表格已接收 " + ev.table.rows.length + " 行；下方每页显示最多 50 行，复制 TSV 包含全部已接收行");
  return lines.join("\n") || "工具已返回；可用摘要字段不足，详情见技术信息。不能据此确认可读性、无风险或完整总量。";
}
function resultCopyText(ev, summary = toolSummary(ev)) {
  return ev.tool ? summary + "\n\n原始输出：\n" + String(ev.text ?? "") : String(ev.text ?? "");
}
function appendResultInfo(pane, ev, summary = toolSummary(ev)) {
  const text = document.createElement("section"); text.className = "result-human-summary";
  text.setAttribute("aria-label", ev.tool ? "文件处理结果" : "任务回复");
  summary.split("\n").forEach((line, index) => {
    if (!line.trim()) return;
    const row = document.createElement(ev.tool && index === 0 ? "h3" : "p");
    row.textContent = line;
    if (line.startsWith("⚠")) row.className = "result-risk";
    text.append(row);
  });
  pane.prepend(text);
  window.OfficeImages?.appendImages(pane, ev.images);
  if (!ev.tool) return;
  const details = document.createElement("details"); details.className = "result-source";
  const title = document.createElement("summary"); title.textContent = "技术信息（原始输出）";
  const raw = document.createElement("pre"); raw.textContent = String(ev.text ?? "");
  details.append(title, raw);
  if (ev.data !== undefined) {
    const label = document.createElement("p"); label.textContent = "结构化返回（JSON）";
    const data = document.createElement("pre"); data.textContent = JSON.stringify(ev.data, null, 2);
    details.append(label, data);
  }
  pane.append(details); // 默认关闭，原文一字不丢；不把展示 JSON 变成执行权限。
}

function renderTable(t, start = 0, end = start + resultPageSize) {
  let html = "";
  if (t.caption) html += '<div class="table-caption">' + escapeHtml(t.caption) + '</div>';
  html += '<table class="result-table"><thead><tr>';
  if (t.row_labels) html += '<th class="rowhead"></th>';
  (t.headers || []).forEach((h) => { html += '<th>' + escapeHtml(h) + '</th>'; });
  html += '</tr></thead><tbody>';
  (t.rows || []).slice(start, end).forEach((row, i) => {
    html += '<tr>';
    if (t.row_labels) html += '<td class="rowhead">' + escapeHtml(t.row_labels[start + i] ?? "") + '</td>';
    (row || []).forEach((cell) => { html += '<td>' + escapeHtml(cell) + '</td>'; });
    html += '</tr>';
  });
  html += '</tbody></table>';
  return html;
}

function renderResult(ev) {
  const previous = document.activeElement;
  const replacingFocus = previous !== el("result-pane") && el("result-pane").contains(previous);
  resultEvent = ev;
  resultSummary = toolSummary(ev);
  resultText = String(ev.text ?? "");
  resultTable = ev.table && Array.isArray(ev.table.rows) ? ev.table : null;
  resultPage = 0;
  el("result-pane").classList.remove("empty-state");
  el("result-controls").hidden = !resultTable;
  el("btn-copy-tsv").disabled = !resultTable;
  if (resultTable) renderResultPage();
  else {
    el("result-pane").replaceChildren();
    appendResultInfo(el("result-pane"), ev, resultSummary);
  }
  if (replacingFocus) focusVisible(el("result-pane"));
  else repairFocus(previous);
}

function renderResultPage() {
  if (!resultTable) return;
  const previous = document.activeElement;
  const replacingFocus = previous !== el("result-pane") && el("result-pane").contains(previous);
  const total = resultTable.rows.length;
  const pages = Math.max(1, Math.ceil(total / resultPageSize));
  resultPage = Math.max(0, Math.min(resultPage, pages - 1));
  const start = resultPage * resultPageSize;
  const end = Math.min(total, start + resultPageSize);
  el("result-range").textContent = "已接收总量 " + total + " 行 · 当前 " + (total ? start + 1 : 0) + "–" + end + " 行 · 第 " + (resultPage + 1) + "/" + pages + " 页";
  el("btn-result-prev").disabled = resultPage === 0;
  el("btn-result-next").disabled = resultPage >= pages - 1;
  const pane = el("result-pane");
  pane.innerHTML = total ? renderTable(resultTable, start, end) : "";
  if (!total) {
    const empty = document.createElement("p");
    empty.textContent = "本次结构化结果没有数据行。";
    pane.appendChild(empty);
  }
  appendResultInfo(pane, resultEvent, resultSummary);
  pane.scrollTop = 0;
  if (replacingFocus) focusVisible(pane);
  else repairFocus(previous);
}

function resultAsTSV(t) {
  // 引号保护制表符、换行及双引号；不把前端页码/说明混入数据。
  const field = (value) => {
    const text = String(value ?? "");
    return /[\t\r\n"]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  };
  const lines = [];
  if (t.headers && t.headers.length) lines.push((t.row_labels ? ["行", ...t.headers] : t.headers).map(field).join("\t"));
  t.rows.forEach((row, i) => {
    lines.push((t.row_labels ? [t.row_labels[i] ?? "", ...row] : row).map(field).join("\t"));
  });
  return lines.join("\r\n");
}
async function copyResultTSV() {
  if (!resultTable) return;
  await copyText(resultAsTSV(resultTable), "全部已接收表格（" + resultTable.rows.length + " 行，TSV）");
}

/* C1：定位流内卡片，不再切换/隐藏另一种内容。 */
function selectPanel(name, focus = false) {
  if (focus) {
    const card = el(name + "-card");
    if (card.closest("#card-store")) return;
    focusVisible(el(name + "-pane"));
  }
}

/* C2：记录仍逐条计数；按真实 task 分轮，连续过程段折叠，不跨卡片重排事件。 */
let currentRound = null;
let roundSequence = 0;
let authorizedFileName = "尚未选择";
let localMainRound = null;
function localMainMessage(question) {
  currentRound = newFlowRound(question);
  localMainRound = currentRound;
  currentRound.state = '正在发送';
  appendLog('user-message', question);
  const node = document.createElement('p');node.className='log-line answer';node.setAttribute('role','status');
  node.textContent='正在发送…';currentRound.localReply=node;
  appendLog('answer',node.textContent,node);
  const owner=currentRound;
  setTimeout(()=>{if(localMainRound===owner)setMainReply('消息还在等待应用接收',owner);},1800);
  return owner;
}
function setMainReply(text,round=currentRound) {
  if(!round?.localReply?.isConnected)return false;
  round.localReply.textContent=text;
  const record=logRecords.concat(logQueue.slice(logQueueHead)).find(value=>value?.node===round.localReply);
  if(record)record.text=text;
  return true;
}
function newFlowRound(question = "") {
  return { number: ++roundSequence, question: String(question), file: authorizedFileName,
    state: question ? "指令已受理" : "窗口记录", expanded: false, node: null, trace: null };
}
function flowRoundNode(round) {
  if (!round.node || !round.node.isConnected) {
    round.node = document.createElement("article");
    round.node.className = "assistant-round";
    round.node.dataset.round = String(round.number);
    const meta = document.createElement("div"); meta.className = "reply-meta";
    const state = document.createElement("span"); state.className = "reply-state";
    const source = document.createElement("span"); source.className = "reply-source";
    source.textContent = "来源：" + round.file;
    meta.append(state, source); round.node.append(meta); round.trace = null;
    el("log-pane").append(round.node);
  }
  setText(round.node.querySelector(".reply-state"), round.state);
  return round.node;
}
function updateRoundState(text) {
  if (!currentRound) return;
  currentRound.state = String(text || "状态未说明");
  if (currentRound.node?.isConnected) setText(currentRound.node.querySelector(".reply-state"), currentRound.state);
}
function refreshTrace(trace) {
  if (!trace?.isConnected) return;
  const count = trace.querySelector(".trace-lines").children.length;
  if (!count) { trace.remove(); return; }
  setText(trace.querySelector("summary"), "处理过程 · " + count + " 条记录（仅本窗口保留）");
}
function appendFlowRecord(record) {
  const pane = el("log-pane");
  let node = record.node;
  if (!node) {
    node = document.createElement("div");
    node.className = record.tag === "user-message" ? "user-message" : "log-line";
    if (record.tag !== "user-message" && /^[a-z0-9_-]+$/i.test(record.tag)) node.classList.add(record.tag);
    if (record.tag === "error") {
      const notice = document.createElement("p"); notice.textContent = humanErrorNotice(record.text);
      const details = document.createElement("details"); details.className = "error-source";
      const summary = document.createElement("summary"); summary.textContent = "技术信息（错误详情）";
      const raw = document.createElement("pre"); raw.textContent = record.text;
      details.append(summary, raw); node.append(notice, details);
    } else node.textContent = record.text;
  }
  record.node = node;
  if (record.tag === "user-message") { pane.append(node); return; }
  const round = record.round;
  const parent = flowRoundNode(round);
  if (record.trace) {
    if (!round.trace?.isConnected) {
      const trace = document.createElement("details"); trace.className = "processing-trace";
      trace.open = round.expanded;
      const summary = document.createElement("summary"), lines = document.createElement("div");
      lines.className = "trace-lines"; trace.append(summary, lines); parent.append(trace); round.trace = trace;
      trace.addEventListener("toggle", () => {
        if (round.expanded === trace.open) return;
        round.expanded = trace.open;
        round.node?.querySelectorAll(".processing-trace").forEach(segment => { segment.open = round.expanded; });
      });
    }
    round.trace.querySelector(".trace-lines").append(node); refreshTrace(round.trace);
  } else {
    round.trace = null; // 卡片／安全消息结束当前连续段，后续日志保持真实位置。
    parent.append(node);
  }
}
function removeFlowRecord(record) {
  const node = record.node;
  if (!node) return;
  const trace = node.closest(".processing-trace"), group = node.closest(".assistant-round");
  if (node === el("plan-card") || node === el("result-card")) el("card-store").append(node);
  else node.remove();
  refreshTrace(trace);
  if (group && group.children.length === 1) group.remove();
}
function addSnapshotActions(snapshot, record, table = null) {
  const actions = document.createElement("div"); actions.className = "message-actions";
  actions.setAttribute("role", "group"); actions.setAttribute("aria-label", "历史消息复制（只读）");
  const copy = document.createElement("button"); copy.type = "button"; copy.className = "btn link message-copy";
  copy.textContent = "复制文本"; copy.disabled = busy || requesting || modalOpen;
  copy.onclick = () => copyText(record.text, "历史消息文本"); actions.append(copy);
  if (table) {
    const tsv = document.createElement("button"); tsv.type = "button"; tsv.className = "btn link message-copy";
    tsv.textContent = "复制全部 TSV"; tsv.disabled = busy || requesting || modalOpen;
    tsv.onclick = () => copyText(resultAsTSV(table), "历史表格全部已接收行（TSV）"); actions.append(tsv);
  }
  snapshot.append(actions);
}

function initConversation() {
  const composer = document.createElement("section");
  composer.className = "composer";
  composer.setAttribute("aria-label", "附件与任务输入");
  composer.append(document.querySelector(".file-strip"));
  const options = document.createElement("details");
  options.className = "attachment-options";
  const summary = document.createElement("summary");
  summary.textContent = "拖入文件 / 最近文档 / 粘贴路径";
  options.append(summary, el("file-entry"));
  const vision = document.createElement("div"); vision.className = "vision-actions";
  const localImage = document.createElement("button"); localImage.id = "vision-local";
  localImage.type = "button"; localImage.className = "btn secondary"; localImage.textContent = "选择图片分析…";
  localImage.title = "选择本地 PNG/JPEG，并在确认后发送给视觉模型";
  const documentImage = document.createElement("button"); documentImage.id = "vision-document";
  documentImage.type = "button"; documentImage.className = "btn secondary"; documentImage.textContent = "分析文档图片…";
  documentImage.title = "从当前 DOCX/XLSX 的图片清单选择一张分析";
  const imageList = document.createElement("select"); imageList.id = "vision-list";
  imageList.setAttribute("aria-label", "当前文档中的图片"); imageList.hidden = true;
  vision.append(localImage, imageList, documentImage);
  composer.append(options, document.querySelector(".prompt-bar"), vision, el("input-help"));
  localImage.onclick = () => requestAction("正在分析图片…", () => analyzeVision("local"));
  documentImage.onclick = () => requestAction("正在分析图片…", () => analyzeVision("document"));
  el("work-view").prepend(composer);
  el("work-view").insertBefore(el("pending-notice"), el("content-panes"));
  document.querySelectorAll("[data-prompt]").forEach(button => {
    button.addEventListener("click", () => {
      if (busy || requesting || modalOpen) return;
      el("prompt").value = button.dataset.prompt;
      focusVisible(el("prompt"));
    });
  });
  refreshEmptyFlow();
}

// 保留脚本绑定的当前卡片；替换前将旧内容固化为无写入入口的只读卡片。
function archiveFlowCard(id) {
  const card = el(id);
  const record = logRecords.concat(logQueue.slice(logQueueHead)).find(item => item && item.node === card);
  if (!record) return;
  const snapshot = card.cloneNode(true);
  snapshot.classList.remove("plan-pending");
  snapshot.classList.add("archived-card");
  snapshot.querySelectorAll("button, .plan-actions, .message-actions").forEach(node => node.remove());
  [snapshot, ...snapshot.querySelectorAll("[id]")].forEach(node => node.removeAttribute("id"));
  snapshot.querySelectorAll("[aria-labelledby], [aria-controls]").forEach(node => {
    node.removeAttribute("aria-labelledby"); node.removeAttribute("aria-controls");
  });
  const label = document.createElement("p");
  label.className = "archive-note";
  label.textContent = "此前的内容 · 仅供回看，不代表现在有读取权限，也不能应用或撤销";
  snapshot.append(label);
  if (id === "result-card" && resultTable) {
    const table = resultTable, event = resultEvent, summary = resultSummary;
    const body = snapshot.querySelector(".pane-body");
    const controls = snapshot.querySelector(".result-controls");
    let page = resultPage;
    const previous = document.createElement("button"), next = document.createElement("button");
    previous.className = next.className = "btn secondary";
    previous.textContent = "上一页"; next.textContent = "下一页";
    const range = document.createElement("span");
    controls.replaceChildren(range, previous, next);
    function draw() {
      const focused = document.activeElement;
      const pages = Math.max(1, Math.ceil(table.rows.length / resultPageSize));
      previous.disabled = page === 0; next.disabled = page === pages - 1;
      range.textContent = "已接收 " + table.rows.length + " 行 · " + (page + 1) + "/" + pages + " 页";
      body.innerHTML = renderTable(table, page * resultPageSize);
      appendResultInfo(body, event, summary);
      if (focused && !canFocus(focused) && (document.activeElement === focused || document.activeElement === document.body)) focusVisible(body);
    }
    previous.onclick = () => { page--; draw(); };
    next.onclick = () => { page++; draw(); };
    draw();
  }
  if (id === "result-card" && !resultTable && resultEvent?.images) {
    snapshot.querySelector(".image-results")?.remove();
    window.OfficeImages?.appendImages(snapshot.querySelector(".pane-body"), resultEvent.images);
  }
  addSnapshotActions(snapshot, record, id === "result-card" ? resultTable : null);
  record.node = snapshot;
  if (el("log-pane").contains(card)) card.replaceWith(snapshot);
  el("card-store").append(card);
}
function enqueueCard(id, text) {
  const card = el(id);
  card.hidden = false;
  appendLog("ai", text, card);
}
function refreshEmptyFlow() {
  el("conversation-empty").hidden = Boolean(logRecords.length || logQueue.length > logQueueHead);
  el("work-view").classList.toggle("empty-flow", !el("conversation-empty").hidden);
}

/* 历史是独立只读视图；工作 DOM 保留，不复制/重放历史事件。 */
let currentView = "work";
let workScrollTop = 0;
let workScrollContainer = null;
let historySessions = [];
let historySelected = null;
let historyCurrentDir = null;
let historyRequest = 0;
let historyLoading = false;
let historyWarning = "";
let historyError = "";
let historySelectingFile = false;
const historyHidden = new Set(); // 仅本窗口记忆；不写 localStorage 或磁盘。

function setNavigation(view) {
  currentView = view === "history" ? "history" : "work";
  const historical = currentView === "history";
  // 保留当前任务全部 DOM 和安全状态；只切换显示，绝不重放历史事件。
  el("work-view").hidden = historical;
  el("history-reader").hidden = !historical;
  el("history-view").hidden = false;
  el("workspace").dataset.view = currentView;
  el("nav-work").classList.toggle("active", !historical);
  el("nav-history").classList.toggle("active", historical);
  if (historical) { el("nav-work").removeAttribute("aria-current"); el("nav-history").setAttribute("aria-current", "page"); }
  else { el("nav-work").setAttribute("aria-current", "page"); el("nav-history").removeAttribute("aria-current"); }
  if (!window.OfficeWorkbench) el("page-title").textContent = historical ? "历史会话 · 只读" : "当前任务";
  window.dispatchEvent(new CustomEvent('officeapp:navigation', {detail:{view:currentView}}));
  refreshButtons();
}
function navigate(view) {
  if (modalOpen) return;
  const enteringHistory = view === 'history' && currentView !== 'history';
  setNavigation(view);
  if (view === "history") {
    if (enteringHistory && !historyLoading && !historySelectingFile) refreshHistory();
    if (window.OfficeWorkbench) closeDrawer(false);
    else { openDrawer(); focusVisible(el("btn-history-refresh")); }
  }
  else {
    historySelected = null; renderHistoryDetail();
    if (view === "log") el("technical-details").open = true;
    closeDrawer(false); focusVisible(el("prompt"));
  }
}
el("nav-work").addEventListener("click", () => { if (!window.OfficeWorkspace) navigate("work"); });
el("nav-log").addEventListener("click", () => navigate("log"));
el("nav-history").addEventListener("click", () => { if (!window.OfficeWorkspace) navigate("history"); });
el("btn-history-current").addEventListener("click", () => window.OfficeWorkspace ? window.OfficeWorkspace.open('work') : navigate('work'));
el("btn-history-refresh").addEventListener("click", refreshHistory);
el("btn-history-hide").addEventListener("click", hideHistorySession);
el("history-search-toggle").addEventListener("click", () => {
  const panel = el("history-search-panel"), expanded = panel.hidden;
  panel.hidden = !expanded;
  el("history-search-toggle").setAttribute("aria-expanded", String(expanded));
  if (expanded) el("history-search").focus();
  else { historyFilter = ""; el("history-search").value = ""; renderHistory(); focusVisible(el("history-search-toggle")); }
});
el("history-search").addEventListener("input", () => { historyFilter = el("history-search").value; renderHistory(); });
document.querySelector("main").addEventListener("focusin", (e) => {
  if (currentView === "history" || window.OfficeWorkspace) return;
  setNavigation("work");
});

function metadataText(value) {
  return value == null ? "未记录" : typeof value === "object" ? JSON.stringify(value, null, 2) : String(value);
}
function translateHistoryText(value) {
  const text = String(value ?? "");
  const trimmed = text.trim();
  const internalKeys = ["classification", "skipped", "question", "goal"];
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    try {
      const parsed = JSON.parse(trimmed);
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed) && internalKeys.some(key => Object.prototype.hasOwnProperty.call(parsed, key))) {
        if (Object.prototype.hasOwnProperty.call(parsed, "classification")) return "这条记录当时没能确认需求";
        if (parsed.skipped === true) return "这条记录当时跳过了部分内容";
        if (Object.prototype.hasOwnProperty.call(parsed, "skipped")) return "这条记录包含需求处理信息";
        return "这条记录包含当时的需求信息";
      }
    } catch (_) { /* 旧记录可能保存了不完整的内部对象，下面会剥除对象块。 */ }
  }
  if (/\{"[^"\n]+"\s*:/.test(text)) {
    const objectBlock = /\{[\s\S]*\}/.exec(text);
    return objectBlock ? text.replace(objectBlock[0], "（当时的内部信息已省略）") : text.replace(/\{[\s\S]*$/, "（当时的内部信息已省略）");
  }
  return text;
}
function appendHistorySource(parent, original) {
  const details = document.createElement("details"); details.className = "history-source";
  const summary = document.createElement("summary"); summary.textContent = "查看当时的内部信息";
  const raw = document.createElement("pre"); raw.textContent = original;
  details.append(summary, raw); parent.append(details);
}
function appendHistoryText(parent, value, className = "history-event-text") {
  const original = metadataText(value), translated = translateHistoryText(original);
  const text = document.createElement("pre"); text.className = className;
  text.textContent = translated || "未记录"; parent.append(text);
  if (translated !== original) appendHistorySource(parent, original);
  return translated;
}
function isCurrentSession(session) {
  return session.current === true || (historyCurrentDir !== null && session.session_dir === historyCurrentDir);
}
let historyLoadToken = 0;
let historyLoadedDir = null;
let historyTranscriptState = null;
let historyFilter = "";
function notAuthenticated(response) {
  return Boolean(response && (response.reason === "not_authenticated" || response.error === "not_authenticated"));
}
function historyStatus(value) {
  return ({ idle: "就绪", running: "处理中", completed: "已完成", done: "已完成", failed: "失败", error: "错误",
    pending: "待确认", awaiting_confirmation: "等待确认", waiting_confirmation: "等待确认", cancelled: "已取消",
    stopped: "已停止", budget: "内容太多放不下（已安全停止）", limit: "步数用完（已安全停止）",
    confirmation_required: "等待你确认是否写入", budget_exceeded: "内容太多放不下", max_iterations: "步数已用完",
    verified: "副本已验证", rolled_back: "已撤销", rolling_back: "撤销待核实", uncertain: "结果待核实" })[value] || metadataText(value);
}
function historyDate(value) {
  if (typeof value !== "string" || !value) return metadataText(value);
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString("zh-CN", { hour12: false });
}
function resetHistoryTranscript() {
  ++historyLoadToken; historyLoadedDir = null; historyTranscriptState = null;
  el("history-transcript")?.replaceChildren();
}
function appendReadOnlyTable(host, table) {
  let page = 0;
  const controls = document.createElement("div"), range = document.createElement("span");
  const prev = document.createElement("button"), next = document.createElement("button"), body = document.createElement("div");
  prev.type = next.type = "button"; prev.className = next.className = "btn secondary";
  prev.textContent = "上一页"; next.textContent = "下一页";
  controls.append(range, prev, next); host.append(controls, body);
  function draw() {
    const total = table.rows.length, pages = Math.max(1, Math.ceil(total / resultPageSize));
    page = Math.max(0, Math.min(page, pages - 1));
    const start = page * resultPageSize, end = Math.min(total, start + resultPageSize);
    range.textContent = "已接收 " + total + " 行 · 当前 " + (total ? start + 1 : 0) + "–" + end + " 行 · 第 " + (page + 1) + "/" + pages + " 页";
    prev.disabled = page === 0; next.disabled = page === pages - 1;
    body.innerHTML = renderTable(table, start, end); // 所有数据经 escapeHtml；独立页码不触碰当前表格。
  }
  prev.onclick = () => { page--; draw(); }; next.onclick = () => { page++; draw(); }; draw();
}
function renderHistoryTranscript() {
  const host = el("history-transcript");
  if (!host) return;
  host.replaceChildren();
  if (!historySelected) return;
  const state = historyTranscriptState;
  const note = document.createElement("p"); note.className = "history-note";
  note.textContent = "历史记录仅供查看与复制；不恢复任务、不授予读取许可、不提供写入或撤销入口。";
  host.append(note);
  const message = document.createElement("p"); host.append(message);
  if (!state) { message.textContent = "请选择一条记录查看内容。"; return; }
  if (state.loading) { message.textContent = "正在读取历史记录…"; return; }
  if (state.error) { message.textContent = "历史记录加载失败；请重新选择会话或刷新重试。"; message.style.color = "var(--danger)"; appendArchivedError(host, state.error); return; }
  if (!state.events) { appendHistoryText(message, state.warning || "这条记录没有完整的对话内容。"); return; }
  message.textContent = state.truncated ? "这条记录已截断，以下内容不完整。" : "以下是这条记录的内容。";
  const hasTask = state.events.some(ev => ev && ev.type === "task");
  state.events.forEach(ev => {
    if (!ev || typeof ev !== "object") return;
    if (hasTask && ev.type === "log" && ev.tag === "user") return;
    const card = document.createElement("article"); card.className = "card archived-card";
    const title = document.createElement("h3"), body = document.createElement("div");
    const labels = { task: "当时的要求", answer: "模型回复", result: "读取结果", plan: "当时的修改计划（已失效，不能应用）", output: "当时的输出（已失效，不能撤销）",
      file: "当时打开的文件（不代表现在也在用）", status: "状态", log: "处理记录", checkpoint: "进度记录（不能接着跑）", diagram: "生成的图表",
      read_confirm: "当时询问过是否允许读取（现已失效）", badge: "状态标记", progress: "当时的进度", busy: "处理状态" };
    title.textContent = ev.type==='output'&&state.recordKind==='manual_copy'?'副本事务记录（不恢复撤销批准）':labels[ev.type] || "记录"; card.append(title, body);
    let copy = String(ev.text ?? "");
    let table = null;
    if (ev.type === "result") {
      const summary = toolSummary(ev); copy = resultCopyText(ev, summary);
      table = ev.table && Array.isArray(ev.table.rows) ? ev.table : null;
      if (table) appendReadOnlyTable(body, table);
      appendResultInfo(body, ev, summary);
    } else if (ev.type === "file") {
      copy = "文件：" + String(ev.name ?? "未记录") + "\n" + toolSummary({ tool: "inspect", data: ev.inspect });
      appendHistoryText(body, copy);
    } else if (ev.tag === "error" || ev.type === "error" || (ev.level === "danger" && ev.tag !== "safety")) {
      appendArchivedError(body, copy); body.style.color = "var(--danger)";
    } else if (ev.type === "read_confirm") {
      copy = "当时曾询问是否允许读取：" + String(ev.tool ?? "未记录") + "。这条记录不会给你任何现在的读取权限。";
      body.textContent = copy;
    } else {
      if (!copy && ev.type === "badge") copy = String(ev.label ?? "");
      if (!copy && ev.type === "busy") copy = ev.busy ? "当时正在处理" : "当时处理结束";
      if (!copy && ev.type === "progress") copy = "当时已处理 " + humanValue(ev.iterations) + " 轮。";
      appendHistoryText(body, copy || "事件未提供文本内容。");
      if (ev.level === "danger" || ["error", "warn", "safety"].includes(ev.tag)) body.style.color = ev.tag === "error" || ev.level === "danger" ? "var(--danger)" : "var(--warn)";
    }
    addSnapshotActions(card, { text: copy }, table);
    host.append(card);
  });
}
async function loadHistorySession(session) {
  const token = ++historyLoadToken, request = historyRequest, dir = session.session_dir;
  historyLoadedDir = dir;
  if (session.history_available !== true) {
    historyTranscriptState = { warning: "此会话未提供完整记录，只有元数据 / 结果摘要，不能恢复原始对话。" };
    renderHistoryTranscript(); return;
  }
  historyTranscriptState = { loading: true }; renderHistoryTranscript();
  try {
    if (!api || typeof api.load_session !== "function") throw new Error("服务未提供只读 load_session API");
    window.dispatchEvent(new Event("officediagram:invalidate"));
    const response = await api.load_session(dir); // dir 仅来自 list_sessions；后端负责路径校验。
    if (token !== historyLoadToken || request !== historyRequest || dir !== historySelected || historySelectingFile) return;
    if (notAuthenticated(response)) { resetHistoryTranscript(); return; }
    if (!response || response.ok !== true || !Array.isArray(response.events)) throw new Error(response?.error || response?.reason || "历史返回格式无效");
    historyTranscriptState = { events: response.events, recordKind:response.record_kind, truncated: response.history_truncated === true || response.truncated === true, warning: String(response.warning || "") };
  } catch (e) {
    if (token !== historyLoadToken || request !== historyRequest || dir !== historySelected) return;
    historyTranscriptState = { error: e?.message || String(e) };
  }
  if (token === historyLoadToken && request === historyRequest && dir === historySelected) renderHistoryTranscript();
}
function renderHistoryDetail() {
  const session = historySessions.find((item) => item.session_dir === historySelected && !historyHidden.has(item.session_dir));
  el("history-detail").hidden = !session;
  el("history-metadata").replaceChildren();
  el("history-summary").textContent = "";
  el("history-summary").parentElement.querySelectorAll(":scope > .history-source").forEach(node => node.remove());
  el("btn-history-hide").disabled = !session || busy || requesting || modalOpen || historyLoading;
  if (!session) { resetHistoryTranscript(); return; }
  if(session.record_kind==='manual_copy'){
    const term=document.createElement('dt'),body=document.createElement('dd');term.textContent='副本';el('history-metadata').append(term,body);
    const info=document.createElement('p');info.textContent=session.format+' · '+historyStatus(session.status)+' · '+historyDate(session.created_at)+'（审计更新时间）';body.append(info);
    const details=document.createElement('details'),summary=document.createElement('summary'),path=document.createElement('p');summary.textContent='源文件与副本位置';path.textContent=session.file+' → '+session.output;details.append(summary,path);body.append(details);
    const open=document.createElement('button');open.type='button';open.textContent='重新打开副本';open.disabled=session.status!=='verified'||busy||requesting;
    const feedback=document.createElement('p');feedback.setAttribute('role','status');
    open.onclick=async()=>{open.disabled=true;feedback.textContent='正在核对副本身份…';try{const opened=await window.OfficeFloating.openHistoryOutput(session.transaction_id);feedback.textContent=opened?'请在工作区确认读取':'已取消打开，当前草稿保留';}catch(error){feedback.textContent=error.message;}finally{if(open.isConnected)open.disabled=session.status!=='verified';}};
    body.append(open,feedback);el('history-summary').textContent=session.result_summary;
    if(!historyLoading&&historyLoadedDir!==session.session_dir)loadHistorySession(session);
    return;
  }
  [["会话目录", session.session_dir], ["文件", session.file], ["指令", session.question],
    ["创建时间", historyDate(session.created_at)], ["状态", historyStatus(session.status)],
    ["历史记录", session.history_available === true ? "可读取" + (session.history_truncated === true ? "（已截断）" : "") : "完整记录未提供"],
    ["事件数量", session.event_count], ["当前会话", isCurrentSession(session) ? "是" : "否"]].forEach(([label, value]) => {
    const term = document.createElement("dt"), description = document.createElement("dd");
    term.textContent = label; appendHistoryText(description, value, "history-metadata-value"); el("history-metadata").append(term, description);
  });
  const summary = metadataText(session.result_summary ? session.result_summary : "未记录结果摘要。");
  el("history-summary").textContent = translateHistoryText(summary);
  if (el("history-summary").textContent !== summary) appendHistorySource(el("history-summary").parentElement, summary);
  if (!historyLoading && historyLoadedDir !== session.session_dir) loadHistorySession(session);
}
function renderHistory() {
  const list = el("history-list");
  const focusedSession = list.contains(document.activeElement) ? document.activeElement.dataset.session : null;
  list.textContent = "";
  list.setAttribute("aria-busy", String(historyLoading));
  el("btn-history-refresh").disabled = !connected || historyLoading || historySelectingFile || modalOpen;
  const available = historySessions.filter((session) => !historyHidden.has(session.session_dir));
  if (available.length < 3) {
    historyFilter = "";
    el("history-search").value = "";
    el("history-search-panel").hidden = true;
    el("history-search-toggle").setAttribute("aria-expanded", "false");
  }
  el("history-search-toggle").hidden = available.length < 3;
  const query = historyFilter.trim().toLocaleLowerCase();
  el("history-view").classList.toggle("multiple-tasks", available.length !== 1);
  const visible = available.filter(session => !query || [translateHistoryText(session.question), session.file, historyStatus(session.status), historyDate(session.created_at)].some(value => String(value ?? "").toLocaleLowerCase().includes(query)));
  el("history-status").classList.toggle("state-error", Boolean(historyError));
  const status = historyLoading ? "正在读取记录…" :
    historyError ? "记录列表加载失败，可点击刷新重试。" :
    historyWarning ? translateHistoryText(historyWarning).split(/[。！？]/, 1)[0] + "。" :
    !visible.length ? (query ? "没有匹配的记录。" : available.length ? "本窗口已隐藏全部记录。" : "还没有处理记录。") :
    historyHidden.size ? "部分记录已在本窗口隐藏。" : "";
  el("history-status").textContent = status;
  el("history-status").hidden = !status;
  visible.forEach((session) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "history-item";
    button.dataset.session = session.session_dir;
    button.classList.toggle("history-current", isCurrentSession(session));
    button.setAttribute("aria-pressed", String(session.session_dir === historySelected));
    const title = document.createElement("strong");
    title.className = "history-item-title";
    title.textContent = translateHistoryText(String(session.question || "未记录指令"));
    title.title = title.textContent;
    button.append(title);
    const stateLabels = { failed: "失败", error: "失败", pending: "待确认", awaiting_confirmation: "待确认",
      waiting_confirmation: "待确认", confirmation_required: "待确认", running: "处理中", idle: "未完成",
      cancelled: "未完成", stopped: "未完成", budget: "未完成", limit: "未完成", budget_exceeded: "未完成", max_iterations: "未完成" };
    const shortState = stateLabels[session.status] || (session.record_kind==='manual_copy'?session.format+' · '+historyStatus(session.status):null);
    if (shortState) {
      const state = document.createElement("span"); state.className = "history-item-state";
      state.dataset.attention = String(!["completed", "verified", "rolled_back"].includes(session.status));
      title.dataset.hasState = "true";
      state.textContent = shortState; button.append(state);
    }
    button.addEventListener("click", () => {
      resetHistoryTranscript();
      historySelected = session.session_dir;
      // 仅更新历史面板，不动当前 plan/result/progress/model 或确认状态。
      Array.from(list.children).forEach((item) => item.setAttribute("aria-pressed", String(item === button)));
      setNavigation("history");
      renderHistoryDetail();
      closeDrawer(false);
      focusVisible(el("history-reader"));
    });
    list.appendChild(button);
  });
  renderHistoryDetail();
  if (focusedSession && (document.activeElement === document.body || document.activeElement === list)) {
    const replacement = Array.from(list.children).find(item => item.dataset.session === focusedSession);
    focusVisible(replacement || el("btn-history-refresh"));
  }
}
function invalidateHistory() {
  resetHistoryTranscript();
  ++historyRequest;
  historyLoading = false;
  historySessions = []; historySelected = null; historyCurrentDir = null;
  historyWarning = ""; historyError = "";
  renderHistory();
}
async function refreshHistory() {
  resetHistoryTranscript();
  const request = ++historyRequest;
  if (historySelectingFile) return;
  historyLoading = true; historyError = "";
  renderHistory();
  try {
    if (!connected || !api || typeof api.list_sessions !== "function") throw new Error("当前服务未提供只读 list_sessions API");
    const response = await api.list_sessions();
    if (request !== historyRequest) return;
    if (notAuthenticated(response)) { historySessions = []; historySelected = null; historyCurrentDir = null; historyWarning = ""; return; }
    if (!response || response.ok !== true || !Array.isArray(response.sessions)) throw new Error(response && response.error ? String(response.error) : "历史列表返回格式无效");
    historySessions = response.sessions.filter((item) => item && typeof item.session_dir === "string");
    historyCurrentDir = typeof response.current_session_dir === "string" ? response.current_session_dir : null;
    historyWarning = String(response.warning || "");
  } catch (e) {
    if (request !== historyRequest) return;
    historySessions = []; historySelected = null; historyCurrentDir = null;
    historyError = e && e.message ? e.message : String(e);
  } finally {
    if (request === historyRequest) { historyLoading = false; renderHistory(); }
  }
}
async function hideHistorySession() {
  if (busy || requesting || modalOpen || historyLoading || !historySelected) return;
  const sessionDir = historySelected, request = historyRequest;
  const ok = await confirmModal("仅在本窗口隐藏", "只是不在列表里显示，不会删除任何记录或已生成的副本；重新打开窗口就会恢复。是否隐藏？");
  if (!ok || request !== historyRequest || sessionDir !== historySelected || busy || requesting || historySelectingFile) return;
  historyHidden.add(sessionDir); historySelected = null;
  renderHistory(); focusVisible(el("btn-history-refresh"));
}

/* ---------- 渲染事件 ---------- */
let answerDraft = null;
function clearAnswerDraft() {
  answerDraft?.node?.remove();
  answerDraft = null;
}
function answerDraftEvent(ev) {
  if (!/^[a-f0-9]{32}$/.test(ev.stream_id || "")) return;
  if (ev.type === "answer_draft_start") {
    clearAnswerDraft();
    if (!currentRound) currentRound = newFlowRound();
    const node = document.createElement("section"); node.className = "answer-draft";
    const label = document.createElement("strong"); label.textContent = "正在生成回复（尚未完成）";
    const content = document.createElement("div");
    node.append(label, content); flowRoundNode(currentRound).append(node);
    answerDraft = {id: ev.stream_id, node, content, text: ""};
  } else if (answerDraft?.id === ev.stream_id) {
    if (ev.type === "answer_draft_cancel") clearAnswerDraft();
    else if (ev.type === "answer_draft_delta" && typeof ev.text === "string") {
      if (answerDraft.text.length + ev.text.length > 40000) { clearAnswerDraft(); return; }
      answerDraft.text += ev.text; answerDraft.content.textContent = answerDraft.text;
    }
  }
}
function handleEvent(ev) {
  window.dispatchEvent(new CustomEvent("officedaily:event", {detail: ev}));
  if (["task", "answer", "idle", "file", "error"].includes(ev.type)) clearAnswerDraft();
  const previous = document.activeElement;
  switch (ev.type) {
    case "answer_draft_start": case "answer_draft_delta": case "answer_draft_cancel": answerDraftEvent(ev); break;
    case "log":
      // bridge 同时发送 task 与 user 日志；用户气泡只由受理 task 创建，避免重复。
      if (ev.tag !== "user") appendLog(ev.tag === "warn" && isTechnicalWarning(ev.text) ? "status" : ev.tag || "body", ev.text);
      // 仅用已有明确标记分层，不猜测任务终态；不落盘提示不被后续任务警告覆盖。
      if (ev.tag === "safety" || ev.tag === "warn" || String(ev.text || "").includes("不落盘")) {
        const warning = String(ev.text || "");
        if (warning.includes("不落盘")) fileWarning = warning;
        else if (ev.tag === "safety") taskSafety = warning;
        else taskWarning = isTechnicalWarning(warning) ? "有处理状态提醒，展开「运行详情」可以看到。" : warning;
        renderWarnings();
      }
      break;
    case "status": {
      const source=statusFeedback?.source;
      const sameTurn=source?.channel==='chat'&&source.session_id===ev.session_id&&source.turn_id===ev.turn_id;
      if(ev.channel==='chat'&&window.OfficeDaily?.completed?.(ev.session_id,ev.turn_id)&&statusFeedback&&!sameTurn)break;
      setStatus(ev.level || "info", ev.text, ev.channel === 'chat' ? {channel:'chat',session_id:ev.session_id,turn_id:ev.turn_id} : null); appendLog(ev.level === "danger" ? "error" : ev.level === "warn" && !isTechnicalWarning(ev.text) ? "warn" : "status", ev.text); break;
    }
    case "badge": setBadge(ev.level || "info", ev.label || ""); break;
    case "progress": setProgress(ev); appendLog("progress", "处理进度：" + JSON.stringify(ev)); break;
    case "answer": if(!setMainReply(ev.text))appendLog("answer", ev.text); updateRoundState("模型已回复"); break;
    case "plan":
      archiveFlowCard("plan-card");
      planVersion++;
      pending = true; verified = false;
      setBadge("warn", "● 还没写入 · 等你确认");
      el("plan-title").textContent = "修改方案（还没写入）";
      el("plan-pane").textContent = ev.text;
      el("plan-pane").classList.remove("empty-state");
      el("plan-card").classList.add("plan-pending");
      el("plan-pane").scrollTop = 0;
      enqueueCard("plan-card", ev.text);
      refreshButtons();
      break;
    case "output":
      archiveFlowCard("plan-card");
      planVersion++;
      pending = false; verified = !ev.undone;
      setBadge(ev.undone ? "warn" : "ok", ev.undone ? "已撤销副本" : "副本已生成");
      el("plan-title").textContent = "修改结果";
      el("plan-pane").textContent = ev.text;
      el("plan-pane").classList.remove("empty-state");
      el("plan-card").classList.remove("plan-pending");
      el("plan-pane").scrollTop = 0;
      enqueueCard("plan-card", ev.text);
      refreshButtons();
      break;
    case "diagram":
      appendDiagramMessage(ev);
      setStatus("info", "图已经画好了（还没存成文件）；想留下它，请点下面的保存按钮。");
      setBadge("info", "图已生成");
      break;
    case "result":
      archiveFlowCard("result-card");
      renderResult(ev);
      el("result-pane").scrollTop = 0;
      enqueueCard("result-card", resultCopyText(ev, resultSummary));
      break;
    case "checkpoint":
      el("checkpoint").textContent = ev.text;
      break;
    case "file":
      el("vision-list").replaceChildren(); el("vision-list").hidden = true;
      el("vision-document").textContent = "分析文档图片…";
      invalidateClarification();
      window.OfficeImageGen?.retirePending("已更换文件；旧生图操作失效，请重新规划。");
      releaseDiagramURLs();
      currentRound = null; authorizedFileName = String(ev.name || "未说明文件");
      archiveFlowCard("plan-card");
      archiveFlowCard("result-card");
      appendLog("meta", "已更换为：" + ev.name + "；此前的卡片仅供回看。");
      // 后端成功换文件会清除计划和输出身份；前端同步，取消选择不走此分支。
      pending = false; verified = false;
      hasFile = true;
      el("file-empty").hidden = true;
      el("file-options").open = false;
      el("plan-card").classList.remove("plan-pending");
      el("plan-title").textContent = "修改计划 / 输出";
      el("plan-pane").textContent = "这个文件还没有修改方案。保存前会再问你一次，原文件不会被动。";
      el("plan-pane").classList.add("empty-state");
      resultTable = null; resultText = ""; resultPage = 0; resultEvent = null; resultSummary = "";
      el("result-controls").hidden = true;
      el("btn-copy-tsv").disabled = true;
      el("result-pane").textContent = "当前文件尚无读取结果。输入任务后点击发送。";
      el("result-pane").classList.add("empty-state");
      el("checkpoint").textContent = "这个文件还没有进度记录；暂时不能接着跑。";
      fileWarning = ""; taskWarning = ""; taskSafety = "";
      planVersion++;
      renderWarnings();
      setBadge("info", "已选择文档");
      setStatus("info", "可读取内容或预览修改；保存副本需单独确认。原文件保留。");
      selectPanel("result");
      refreshButtons();
      el("file-name").textContent = "当前文件：" + ev.name;
      el("details-file").textContent = "当前文件：" + ev.name;
      el("file-dir").textContent = "当前文件（只对这一个文件生效）：" + ev.dir;
      appendLog("meta", "文件初检：\n" + toolSummary({ tool: "inspect", data: ev.inspect }));
      invalidateHistory();
      if (!historySelectingFile) refreshHistory();
      break;
    case "task":
      stopRequested = false;
      decisionLimit = null;
      el("iter-fill").parentElement.hidden = true;
      invalidateClarification();
      window.OfficeImageGen?.retirePending("新任务已受理；旧生图操作失效，请重新规划。");
      const echoed=localMainRound?.question===ev.text?localMainRound:null;
      currentRound = echoed||newFlowRound(ev.text);
      localMainRound=null;
      if(echoed)setMainReply('我会先核对当前文档和处理范围，再按你的要求执行。');
      setStatus("info", ev.channel==='chat'?"正在查看本轮问题。":"已收到，正在处理。写入前还会再问你一次。", ev.channel==='chat'?{channel:'chat',session_id:ev.session_id,turn_id:ev.turn_id}:null);
      setText(el("error-detail"), "当前任务尚未收到错误");
      setBadge("info", "已受理指令");
      el("checkpoint").textContent = "这次任务还没有进度记录；暂时不能接着跑。";
      el("iter-value").textContent = "尚未收到"; el("budget-value").textContent = "尚未收到";
      el("iter-fill").style.width = "0%"; el("budget-fill").style.width = "0%";
      el("budget-fill").className = "fill";
      planVersion++;
      if (pending) {
        pending = false; // bridge.ask 已使 pending_id 失效，旧预览仅供查看。
        el("plan-card").classList.remove("plan-pending");
        el("plan-title").textContent = "此前修改计划（新任务已开始，不能确认）";
      }
      taskWarning = ""; taskSafety = ""; // 旧任务提示仍保留在窗口记录中。
      renderWarnings();
      refreshButtons();
      if(!echoed)appendLog("user-message", ev.text);
      refreshHistory();
      break;
    case "busy":
      busy = ev.busy;
      if (!busy) stopRequested = false;
      el("busy-pill").style.display = busy ? "" : "none";
      el("busy-label").textContent = busy ? "正在处理模型任务" : "就绪";
      refreshButtons();
      if (!busy) refreshHistory();
      if (!busy) window.dispatchEvent(new Event('officemodels:request-finished'));
      break;
    case "clarify":
      renderClarification(ev);
      break;
    case "read_confirm":
      handleReadConfirm(ev);
      break;
    // P3 生图：生成 / 保存 / 撤销是三个互不代用的独立确认，内联在对话轮次里。
    case "image_generation_started":
      setBadge("info", "正在生成图片"); setStatus("info", ev.text); break;
    case "image_plan":
    case "image_result":
    case "image_save_plan":
    case "image_saved":
    case "image_rollback_plan":
    case "image_rolled_back":
      window.OfficeImageGen?.handleEvent(ev, (name, ...args) =>
        (api && typeof api[name] === "function") ? api[name](...args) : Promise.resolve({ok: false, error: "not_authenticated"}),
        mountImageMessage);
      break;
    // 检测结果只在设置窗口里回显一行结论；详细报告仍留在对话流的运行记录里。
    case "office_probe":
      officeProbing = ev.state === "running";
      settingsProbeStatus(ev.state, ev.state === "running" ? "正在检测本机 Office 环境…" :
        ev.state === "complete" ? "检测完成；可用性与限制请查看本次检测报告。" : "检测失败；请查看检测报告后处理。");
      refreshButtons();
      break;
    case "text":
      if (requestLabel.includes("检测")) settingsProbeStatus("", "检测完成：" + String(ev.text || "").slice(0, 300));
      break;
    default:
      break;
  }
  repairFocus(previous);
  window.dispatchEvent(new CustomEvent('officeapp:state', {detail:{hasFile, filename:authorizedFileName, busy:busy || requesting, pending, verified}}));
}

// 事件处理仍同步、有序；仅日志 DOM 延迟到下一帧，每帧最多创建 250 条。
const logLimit = 2000;
const logBatchSize = 250;
let logQueue = [];
let logQueueHead = 0;
let logRecords = [];
let logFolded = 0;
let logFrame = null;
let logFollowing = true;
let logScrollTarget = null;
let logPlaceholder = null;

function appendLog(tag, text, node = null) {
  if (!currentRound && tag !== "user-message") currentRound = newFlowRound();
  const trace = !node && tag !== "user-message" && !["warn", "error", "safety", "answer"].includes(tag) && !String(text).includes("不落盘");
  logQueue.push({ tag, text: String(text ?? ""), node, round: currentRound, trace });
  refreshEmptyFlow();
  if (logFrame === null) logFrame = window.requestAnimationFrame(flushLogs);
}

/* 图表在对话流内联渲染：SQL 源码折起，预览只在内存 blob 中，不落盘。 */
const diagramObjectURLs = [];
function releaseDiagramURLs() {
  while (diagramObjectURLs.length) URL.revokeObjectURL(diagramObjectURLs.pop());
}
function appendDiagramMessage(ev) {
  const svg = typeof ev.svg === "string" ? ev.svg : "";
  const source = typeof ev.source === "string" ? ev.source : "";
  if (!svg || !/^<svg\s/.test(svg) || svg.length > 1_100_000) {
    appendLog("warn", "图表结果格式异常，已拒绝显示。");
    return;
  }
  if (!currentRound) currentRound = newFlowRound();
  const card = document.createElement("section");
  card.className = "diagram-message";
  const title = document.createElement("div");
  title.className = "diagram-message-title";
  const labels = { flowchart: "流程图", classDiagram: "类图", sequenceDiagram: "时序图",
    stateDiagramV2: "状态图", stateDiagram: "状态图", erDiagram: "实体关系图" };
  title.textContent = (labels[ev.kind] || "图表") + " · 已生成（尚未保存文件）";
  card.append(title);
  const figure = document.createElement("img");
  figure.className = "diagram-figure";
  figure.alt = "本地渲染的图表预览；语义需人工核对";
  // <img> 加载 SVG 不会执行其中脚本；blob 仅在内存，登录/换文件时回收。
  const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
  diagramObjectURLs.push(url);
  figure.src = url;
  figure.addEventListener("error", () => {
    figure.remove();
    const note = document.createElement("p");
    note.className = "diagram-note";
    note.textContent = "图表图片未能显示，请重试。";
    card.append(note);
  });
  card.append(figure);
  if (source) {
    const details = document.createElement("details");
    details.className = "diagram-source-details";
    const summary = document.createElement("summary");
    summary.textContent = "查看图表源码";
    const pre = document.createElement("pre");
    pre.className = "diagram-source-text";
    pre.textContent = source;
    details.append(summary, pre);
    card.append(details);
  }
  const actions = document.createElement("div");
  actions.className = "diagram-message-actions";
  const copy = document.createElement("button");
  copy.type = "button"; copy.className = "btn link"; copy.textContent = "复制源码";
  copy.disabled = !source;
  copy.addEventListener("click", async () => {
    if (!source) return;
    try { await navigator.clipboard.writeText(source); copy.textContent = "已复制"; }
    catch (_) { copy.textContent = "复制失败，请手动选择"; }
  });
  actions.append(copy);
  card.append(actions);
  // 保存 / 撤销按钮跟着这条消息走；新建图表时旧图自动失去导出权限。
  const exportActions = document.createElement("div");
  exportActions.className = "diagram-message-actions";
  card.append(exportActions);
  appendLog("diagram", "", card);
  window.OfficeDiagramExport?.mount?.(exportActions);
  // 交给导出模块：保存 SVG 副本仍需用户独立确认，这里只登记候选元数据。
  window.OfficeDiagramPreview?.ready?.(ev.metadata);
}

/* 生图结果同样内联在对话流里：每张消息自带生成/保存/撤销确认，旧消息不再有权限。
   mount 只在轮到该消息时才把它接进当前轮次，避免一次性灌入大量节点。 */
function mountImageMessage(card) {
  if (!currentRound) currentRound = newFlowRound();
  appendLog("imagegen", "", card);
}

function setLogFollowing(following) {
  logFollowing = following;
  el("btn-log-latest").hidden = following;
}
function scrollLogTo(top) {
  const pane = el("log-pane");
  pane.scrollTop = top;
  logScrollTarget = pane.scrollTop;
}
function flushLogs() {
  logFrame = null;
  const pane = el("log-pane");
  const count = Math.min(logBatchSize, logQueue.length - logQueueHead);
  if (!count) return;
  const removeCount = Math.max(0, logRecords.length + count - logLimit);
  // 待确认计划/可撤销输出不能被日志洪峰回收；仍计入 2000 条总上限。
  const candidates = logRecords.filter(record => !(record.node === el("plan-card") && (pending || verified)));
  const removed = new Set(candidates.slice(0, removeCount));
  const anchor = logRecords.filter(record => !removed.has(record)).map(record => {
    const trace = record.node?.closest(".processing-trace");
    return trace && !trace.open ? trace.querySelector("summary") : record.node;
  }).find(node => node?.isConnected && node.getClientRects().length &&
    node.getBoundingClientRect().bottom >= pane.getBoundingClientRect().top &&
    node.getBoundingClientRect().top <= pane.getBoundingClientRect().bottom);
  const anchorTop = !logFollowing && anchor ? anchor.getBoundingClientRect().top : null;
  const oldTop = pane.scrollTop;
  const previousFocus = document.activeElement;
  removed.forEach(removeFlowRecord);
  logRecords = logRecords.filter(record => !removed.has(record));
  logFolded += removed.size;
  if (logFolded) {
    if (!logPlaceholder) {
      logPlaceholder = document.createElement("div");
      logPlaceholder.className = "log-folded";
      pane.prepend(logPlaceholder);
    }
    logPlaceholder.textContent = "已折叠较早记录 " + logFolded + " 条（保留至多 " + logLimit + " 条，当前待确认／可撤销安全卡片优先保留；不一定全是最近记录）";
  }
  for (let i = 0; i < count; i++) {
    const record = logQueue[logQueueHead];
    logQueue[logQueueHead++] = null; // 已消费槽位释放文本引用，避免追赶期间保留折叠记录。
    appendFlowRecord(record);
    logRecords.push(record);
  }
  if (logFollowing) scrollLogTo(pane.scrollHeight);
  else if (anchorTop !== null) scrollLogTo(oldTop + anchor.getBoundingClientRect().top - anchorTop);
  else scrollLogTo(0);
  if (logQueueHead < logQueue.length) logFrame = window.requestAnimationFrame(flushLogs);
  else { logQueue = []; logQueueHead = 0; }
  refreshEmptyFlow();
  repairFocus(previousFocus);
}

function resetLogs() {
  clearAnswerDraft();
  invalidateClarification();
  if (logFrame !== null) window.cancelAnimationFrame(logFrame);
  logFrame = null;
  logQueue = []; logQueueHead = 0; logRecords = []; logFolded = 0;
  logPlaceholder = null;
  ["plan-card", "result-card"].forEach(id => el("card-store").append(el(id)));
  el("log-pane").textContent = "";
  currentRound = null;
  setLogFollowing(true);
  scrollLogTo(0);
  // 清空显示不是重置后端：当前安全身份与可操作卡片必须保留。
  if (pending || verified) enqueueCard("plan-card", el("plan-pane").textContent);
  refreshEmptyFlow();
}
function copyLogs() {
  // 与回收一致：包含 queued，最多 2000 条且有效安全卡片优先；不是全审计。
  const all = logRecords.concat(logQueue.slice(logQueueHead)).filter(Boolean);
  let overflow = Math.max(0, all.length - logLimit);
  const records = all.filter(record => {
    if (overflow && !(record.node === el("plan-card") && (pending || verified))) { overflow--; return false; }
    return true;
  });
  const folded = logFolded + all.length - records.length;
  const prefix = folded ? "已回收较早记录 " + folded + " 条；以下为保留的 " + records.length + " 条窗口记录（安全卡片优先，未必全是最新记录）。\n" : "";
  return copyText(prefix + records.map((record) => record.text).join("\n"), "这个窗口里的记录（最多 " + logLimit + " 条，不是完整存档）");
}

el("log-pane").addEventListener("scroll", () => {
  const pane = el("log-pane");
  if (logScrollTarget !== null && Math.abs(pane.scrollTop - logScrollTarget) < 1) {
    logScrollTarget = null;
    return;
  }
  logScrollTarget = null;
  setLogFollowing(pane.scrollHeight - pane.clientHeight - pane.scrollTop <= 8);
}, { passive: true });
el("btn-log-latest").addEventListener("click", () => {
  setLogFollowing(true);
  scrollLogTo(el("log-pane").scrollHeight);
  el("log-pane").focus({ preventScroll: true });
});

function isTechnicalWarning(text) {
  return /^状态\s/.test(String(text || "")) && /tokens|token|预算/i.test(String(text)) &&
    !/不落盘|未提交|未写入|确认|授权|拒绝|失败|风险|保存|撤销/.test(String(text));
}
function humanErrorNotice(text) {
  const raw = String(text ?? "");
  if (/image_configuration_(?:missing|invalid)/.test(raw)) {
    return "云端生图没有配置好，本次没有请求生图服务。请检查项目配置文件中的生图提供方、密钥、地址和模型，重启后再试；其他步骤的结果仍需单独核对。原始信息可在下方展开。";
  }
  let reason = "这次没有做完。";
  if (/timeout|timed out|超时/i.test(raw)) reason = "等太久了，请检查服务连接后再试一次。";
  else if (/401|403|unauthorized|authentication|凭据|密钥/i.test(raw)) reason = "登录或密钥有问题，请核对配置，或重新登录试试。";
  else if (/permission|access denied|拒绝访问|权限/i.test(raw)) reason = "这个文件或服务打不开，可能是权限不够。";
  else if (/connect|network|连接|网络/i.test(raw)) reason = "暂时连不上服务，请检查网络或本机服务是否在运行。";
  return reason + " 光看这条报错无法判断有没有写入成功；请查看当前的计划、输出和记录，不要重复确认。展开下面可以看到原始信息。";
}
function appendArchivedError(parent, text) {
  const notice = document.createElement("p"); notice.textContent = humanErrorNotice(text);
  const details = document.createElement("details"); details.className = "error-source";
  const title = document.createElement("summary"); title.textContent = "原始报错信息";
  const raw = document.createElement("pre"); raw.textContent = String(text ?? "");
  details.append(title, raw); parent.append(notice, details);
}
function setStatus(level, text, source = null) {
  statusFeedback = {source};
  // 错误只由 alert 播报，避免 status 与 alert 同时复述同一失败。
  el("status-line").setAttribute("aria-live", "off");
  setText(el("status-line"), level === "danger" ? humanErrorNotice(text) : text);
  el("status-line").style.color =
    level === "danger" ? "var(--danger)" :
    level === "warn" ? "var(--warn)" :
    level === "ok" ? "var(--ok)" : "var(--ink)";
  el("status-card").classList.toggle("state-error", level === "danger");
  el("status-card").classList.toggle("state-warning", level === "warn");
  const technicalWarning = level === "warn" && isTechnicalWarning(text);
  const warning = level === "warn" ? (technicalWarning ? "有处理状态提醒，展开「运行详情」可以看到；写入前仍会再问你一次。" : text) : "";
  const notice = level === "danger" ? humanErrorNotice(text) : warning;
  setText(el("error-line"), notice);
  if (level === "danger" || level === "warn") setText(el("error-detail"), text);
  el("error-line").style.display = notice ? "" : "none";
}

window.OfficeRequestFeedback = Object.freeze({
  snapshot: () => statusFeedback,
  resolveBudget: previous => {
    if (!previous || previous !== statusFeedback || previous.source !== 'chat-budget') return false;
    setStatus('info', '当前输入的发送范围已重新核对；尚未调用模型。');
    setText(el('error-detail'), '');
    return true;
  }
});

function setBadge(level, label) {
  const pill = el("badge");
  pill.className = "pill " + (["info", "ok", "warn", "danger"].includes(level) ? level : "info");
  el("badge-label").textContent = String(label).replace(/^●\s*/, "");
  updateRoundState(label);
}

function setProgress(ev) {
  const number = value => value !== null && value !== undefined && value !== "" && Number.isFinite(Number(value)) ? Number(value) : null;
  const iter = number(ev.iterations), cap = number(ev.capacity), tok = number(ev.tokens);
  const limit = number(ev.decision_limit);
  decisionLimit = limit !== null && Number.isInteger(limit) && limit >= 1 && limit <= 64 ? limit : null;
  el("iter-value").textContent = iter === null ? "未提供" : "第 " + iter + " 步" + (decisionLimit ? "（最多 " + decisionLimit + " 步）" : "");
  el("iter-fill").parentElement.hidden = decisionLimit === null;
  el("iter-fill").style.width = iter === null || decisionLimit === null ? "0%" : Math.max(0, Math.min(100, (iter / decisionLimit) * 100)) + "%";
  const fill = el("budget-fill");
  fill.style.width = "0%"; fill.className = "fill";
  el("budget-value").textContent = "未提供";
  if (cap !== null && cap > 0 && tok !== null) {
    const frac = Math.max(0, Math.min(1, tok / cap));
    el("budget-value").textContent = "已用 " + Math.round(frac * 100) + "%";
    const fill = el("budget-fill");
    fill.style.width = Math.round(frac * 100) + "%";
    fill.className = "fill" + (frac >= 0.9 ? " danger" : frac >= 0.8 ? " warn" : "");
  }
}

function invalidateClarification(expected = activeClarification) {
  if (!activeClarification || activeClarification !== expected) return;
  activeClarification.card.querySelectorAll("button, input").forEach(control => { control.disabled = true; });
  activeClarification = null;
}
function refreshClarification() {
  if (!activeClarification) return;
  const disabled = !connected || busy || requesting || modalOpen || currentView === "history";
  activeClarification.card.querySelectorAll("button, input").forEach(control => { control.disabled = disabled; });
}
function renderClarification(ev) {
  invalidateClarification();
  if (!Array.isArray(ev.questions) || ev.questions.length < 1 || ev.questions.length > 3) return;
  const card = document.createElement("section"); card.className = "clarification-card";
  const title = document.createElement("h3"); title.textContent = "动手前确认需求"; card.append(title);
  const preamble = document.createElement("p"); preamble.textContent = ev.preamble; card.append(preamble);
  const answers = {}, fields = [];
  for (const question of ev.questions) {
    const group = document.createElement("fieldset");
    const legend = document.createElement("legend"); legend.textContent = question.question; group.append(legend);
    const chips = [];
    for (const option of question.options || []) {
      const chip = document.createElement("button"); chip.type = "button"; chip.textContent = option.label;
      chip.setAttribute("aria-pressed", "false");
      chip.addEventListener("click", () => {
        answers[question.id] = option.value;
        chips.forEach(other => other.setAttribute("aria-pressed", String(other === chip)));
      });
      chips.push(chip); group.append(chip);
    }
    if (question.allow_free_text) {
      const label = document.createElement("label"); label.textContent = "补充说明（" + question.question + "）：";
      const input = document.createElement("input"); input.type = "text"; input.maxLength = 500;
      label.append(input); group.append(label); fields.push([question.id, input]);
    }
    const defaults = document.createElement("p"); defaults.textContent = "跳过或未填写时：" + question.default_label;
    group.append(defaults); card.append(group);
  }
  const message = document.createElement("p"); message.setAttribute("role", "status"); card.append(message);
  const identity = { card, requestId: ev.request_id }; activeClarification = identity;
  async function submit(skipped) {
    if (activeClarification !== identity || !connected || busy || requesting || modalOpen || currentView === "history") return;
    const values = skipped ? {} : { ...answers };
    if (!skipped) fields.forEach(([id, input]) => { if (input.value.trim()) values[id] = input.value.trim(); });
    requesting = true; requestLabel = "提交需求信息（不是批准）"; refreshButtons();
    try {
      const result = await api.answer_clarification(identity.requestId, values, skipped);
      if (result.ok) { message.textContent = "收到你的回答，正在规划；写入前还会再问你一次。"; invalidateClarification(identity); }
      else { message.textContent = result.error || "追问已失效或无法提交，请重新发送任务。"; if (result.reason === "stale_clarification") invalidateClarification(identity); }
    } catch (_) { message.textContent = "提交未完成，请核对处理记录后重试。"; }
    finally { requesting = false; requestLabel = ""; refreshButtons(); }
  }
  const next = document.createElement("button"); next.type = "button"; next.textContent = "按回答继续规划";
  next.addEventListener("click", () => submit(false)); card.append(next);
  const skip = document.createElement("button"); skip.type = "button"; skip.textContent = ev.skip_label;
  skip.addEventListener("click", () => submit(true)); card.append(skip);
  card.querySelectorAll("button").forEach(button => { button.className = "btn secondary"; });
  appendLog("clarification", ev.preamble, card); setStatus("info", "请补充需求或按显示的默认值跳过；这不是执行批准。");
  refreshClarification();
}

function refreshButtons() {
  refreshClarification();
  const previous = document.activeElement;
  el("btn-history-hide").disabled = busy || requesting || modalOpen || historyLoading || !historySelected;
  el("btn-history-refresh").disabled = !connected || historyLoading || historySelectingFile || modalOpen;
  const working = busy || requesting;
  const blocked = !connected || working || modalOpen;
  el("btn-confirm").disabled = blocked || !pending;
  el("btn-confirm").hidden = !pending;
  el("btn-undo").disabled = blocked || !verified;
  el("btn-undo").hidden = !verified;
  el("pending-notice").hidden = !pending;
  el("btn-ask").disabled = blocked;
  el("vision-local").disabled = blocked;
  el("vision-document").disabled = blocked || !hasFile;
  el("vision-list").disabled = blocked || !hasFile;
  el("btn-resume").disabled = true;
  el("btn-select2").textContent = hasFile ? "换一个文件…" : "+ 选择文档";
  el("prompt").disabled = modalOpen || (working && !window.OfficeWorkbench);
  el("processing-message").hidden = !working;
  document.querySelectorAll("[data-prompt]").forEach(button => { button.disabled = working || modalOpen; });
  ["btn-select", "btn-select2", "btn-path-select"].forEach((id) => { el(id).disabled = blocked; });
  el("file-path").disabled = blocked;
  el("file-drop").setAttribute("aria-disabled", String(blocked));
  if (blocked) el("file-drop").classList.remove("drag-active");
  el("recent-files").querySelectorAll("button").forEach((button) => { button.disabled = blocked; });
  ["btn-probe", "settings-probe"].forEach((id) => { el(id).disabled = blocked || officeProbing; });
  el("btn-ask").textContent = working ? "处理中…" : "发送";
  el("btn-stop").hidden = !busy;
  el("btn-stop").disabled = !connected || stopRequested;
  el("busy-pill").style.display = working ? "" : "none";
  el("busy-label").textContent = requesting ? requestLabel : decisionLimit ? "正在处理（最多 " + decisionLimit + " 步）" : "正在处理";
  el("content-panes").setAttribute("aria-busy", String(working));
  el("next-step").textContent = !connected ? "正在连接本机服务…" :
    working ? "请等这一步结束；已有内容可以照常查看。" :
    !hasFile ? "下一步：选择一个文档。" :
    pending ? "下一步：看完整计划，再确认是否保存副本。" :
    verified ? "副本已保存；可以查看它在哪里，或继续下一个任务。" :
    "下一步：写下你要做什么，然后发送。";
  const activity = el("activity");
  activity.classList.toggle("is-working", working);
  setText(activity, !connected ? "正在连接本机服务；连上后就可以选文件了。" :
    working ? (requesting ? requestLabel : "正在处理；已有的计划和读取结果会保留，请稍等。") :
    !hasFile ? "还没有打开文件：请先点「选择文档」，再写下要求。" :
    pending ? "有一份计划在等你确认：读取结果不会把它覆盖；确认后才会写入。" :
    "就绪：可以读取内容或预览修改；真正写入前还需要你确认。");
  document.querySelectorAll(".message-actions button, #btn-copy-log, #btn-clear, #btn-copy-plan, #btn-copy-result").forEach(button => { button.disabled = working || modalOpen; });
  el("btn-copy-tsv").disabled = working || modalOpen || !resultTable;
  if (drawerOpen && !canFocus(previous)) focusVisible(el("sidebar-toggle"));
  repairFocus(previous);
}

/* 仅封装已有 API 请求：显示加载/错误并阻止重复提交，不改变后端闸门。 */
async function requestAction(label, action, argument = null) {
  if (!connected || busy || requesting || modalOpen) return;
  if (currentView === "history" && action !== selectFile) return;
  if ((action === confirmWrite && !pending) || (action === undo && !verified)) return;
  actionFocus = document.activeElement;
  requesting = true; requestLabel = label;
  refreshButtons();
  try { await action(argument); }
  catch (e) {
    const text = "操作未完成：" + (e && e.message ? e.message : String(e)) + "。请查看记录并重试；不要重复确认写入。";
    setStatus("danger", text);
    appendLog("error", text);
  } finally {
    requesting = false; requestLabel = "";
    refreshButtons();
    if (deferredModalFocus && document.activeElement === deferredModalFocus.interim)
      focusVisible(deferredModalFocus.target);
    deferredModalFocus = null; actionFocus = null;
  }
}

/* 当前文件的一次读取许可；写入确认和批量/真实Office读取不共享许可。 */
async function handleReadConfirm(ev) {
  const session = ev.scope === "selected_file_session";
  const policy = session
    ? "同意后，这次会话里可以反复读取当前这个文件，不再每次都问。\n换文件或重启应用后失效；不会顺带允许读别的文件、调用真实 Office 或写入。\n读取结果和图片可能留在这个文件旁边的只读记录里；图片内容不会发给模型，也不会写回原文件。"
    : "只允许下面这一次读取，不会顺带扩大别的范围。";
  const body = policy + "\n\n要执行的读取：" + ev.tool + "\n" +
    JSON.stringify(ev.args || {}, null, 2) + "\n\n是否允许？";
  const ok = await confirmModal(session ? "允许本次会话反复读取此文件" : "确认这一次读取", body);
  await api.answer_read(ok, ev.request_id || null);
}

/* ---------- 确认模态（写操作前，默认焦点在取消） ---------- */
function confirmModal(title, body) {
  if (modalOpen) return Promise.resolve(false);
  return new Promise((resolve) => {
    const previousFocus = actionFocus || document.activeElement;
    // 保留抽屉身份：workspace inert 覆盖主区和侧栏，取消可回原历史入口。
    modalOpen = true;
    el("workspace").inert = true;
    el("modal-title").textContent = title;
    el("modal-body").textContent = body;
    el("modal").style.display = "grid";
    refreshButtons();
    el("modal-cancel").focus();
    const ok = () => { close(false); resolve(true); };
    const cancel = () => { close(true); resolve(false); };
    function onKey(e) {
      if (e.key === "Escape") {
        e.preventDefault(); e.stopImmediatePropagation();
        close(true); resolve(false);
      } else if (e.key === "Tab") {
        e.preventDefault();
        (document.activeElement === el("modal-cancel") ? el("modal-ok") : el("modal-cancel")).focus();
      } else if (e.ctrlKey && e.key === "Enter") {
        e.preventDefault(); e.stopImmediatePropagation();
      }
    }
    function close(cancelled) {
      el("modal").style.display = "none";
      modalOpen = false;
      el("workspace").inert = false;
      el("modal-ok").removeEventListener("click", ok);
      el("modal-cancel").removeEventListener("click", cancel);
      document.removeEventListener("keydown", onKey, true);
      refreshButtons();
      focusVisible(canFocus(previousFocus) ? previousFocus : focusFallback(previousFocus));
      if (cancelled && requesting)
        deferredModalFocus = { target: previousFocus, interim: document.activeElement };
    }
    document.addEventListener("keydown", onKey, true);
    el("modal-ok").addEventListener("click", ok);
    el("modal-cancel").addEventListener("click", cancel);
  });
}

/* F8：候选只用于提示，系统文件框才是真正打开文件的入口。存储内容永不重放为 file 事件。 */
const recentFileKey = "office-agent.recent-files.v1";
let recentFiles = [];
function fileEntryStatus(text) { el("file-entry-status").textContent = text; }
function validOfficeName(value) { return /\.(docx|xlsx|doc|xls)$/i.test(value); }
function pathCandidate(raw) {
  const value = String(raw || "").trim().replace(/^"([^"]+)"$/, "$1");
  if (!value || value.length > 4096 || /[\x00-\x1f<>|?*]/.test(value) ||
      !(/^[A-Za-z]:[\\/]/.test(value) || /^\\\\[^\\/]+[\\/][^\\/]+[\\/]/.test(value)) ||
      /(^|[\\/])\.{1,2}([\\/]|$)/.test(value) ||
      value.slice(2).includes(":" ) || !validOfficeName(value)) return null;
  return value;
}
function saveRecentFiles() {
  try { localStorage.setItem(recentFileKey, JSON.stringify(recentFiles)); }
  catch (e) { fileEntryStatus("最近记录存不下来，只能保留在本次窗口里；选文件的流程不受影响。"); }
}
function renderRecentFiles() {
  const host = el("recent-files");
  host.replaceChildren();
  if (!recentFiles.length) { host.textContent = "还没有最近打开过的文件。这里只记路径，不保存文档内容。"; return; }
  recentFiles.forEach((path) => {
    const button = document.createElement("button");
    button.type = "button"; button.className = "btn secondary";
    button.textContent = path + " · 重新打开…";
    button.disabled = !connected || busy || requesting || modalOpen;
    button.addEventListener("click", () => authorizeCandidate(path));
    host.appendChild(button);
  });
}
function rememberAuthorizedFile(result) {
  if (!result || !result.ok || typeof result.dir !== "string" || typeof result.name !== "string") return;
  const path = pathCandidate(result.dir.replace(/[\\/]$/, "") + "\\" + result.name);
  if (!path) return;
  recentFiles = [path, ...recentFiles.filter((item) => item.toLowerCase() !== path.toLowerCase())].slice(0, 8);
  saveRecentFiles(); renderRecentFiles();
}
function initFileEntries() {
  try {
    const saved = localStorage.getItem(recentFileKey);
    const rows = saved && saved.length <= 40000 ? JSON.parse(saved) : [];
    recentFiles = Array.isArray(rows) ? [...new Set(rows.filter((item) => typeof item === "string" && pathCandidate(item)).map(pathCandidate))].slice(0, 8) : [];
  } catch (e) { recentFiles = []; }
  renderRecentFiles();
}
function authorizeCandidate(raw, nameOnly = false) {
  if (!connected || busy || requesting || modalOpen) {
    fileEntryStatus("还没有连上本机服务，或正在处理上一个任务；稍后再试。"); return;
  }
  const candidate = nameOnly ? String(raw || "") : pathCandidate(raw);
  if (!candidate || (nameOnly && (!validOfficeName(candidate) || /[\\/\x00-\x1f]/.test(candidate)))) {
    fileEntryStatus("这个路径用不了：请提供一个本机完整路径，且以 .docx/.xlsx/.doc/.xls 结尾；不能是网址或相对路径。"); return;
  }
  el("file-path").value = nameOnly ? "" : candidate;
  fileEntryStatus("准备打开：" + candidate + "。这一步只是填表，还没真正打开；请在系统文件框里确认，取消不会改变现在打开的文件。");
  // 不向 bridge 传入路径，亦不手动设置 hasFile/path；只有 select_file 的结果有效。
  requestAction("正在打开文件…", selectFile, candidate);
}
function bindFileEntries() {
  const drop = el("file-drop");
  const choose = () => requestAction("正在打开文件…", selectFile);
  drop.addEventListener("click", choose);
  drop.addEventListener("keydown", (e) => {
    if ((e.key === "Enter" || e.key === " ") && !e.isComposing) { e.preventDefault(); e.stopPropagation(); choose(); }
  });
  el("btn-path-select").addEventListener("click", () => authorizeCandidate(el("file-path").value));
  el("file-path").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); e.stopPropagation(); authorizeCandidate(e.target.value); }
  });
  el("btn-recent-clear").addEventListener("click", () => {
    recentFiles = []; saveRecentFiles(); renderRecentFiles();
    fileEntryStatus("已清空最近路径列表；当前打开的文件和记录都不受影响。");
  });
  // 阻止文件拖拽让WebView导航离开本应用；文本粘贴不受影响。
  ["dragover", "drop"].forEach((type) => document.addEventListener(type, (e) => {
    if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files")) e.preventDefault();
  }));
  document.addEventListener("dragenter", (e) => {
    if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files")) {
      e.preventDefault(); if (connected && !busy && !requesting && !modalOpen) drop.classList.add("drag-active");
    }
  });
  document.addEventListener("dragleave", (e) => {
    if (!e.relatedTarget) drop.classList.remove("drag-active");
  });
  document.addEventListener("drop", (e) => {
    if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes("Files")) return;
    e.preventDefault(); drop.classList.remove("drag-active");
    const files = Array.from(e.dataTransfer ? e.dataTransfer.files || [] : []);
    if (files.length !== 1) { fileEntryStatus("一次只能拖入一个 Office 文档；文件夹和多个文件都不行。"); return; }
    authorizeCandidate(files[0].name, true); // 浏览器 File.name 不是本地路径。
  });
}
initFileEntries();

/* ---------- 动作 ---------- */
async function selectFile(candidate = null) {
  window.dispatchEvent(new Event("officediagram:invalidate"));
  historySelectingFile = true;
  invalidateHistory(); // 换文件前即作废旧响应，不能让旧文件列表回流。
  try {
    const r = await api.select_file();
    if (notAuthenticated(r)) return;
    if (r && r.ok) {
      rememberAuthorizedFile(r);
      const selected = String(r.dir || "").replace(/[\\/]$/, "") + "\\" + String(r.name || "");
      const matches = !candidate || (pathCandidate(candidate) ?
        selected.replace(/\//g, "\\").toLowerCase() === candidate.replace(/\//g, "\\").toLowerCase() : r.name === candidate);
      fileEntryStatus(matches ? "已打开：" + selected :
        "实际打开的是：" + selected + "，和刚才填的不一样；可能是原文件被移动了，或者你在文件框里选了别的。");
    } else {
      const reason = r && r.reason ? r.reason : "未取得选择结果";
      appendLog("warn", "选文件：" + reason);
      fileEntryStatus(reason === "cancelled" ? "已取消，现在打开的文件没有变。文件可能已移动，请重新选择。" :
        "没能打开：" + reason + "。请点「选择文档」重试。");
    }
  } finally {
    historySelectingFile = false;
    refreshHistory(); // 只读异步刷新，不占用当前任务 busy/requesting。
  }
}
async function analyzeVision(source) {
  const question = el("prompt").value.trim();
  if (!question || question.length > 2000) {
    setStatus("warn", "请先输入不超过 2000 字的图片问题。"); focusVisible(el("prompt")); return;
  }
  const list = el("vision-list");
  if (source === "document" && list.hidden) {
    const inventory = await api.list_vision_images();
    if (notAuthenticated(inventory)) return;
    if (!inventory?.ok) { setStatus("warn", inventory?.error || "无法读取当前文档的图片清单。"); return; }
    list.replaceChildren();
    inventory.images.forEach(image => {
      const option = document.createElement("option"); option.textContent = image.label;
      option.value = image.relationship; option.dataset.part = image.part; list.append(option);
    });
    if (!list.options.length) { setStatus("info", "当前文档没有可选的内部图片。"); return; }
    list.dataset.version = inventory.version; list.hidden = false;
    el("vision-document").textContent = "分析所选图片";
    setStatus("info", "请选择文档图片，再点「分析所选图片」。");
    return;
  }
  const label = source === "local" ? "本地图片" : list.selectedOptions[0]?.textContent;
  if (!label) { setStatus("warn", "请先选择图片。"); return; }
  const confirmed = await confirmModal("确认图片理解",
    "将仅把所选" + label + "和当前问题发送给 DeepSeek 视觉模型；可能产生费用。不会写入文档。是否继续？");
  if (!confirmed) return;
  const option = list.selectedOptions[0];
  const result = await api.understand_image(source, question,
    source === "document" ? option.value : "",
    source === "document" ? option.dataset.part : "",
    source === "document" ? list.dataset.version : "");
  if (notAuthenticated(result)) return;
  if (!result?.ok) {
    if (result?.reason !== "cancelled") setStatus("warn", result?.error || "图片理解未完成，请检查图片和视觉模型配置。");
    return;
  }
  appendLog("user-message", "图片理解（" + label + "）：" + question);
  appendLog("answer", result.answer);
  el("prompt").value = "";
  setStatus("info", "图片理解完成；结果仅在当前窗口显示，不会写回文档。");
}

async function ask() {
  const selector = typeof window !== 'undefined' ? window.OfficeModelSelector : null;
  if (selector && !selector.readyForRequest()) return;
  if (!hasFile && (!selector || selector.mode() === 'agent')) {
    setStatus("warn", "请先点「选择文档」打开一个文件，刚才输入的要求已保留。");
    focusVisible(el("btn-select2"));
    return;
  }
  window.dispatchEvent(new Event("officediagram:invalidate"));
  const q = el("prompt").value;
  const local=selector?.mode()==='agent'?localMainMessage(q):null;
  if(local)await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
  const r = window.OfficeDaily ? await window.OfficeDaily.send(q) : await api.ask(q);
  if (r?.reason === 'stale') return;
  if (notAuthenticated(r)) return;
  if (r && !r.ok) {
    if(local){localMainRound=null;setMainReply(r.error||'请求未受理，原话已保留；请核对当前文件和模型。',local);}
    if (r.reason === "no_file") setStatus("warn", "请先点「选择文档」打开一个文件。");
    else if (r.reason === "busy") setStatus("warn", "上一个任务还在处理，等它结束再发。");
    else if (r.reason === "empty") setStatus("warn", "请先写下你要做什么。");
    else if (r.reason === 'too_long') setStatus('warn', '输入有 '+r.characters+' 字符，最多 '+r.limit+' 字符；请缩短后再发，原文保留。');
    else if (r.reason === 'chat_context_changed') setStatus('warn', '发送范围已改变，请重新预览；输入保留，没有调用模型。');
    else if (r.reason === 'chat_submission_uncertain') setStatus('warn', '未确认请求是否已受理；输入保留，请先核对会话记录，避免重复发送。');
    else if (r.reason === 'model_context_budget') setStatus('warn', '当前输入超出上下文预算；输入已保留，没有请求模型。请缩短输入或核对上下文配置后重新预览。', 'chat-budget');
    else if (String(r.reason || '').startsWith('model_')) setStatus('warn', '模型选择已失效或配置已改变，请重新选择；输入已保留。');
    else setStatus('warn', '请求未受理，请核对模型配置或生图参数；输入已保留。');
  } else if (r && r.ok) {
    if (el("prompt").value === q && (!window.OfficeDaily || window.OfficeDaily.mayClear(q))) el("prompt").value = "";
    window.OfficeDaily?.sent(q);
  } else {
    setStatus("warn", "未确认指令是否已受理；输入已保留，请先查看任务记录，避免重复发送。");
  }
}
async function stopTask() {
  if (!busy || stopRequested || !api?.cancel_task) return;
  stopRequested = true;
  refreshButtons();
  try {
    const result = await api.cancel_task();
    if (result?.ok) setStatus("warn", result.local_only ? "已停止本地显示；远端可能继续处理或计费，等待请求释放后可主动重试。" : "正在停止：当前模型请求返回后，不会再执行后续模型或工具步骤。");
    else if (!notAuthenticated(result)) {
      stopRequested = false;
      setStatus("warn", "当前没有可停止的模型任务，可稍后再试。");
    }
  } catch (_) {
    stopRequested = false;
    setStatus("warn", "停止请求未送达，请核对当前状态后重试。");
  }
  refreshButtons();
}
async function confirmWrite() {
  if (!pending) return;
  const version = planVersion;
  const body = el("plan-pane").textContent;
  const ok = await confirmModal("确认真实写入",
    "将先备份并创建输出副本（原文件不动）。\n\n" + body + "\n\n是否确认执行？");
  if (!ok) { appendLog("meta", "已取消：未做任何写入，原文件保持不变。"); return; }
  if (!pending || version !== planVersion) {
    setStatus("warn", "确认期间计划已变化或失效；未提交写入，请重新查看当前计划。");
    return;
  }
  await api.execute();
}
async function undo() {
  if (!verified) return;
  const version = planVersion;
  const ok = await confirmModal("确认撤销",
    "只删除上次生成、且内容没有被改动过的输出副本；原文件和备份都保留。\n确认撤销？");
  if (!ok) { appendLog("meta", "已取消撤销；输出副本保留。"); return; }
  if (!verified || version !== planVersion) { setStatus("warn", "这份输出已经变了，没有提交撤销；请先看当前输出。"); return; }
  await api.undo();
}
async function clearLog() {
  if (busy || requesting || modalOpen) return;
  const ok = await confirmModal("清空显示",
    "只清空这个窗口里的对话显示；正在等你确认的计划和可撤销的输出都还在。当前打开的文件、后台任务、文件旁边的记录和已生成的副本都不受影响。是否清空？");
  if (!ok || busy || requesting) return;
  resetLogs();
}
async function copyText(text, label) {
  if (busy || requesting || modalOpen) return;
  try {
    if (!navigator.clipboard) throw new Error("当前窗口未提供剪贴板权限");
    await navigator.clipboard.writeText(text);
    appendLog("meta", "已复制" + label + "，共 " + text.length + " 个字符。");
  } catch (e) {
    setStatus("warn", "复制未完成：" + (e && e.message ? e.message : String(e)) + "。可选中结果文字手动复制。");
    appendLog("warn", "复制失败：" + label);
  }
}
async function copyPane(paneId) {
  await copyText(el(paneId).textContent, "面板文本");
}
async function probe() { await api.probe(); }
function settingsProbeStatus(state, text) {
  const node = el("settings-probe-status");
  node.className = "settings-hint" + (state ? " state-" + state : "");
  node.textContent = text;
}

/* ---------- 绑定 ---------- */
function bind() {
  bindFileEntries();
  el("btn-select").addEventListener("click", () => requestAction("正在打开文件…", selectFile));
  el("btn-select2").addEventListener("click", () => requestAction("正在打开文件…", selectFile));
  el("btn-ask").addEventListener("click", () => requestAction("正在提交…", ask));
  el("btn-stop").addEventListener("click", stopTask);
  el("btn-confirm").addEventListener("click", () => requestAction("正在确认写入…", confirmWrite));
  el("btn-undo").addEventListener("click", () => requestAction("正在确认撤销…", undo));
  el("btn-resume").addEventListener("click", () => requestAction("正在继续上次进度…", async () => {
    if (api.resume) await api.resume();
    else setStatus("warn", "这一版还不支持接着上次跑；请重新发一次要求。");
  }));
  el("btn-clear").addEventListener("click", clearLog);
  el("btn-copy-log").addEventListener("click", copyLogs);
  el("btn-copy-plan").addEventListener("click", () => copyPane("plan-pane"));
  el("btn-copy-result").addEventListener("click", () => copyText(resultEvent ? resultCopyText(resultEvent, resultSummary) : el("result-pane").textContent, "结果摘要与原始输出"));
  el("btn-copy-tsv").addEventListener("click", copyResultTSV);
  el("btn-result-prev").addEventListener("click", () => { resultPage--; renderResultPage(); });
  el("btn-result-next").addEventListener("click", () => { resultPage++; renderResultPage(); });
  el("btn-show-plan").addEventListener("click", () => selectPanel("plan", true));
  el("btn-pending-plan").addEventListener("click", () => selectPanel("plan", true));
  el("btn-show-result").addEventListener("click", () => selectPanel("result", true));
  el("btn-probe").addEventListener("click", () => requestAction("正在检测本机 Office…", probe));
  el("settings-probe").addEventListener("click", () => {
    if (!connected || busy || requesting || modalOpen || officeProbing || currentView === "history") return;
    settingsProbeStatus("", "正在检测本机 Office 环境…");
    requestAction("正在检测本机 Office…", probe);
  });
  el("prompt").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.isComposing && !e.shiftKey && !e.ctrlKey && !modalOpen && !window.OfficeSettings?.isOpen?.() && !drawerOpen && currentView !== 'history' && !document.querySelector?.('dialog[open]')) {
      e.preventDefault();
      requestAction("正在提交指令…", ask);
    }
  });
  document.addEventListener("keydown", (e) => {
    if (modalOpen || e.isComposing || e.defaultPrevented || document.querySelector?.('dialog[open]')) return;
    // 设置窗口是独立一层：键盘归它自己管，不把 Ctrl+Enter 提交给当前任务。
    if (window.OfficeSettings?.isOpen?.()) {
      if (e.ctrlKey && e.key === "Enter") e.preventDefault();
      return;
    }
    if(window.OfficeFloating?.active&&!document.getElementById('floating-root')?.hidden)return;
    if (currentView === "history" || drawerOpen || el("history-view").contains(document.activeElement)) {
      if (e.ctrlKey && e.key === "Enter") e.preventDefault();
      if (e.key === "Escape" && !drawerOpen) { e.preventDefault(); navigate("work"); }
      return; // 历史视图不能经全局快捷键提交当前任务。
    }
    const documentWorkspace = document.getElementById?.("document-workspace");
    if (documentWorkspace && e.target !== el('prompt')) {
      if (e.ctrlKey && e.key === "Enter") e.preventDefault();
      if (e.key === "Escape") {
        e.preventDefault();
        window.dispatchEvent(new CustomEvent("officeworkspace:escape"));
      }
      return;
    }
    if (e.ctrlKey && e.key === "Enter") {
      e.preventDefault();
      requestAction("正在提交指令…", ask);
    } else if (e.key === "Escape") {
      e.preventDefault(); el("prompt").focus();
    }
  });
}

/* C3–C5：只读侧栏与渐进披露；窄窗抽屉和最终确认互斥。 */
let drawerOpen = false;
const narrowSidebar = window.matchMedia("(max-width: 899px)");
function openDrawer() {
  if (!narrowSidebar.matches || modalOpen) return;
  drawerOpen = true;
  document.querySelector(".sidebar").classList.add("drawer-open");
  el("sidebar-backdrop").hidden = false;
  document.querySelector("main").inert = true;
  el("sidebar-toggle").setAttribute("aria-expanded", "true");
}
function closeDrawer(restore = true) {
  if (!drawerOpen) return;
  drawerOpen = false;
  document.querySelector(".sidebar").classList.remove("drawer-open");
  el("sidebar-backdrop").hidden = true;
  document.querySelector("main").inert = false;
  el("sidebar-toggle").setAttribute("aria-expanded", "false");
  if (restore && !modalOpen) focusVisible([el("sidebar-toggle"), el("nav-work"), el("prompt")].find(canFocus));
}
function initProgressiveUI() {
  const sidebar = document.querySelector(".sidebar"), nav = sidebar.querySelector("nav");
  sidebar.id = "task-sidebar";
  el("history-view").hidden = false;
  const reader = document.createElement("section"); reader.id = "history-reader"; reader.hidden = true;
  reader.setAttribute("aria-label", "历史会话只读阅读区"); reader.setAttribute("tabindex", "-1");
  const readerNote = document.createElement("p"); readerNote.className = "history-note";
  readerNote.textContent = "从侧栏选择历史会话。此处只读，不恢复任务或批准；返回当前工作可继续原任务。";
  const returnButton = document.createElement("button"); returnButton.className = "btn secondary";
  returnButton.textContent = "回到当前工作"; returnButton.onclick = () => navigate("work");
  reader.append(returnButton, readerNote, el("history-detail"));
  el("work-view").after(reader);
  const historyRegion = document.createElement("div"); historyRegion.id = "history-region"; historyRegion.className = "history-region";
  nav.insertBefore(historyRegion, el("nav-settings").closest(".nav-group"));
  historyRegion.append(el("history-view"));
  const toggle = document.createElement("button");
  toggle.id = "sidebar-toggle"; toggle.className = "iconbtn sidebar-toggle";
  toggle.type = "button"; toggle.textContent = "☰"; toggle.setAttribute("aria-label", "展开或关闭任务侧栏");
  toggle.setAttribute("aria-controls", "task-sidebar"); toggle.setAttribute("aria-expanded", "false");
  sidebar.prepend(toggle);
  const close = document.createElement("button"); close.type = "button"; close.id = "sidebar-close";
  close.className = "btn secondary sidebar-close"; close.textContent = "关闭侧栏";
  sidebar.append(close);
  const backdrop = document.createElement("button"); backdrop.id = "sidebar-backdrop";
  backdrop.className = "sidebar-backdrop"; backdrop.hidden = true; backdrop.tabIndex = -1;
  backdrop.setAttribute("aria-label", "关闭任务侧栏"); el("workspace").append(backdrop);
  toggle.onclick = () => { if (drawerOpen) closeDrawer(); else { openDrawer(); focusVisible(el("btn-history-refresh")); } };
  close.onclick = () => closeDrawer(); backdrop.onclick = () => closeDrawer();
  narrowSidebar.addEventListener("change", () => { const previous = document.activeElement; const hadFocus = sidebar.contains(previous); closeDrawer(false); if (hadFocus && !canFocus(previous)) focusVisible([el("sidebar-toggle"), el("nav-work"), el("prompt")].find(canFocus)); });
  document.addEventListener("keydown", event => {
    if (!drawerOpen || modalOpen || event.isComposing) return;
    if(window.OfficeFloating?.active&&!document.getElementById('floating-root')?.hidden)return;
    if (event.key === "Escape") { event.preventDefault(); event.stopImmediatePropagation(); closeDrawer(); }
    else if (event.key === "Tab") {
      const targets = Array.from(sidebar.querySelectorAll("button, input, summary, [tabindex='0']")).filter(canFocus);
      const first = targets[0], last = targets[targets.length - 1];
      if (!targets.length) return;
      if (!sidebar.contains(document.activeElement) || (event.shiftKey && document.activeElement === first) || (!event.shiftKey && document.activeElement === last)) {
        event.preventDefault(); (event.shiftKey ? last : first).focus();
      }
    } else if (event.ctrlKey && event.key === "Enter") { event.preventDefault(); event.stopImmediatePropagation(); }
  }, true);
  const top = document.querySelector(".topbar");
  document.querySelector("main").append(top); // DOM：附件/输入 → 流 → 详情；视觉用 order 保持顶栏。
  const technical = el("technical-details"); technical.classList.add("top-details");
  technical.querySelector("summary").textContent = "详情";
  const info = technical.querySelector(".technical-info");
  [el("status-line"), el("next-step"), el("activity"), el("busy-pill"), el("model-info")].filter(Boolean).forEach(node => { node.setAttribute("aria-live", "off"); node.removeAttribute("role"); info.prepend(node); });
  const file = document.createElement("div"); file.id = "details-file"; file.className = "filedir";
  const errorTitle = document.createElement("h2"); errorTitle.textContent = "最近错误原文";
  const error = document.createElement("pre"); error.id = "error-detail"; error.textContent = "尚未收到错误";
  info.prepend(file); info.append(document.querySelector(".attachment-options"), errorTitle, error); top.append(technical);
  const more = document.createElement("details"); more.id = "more-menu"; more.className = "top-details more-menu";
  const summary = document.createElement("summary"); summary.textContent = "更多";
  const menu = document.createElement("div"); menu.className = "more-actions";
  [el("btn-copy-log"), el("btn-clear"), el("btn-resume")].forEach(node => { node.hidden = false; menu.append(node); });
  more.append(summary, menu); top.append(more);
  const tools = el("tools-menu"), topMenus = [technical, more, tools];
  topMenus.forEach(details => details.addEventListener("toggle", () => { if (details.open) { if (details === tools) el("tools-menu-toggle").textContent = "工具"; topMenus.filter(other => other !== details).forEach(closeTopDetails); } }));
  function closeTopDetails(details) {
    const focused = document.activeElement;
    details.open = false;
    if (details.contains(focused) && !canFocus(focused)) focusVisible(details.querySelector("summary"));
  }
  document.addEventListener("keydown", event => {
    if (event.key !== "Escape" || modalOpen || drawerOpen || !topMenus.some(details => details.open)) return;
    const active = topMenus.find(details => details.open && details.contains(document.activeElement)) || topMenus.find(details => details.open);
    topMenus.forEach(closeTopDetails);
    focusVisible(active.querySelector("summary"));
    event.preventDefault(); event.stopImmediatePropagation();
  }, true);
  document.addEventListener("click", event => { topMenus.forEach(details => { if (details.open && !details.contains(event.target)) closeTopDetails(details); }); });
  [["plan-card", ["btn-copy-plan"]], ["result-card", ["btn-copy-result", "btn-copy-tsv"]]].forEach(([id, ids]) => {
    const actions = document.createElement("div"); actions.className = "message-actions";
    ids.forEach(button => actions.append(el(button))); el(id).append(actions);
  });
  el("btn-confirm").textContent = "应用更改…";
  el("btn-undo").textContent = "撤销输出…";
  setNavigation("work");
}

/* ---------- 轮询 ---------- */
let pollingEvents=false;
async function poll() {
  if (window.OfficeFloating?.active) return;
  if (floatingEventOwner) return;
  if (!api || !window.OfficeAuth?.isAuthenticated()) return;
  if(pollingEvents||document.hidden)return;
  pollingEvents=true;
  try {
    const events = await api.poll();
    if (Array.isArray(events)) events.forEach(handleEvent);
  } catch (e) { /* API 尚未就绪或窗口关闭 */ }
  finally { pollingEvents=false; }
}

initConversation();
initProgressiveUI();
refreshButtons();
let businessBound = false, pollTimer = null, connecting = false, floatingEventOwner = false, connectionEpoch = 0;
async function connectAuthenticated() {
  if (connecting || !window.pywebview?.api || !window.OfficeAuth) return;
  connecting = true;
  const mine=connectionEpoch;
  try {
    if (!await window.OfficeAuth.ensure()) return;
    const candidateApi=window.OfficeAuth.wrapApi(window.pywebview.api);
    let owner=false;
    if(typeof window.pywebview.api.floating_capabilities==='function'){
      try {const capabilities=await candidateApi.floating_capabilities();owner=capabilities?.ok===true&&capabilities.event_owner==='floating_snapshot';}
      catch(_){owner=false;}
    }
    if(mine!==connectionEpoch||!window.OfficeAuth.isAuthenticated())return;
    api=candidateApi;floatingEventOwner=owner;
    void restoreTheme();
    if (floatingEventOwner && window.OfficeFloating) window.OfficeFloating.active = true;
    connected = Boolean(api);
    window.dispatchEvent(new Event('officemodels:host-ready'));
    if (!businessBound) { bind(); businessBound = true; }
    refreshButtons();
    await refreshHistory();
    if (pollTimer === null && !floatingEventOwner) pollTimer = setInterval(poll, 200);
  } finally { connecting = false;if(mine!==connectionEpoch&&window.OfficeAuth?.isAuthenticated())queueMicrotask(connectAuthenticated); }
}
window.addEventListener("pywebviewready", connectAuthenticated);
window.addEventListener("officeauth:authenticated", connectAuthenticated);
window.addEventListener("officeauth:unauthenticated", () => {
  connected = false; api = null;
  connectionEpoch++;
  floatingEventOwner = false;
  el("tools-menu").open = false; el("tools-menu-toggle").textContent = "工具";
  if (pollTimer !== null) { clearInterval(pollTimer); pollTimer = null; }
  refreshButtons();
});
if (window.pywebview?.api) connectAuthenticated();
