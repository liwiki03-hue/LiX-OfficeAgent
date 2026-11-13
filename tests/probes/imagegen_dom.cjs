/* Static DOM-stub probe for imagegen.js: proves the UI cannot authorize anything by itself.
   No browser, no network, no storage.
   Run: node IMAGE_P3_imagegen_dom_probe.js <static-dir> <out.json>

   生图已改为对话驱动：不再有 #imagegen-body 宿主，事件由 app.js 提供的 mount()
   把消息节点接进对话流。本探针用同样的契约调用 handleEvent，从而验证
   「渲染不自动提交」「必须两次点击」「退出登录立刻失效」这三条不变量仍然成立。 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const staticDir = process.argv[2];
const outFile = process.argv[3];
const report = {ok: false, checks: {}, error: null, submitted_operations: [],
                automatic_calls_after_render: 0, automatic_calls_after_dialog_open: 0,
                unknown_event_ignored: false, logout_disabled_buttons: 0,
                network_or_storage_used: false, ignored_selectors: []};
const calls = [];

class El {
  constructor(tag) {
    this.tagName = String(tag || "div").toUpperCase();
    this.children = [];
    this.parent = null;
    this.className = "";
    this.textContent = "";
    this.title = "";
    this.hidden = false;
    this.disabled = false;
    this.focused = false;
    this.dataset = {};
    this.attributes = {};
    this.listeners = {};
    this.scrollTop = 0;
    this.scrollHeight = 0;
    this.open = false;
    this.onclick = null;
  }
  append(...nodes) { nodes.forEach(node => { node.parent = this; this.children.push(node); }); }
  replaceChildren(...nodes) { this.children = []; this.append(...nodes); }
  remove() { if (this.parent) this.parent.children = this.parent.children.filter(item => item !== this); this.parent = null; }
  setAttribute(key, value) { this.attributes[key] = String(value); }
  removeAttribute(key) { delete this.attributes[key]; }
  addEventListener(type, fn) { (this.listeners[type] = this.listeners[type] || []).push(fn); }
  fire(type, event) { (this.listeners[type] || []).forEach(fn => fn(event || {})); }
  get isConnected() { return true; }
  showModal() { this.open = true; }
  close() { this.open = false; this.fire("close"); }
  focus() { this.focused = true; }
  all() { return this.children.flatMap(child => [child].concat(typeof child.all === "function" ? child.all() : [])); }
  querySelectorAll(selector) {
    const want = String(selector).toUpperCase();
    return this.all().filter(node => node.tagName === want);
  }
}

const registry = new Map();
/* 对话流宿主：mountImageMessage 会把每条消息追加到这里。 */
const flow = new El("div");
const document = {
  body: new El("body"),
  createElement: tag => new El(tag),
  createTextNode: value => ({textContent: String(value), nodeValue: String(value), parent: null,
                             all: () => [], remove() {}}),
  getElementById: id => registry.get(id) || null,
  querySelectorAll: selector => {
    if (selector === ".imagegen-item .imagegen-actions button") {
      return flow.all().filter(node => node.tagName === "BUTTON"
        && String(node.parent && node.parent.className).includes("imagegen-actions"));
    }
    report.ignored_selectors.push(selector);
    return [];
  },
};
const listeners = {};
const window = {
  addEventListener: (type, fn) => { (listeners[type] = listeners[type] || []).push(fn); },
  OfficeImages: {appendImages: (host, images) => {
    report.office_images_appended = (report.office_images_appended || 0) + (Array.isArray(images) ? images.length : 0);
    host.append(new El("figure"));
  }},
};

function load(file) {
  const source = fs.readFileSync(path.join(staticDir, file), "utf8");
  if (file === "imagegen.js") {
    report.network_or_storage_used = ["fetch(", "XMLHttpRequest", "localStorage", "sessionStorage",
      "eval(", "new Function", "innerHTML", "http://", "https://"].some(token => source.includes(token));
  }
  vm.runInContext(source, sandbox, {filename: file});
}

function buttonByText(root, needle) {
  return root.all().find(node => node.tagName === "BUTTON" && String(node.textContent).includes(needle));
}
function tick() { return new Promise(resolve => setTimeout(resolve, 0)); }
/* 与 app.js 的 mountImageMessage 等价：把消息接进对话流当前轮次。 */
function mount(card) { flow.append(card); }

const sandbox = {window, document, console, setTimeout, clearTimeout, atob, btoa};
vm.createContext(sandbox);

async function main() {
try {
  load("images.js");
  load("imagegen.js");
  const api = window.OfficeImageGen;
  report.checks.module_loaded = typeof (api && api.handleEvent) === "function";
  report.checks.images_module_loaded = typeof (window.OfficeImages && window.OfficeImages.appendImages) === "function";

  const call = async (name, ...args) => {
    calls.push(name + (args.length ? ":" + String(args[args.length - 1]) : ""));
    return {ok: true, status: "generated", result: {}};
  };
  const planEvent = {type: "image_plan", plan: {plan_id: "plan-1",
    effective_prompt: "Requested visual style: natural.\nprobe", size: "1024x1024", style: "natural", n: 1,
    approval_ttl_seconds: 300, configuration: {provider: "openai_compatible", model: "synthetic-model",
    base_url: "https://example.invalid/v1"}}};

  api.handleEvent(planEvent, call, mount);
  report.automatic_calls_after_render = calls.length;
  const texts = flow.all().map(node => String(node.textContent));
  report.checks.plan_card_rendered = flow.children.length === 1
    && texts.some(value => value.includes("发给模型的提示词"))
    && texts.some(value => value.includes("无法收回"))
    && texts.some(value => value.includes("1024x1024"));
  const card = flow.children[0];
  const generate = buttonByText(card, "确认生成");
  const cancelPlan = buttonByText(card, "取消，不要这张图");
  report.checks.two_separate_buttons = Boolean(generate && cancelPlan);

  // Opening the dialog: the click handler stays pending until a dialog button is clicked,
  // so it must not be awaited here (that would stall the probe).
  void generate.onclick();
  await tick();
  const dialog = document.body.children.find(node => node.tagName === "DIALOG" && node.open);
  report.automatic_calls_after_dialog_open = calls.length;
  report.checks.dialog_requires_second_click = Boolean(dialog) && calls.length === 0;
  report.checks.dialog_default_focus_is_cancel = dialog.children
    .flatMap(child => child.all()).some(node => node.tagName === "BUTTON" && node.focused
      && String(node.textContent) === "取消");
  void buttonByText(dialog, "确认生成").onclick();
  await tick();
  report.checks.generation_submitted_after_two_clicks = calls.includes("confirm_image_generation:true")
    && !document.body.children.some(node => node.tagName === "DIALOG");

  api.handleEvent(planEvent, call, mount);
  const freshCancel = buttonByText(flow.children[flow.children.length - 1], "取消，不要这张图");
  void freshCancel.onclick();
  await tick();
  report.checks.cancel_submits_explicit_false = calls.includes("confirm_image_generation:false");

  // A real 1x1 PNG so the P1 validating renderer (images.js) is exercised end to end.
  const VALID_PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WlM9x8AAAAASUVORK5CYII=";
  api.handleEvent({type: "image_result", plan_id: "plan-1", text: "已生成", images: [{data_url: VALID_PNG,
    alt: "云端生成的图片（尚未保存）", caption: "12x8 · 仅内存，未保存"}]}, call, mount);
  const resultCard = flow.children[flow.children.length - 1];
  report.checks.result_rendered_images = resultCard.all().some(node =>
    String(node.className).includes("image-results")) &&
    resultCard.all().some(node => String(node.className).includes("image-enlarge"));
  void buttonByText(flow.children[flow.children.length - 1], "保存成新的图片文件").onclick();
  await tick();
  report.checks.save_planned_from_result = calls.includes("plan_image_save:plan-1");

  api.handleEvent({type: "image_save_plan", plan: {save_id: "save-1", path: "/authorized/office-image-save-1.png",
    bytes: 78, sha256: "a".repeat(64)}}, call, mount);
  const savedCard = flow.children[flow.children.length - 1];
  report.checks.save_plan_discloses_new_file = savedCard.all()
    .some(node => String(node.textContent).includes("不会覆盖原文档"));
  void buttonByText(savedCard, "确认保存新文件").onclick();
  await tick();
  const saveDialog = document.body.children.find(node => node.tagName === "DIALOG" && node.open);
  void buttonByText(saveDialog, "确认保存新文件").onclick();
  await tick();
  report.checks.save_submitted_after_dialog = calls.includes("confirm_image_save:true");

  api.handleEvent({type: "image_saved", save_id: "save-1", text: "已保存新文件"}, call, mount);
  void buttonByText(flow.children[flow.children.length - 1], "撤销这个文件").onclick();
  await tick();
  report.checks.rollback_planned_from_saved = calls.includes("plan_image_rollback:save-1");

  api.handleEvent({type: "image_rollback_plan", plan: {rollback_id: "rollback-1",
    path: "/authorized/office-image-save-1.png", sha256: "a".repeat(64)}}, call, mount);
  const rollbackCard = flow.children[flow.children.length - 1];
  report.checks.rollback_discloses_no_fee_reversal = rollbackCard.all()
    .some(node => String(node.textContent).includes("不退费用"));
  void buttonByText(rollbackCard, "确认删除这个文件").onclick();
  await tick();
  const rollbackDialog = document.body.children.find(node => node.tagName === "DIALOG" && node.open);
  void buttonByText(rollbackDialog, "确认删除这个文件").onclick();
  await tick();
  report.checks.rollback_submitted_after_dialog = calls.includes("confirm_image_rollback:true");

  api.handleEvent({type: "image_rolled_back", text: "已删除自有输出"}, call, mount);
  report.checks.rolled_back_card = flow.all().some(node => String(node.textContent).includes("没有退回"));

  const before = flow.children.length;
  api.handleEvent({type: "unknown_event_kind"}, call, mount);
  report.unknown_event_ignored = flow.children.length === before;

  api.handleEvent(planEvent, call, mount); // leave one live card so the logout check has buttons to disable
  const pendingBefore = document.querySelectorAll(".imagegen-item .imagegen-actions button")
    .filter(button => !button.disabled).length;
  (listeners["officeauth:unauthenticated"] || []).forEach(fn => fn({}));
  const stillEnabled = document.querySelectorAll(".imagegen-item .imagegen-actions button")
    .filter(button => !button.disabled).length;
  report.logout_disabled_buttons = pendingBefore - stillEnabled;
  report.checks.logout_disables_pending_buttons = stillEnabled === 0 && pendingBefore >= 2;

  report.submitted_operations = calls;
  report.ok = Object.values(report.checks).every(Boolean);
  if (!report.ok) report.error = "one or more checks failed";
} catch (error) {
  report.error = String((error && error.stack) || error);
}
}

main().then(() => {
  delete report.office_images_appended;
  fs.writeFileSync(outFile, JSON.stringify(report, null, 2), "utf8");
  if (!report.ok) { console.error(report.error || "probe failed"); process.exit(1); }
  console.log(JSON.stringify({ok: report.ok, checks: Object.keys(report.checks).length,
                              calls: report.submitted_operations.length}));
});
