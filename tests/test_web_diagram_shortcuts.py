"""Extracted production keyboard handlers in a Node DOM stub, not real keyboard proof.

The standalone diagram card is gone, so Ctrl+Enter no longer needs a card-scoped
exemption: submitting always means "submit the document task".
"""
import shutil
import subprocess
from pathlib import Path

STATIC = Path(__file__).resolve().parents[1] / "office_agent/webui/static/app.js"
BUNDLED = Path.home() / ".dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe"
NODE = shutil.which("node") or (str(BUNDLED) if BUNDLED.is_file() else None)


def test_diagram_card_exemption_is_removed(tmp_path):
    assert NODE is not None, "Node is required for deterministic production handler checks"
    script = tmp_path / "shortcuts.js"
    script.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const raw = fs.readFileSync(process.argv[2], 'utf8');
// The deleted form must not linger anywhere in the production handlers.
assert(!raw.includes('#diagram-card'), '图表卡片豁免应随表单一起移除');
const source = raw.replace(/\r\n/g, '\n');
const start = source.indexOf('  el("prompt").addEventListener("keydown", (e) => {');
const end = source.indexOf('\n}\n', start);
assert(start >= 0 && end > start, 'production keyboard handlers must exist');
const handlers = {}, requests = [], navigation = [];
let asks = 0, focuses = 0;
const prompt = {addEventListener:(type, fn)=>{handlers.prompt = fn;}, focus:()=>{focuses++;}};
let settingsOpen = false;
const context = {
  modalOpen:false, currentView:'work', drawerOpen:false,
  document:{activeElement:null, addEventListener:(type, fn)=>{assert.equal(type, 'keydown'); handlers.global = fn;}},
  el:id=>id === 'prompt' ? prompt : {contains:()=>false},
  ask:()=>{asks++;}, requestAction:(label, fn)=>{requests.push(label); fn();},
  navigate:view=>navigation.push(view),
  window:{OfficeSettings:{isOpen:()=>settingsOpen}}
};
vm.runInNewContext(source.slice(start, end), context);
function event(target, extra={}) {
  return {target, key:'Enter', ctrlKey:true, isComposing:false, prevented:0,
    preventDefault(){this.prevented++;}, ...extra};
}
// Ctrl+Enter submits the document task wherever focus sits.
for (const target of [{id:'anywhere', closest:()=>null}, {closest:()=>null}, {}, null]) {
  const e = event(target); handlers.global(e); assert.equal(e.prevented, 1);
}
assert.equal(asks, 4); assert.equal(requests.length, 4);
// In the prompt, plain Enter submits.
const enter = event(prompt, {ctrlKey:false}); handlers.prompt(enter);
assert.equal(enter.prevented, 1); assert.equal(asks, 5); assert.equal(requests.length, 5);
// Ctrl+Enter in the prompt is ignored by the prompt handler (it requires !ctrlKey)
// and handled by the document handler exactly once — no double submission.
const ctrlEnter = event(prompt); handlers.prompt(ctrlEnter); handlers.global(ctrlEnter);
assert.equal(ctrlEnter.prevented, 1, '只由全局 handler 阻止默认行为');
assert.equal(asks, 6, '同一次按键只提交一次任务，不因两个 handler 而重复提交');
assert.equal(requests.length, 6);
// Escape still focuses the document prompt.
const escape = event({closest:()=>({})}, {key:'Escape', ctrlKey:false}); handlers.global(escape);
assert.equal(escape.prevented, 1); assert.equal(focuses, 1);
// Suppressed contexts never submit and never navigate.
for (const state of ['modalOpen', 'isComposing', 'drawerOpen', 'history']) {
  context.modalOpen = state === 'modalOpen'; context.drawerOpen = state === 'drawerOpen';
  context.currentView = state === 'history' ? 'history' : 'work';
  handlers.global(event({closest:()=>null}, {isComposing:state === 'isComposing'}));
  assert.equal(asks, 6);
}
// 设置窗口打开时是独立一层：既不提交任务，也不把 Esc 转给对话输入框。
context.modalOpen = false; context.drawerOpen = false; context.currentView = 'work';
settingsOpen = true;
const inSettings = event({closest:()=>null});
handlers.global(inSettings);
assert.equal(inSettings.prevented, 1, '设置窗口里 Ctrl+Enter 只被拦下，不提交');
assert.equal(asks, 6);
const escapeSettings = event({closest:()=>null}, {key:'Escape', ctrlKey:false});
handlers.global(escapeSettings);
assert.equal(focuses, 1, '设置窗口自己处理 Esc 关闭，全局 handler 不再抢焦点');
settingsOpen = false;
assert.equal(navigation.length, 0);
console.log('diagram shortcut DOM stub: passed');
''', encoding="utf-8")
    result = subprocess.run(
        [NODE, str(script), str(STATIC)], capture_output=True, text=True,
        encoding="utf-8", timeout=10,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "diagram shortcut DOM stub: passed" in result.stdout
