"""Direction 3 contracts: static structure and deterministic DOM behavior, no providers/Office."""
import json
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'
BUNDLED = Path.home() / '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe'
NODE = shutil.which('node') or (str(BUNDLED) if BUNDLED.is_file() else None)


class Structure(HTMLParser):
    def __init__(self, html):
        super().__init__()
        self.stack = []
        self.nodes = {}
        self.presets = []
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if 'id' in attrs:
            assert attrs['id'] not in self.nodes, f"duplicate ID: {attrs['id']}"
            self.nodes[attrs['id']] = (tag, attrs, tuple(self.stack))
        if 'data-prompt' in attrs:
            self.presets.append(attrs)
        if tag not in {'input', 'img', 'meta', 'link', 'br', 'hr'}:
            self.stack.append(attrs.get('id', tag))

    def handle_endtag(self, tag):
        if self.stack:
            self.stack.pop()


def test_empty_state_tools_and_safety_structure():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    dom = Structure(html)
    assert len(dom.presets) == 4
    assert '你想对这份文档做什么？' in html
    assert dom.nodes['tools-menu'][0] == 'details'
    assert 'open' not in dom.nodes['tools-menu'][1]
    assert dom.nodes['tools-menu-toggle'][0] == 'summary'
    assert dom.nodes['tools-menu-toggle'][1]['aria-controls'] == 'tools-panel'
    # 图表与生图已改为对话驱动：不再有类型下拉框或需求输入框，改为说明入口。
    assert 'imagegen-card' not in dom.nodes
    assert 'diagram-card' not in dom.nodes
    for control in ('diagram-kind', 'diagram-goal', 'diagram-generate', 'diagram-cancel',
                    'diagram-source', 'diagram-image', 'diagram-save', 'diagram-rollback',
                    'diagram-status', 'diagram-export-status', 'diagram-render', 'diagram-copy',
                    'imagegen-card', 'imagegen-body', 'imagegen-state'):
        assert control not in dom.nodes, f"表单控件 {control} 应已移除"
    assert 'tools-panel' in html and 'tools-menu' in html  # 工具说明仍挂在侧边工具面板里
    assert 'hidden' in dom.nodes['workspace'][1] and 'inert' in dom.nodes['workspace'][1]
    assert 'open' not in dom.nodes['technical-details'][1]
    # 对话驱动的说明必须出现在界面上，用户可以据此知道怎么触发。
    for hint in ('画一个', '生成一张'):
        assert hint in html, f"缺少对话触发示例：{hint}"
    css = (STATIC / 'theme.css').read_text(encoding='utf-8')
    assert '.empty-flow .flow-toolbar, .empty-flow .conversation-stream { display: none; }' in css
    assert '.empty-flow #status-card:not(.state-error):not(.state-warning)' in css
    assert '.sidebar .history-region' in css
    assert '.diagram-message' in css and '.diagram-figure' in css
    source = (STATIC / 'app.js').read_text(encoding='utf-8')
    assert 'info.append(document.querySelector(".attachment-options"), errorTitle, error)' in source
    assert 'function appendDiagramMessage(' in source
    assert 'window.OfficeDiagramExport?.mount?.(exportActions)' in source


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM-stub dependency')
def test_default_dark_theme_preserves_explicit_preferences(tmp_path):
    probe = tmp_path / 'theme_probe.js'
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[2], 'utf8');
const themeSource = source.slice(source.indexOf('const themes ='), source.indexOf('function escapeHtml('));
function boot(saved, osDark, storageBlocked=false) {
  const ids = {}, values = {}, events = {};
  const el = id => ids[id] ||= {setAttribute(key,value) { this[key]=value; },
                                addEventListener(key,callback) { events[id+':'+key]=callback; }};
  const document = {documentElement:{dataset:{}}};
  let osQueries = 0;
  const ctx = {document, el, api:null, localStorage: {
    getItem() { if (storageBlocked) throw new Error('no storage'); return saved; },
    setItem(key,value) { if (storageBlocked) throw new Error('no storage'); values[key]=value; }
  }, window:{matchMedia: () => { osQueries++; return {matches:osDark}; }}};
  vm.createContext(ctx); vm.runInContext(themeSource, ctx);
  return {ctx, ids, values, events, document, osQueries:()=>osQueries};
}
for (const osDark of [false, true]) {
  const fresh = boot(null, osDark);
  assert.equal(fresh.document.documentElement.dataset.theme,'dark');
  assert.equal(fresh.osQueries(),0);
  assert(fresh.ids['btn-theme'].title.includes('当前主题：深色'));
  fresh.events['settings-theme:click']();
  assert.equal(fresh.document.documentElement.dataset.theme,'light');
  assert.equal(fresh.values['office-agent.theme'],'light');
  assert(fresh.ids['settings-theme'].textContent.includes('当前：浅色'));
}
for (const choice of ['light','dark']) {
  const saved = boot(choice, false);
  assert.equal(saved.document.documentElement.dataset.theme,choice);
}
const legacy = boot('highcontrast', false);
assert.equal(legacy.document.documentElement.dataset.theme,'dark');
assert.equal(legacy.values['office-agent.theme'],'dark');
legacy.events['btn-theme:click']();
assert.equal(legacy.document.documentElement.dataset.theme,'light');
legacy.events['btn-theme:click']();
assert.equal(legacy.document.documentElement.dataset.theme,'dark');
assert.equal(boot('highcontrast', false, true).document.documentElement.dataset.theme,'dark');
assert.equal(boot('unexpected', false).document.documentElement.dataset.theme,'dark');
const blocked = boot(null, false, true);
assert.equal(blocked.document.documentElement.dataset.theme,'dark');
blocked.events['btn-theme:click']();
assert.equal(blocked.document.documentElement.dataset.theme,'light');
console.log('ok');
''', encoding='utf-8')
    result = subprocess.run([NODE, str(probe), str(STATIC / 'app.js')], capture_output=True,
                            text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == 'ok'


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM-stub dependency')
def test_menu_empty_and_multiple_task_dom_regressions(tmp_path):
    probe = tmp_path / 'probe.js'
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[2], 'utf8');
class Element {
 constructor() { this.children=[]; this.dataset={}; this.textContent=''; this.hidden=false; this.open=false; this.events={}; this.classes=new Set(); this.classList={toggle:(name,on)=>on?this.classes.add(name):this.classes.delete(name)}; }
 append(...nodes) { this.children.push(...nodes); }
 appendChild(node) { this.append(node); }
 contains(node) { return node===this || this.children.includes(node); }
 setAttribute(name,value) { this[name]=value; }
 addEventListener(name,callback) { this.events[name]=callback; }
 querySelector() { return this.summary ||= new Element(); }
}
const ids={}, el=id=>ids[id] ||= new Element(), events={};
const document={activeElement:null,body:{},createElement:()=>new Element(),addEventListener:(name,fn)=>events[name]=fn};
let focused;
const ctx={document,el,logRecords:[],logQueue:[],logQueueHead:0,connected:true,historyLoading:false,historySelectingFile:false,modalOpen:false,drawerOpen:false,historyError:'',historyWarning:'',historyFilter:'',historyHidden:new Set(),historySessions:[],historySelected:null,isCurrentSession:()=>false,metadataText:x=>x,translateHistoryText:x=>x,historyDate:x=>x||'today',historyStatus:x=>x||'done',renderHistoryDetail:()=>{},focusVisible:x=>focused=x,canFocus:()=>false,technical:new Element(),more:new Element()};
vm.createContext(ctx);
vm.runInContext(source.slice(source.indexOf('function refreshEmptyFlow('),source.indexOf('/* 历史是独立只读视图')),ctx);
ctx.refreshEmptyFlow(); assert(el('work-view').classes.has('empty-flow')); assert(!el('conversation-empty').hidden);
ctx.logQueue.push({}); ctx.refreshEmptyFlow(); assert(!el('work-view').classes.has('empty-flow')); assert(el('conversation-empty').hidden);
ctx.logQueue=[]; ctx.logRecords.push({}); ctx.refreshEmptyFlow(); assert(el('conversation-empty').hidden);
vm.runInContext(source.slice(source.indexOf('function renderHistory()'),source.indexOf('function invalidateHistory(')),ctx);
ctx.historySessions=[{session_dir:'one',question:'task one',file:'one.xlsx'}]; ctx.renderHistory();
assert(el('history-list').children[0].children.length===1); assert(!el('history-view').classes.has('multiple-tasks'));
el('history-list').children=[]; ctx.historySessions.push({session_dir:'two',question:'task two',file:'two.xlsx'}); ctx.renderHistory();
assert(el('history-view').classes.has('multiple-tasks')); assert(el('history-list').children.every(row=>row.children.length===1));
ctx.historySessions.push({session_dir:'three',question:'another task',file:'three.xlsx'}); ctx.renderHistory();
el('history-list').children=[]; ctx.historyFilter='task one'; ctx.renderHistory(); assert(el('history-list').children.length===1); assert(el('history-list').children[0].children.length===1);
vm.runInContext(source.slice(source.indexOf('  const tools = el("tools-menu")'),source.indexOf('  [["plan-card", ["btn-copy-plan"]]')),ctx);
const tools=el('tools-menu'); el('tools-menu-toggle').textContent='工具 · 生图待查看'; ctx.technical.open=true; tools.open=true; tools.events.toggle(); assert(!ctx.technical.open && tools.open); assert.equal(el('tools-menu-toggle').textContent,'工具');
document.activeElement=tools; let prevented=false, stopped=false; events.keydown({key:'Escape',preventDefault:()=>prevented=true,stopImmediatePropagation:()=>stopped=true});
assert(!tools.open && prevented && stopped && focused===tools.querySelector('summary'));
tools.open=true; ctx.modalOpen=true; events.keydown({key:'Escape'}); assert(tools.open); // confirmation retains keyboard ownership
ctx.modalOpen=false; events.click({target:{}}); assert(!tools.open);
console.log(JSON.stringify({ok:true,menu:true,empty:true,multiple:true}));
''', encoding='utf-8')
    result = subprocess.run([NODE, str(probe), str(STATIC / 'app.js')], capture_output=True,
                            text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout) == {'ok': True, 'menu': True, 'empty': True, 'multiple': True}


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM-stub dependency')
def test_send_without_file_keeps_prompt_and_does_not_submit(tmp_path):
    probe = tmp_path / 'no_file.js'
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[2], 'utf8');
const prompt = {value: '检查第一段'}, select = {};
let notice, focused, calls = 0;
const ctx = {hasFile: false, el: id => id === 'prompt' ? prompt : select,
  setStatus: (level, text) => notice = {level, text}, focusVisible: node => focused = node,
  api: {ask: () => {calls++; throw Error('no-file task submitted');}}};
vm.createContext(ctx);
vm.runInContext(source.slice(source.indexOf('async function ask()'), source.indexOf('async function confirmWrite()')), ctx);
(async () => {
  await ctx.ask();
  assert.equal(calls, 0);
  assert.equal(prompt.value, '检查第一段');
  assert.equal(notice.level, 'warn');
  assert(notice.text.includes('已保留'));
  assert.strictEqual(focused, select);
  assert(source.includes('el("btn-ask").disabled = blocked;'));
})().catch(error => {console.error(error); process.exitCode = 1;});
''', encoding='utf-8')
    result = subprocess.run([NODE, str(probe), str(STATIC / 'app.js')], capture_output=True,
                            text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stderr
