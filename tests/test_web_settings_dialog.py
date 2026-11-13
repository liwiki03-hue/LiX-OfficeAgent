"""设置窗口：从任务栏拆出后，独立窗口的显示、分区、焦点与隔离契约。

只跑静态结构与 Node DOM 桩，不启动 pywebview、不连后端。
"""
import json
import subprocess
from html.parser import HTMLParser
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'
from test_web_simplify import NODE  # noqa: E402

PANES = ('settings-pane-model', 'settings-pane-skills', 'settings-pane-rag',
         'settings-pane-machine', 'settings-pane-account')
TABS = ('tab-model', 'tab-skills', 'tab-rag', 'tab-machine', 'tab-account')


class Dom(HTMLParser):
    def __init__(self):
        super().__init__()
        self.stack, self.nodes = [], {}

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if 'id' in attrs:
            assert attrs['id'] not in self.nodes, f"duplicate ID: {attrs['id']}"
            self.nodes[attrs['id']] = (tag, attrs, tuple(self.stack))
        if tag not in {'input', 'img', 'meta', 'link', 'br', 'hr'}:
            self.stack.append(attrs.get('id', tag))

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in {'input', 'img', 'meta', 'link', 'br', 'hr'}:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        if self.stack:
            self.stack.pop()


def _dom():
    dom = Dom()
    dom.feed((STATIC / 'index.html').read_text(encoding='utf-8'))
    return dom


def test_settings_content_no_longer_sits_in_the_task_rail():
    dom = _dom()
    # 侧栏不再有承载配置内容的折叠容器，只剩一个入口按钮。
    assert 'settings' not in dom.nodes
    assert 'model-settings' not in dom.nodes
    for gone in ('nav-probe', 'nav-theme', 'theme-label'):
        assert gone not in dom.nodes, f"{gone} 应随设置内容一起迁走"
    entry = dom.nodes['nav-settings']
    # 入口按钮仍在侧栏（aside）里，且归入底部组。
    assert 'aside' in entry[2], '设置入口应留在侧栏底部'
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    assert 'nav-group nav-group-tail' in html, '设置入口应挂在侧栏底部的独立分组里'
    assert entry[1]['aria-haspopup'] == 'dialog'
    assert entry[1]['aria-controls'] == 'settings-dialog'
    # 任何一个设置控件都不应再挂在侧栏下。
    for node in ('settings-key', 'skills-panel', 'rag-panel', 'auth-logout', 'settings-probe'):
        assert 'aside' not in dom.nodes[node][2], f"{node} 仍在任务栏里"
        assert 'settings-dialog' in dom.nodes[node][2]


def test_every_pane_is_reachable_from_exactly_one_tab():
    dom = _dom()
    assert len(PANES) == len(TABS)
    for tab, pane in zip(TABS, PANES):
        assert dom.nodes[tab][0] == 'button'
        assert dom.nodes[tab][1]['aria-controls'] == pane
        assert dom.nodes[pane][1]['aria-labelledby'] == tab
        assert 'settings-dialog' in dom.nodes[pane][2]
    # 首屏只展开一个分区，其余必须 hidden，否则窗口会同时渲染全部内容。
    hidden = [pane for pane in PANES if 'hidden' in dom.nodes[pane][1]]
    assert len(hidden) == len(PANES) - 1, '除默认分区外都应默认收起'
    assert 'settings-pane-model' not in hidden, '默认分区应是模型与检索'


def test_skills_and_retrieval_have_scrollable_visible_contents():
    dom = _dom()
    for panel in ('skills-panel', 'rag-panel'):
        assert dom.nodes[panel][0] == 'details'
        assert 'open' in dom.nodes[panel][1]
    css = (STATIC / 'theme.css').read_text(encoding='utf-8')
    assert '.settings-panel { border: 0; height: 100%; min-height: 0; }' in css
    assert '.settings-panel > .settings-form { padding: 16px; height: 100%; overflow-y: auto; }' in css


def test_dialog_is_a_real_modal_layer():
    dom = _dom()
    tag, attrs, _ = dom.nodes['settings-dialog']
    assert tag == 'section' and attrs['role'] == 'dialog'
    assert attrs['aria-modal'] == 'true'
    assert attrs['aria-labelledby'] == 'settings-dialog-title'
    assert 'hidden' in attrs, '窗口默认不显示'
    assert dom.nodes['settings-backdrop'][1]['hidden'] is not None or 'hidden' in dom.nodes['settings-backdrop'][1]
    assert dom.nodes['settings-close'][0] == 'button'
    assert '关闭' in dom.nodes['settings-close'][1]['aria-label']


def test_moved_controls_keep_their_ids_and_contracts():
    """迁移只换位置：id、类型、禁用默认值一律不变，controller 无需改写。"""
    dom = _dom()
    assert dom.nodes['settings-key'][1]['type'] == 'password'
    assert dom.nodes['settings-key'][1]['autocomplete'] == 'off'
    for control in ('settings-save', 'settings-rollback', 'skills-create', 'skills-edit',
                    'skills-delete', 'skills-toggle', 'skills-import', 'skills-rollback',
                    'rag-index'):
        assert 'disabled' in dom.nodes[control][1], f"{control} 必须默认禁用"
    for control in ('rag-question', 'rag-topk'):
        assert 'disabled' in dom.nodes[control][1]
    assert dom.nodes['rag-question'][1]['maxlength'] == '200'
    assert dom.nodes['rag-topk'][1]['value'] == '5'
    assert dom.nodes['settings-probe'][0] == 'button'
    assert dom.nodes['settings-theme'][0] == 'button'
    assert dom.nodes['auth-logout'][0] == 'button'


def test_dialog_module_is_loaded_and_stays_side_effect_free():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    assert html.count('settings_dialog.js') == 1
    for controller in ('settings_dialog.js', '"settings.js"', '"skills.js"', '"rag.js"'):
        assert html.count(controller) == 1, f"{controller} 应只加载一次"
    # 窗口容器必须先于所有住进它的控制器定义，否则控制器找不到宿主元素。
    assert html.index('settings_dialog.js') < html.index('"skills.js"')
    assert html.index('settings_dialog.js') < html.index('"rag.js"')
    source = (STATIC / 'settings_dialog.js').read_text(encoding='utf-8')
    for prohibited in ('innerHTML', 'localStorage', 'sessionStorage', 'fetch(', 'console.', 'eval(', 'pywebview'):
        assert prohibited not in source, f"设置窗口不应出现 {prohibited}"
    assert "role=" not in source, '无障碍语义写在 HTML 里，脚本不重建'


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM-stub dependency')
def test_dialog_focus_partition_and_escape_protocol(tmp_path):
    probe = tmp_path / 'settings_dialog_probe.js'
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
class El {
  constructor(id) {
    this.id = id; this.children = []; this.hidden = false; this.disabled = false;
    this.classes = new Set(); this.attrs = {}; this.listeners = {}; this.textContent = '';
    this.dataset = {}; this.focused = 0;
    this.classList = { toggle: (n, on) => on ? this.classes.add(n) : this.classes.delete(n) };
  }
  append(...n) { this.children.push(...n); }
  addEventListener(t, fn) { this.listeners[t] = fn; }
  setAttribute(k, v) { this.attrs[k] = String(v); }
  removeAttribute(k) { delete this.attrs[k]; }
  getAttribute(k) { return this.attrs[k]; }
  focus() { this.focused++; global.active = this; document.activeElement = this; }
  contains(node) { return node === this || this.children.includes(node); }
  querySelectorAll() { return this.focusables; }
}
const nodes = {}, node = id => nodes[id] ||= new El(id);
const active = () => global.active;
for (const id of ['settings-dialog', 'settings-backdrop', 'settings-close', 'nav-settings',
                  'tab-model', 'tab-skills', 'tab-rag', 'tab-machine', 'tab-account',
                  'settings-pane-model', 'settings-pane-skills', 'settings-pane-rag',
                  'settings-pane-machine', 'settings-pane-account']) node(id);
const dialog = node('settings-dialog');
const tabs = ['tab-model', 'tab-skills', 'tab-rag', 'tab-machine', 'tab-account'];
const panes = tabs.map(t => node(t.replace('tab-', 'settings-pane-')));
dialog.focusables = tabs.map(node);
dialog.children.push(...dialog.focusables);
const document = {
  activeElement: null,
  listeners: {},
  addEventListener: (t, fn) => document.listeners[t] = fn,
  dispatchEvent: () => true,
  getElementById: id => nodes[id] || null,
  readyState: 'complete'
};
const context = { document, window: {}, CustomEvent: class { constructor(t, o) { this.type = t; Object.assign(this, o && o.detail ? { detail: o.detail } : {}); } } };
context.window.document = document;
vm.runInNewContext(fs.readFileSync(process.argv[2], 'utf8'), context);
const api = context.window.OfficeSettings;
const click = id => node(id).listeners.click();
const key = (k, extra = {}) => { const e = { key: k, shiftKey: false, preventDefault() { this.prevented = (this.prevented || 0) + 1; }, stopImmediatePropagation() { this.stopped = true; }, ...extra }; document.listeners.keydown(e); return e; };
assert(api && typeof api.show === 'function' && typeof api.close === 'function');
assert(api.isOpen() === false, '初始必须是关闭态');
// 初始态由 index.html 的 hidden 属性决定；stub 里手工对齐，再走正常开关流程。
dialog.hidden = true; node('settings-backdrop').hidden = true;
api.close();
assert(dialog.hidden === true && node('settings-backdrop').hidden === true, '关闭态不得自行打开');
// 打开：窗口与遮罩显示，焦点落到第一个分区，默认只展开一个。
document.activeElement = node('nav-settings');
api.show();
assert(api.isOpen() === true);
assert(dialog.hidden === false && node('settings-backdrop').hidden === false);
assert(active() === node('tab-model'), '焦点应进入设置窗口');
assert(panes.filter(p => !p.hidden).length === 1 && panes[0].hidden === false);
assert(node('tab-model').attrs['aria-current'] === 'true');
// 切分区：只有目标可见，其余收起，aria-current 只有一个。
click('tab-rag');
assert(panes[2].hidden === false && panes[0].hidden === true);
assert(node('tab-rag').attrs['aria-current'] === 'true' && !node('tab-model').attrs['aria-current']);
click('tab-skills');
assert(panes[1].hidden === false && panes[2].hidden === true);
// Tab 在窗口内打转，不会跑到窗口外。
document.activeElement = dialog.focusables[4];
key('Tab');
assert(active() === dialog.focusables[0], '末尾 Tab 应回到第一个可聚焦项');
document.activeElement = dialog.focusables[0];
key('Tab', { shiftKey: true });
assert(active() === dialog.focusables[4], '开头 Shift+Tab 应跳到最后一个');
document.activeElement = dialog.focusables[1];
const middle = key('Tab');
assert(!middle.prevented, '窗口中间的 Tab 交给浏览器按正常顺序走');
// Esc 关闭并把焦点交回入口按钮；事件被拦下，不泄漏给全局快捷键。
document.activeElement = node('settings-close');
const escape = key('Escape');
assert(escape.prevented === 1 && escape.stopped === true, 'Esc 必须被设置窗口独占处理');
assert(api.isOpen() === false && dialog.hidden === true);
assert(active() === node('nav-settings'), '关闭后焦点回到设置入口');
// 已关闭时 Esc 不再做任何事。
const after = key('Escape');
assert(!after.prevented && api.isOpen() === false);
// 遮罩点击关闭。
api.show();
assert(api.isOpen() === true);
node('settings-backdrop').listeners.click({ target: node('settings-backdrop') });
assert(api.isOpen() === false);
console.log(JSON.stringify({ ok: true, focus: true, panes: true, escape: true }));
''', encoding='utf-8')
    result = subprocess.run([NODE, str(probe), str(STATIC / 'settings_dialog.js')],
                            capture_output=True, text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stdout + result.stderr
    assert json.loads(result.stdout) == {'ok': True, 'focus': True, 'panes': True, 'escape': True}
