"""Human-readable results/error cards: Node DOM stub, not a native visual acceptance."""
import json
import shutil
import subprocess
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / 'office_agent' / 'webui' / 'static'
BUNDLED = Path.home() / '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe'
NODE = shutil.which('node') or (str(BUNDLED) if BUNDLED.is_file() else None)


def test_results_keep_raw_output_in_default_collapsed_technical_information():
    source = (STATIC / 'app.js').read_text(encoding='utf-8')
    section = source.split('function appendResultInfo(', 1)[1].split('function renderTable(', 1)[0]
    assert 'details.open = true' not in section
    assert 'raw.textContent = String(ev.text ?? "")' in section
    assert 'data.textContent = JSON.stringify(ev.data, null, 2)' in section
    assert 'innerHTML' not in section


def test_errors_preserve_raw_detail_without_claiming_no_write():
    source = (STATIC / 'app.js').read_text(encoding='utf-8')
    assert 'raw.textContent = record.text' in source
    assert 'level === "danger" ? humanErrorNotice(text) : text' in source
    assert 'level === "danger" ? humanErrorNotice(text) : warning' in source
    assert '原始报错信息' in source


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM-stub dependency')
def test_dom_stub_results_errors_and_risks(tmp_path):
    probe = tmp_path / 'probe.js'
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[2], 'utf8');
class Element {
  constructor(tag) { this.tag = tag; this.children = []; this.style = {}; this.textContent = ''; this.open = false; }
  append(...items) { this.children.push(...items); }
  prepend(...items) { this.children.unshift(...items); }
  setAttribute(name, value) { this[name] = value; }
}
const context = { document: { createElement: tag => new Element(tag) }, window: {} };
vm.createContext(context);
const summaries = source.slice(source.indexOf('function dataObject('), source.indexOf('function renderTable('));
const errors = source.slice(source.indexOf('function humanErrorNotice('), source.indexOf('function setStatus('));
vm.runInContext(summaries + errors, context);
const ev = {tool: 'inspect', text: '<script>raw exactly</script>', data: {readable: false, errors: ['risk detail'], protected: true}};
const pane = new Element('div');
context.appendResultInfo(pane, ev);
const details = pane.children.find(n => n.tag === 'details');
assert(details && details.open === false);
assert(details.children.find(n => n.tag === 'pre').textContent === ev.text);
assert(details.children[3].textContent === JSON.stringify(ev.data, null, 2));
const human = pane.children[0].children.map(n => n.textContent).join('');
assert(human.includes('不能正常读取') && human.includes('risk detail'));
assert(!human.includes('<script>'));
assert(pane.children[0].tag === 'section');
assert(pane.children[0].children[0].tag === 'h3');
assert(pane.children[0].children.every(n => n.tag !== 'pre'));
assert(pane.children[0].children.some(n => n.className === 'result-risk'));
for (const raw of ['Traceback: synthetic_internal_code', 'http timeout', '401 unauthorized', 'PermissionError: access denied', 'network connection failure']) {
  const notice = context.humanErrorNotice(raw);
  assert(!notice.includes(raw));
  assert(notice.includes('无法判断有没有写入成功'));
  assert(notice.includes('不要重复确认'));
}
for (const code of ['image_configuration_missing', 'image_configuration_invalid']) {
  const notice = context.humanErrorNotice('Tool failed: ' + code + ' IMAGE_API_KEY=private');
  assert(notice.includes('本次没有请求生图服务'));
  assert(notice.includes('重启后再试'));
  assert(!notice.includes('private') && !notice.includes('IMAGE_API_KEY'));
}
const archive = new Element('div');
context.appendArchivedError(archive, 'Traceback: archived_internal_code');
assert(archive.children[0].tag === 'p' && !archive.children[0].textContent.includes('archived_internal_code'));
assert(archive.children[0].textContent.includes('无法判断有没有写入成功'));
assert(archive.children[1].tag === 'details' && archive.children[1].open === false);
assert(archive.children[1].children[1].textContent === 'Traceback: archived_internal_code');
assert(archive.children[1].children[0].textContent === '原始报错信息');
assert(source.includes('appendArchivedError(host, state.error)'));
assert(source.includes('appendArchivedError(body, copy)'));
console.log(JSON.stringify({ok: true, collapsed: true, raw_retained: true, risk_visible: true}));
''', encoding='utf-8')
    run = subprocess.run([NODE, str(probe), str(STATIC / 'app.js')], capture_output=True,
                         text=True, encoding='utf-8', timeout=20)
    assert run.returncode == 0, run.stderr
    assert json.loads(run.stdout) == {'ok': True, 'collapsed': True, 'raw_retained': True, 'risk_visible': True}
