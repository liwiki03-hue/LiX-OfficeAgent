"""Image-understanding UI flow with a deterministic Node DOM stub; no native GUI."""
import shutil
import subprocess
from pathlib import Path

import pytest

SOURCE = Path(__file__).resolve().parents[1] / "office_agent/webui/static/app.js"
BUNDLED = Path.home() / ".dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe"
NODE = shutil.which("node") or (str(BUNDLED) if BUNDLED.is_file() else None)


@pytest.mark.skipif(NODE is None, reason="Node unavailable: deterministic JavaScript flow probe")
def test_document_vision_requires_selection_then_confirmation(tmp_path):
    probe = tmp_path / "vision-flow.js"
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(process.argv[2], 'utf8');
const snippet = src.slice(src.indexOf('async function analyzeVision('), src.indexOf('async function ask()'));
const calls = [], messages = [];
const select = { hidden: true, dataset: {}, options: [], selectedOptions: [],
  replaceChildren() { this.options = []; this.selectedOptions = []; },
  append(option) { this.options.push(option); this.selectedOptions = [option]; } };
const els = { prompt: {value:'图里是什么？'}, 'vision-list': select,
  'vision-document': {textContent:'分析文档图片…'} };
let confirmed = false;
const ctx = {el: id => els[id], document: {createElement: () => ({dataset:{}})},
  api: {list_vision_images: async () => { calls.push('list'); return {ok:true, version:'v1',
    images:[{relationship:'rId1',part:'word/document.xml',label:'正文图片 1'}]}; },
    understand_image: async (...args) => { calls.push(args); return {ok:true, answer:'红色方块'}; }},
  notAuthenticated: () => false, confirmModal: async () => { calls.push('confirm'); return confirmed; },
  setStatus: (...args) => messages.push(args), appendLog: (...args) => messages.push(args),
  focusVisible: () => {} };
vm.createContext(ctx); vm.runInContext(snippet, ctx);
(async () => {
  await ctx.analyzeVision('document');
  assert.deepEqual(calls,['list']);
  assert.equal(select.hidden,false);
  assert.equal(els['vision-document'].textContent,'分析所选图片');
  await ctx.analyzeVision('document');
  assert.deepEqual(calls,['list','confirm']);
  confirmed = true;
  await ctx.analyzeVision('document');
  assert.equal(calls.length,4);
  assert.equal(JSON.stringify(calls[3]), JSON.stringify(['document','图里是什么？','rId1','word/document.xml','v1']));
  assert.equal(els.prompt.value,'');
  assert(messages.some(item => item[0] === 'answer' && item[1] === '红色方块'));
})().catch(error => { console.error(error); process.exitCode = 1; });
''', encoding="utf-8")
    result = subprocess.run([NODE, str(probe), str(SOURCE)], capture_output=True, text=True, timeout=15)
    assert result.returncode == 0, result.stdout + result.stderr
