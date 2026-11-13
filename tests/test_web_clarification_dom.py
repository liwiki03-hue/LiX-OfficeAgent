"""Deterministic Node DOM stub; not real keyboard/clipboard/native evidence."""
import json
import shutil
import subprocess
from pathlib import Path
import pytest
from office_agent.clarification import classify_intent, clarification_card

BUNDLED = Path.home() / '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe'
NODE = shutil.which('node') or (str(BUNDLED) if BUNDLED.is_file() else None)
STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static/app.js'


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM-stub dependency')
def test_card_chips_text_skip_stale_and_no_authorization(tmp_path):
    fixture = clarification_card(classify_intent('生成图表'))
    fixture['request_id'] = 'owned-request'
    fixture['preamble'] = '<script>literal, not markup</script>'
    payload = tmp_path / 'card.json'
    payload.write_text(json.dumps(fixture, ensure_ascii=False), encoding='utf-8')
    script = tmp_path / 'card.js'
    script.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
class Element {
  constructor(tag) { this.tag = tag; this.children = []; this.listeners = {}; this.textContent = ''; this.value = ''; this.disabled = false; }
  append(...items) { this.children.push(...items); }
  setAttribute(k,v) { this[k] = v; }
  addEventListener(k,v) { this.listeners[k] = v; }
  querySelectorAll(selector) {
    const tags = selector.split(',').map(x => x.trim());
    return this.children.flatMap(x => [...(tags.includes(x.tag) ? [x] : []), ...x.querySelectorAll(selector)]);
  }
  click() { if (!this.disabled && this.listeners.click) return this.listeners.click(); }
}
const calls = [], cards = [];
let reply = {ok:true};
const context = { document: {createElement: tag => new Element(tag)}, connected:true, busy:false,
  requesting:false, modalOpen:false, currentView:'current', requestLabel:'',
  api:{answer_clarification: async(...args) => {calls.push(args); return reply;}, execute:()=>{throw Error('write escaped');}},
  appendLog:(_tag,_text,card)=>cards.push(card), setStatus:()=>{}, refreshButtons:()=>{} };
vm.createContext(context);
const source = fs.readFileSync(process.argv[2], 'utf8');
const section = source.slice(source.indexOf('function invalidateClarification('), source.indexOf('function refreshButtons('));
assert(!section.includes('innerHTML')); assert(!section.includes('api.execute')); assert(!section.includes('answer_read'));
vm.runInContext('let activeClarification = null;'+section, context);
const fixture = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
(async()=>{
  context.renderClarification(fixture);
  const first = cards[0], buttons = first.querySelectorAll('button'), inputs = first.querySelectorAll('input');
  assert.strictEqual(first.children[1].textContent, fixture.preamble);
  assert.strictEqual(first.querySelectorAll('fieldset').length, 3);
  assert(first.querySelectorAll('p').filter(p=>p.textContent.startsWith('跳过或未填写时：')).length === 3);
  assert(inputs.every(input=>input.maxLength === 500));
  assert(first.querySelectorAll('label').every(label=>label.textContent.includes('补充说明（')));
  buttons.find(b=>b.textContent==='按我补充的目标制作待确认计划').click();
  inputs[1].value='B2<script>literal</script>';
  const next = buttons.find(b=>b.textContent==='按回答继续规划');
  await next.click();
  assert.strictEqual(calls.length, 1); assert.strictEqual(calls[0][0], 'owned-request');
  assert.strictEqual(calls[0][1].scope, inputs[1].value); assert.strictEqual(calls[0][1].goal, 'plan');
  assert.strictEqual(calls[0][2], false); assert(buttons.every(b=>b.disabled));
  await next.click(); assert.strictEqual(calls.length, 1);
  context.renderClarification({...fixture, request_id:'skip-request'});
  const second=cards[1], skip=second.querySelectorAll('button').find(b=>b.textContent===fixture.skip_label);
  await skip.click(); assert.strictEqual(calls.length, 2);
  assert.strictEqual(JSON.stringify(calls[1]), '["skip-request",{},true]');
  context.renderClarification({...fixture, request_id:'old'});
  const old=cards[2]; context.renderClarification({...fixture, request_id:'new'});
  assert(old.querySelectorAll('button').every(b=>b.disabled));
  const current=cards[3]; context.busy=true; context.refreshClarification();
  assert(current.querySelectorAll('button').every(b=>b.disabled));
  context.busy=false; context.currentView='history'; context.refreshClarification();
  assert(current.querySelectorAll('button').every(b=>b.disabled));
  context.currentView='current'; context.connected=false; context.refreshClarification();
  assert(current.querySelectorAll('button').every(b=>b.disabled));
  context.connected=true; context.refreshClarification();
  assert(current.querySelectorAll('button').every(b=>!b.disabled));
  console.log('clarification DOM stub: passed');
})().catch(error=>{console.error(error);process.exitCode=1;});
''', encoding='utf-8')
    result = subprocess.run([NODE, str(script), str(STATIC), str(payload)], capture_output=True, text=True, timeout=10, encoding='utf-8')
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'clarification DOM stub: passed' in result.stdout
