"""Node adapter protocol stubs; not native rendering or business-quality proof."""
import shutil
import subprocess
from pathlib import Path
import pytest

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'
BUNDLED = Path.home() / '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe'
NODE = shutil.which('node') or (str(BUNDLED) if BUNDLED.is_file() else None)


def test_renderer_is_fixed_local_display_only_and_production_not_enabled():
    adapter = (STATIC / 'diagram_renderer.js').read_text(encoding='utf-8')
    worker = (STATIC / 'diagram_worker.js').read_text(encoding='utf-8')
    assert 'diagram_renderer.js' not in (STATIC / 'index.html').read_text(encoding='utf-8')
    assert "setAttribute('sandbox', 'allow-scripts')" in adapter
    assert 'allow-same-origin' not in adapter
    for policy in ("default-src 'none'", "connect-src 'none'", "img-src 'none'", "object-src 'none'", "frame-src 'none'", "base-uri 'none'", "form-action 'none'"):
        assert policy in adapter
    assert 'source}' not in adapter.split('frame.srcdoc =')[1]
    assert 'innerHTML' not in adapter
    for script in (adapter, worker):
        assert 'api.' not in script and 'fetch(' not in script and 'XMLHttpRequest' not in script
    assert 'securityLevel:\'strict\'' in worker and 'htmlLabels:false' in worker
    assert "typeof window.pywebview === 'undefined'" in worker
    assert "event.source !== parent" in worker
    assert "error?.message === 'unsafe_svg'" in worker


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic protocol-stub dependency')
def test_renderer_message_binding_no_source_in_markup_and_cleanup(tmp_path):
    script = tmp_path / 'protocol.js'
    script.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const code=fs.readFileSync(process.argv[2],'utf8');
function harness(url='http://127.0.0.1:3210/static/index.html') {
  const listeners=new Set(),timers=new Map(),frames=[],sent=[]; let serial=0;
  const window={addEventListener:(n,f)=>listeners.add(f),removeEventListener:(n,f)=>listeners.delete(f)};
  const document={baseURI:url,createElement:()=>{const f={attrs:{},style:{},removed:false,contentWindow:{postMessage:d=>sent.push(d)},setAttribute(k,v){this.attrs[k]=v},remove(){this.removed=true}};frames.push(f);return f},body:{append:f=>{f.appended=true}}};
  const context={window,document,URL,Uint8Array,AbortSignal,crypto:{getRandomValues:a=>{a.fill(++serial);return a}},setTimeout:(f,ms)=>{timers.set(serial,{f,ms});return serial},clearTimeout:id=>timers.delete(id)};
  vm.runInNewContext(code,context);
  return {api:window.OfficeDiagramRenderer,frames,sent,listeners,timers,fire(data,source=frames.at(-1).contentWindow,origin='null'){for(const f of [...listeners])f({data,source,origin})}};
}
const channel='office-diagram-v1',source='flowchart TD\nUser[User]';
(async()=>{
  const h=harness(),p=h.api.render(source);const f=h.frames[0];
  assert.equal(f.attrs.sandbox,'allow-scripts');assert(f.srcdoc.includes("connect-src 'none'"));assert(!f.srcdoc.includes(source));
  const ready={channel,type:'ready',bridgeAbsent:true};
  h.fire(ready,{},'null');h.fire(ready,f.contentWindow,'http://evil.invalid');assert.equal(h.sent.length,0);
  h.fire(ready);h.fire(ready);assert.equal(h.sent.length,1);assert.equal(h.sent[0].source,source);
  h.fire({channel,type:'result',token:'wrong',svg:'unsafe',bridgeAbsent:true});assert(!f.removed);
  h.fire({channel,type:'result',token:h.sent[0].token,svg:'passive fixture',bridgeAbsent:true});
  const result=await p;assert.equal(result.svg,'passive fixture');assert(f.removed);assert.equal(h.listeners.size,0);assert.equal(h.timers.size,0);
  const timed=h.api.render(source);const expired=assert.rejects(timed,e=>e.code==='render_timeout');[...h.timers.values()][0].f();await expired;assert(h.frames[1].removed);assert.equal(h.listeners.size,0);
  const c=new AbortController(),cancelled=h.api.render(source,{signal:c.signal});const cancelledCheck=assert.rejects(cancelled,e=>e.code==='render_cancelled');c.abort();await cancelledCheck;assert(h.frames[2].removed);assert.equal(h.listeners.size,0);assert.equal(h.timers.size,0);
  const b=harness(),bridged=b.api.render(source);const bc=assert.rejects(bridged,e=>e.code==='renderer_failed');b.fire({...ready,bridgeAbsent:false});await bc;assert(b.frames[0].removed);
  const bad=harness(),unsafe=bad.api.render(source);const uc=assert.rejects(unsafe,e=>e.code==='unsafe_svg');bad.fire(ready);bad.fire({channel,type:'result',token:bad.sent[0].token,svg:'x'.repeat(1048577),bridgeAbsent:true});await uc;
  const error=harness(),rejected=error.api.render(source);const ec=assert.rejects(rejected,e=>e.code==='renderer_failed');error.fire(ready);error.fire({channel,type:'error',token:error.sent[0].token,code:'<script>raw error</script>'});await ec;
  const remote=harness('https://remote.invalid/');await assert.rejects(remote.api.render(source),e=>e.code==='unsupported_source');assert.equal(remote.frames.length,0);
  for(const value of ['', 'x'.repeat(8193), '\u0000', null, 42]) await assert.rejects(h.api.render(value),e=>e.code==='unsupported_source');
  for(const ms of [0,99,8001,NaN,'8000']) await assert.rejects(h.api.render(source,{timeoutMs:ms}),e=>e.code==='unsupported_source');
  const already=new AbortController();already.abort();await assert.rejects(h.api.render(source,{signal:already.signal}),e=>e.code==='render_cancelled');
  for(const options of [null, [], 42, {signal:{}}, {signal:'abort'}]) await assert.rejects(h.api.render(source,options),e=>e.code==='unsupported_source');
  const cap=harness(),one=cap.api.render(source),two=cap.api.render(source);
  const oneCheck=assert.rejects(one,e=>e.code==='render_timeout'),twoCheck=assert.rejects(two,e=>e.code==='render_timeout');
  await assert.rejects(cap.api.render(source),e=>e.code==='renderer_failed');assert.equal(cap.frames.length,2);
  for(const timer of [...cap.timers.values()])timer.f();await Promise.all([oneCheck,twoCheck]);assert.equal(cap.listeners.size,0);assert.equal(cap.timers.size,0);
  assert.equal(h.listeners.size,0);assert.equal(h.timers.size,0);console.log('protocol stub passed');
})().catch(e=>{console.error(e);process.exitCode=1});
''', encoding='utf-8')
    result = subprocess.run([NODE, str(script), str(STATIC / 'diagram_renderer.js')], capture_output=True, text=True, timeout=12, encoding='utf-8')
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'protocol stub passed' in result.stdout


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic SVG-flatten stub dependency')
@pytest.mark.parametrize('backdrop_class', ['labelBkg', 'background', ''])
def test_worker_flatten_preserves_label_text_with_contrasting_passive_colors(tmp_path, backdrop_class):
    script = tmp_path / 'labels.js'
    script.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const code=fs.readFileSync(process.argv[2],'utf8');
assert(code.includes('  let used = false;'));
const NS='http://www.w3.org/2000/svg';
class E {
 constructor(name,attrs={},children=[],css={}) {this.nodeType=1;this.namespaceURI=NS;this.localName=name;this.attrs={...attrs};this.childNodes=[];this.css=css;for(const c of children)this.append(c)}
 get attributes(){return Object.entries(this.attrs).map(([name,value])=>({name,value}))}
 getAttribute(k){return this.attrs[k]??null}
 hasAttribute(k){return k in this.attrs}
 setAttribute(k,v){this.attrs[k]=String(v)}
 append(c){this.childNodes.push(c);c.parentNode=this}
 closest(s){for(let e=this;e;e=e.parentNode)if(s==='defs'?e.localName==='defs':e.getAttribute('class')?.split(/\s+/).includes(s.slice(1)))return e;return null}
 querySelector(s){for(const c of this.childNodes)if(c.nodeType===1){if(c.localName===s)return c;const f=c.querySelector(s);if(f)return f}return null}
}
const raw={fill:'rgb(147, 112, 219)',stroke:'rgb(147, 112, 219)','font-size':'14px'};
const word={nodeType:3,textContent:'places'};
const span=new E('tspan',{},[word],raw),text=new E('text',{},[span],raw);
const backdrop=new E('rect',{class:process.argv[3],x:'0',y:'0',width:'40',height:'16'},[],raw);
const shape=new E('rect',{class:'entityBox',width:'80',height:'50'},[],raw);
const label=new E('g',{class:'edgeLabel'},[new E('g',{class:'label'},[backdrop,text])]);
const root=new E('svg',{viewBox:'0 0 160 100'},[shape,label]);
const parent={postMessage:()=>{}};Object.defineProperty(parent,'document',{get(){throw new Error('cross-origin')}});
const context={window:{addEventListener:()=>{}},parent,TextEncoder,document:{createElementNS:(_ns,n)=>new E(n),createTextNode:t=>({nodeType:3,textContent:t})},getComputedStyle:e=>({fontSize:'14px',fontWeight:'400',getPropertyValue:k=>e.css[k]||''}),XMLSerializer:class{serializeToString(e){return JSON.stringify(e,(k,v)=>k==='parentNode'?undefined:v)}}};
vm.runInNewContext(code.replace('  let used = false;','  globalThis.__flattenForTest=flatten;\n  let used = false;'),context);
const out=JSON.parse(context.__flattenForTest(root));
const labelOut=out.childNodes[1].childNodes[0],bgOut=labelOut.childNodes[0],textOut=labelOut.childNodes[1],spanOut=textOut.childNodes[0];
assert.equal(bgOut.attrs.fill,'#fff');assert.equal(textOut.attrs.fill,'#222');assert.equal(spanOut.attrs.fill,'#222');
assert.equal(textOut.attrs.stroke,'none');assert.equal(bgOut.attrs.stroke,'none');assert.equal(spanOut.childNodes[0].textContent,'places');
assert.equal(bgOut.attrs.width,'40');assert.equal(out.childNodes[0].attrs.fill,raw.fill);assert(!out.attrs.style);
assert.equal(textOut.attrs['font-size'],'14');assert.equal(textOut.attrs['font-family'],'Arial');
root.append(new E('use'));assert.throws(()=>context.__flattenForTest(root),e=>e.message==='unsafe_svg');
console.log('label flatten stub passed');
''', encoding='utf-8')
    result = subprocess.run([NODE, str(script), str(STATIC / 'diagram_worker.js'), backdrop_class], capture_output=True, text=True, timeout=12, encoding='utf-8')
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'label flatten stub passed' in result.stdout
