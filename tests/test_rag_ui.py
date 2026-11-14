"""RAG HTML/source and executable DOM contracts; no native services."""
import os
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path

import pytest
import playwright

STATIC = Path(__file__).resolve().parents[1] / "office_agent" / "webui" / "static"
BUNDLED_NODE = Path(playwright.__file__).resolve().parent / 'driver' / ('node.exe' if os.name == 'nt' else 'node')
NODE = str(BUNDLED_NODE) if BUNDLED_NODE.is_file() else shutil.which("node")


def test_html_mounts_authenticated_rag():
    class DOM(HTMLParser):
        def __init__(self):
            super().__init__()
            self.stack, self.nodes, self.scripts = [], {}, []

        def handle_starttag(self, tag, attrs):
            attrs = dict(attrs)
            if "id" in attrs:
                assert attrs["id"] not in self.nodes
                self.nodes[attrs["id"]] = (tag, attrs, tuple(self.stack))
            if tag == "script":
                self.scripts.append(attrs.get("src"))
            if tag not in {"input", "meta", "link", "br", "img", "hr"}:
                self.stack.append(attrs.get("id", tag))

        def handle_startendtag(self, tag, attrs):
            self.handle_starttag(tag, attrs)
            if tag not in {"input", "meta", "link", "br", "img", "hr"}:
                self.handle_endtag(tag)

        def handle_endtag(self, tag):
            if self.stack:
                self.stack.pop()

    dom = DOM()
    dom.feed((STATIC / "index.html").read_text(encoding="utf-8"))
    assert dom.nodes["rag-panel"][0] == "details"
    # 检索面板已从任务栏迁入独立设置窗口的「文档内容检索」分区。
    assert {"settings-dialog", "settings-pane-rag"} <= set(dom.nodes["rag-panel"][2])
    assert "sidebar" not in dom.nodes["rag-panel"][2]
    assert {"hidden", "inert"} <= dom.nodes["workspace"][1].keys()
    for name in ("status", "snapshot", "hits", "evidence", "question", "topk"):
        assert "rag-panel" in dom.nodes["rag-" + name][2]
    assert dom.nodes["rag-status"][1]["role"] == "status"
    assert dom.nodes["rag-question"][1]["maxlength"] == "200"
    attrs = dom.nodes["rag-topk"][1]
    assert {k: attrs[k] for k in ("type", "min", "max", "value")} == {
        "type": "number", "min": "1", "max": "10", "value": "5"}
    for name in ("load", "index", "query", "read", "next", "rollback-index", "rollback-evidence"):
        tag, attrs, parents = dom.nodes["rag-" + name]
        assert tag == "button" and attrs["type"] == "button" and "rag-panel" in parents
        if name in {"read", "next", "rollback-index", "rollback-evidence"}:
            assert "disabled" in attrs
    assert dom.scripts.count("rag.js") == 1
    assert dom.scripts.index("auth.js") < dom.scripts.index("rag.js")


def test_controller_text_and_native_api_only():
    source = (STATIC / "rag.js").read_text(encoding="utf-8")
    for forbidden in ("innerHTML", "outerHTML", "insertAdjacentHTML", "localStorage",
                      "sessionStorage", "fetch(", "XMLHttpRequest", "eval(", "new Function",
                      "console.", "document.write", "execute_script", "run_script"):
        assert forbidden not in source
    for required in ("textContent", "createElement", "OfficeAuth.wrapApi", "OfficeConfirm",
                     "get_rag_snapshot", "plan_rag_index", "confirm_rag_index", "plan_rag_query",
                     "confirm_rag_query", "get_rag_evidence", "plan_rag_rollback", "confirm_rag_rollback",
                     "officeauth:authenticated", "officeauth:unauthenticated", "officediagram:invalidate",
                     "officerag:invalidate", "pagehide", "MutationObserver", "pywebviewready"):
        assert required in source


PROBE = r"""
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const source=fs.readFileSync(process.argv[1],'utf8'), scenario=process.argv[2];
class Element {
 constructor(id=''){this.id=id;this.value='';this._text='';this.children=[];this.events={};this.disabled=false;}
 get textContent(){return this._text+this.children.map(x=>x.textContent).join('');}
 set textContent(v){this._text=String(v);this.children=[];}
 setAttribute(k,v){this[k]=v;} addEventListener(n,f){(this.events[n] ||= []).push(f);}
 append(...n){this.children.push(...n);} appendChild(n){this.append(n);return n;}
 replaceChildren(...n){this._text='';this.children=n;} querySelectorAll(){return this.children;}
 fire(n){return Promise.all((this.events[n]||[]).map(f=>f({type:n,target:this})));}
}
const actions=['load','index','query','read','next','rollback-index','rollback-evidence'];
const nodes=Object.fromEntries(['rag-panel','rag-status','rag-snapshot','rag-hits','rag-evidence','rag-question','rag-topk','file-name','file-dir',...actions.map(x=>'rag-'+x)].map(id=>[id,new Element(id)]));
nodes['rag-topk'].value='5';nodes['file-name'].textContent='report.txt';nodes['file-dir'].textContent='workspace';
const events={},observers=[],calls=[];let authenticated=true,modal=true,modalHook=()=>{},ensured=0;
const ID='a'.repeat(32),RID='b'.repeat(32),ERID='9'.repeat(32),SID='c'.repeat(32),ESID='e'.repeat(32),HASH='f'.repeat(64);
const hit=()=>({source:'report.txt',location:'line 1',excerpt:'<img onerror=x>',model:'fixture',source_fingerprint:HASH,id:1,doc_id:'d'.repeat(64),distance:0,rank:1});
const ip=()=>({ok:true,plan:{plan_id:ID,doc_id:'d'.repeat(64),source_fingerprint:HASH,chunk_count:1,embedding:{channel:'local',model:'fixture',dimension:2},provider_call_after_confirmation:true,requires_confirmation:true}});
const qp=()=>({ok:true,plan:{plan_id:ID,operation:'save_evidence',hit_count:1,byte_count:1000,confirmation_required:true,wrote_nothing:true,payload_sha256:HASH,expires_in_seconds:300}});
const saved=()=>({ok:true,result:{status:'saved',applied:true,wrote_nothing:false,save_id:SID,model:'fixture',dimension:2,revision:1,chunk_count:1,audit_count:1,uncertain:false}});
const evidence=()=>({ok:true,result:{status:'saved',applied:true,wrote_nothing:false,verified:true,save_id:ESID,evidence:'host-relative/evidence.json',payload_sha256:HASH,hits:[hit()]}});
const snapshot=()=>({ok:true,snapshot:{exists:true,model:'fixture',dimension:2,revision:1,chunk_count:1,audit_count:1,uncertain:false}});
const read=()=>({ok:true,evidence:{text:'json plaintext',offset:0,has_more:true}});
const methods={get_rag_snapshot:async()=>snapshot(),plan_rag_index:async()=>ip(),confirm_rag_index:async()=>saved(),plan_rag_query:async()=>qp(),confirm_rag_query:async()=>evidence(),get_rag_evidence:async()=>read(),plan_rag_rollback:async(kind)=>({ok:true,plan:kind==='index'?{rollback_id:RID,operation:'rollback',confirmation_required:true,revision:1,chunk_count:1,first_absent_restored_empty:false}:{rollback_id:ERID,operation:'remove_evidence',confirmation_required:true,wrote_nothing:true,expires_in_seconds:300}}),confirm_rag_rollback:async()=>({ok:true,result:{status:'rolled_back',restored:true,verified:true,wrote_nothing:false}})};
const raw=Object.fromEntries(Object.keys(methods).map(k=>[k,async(...a)=>{calls.push([k,...a]);return methods[k](...a);} ]));
const window={pywebview:{api:raw},OfficeAuth:{isAuthenticated:()=>authenticated,wrapApi:a=>a,ensure:async()=>{ensured++;authenticated=true;}},addEventListener:(n,f)=>(events[n] ||= []).push(f),OfficeConfirm:{ask:async()=>{calls.push(['modal']);modalHook();return modal;}}};
const document={readyState:'complete',getElementById:id=>nodes[id]||null,createElement:()=>new Element(),addEventListener:()=>{}};
if(scenario==='initial_auth')authenticated=false;
if(scenario==='initial_api')delete window.pywebview;
vm.runInNewContext(source,{window,document,TextEncoder,MutationObserver:class{constructor(fn){this.fn=fn;}observe(el){observers.push([el,this.fn]);}}});
const tick=async()=>{for(let i=0;i<20;i++)await Promise.resolve();};
const click=async(a)=>{await nodes['rag-'+a].fire('click');await tick();};
const fire=async(n)=>{for(const f of events[n]||[])await f();await tick();};
const fill=(q='question')=>{nodes['rag-question'].value=q;nodes['rag-topk'].value='5';};
const query=async()=>{fill();await click('query');assert.equal(nodes['rag-read'].disabled,false,'query fixture must grant evidence');};
const deferred=()=>{let resolve;const promise=new Promise(r=>resolve=r);return {promise,resolve};};
const count=n=>calls.filter(x=>x[0]===n).length;
const retired=()=>{assert.equal(nodes['rag-question'].value,'');assert.equal(nodes['rag-hits'].textContent,'');assert.equal(nodes['rag-evidence'].textContent,'');for(const x of ['read','next','rollback-index','rollback-evidence'])assert(nodes['rag-'+x].disabled);};
const retire=async(kind)=>{if(kind==='auth'){authenticated=false;await fire('officeauth:unauthenticated');}else if(kind==='scope' || kind==='scope_name'){const field=kind==='scope'?'file-dir':'file-name';nodes[field].textContent='other';for(const [el,fn] of observers)if(el===nodes[field])fn();await tick();}else if(kind==='scope_text'){nodes['file-name'].textContent='other';}else if(kind==='revision' || kind==='topk'){const field=kind==='revision'?'rag-question':'rag-topk';nodes[field].value=kind==='revision'?'changed':'7';await nodes[field].fire('input');await tick();}else await fire(kind);};
(async()=>{
assert.equal(calls.length,0,'render must remain idle');
assert.equal(nodes['rag-panel'].open,true,'retrieval must remain visible after initialization');
nodes['rag-panel'].open=false;await fire('officerag:invalidate');
assert.equal(nodes['rag-panel'].open,true,'reset must not blank the retrieval tab');
if(scenario==='snapshot'){
 await click('load');assert.deepEqual(calls,[['get_rag_snapshot']]);assert(nodes['rag-snapshot'].textContent.includes('fixture'));assert(nodes['rag-read'].disabled);assert(nodes['rag-rollback-index'].disabled);
}else if(scenario==='initial_auth' || scenario==='initial_api'){
 for(const a of actions){assert(nodes['rag-'+a].disabled);await click(a);}assert.equal(calls.length,0);
}else if(scenario==='index'){
 await click('index');assert.deepEqual(calls,[['plan_rag_index'],['modal'],['confirm_rag_index',ID,true]]);assert.equal(nodes['rag-rollback-index'].disabled,false);assert(nodes['rag-read'].disabled);
 await click('rollback-index');assert.deepEqual(calls.slice(-3),[['plan_rag_rollback','index',SID],['modal'],['confirm_rag_rollback','index',RID,true]]);
}else if(scenario==='query_read'){
 await query();assert.deepEqual(calls,[['plan_rag_query','question',5],['modal'],['confirm_rag_query',ID,true]]);assert(nodes['rag-hits'].textContent.includes('<img onerror=x>'));assert.equal(nodes['rag-read'].disabled,false);
 await click('read');assert.deepEqual(calls.at(-1),['get_rag_evidence',ESID,0,2000]);assert.equal(nodes['rag-evidence'].textContent,'json plaintext');
 methods.get_rag_evidence=async()=>({ok:true,evidence:{text:'second page',offset:14,has_more:false}});await click('next');assert.deepEqual(calls.at(-1),['get_rag_evidence',ESID,14,2000]);assert.equal(nodes['rag-evidence'].textContent,'second page');assert(nodes['rag-next'].disabled);
}else if(scenario==='distinct'){
 await click('index');await query();await click('rollback-index');assert.equal(nodes['rag-read'].disabled,false);await click('rollback-evidence');assert.deepEqual(calls.filter(x=>x[0]==='plan_rag_rollback'),[['plan_rag_rollback','index',SID],['plan_rag_rollback','evidence',ESID]]);assert.deepEqual(calls.filter(x=>x[0]==='confirm_rag_rollback'),[['confirm_rag_rollback','index',RID,true],['confirm_rag_rollback','evidence',ERID,true]]);assert(nodes['rag-read'].disabled);
}else if(scenario.startsWith('modal_value:')){
 const [,action,value]=scenario.split(':');
 if(action.startsWith('rollback')){await click('index');await query();}else fill();
 modal=({false:false,null:null,zero:0,empty:'',string:'true',object:{},one:1})[value];
 methods.confirm_rag_index=async()=>({ok:false,reason:'confirmation_required'});
 methods.confirm_rag_query=async()=>({ok:true,result:{status:'rejected',applied:false,wrote_nothing:true}});
 methods.confirm_rag_rollback=async(k)=>k==='index'?{ok:false,reason:'confirmation_required'}:{ok:true,result:{status:'rejected',restored:false,wrote_nothing:true}};
 const before=calls.length;await click(action);
 const expected=action==='index'?['confirm_rag_index',ID,false]:action==='query'?['confirm_rag_query',ID,false]:['confirm_rag_rollback',action.slice(9),action==='rollback-index'?RID:ERID,false];
 assert.deepEqual(calls.at(-1),expected);assert.equal(calls.slice(before).filter(x=>x[0]==='modal').length,1);
 if(action==='index')assert(nodes['rag-rollback-index'].disabled);if(action==='query')assert(nodes['rag-read'].disabled);
 const again=calls.length;await click(action);assert(calls.length>again,'safe denial stays reusable');
}else if(scenario.startsWith('denial:')){
 const action=scenario.split(':')[1];
 if(action==='query'){modal=false;methods.confirm_rag_query=async()=>({ok:true,result:{status:'rejected',applied:false,wrote_nothing:true}});fill();}
 else{await click('index');await query();modal=false;methods.confirm_rag_rollback=async(k)=>k==='index'?{ok:false,reason:'confirmation_required'}:{ok:true,result:{status:'rejected',restored:false,wrote_nothing:true}};}
 await click(action);assert.equal(calls.at(-1).at(-1),false);if(action==='query'){assert(nodes['rag-read'].disabled);assert(nodes['rag-rollback-evidence'].disabled);}const before=calls.length;await click(action);assert(calls.length>before,'safe denial stays reusable');
}else if(scenario.startsWith('invalid:')){
 const v=scenario.split(':')[1];fill();const q={empty:'',space:'  ',long:'字'.repeat(201),surrogate:'\ud800',low:'\udc00'};if(v in q)nodes['rag-question'].value=q[v];else nodes['rag-topk'].value=({zero:'0',eleven:'11',fraction:'1.5',nan:'NaN',exponent:'1e0',blank:''})[v];await click('query');assert.equal(calls.length,0);
}else if(scenario==='unicode'){fill('😀'.repeat(200));await click('query');assert.deepEqual(calls[0],['plan_rag_query','😀'.repeat(200),5]);
}else if(scenario.startsWith('retire:')){await click('index');await query();await click('read');await retire(scenario.slice(7));retired();
}else if(scenario.startsWith('late:')){
 const [,stage,kind]=scenario.split(':');const d=deferred();let pending,result;
 if(stage==='index_plan'){methods.plan_rag_index=()=>d.promise;pending=click('index');result=ip();}else if(stage==='index_save'){methods.confirm_rag_index=()=>d.promise;pending=click('index');result=saved();}else if(stage==='snapshot'){methods.get_rag_snapshot=()=>d.promise;pending=click('load');result=snapshot();}else if(stage==='read'){await query();methods.get_rag_evidence=()=>d.promise;pending=click('read');result=read();}else if(stage==='plan'){fill();methods.plan_rag_query=()=>d.promise;pending=click('query');result=qp();}else{fill();methods.confirm_rag_query=()=>d.promise;pending=click('query');result=evidence();}
 await tick();assert.equal(count(({index_plan:'plan_rag_index',index_save:'confirm_rag_index',snapshot:'get_rag_snapshot',read:'get_rag_evidence',plan:'plan_rag_query',save:'confirm_rag_query'})[stage]),1);await retire(kind);d.resolve(result);await pending;if(stage==='index_plan'){assert.equal(count('modal'),0);assert.equal(count('confirm_rag_index'),0);}if(stage.startsWith('index_'))assert(nodes['rag-rollback-index'].disabled);if(stage==='plan'){assert.equal(count('modal'),0);assert.equal(count('confirm_rag_query'),0);}assert.equal(nodes['rag-evidence'].textContent,'');if(stage==='read' && ['revision','topk','scope_text'].includes(kind)){assert(nodes['rag-hits'].textContent.includes('<img onerror=x>'));assert.equal(nodes['rag-read'].disabled,false);}else{assert.equal(nodes['rag-hits'].textContent,'');assert(nodes['rag-read'].disabled);assert(nodes['rag-rollback-evidence'].disabled);}if(stage==='snapshot')assert.equal(nodes['rag-snapshot'].textContent,'');
}else if(scenario.startsWith('modal_guard:')){
 const [,action,kind]=scenario.split(':');if(action.startsWith('rollback')){await click('index');await query();}else fill();modalHook=()=>{if(kind==='auth')authenticated=false;else if(kind==='scope')nodes['file-name'].textContent='different';else{nodes['rag-question'].value='changed';nodes['rag-question'].fire('input');}};const name=action==='index'?'confirm_rag_index':action==='query'?'confirm_rag_query':'confirm_rag_rollback',before=count(name),modals=count('modal');await click(action);assert.equal(count('modal'),modals+1);assert.equal(count(name),before);
}else if(scenario.startsWith('malformed:')){
 const k=scenario.split(':')[1];fill();if(k==='plan_hits')methods.plan_rag_query=async()=>{const p=qp();p.plan.hits=[hit()];return p;};else if(k==='plan_id')methods.plan_rag_query=async()=>{const p=qp();p.plan.plan_id='../bad';return p;};else methods.confirm_rag_query=async()=>{const p=evidence();if(k==='fulltext')p.result.hits[0].fulltext='private full document';else if(k==='excerpt')p.result.hits[0].excerpt='x'.repeat(201);else p.result.save_id='../bad';return p;};await click('query');assert(nodes['rag-read'].disabled);assert(nodes['rag-rollback-evidence'].disabled);assert.equal(nodes['rag-hits'].textContent,'');if(k.startsWith('plan')){assert.equal(count('modal'),0);assert.equal(count('confirm_rag_query'),0);}
}else if(scenario.startsWith('uncertain:')){
 const k=scenario.split(':')[1];if(k==='snapshot'){methods.get_rag_snapshot=async()=>{const p=snapshot();p.snapshot.uncertain=true;return p;};await click('load');}else{methods.confirm_rag_index=async()=>{const p=saved();if(k==='outcome')p.outcome_uncertain=true;else if(k==='nested')p.result.uncertain=true;else if(k==='contradictory')p.result.wrote_nothing=true;else p.result.status='uncertain';return p;};await click('index');}assert(nodes['rag-rollback-index'].disabled);assert(nodes['rag-read'].disabled);await fire('officeauth:unauthenticated');await fire('officeauth:authenticated');const before=calls.length;fill();await click('index');await click('query');assert.equal(calls.length,before,'uncertainty fuse survives auth');
}else if(scenario==='read_limit'){await query();methods.get_rag_evidence=async()=>({ok:true,evidence:{text:'x'.repeat(2001),offset:0,has_more:true}});await click('read');assert(nodes['rag-evidence'].textContent.length<=2000);
}else if(scenario==='safe_error'){methods.plan_rag_index=async()=>{throw Error('SECRET provider token C:\\private\\file');};await click('index');assert(!Object.values(nodes).some(n=>n.textContent.includes('SECRET')));
}else if(scenario==='readiness'){authenticated=false;delete window.pywebview;await fire('officeauth:unauthenticated');for(const a of actions)assert(nodes['rag-'+a].disabled);window.pywebview={api:raw};await fire('pywebviewready');assert.equal(ensured,1);assert.equal(nodes['rag-load'].disabled,false);assert.equal(calls.length,0);
}else if(scenario==='idle'){await fire('officeauth:authenticated');await fire('pywebviewready');assert.equal(calls.length,0);
}else throw Error('unknown scenario '+scenario);
process.stdout.write('RAG_DOM_OK');
})().catch(e=>{process.stderr.write(e.stack);process.exitCode=1;});
"""

SCENARIOS = [
    "idle", "snapshot", "initial_auth", "initial_api", "index", "query_read", "distinct", "unicode", "read_limit", "safe_error", "readiness",
    *[f"modal_value:{a}:{v}" for a in ("index", "query", "rollback-index", "rollback-evidence")
      for v in ("false", "null", "zero", "empty", "string", "object", "one")],
    *[f"denial:{a}" for a in ("query", "rollback-index", "rollback-evidence")],
    *[f"invalid:{v}" for v in ("empty", "space", "long", "surrogate", "low", "zero", "eleven", "fraction", "nan", "exponent", "blank")],
    *[f"retire:{e}" for e in ("auth", "scope", "scope_name", "officediagram:invalidate", "officerag:invalidate", "pagehide")],
    *[f"late:{s}:{k}" for s in ("index_plan", "index_save", "plan", "save", "read", "snapshot")
      for k in ("auth", "scope", "scope_text", "revision", "topk")],
    *[f"modal_guard:{a}:{k}" for a in ("index", "query", "rollback-index", "rollback-evidence") for k in ("auth", "scope", "revision")],
    *[f"malformed:{k}" for k in ("plan_hits", "plan_id", "fulltext", "excerpt", "save_id")],
    *[f"uncertain:{k}" for k in ("snapshot", "outcome", "nested", "contradictory", "status")],
]


@pytest.mark.skipif(NODE is None, reason="No existing Node runtime for executable RAG DOM contract")
@pytest.mark.parametrize("scenario", SCENARIOS)
def test_dom_contract(scenario):
    result = subprocess.run([NODE, "-e", PROBE, str(STATIC / "rag.js"), scenario],
                            capture_output=True, text=True, encoding="utf-8", timeout=15)
    assert result.returncode == 0, result.stderr
    assert result.stdout == "RAG_DOM_OK", "Pending promises must not count as completed probes"
