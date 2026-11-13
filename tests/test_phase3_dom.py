import json
import shutil
import subprocess
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1]/'office_agent/webui/static'
BUNDLED = Path.home()/'.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe'
NODE = shutil.which('node') or (str(BUNDLED) if BUNDLED.is_file() else None)


@pytest.mark.skipif(NODE is None,reason='Node unavailable: deterministic DOM-stub dependency')
def test_phase3_dom_cancellation_auth_fence_and_safe_drafts(tmp_path):
    script = tmp_path/'probe.js'
    script.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const elements=new Map(), handlers=new Map();
class Element {
  constructor(tag='div'){this.tag=tag;this.children=[];this.listeners={};this.value='';this.textContent='';this.disabled=false;this.isConnected=true;}
  addEventListener(name,fn){this.listeners[name]=fn;}
  append(...nodes){this.children.push(...nodes);}
  remove(){this.removed=true;}
}
const get=id=>{if(!elements.has(id))elements.set(id,new Element());return elements.get(id);};
let authenticated=true, confirmCount=0, confirmResult=false, writes=[];
const token='a'.repeat(32);
const service={
  plan_general_memory:async(operation,value)=>({ok:true,result:{plan_id:token,operation,requires_confirmation:true,...value}}),
  confirm_general_memory:async(id,confirmed)=>{writes.push({id,confirmed});return {ok:true,result:{applied:false}};},
  general_memory_snapshot:async()=>({ok:true,result:{items:[]}}),
  plan_task_documents:async(names)=>({ok:true,result:{plan_id:token,files:names,requires_confirmation:true}}),
  confirm_task_documents:async(id,confirmed)=>{writes.push({id,confirmed});return {ok:false};}
};
const context={document:{getElementById:get,createElement:tag=>new Element(tag),createTextNode:text=>({textContent:text})},
  window:{OfficeAuth:{isAuthenticated:()=>authenticated},pywebview:{api:service},confirm:()=>{confirmCount++;return confirmResult;},addEventListener:(name,fn)=>handlers.set(name,fn)},console};
vm.createContext(context);vm.runInContext(fs.readFileSync(process.argv[2],'utf8'),context);
const settle=()=>new Promise(resolve=>setImmediate(resolve));
(async()=>{
 get('general-content').value='长期任务是整理季度汇总';get('general-category').value='task';
 get('general-save').listeners.click();await settle();await settle();
 assert.deepStrictEqual(writes,[{id:token,confirmed:false}]);assert.strictEqual(confirmCount,1);
 writes=[];let resolvePlan;
 service.plan_general_memory=()=>new Promise(resolve=>{resolvePlan=resolve;});
 get('general-save').listeners.click();await settle();
 authenticated=false;handlers.get('officeauth:unauthenticated')();
 resolvePlan({ok:true,result:{plan_id:token,operation:'remember',requires_confirmation:true,content:'长期任务是整理季度汇总',category:'task'}});
 await settle();assert.strictEqual(writes.length,0);assert.strictEqual(confirmCount,1);
 assert.strictEqual(get('general-content').value,'');assert.strictEqual(get('multi-names').value,'');
 authenticated=true;handlers.get('officeauth:authenticated')();
 get('multi-names').value='one.docx\ntwo.xlsx'.replace('\\n','\n');
 get('multi-select').listeners.click();await settle();await settle();
 assert.strictEqual(writes.length,1);assert.strictEqual(writes[0].confirmed,false);
 const app=fs.readFileSync(process.argv[3],'utf8');
 const section=app.slice(app.indexOf('let answerDraft = null;'),app.indexOf('function handleEvent(ev)'));
 const host=new Element();context.currentRound={};context.flowRoundNode=()=>host;context.newFlowRound=()=>({});
 vm.runInContext(section,context);
 context.answerDraftEvent({type:'answer_draft_start',stream_id:token});
 context.answerDraftEvent({type:'answer_draft_delta',stream_id:token,text:'<script>untrusted text</script>'});
 assert.strictEqual(host.children[0].children[1].textContent,'<script>untrusted text</script>');
 context.answerDraftEvent({type:'answer_draft_delta',stream_id:'b'.repeat(32),text:'late'});
 assert(!host.children[0].children[1].textContent.includes('late'));
 context.answerDraftEvent({type:'answer_draft_cancel',stream_id:token});assert(host.children[0].removed);
 assert(!section.includes('innerHTML'));
 console.log(JSON.stringify({ok:true,cancelled:true,late_auth_fenced:true,draft_text_only:true}));
})().catch(error=>{console.error(error);process.exitCode=1;});
''',encoding='utf-8')
    run = subprocess.run([NODE,str(script),str(STATIC/'phase3.js'),str(STATIC/'app.js')],capture_output=True,text=True,encoding='utf-8',timeout=20)
    assert run.returncode == 0, run.stderr
    assert json.loads(run.stdout)['ok'] is True
