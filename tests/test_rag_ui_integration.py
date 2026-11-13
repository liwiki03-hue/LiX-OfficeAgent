"""Actual synthetic Api receipts consumed by JS; replay, not native GUI acceptance."""
import json
import os
import runpy
import shutil
import subprocess
from pathlib import Path

import pytest
import playwright

ROOT = Path(__file__).resolve().parents[1]
BUNDLED = Path(playwright.__file__).resolve().parent / 'driver' / ('node.exe' if os.name == 'nt' else 'node')
NODE = str(BUNDLED) if BUNDLED.is_file() else shutil.which('node')

PROBE = r"""
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const receipts=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
class Element {
 constructor(){this.value='';this.textContent='';this.children=[];this.events={};this.disabled=false;}
 setAttribute(k,v){this[k]=v;} addEventListener(k,v){this.events[k]=v;}
 replaceChildren(...children){this.children=children;}
}
const elements={}; const get=id=>elements[id]||=new Element();
const calls=[];const queue=receipts.map(x=>[x.method,x.args,x.response]);
const api=new Proxy({}, {get:(_,name)=>(...args)=>{calls.push([name,args]);
 const expected=queue.shift();assert(expected,'unexpected API');assert.equal(name,expected[0]);
 assert.deepEqual(args,expected[1]);return Promise.resolve(expected[2]);}});
const window={OfficeAuth:{isAuthenticated:()=>true,wrapApi:x=>x},pywebview:{api},
 addEventListener(){},confirm:()=>true};
const document={readyState:'complete',getElementById:get,createElement:()=>new Element(),addEventListener(){}};
vm.runInNewContext(fs.readFileSync(process.argv[1],'utf8'),{window,document,MutationObserver:class{observe(){}}});
async function click(name){assert.equal(get('rag-'+name).disabled,false,name+' disabled');await get('rag-'+name).events.click();}
(async()=>{
 assert.equal(calls.length,0);
 await click('load'); assert.match(get('rag-snapshot').textContent,/尚未建立/);
 await click('index');assert.equal(get('rag-rollback-index').disabled,false);
 get('rag-question').value='fixture';get('rag-topk').value='1';get('rag-question').events.input();
 await click('query');assert.equal(get('rag-hits').children.length,1);assert.equal(get('rag-read').disabled,false);
 await click('read');assert.equal(Array.from(get('rag-evidence').textContent).length,2000);
 await click('next');assert.ok(Array.from(get('rag-evidence').textContent).length<=2000);
 await click('rollback-evidence');assert.equal(get('rag-read').disabled,true);assert.equal(get('rag-rollback-index').disabled,false);
 await click('rollback-index');assert.equal(get('rag-rollback-index').disabled,true);
 assert.equal(queue.length,0);process.stdout.write('ok');
})().catch(e=>{process.stderr.write(e.stack);process.exitCode=1;});
"""


@pytest.mark.skipif(NODE is None, reason='Existing Node required for actual-Api/JS receipt compatibility')
@pytest.mark.parametrize('unicode_kind',['bmp','supplementary'])
def test_actual_api_receipt_compatibility(tmp_path, unicode_kind):
    make = runpy.run_path(str(ROOT / 'tests' / 'test_rag_api.py'))['make']
    api, app, auth, root, source, embedder, factories = make(tmp_path)
    if unicode_kind == 'supplementary':
        source.write_text('🙂' * 3000,encoding='utf-8')
    calls = []
    def record(method, *args):
        result = getattr(api, method)(*args)
        assert result['ok'], result
        calls.append({'method':method,'args':args,'response':result})
        return result
    record('get_rag_snapshot')
    plan = record('plan_rag_index')['plan']
    indexed = record('confirm_rag_index',plan['plan_id'],True)['result']
    plan = record('plan_rag_query','fixture',1)['plan']
    saved = record('confirm_rag_query',plan['plan_id'],True)['result']
    page = record('get_rag_evidence',saved['save_id'],0,2000)['evidence']
    assert page['has_more'] is True
    record('get_rag_evidence',saved['save_id'],len(page['text']),2000)
    for kind, target in [('evidence',saved['save_id']),('index',indexed['save_id'])]:
        plan = record('plan_rag_rollback',kind,target)['plan']
        record('confirm_rag_rollback',kind,plan['rollback_id'],True)
    evidence = tmp_path / 'receipts.json'
    evidence.write_text(json.dumps(calls,ensure_ascii=False),encoding='utf-8')
    result = subprocess.run([NODE,'-e',PROBE,str(ROOT/'office_agent/webui/static/rag.js'),str(evidence)],
                            capture_output=True,text=True,encoding='utf-8',timeout=15)
    assert result.returncode == 0, result.stderr
    assert result.stdout == 'ok', 'Pending promises cannot prove a completed probe'


SETTINGS_PROBE = r"""
const fs=require('fs'),vm=require('vm'),assert=require('assert'),mode=process.argv[3];
class Element{constructor(){this.value='';this.textContent='';this.children=[];this.events={};this.disabled=false;}
 addEventListener(k,v){this.events[k]=v;}setAttribute(k,v){this[k]=v;}replaceChildren(...x){this.children=x;}}
const nodes={},get=id=>nodes[id]||=new Element(),listeners={},notifications=[];
const id='a'.repeat(32),sid='b'.repeat(32),hash='f'.repeat(64);
const api={
 async plan_rag_index(){return {ok:true,plan:{plan_id:id,doc_id:hash,chunk_count:1,source_fingerprint:hash,embedding:{channel:'local',model:'fixture',dimension:2},requires_confirmation:true,provider_call_after_confirmation:true}};},
 async confirm_rag_index(){return {ok:true,result:{status:'saved',applied:true,wrote_nothing:false,save_id:sid,model:'fixture',dimension:2,revision:1,chunk_count:1,audit_count:1,uncertain:false}};},
 async get_settings(){if(mode==='transport')throw Error('not displayed');return {ok:true,settings:{provider:'deepseek',model:'fixture',base_url:'https://example.invalid',context_window:8192,embedding_channel:'local',embedding_model:'fixture',api_key_configured:false}};},
 async plan_settings_update(){return {ok:true,plan:{plan_id:id,changed_fields:['model']}};},
 async confirm_settings_update(){return {ok:true,result:mode==='uncertain'?{status:'uncertain',applied:null,wrote_nothing:false}:{status:'saved',applied:true,save_id:sid}};},
 async plan_settings_rollback(){return {ok:true,plan:{rollback_id:id}};},
 async confirm_settings_rollback(){return {ok:true,result:mode==='rollback-uncertain'?{status:'uncertain',restored:null,wrote_nothing:false}:{status:'rolled_back',restored:true}};}
};
const window={OfficeAuth:{isAuthenticated:()=>true,wrapApi:x=>x},pywebview:{api},confirm:()=>true,
 addEventListener(k,v){(listeners[k]||=[]).push(v);},
 dispatchEvent(event){if(event.type==='officerag:invalidate')notifications.push(event.ragUncertain);(listeners[event.type]||[]).forEach(fn=>fn(event));}};
const context={window,document:{readyState:'complete',getElementById:get,createElement:()=>new Element(),addEventListener(){}},MutationObserver:class{observe(){}},Event:class{constructor(type){this.type=type;}}};
vm.runInNewContext(fs.readFileSync(process.argv[1],'utf8'),context);
vm.runInNewContext(fs.readFileSync(process.argv[2],'utf8'),context);
(async()=>{
 await get('rag-index').events.click();assert.equal(get('rag-rollback-index').disabled,false);
 get('rag-question').value='private question';
 await get('settings-load').events.click();
 if(mode!=='transport')await get('settings-save').events.click();
 if(mode==='normal'||mode==='rollback-uncertain')await get('settings-rollback').events.click();
 assert.equal(get('rag-rollback-index').disabled,true);assert.equal(get('rag-question').value,'');
 assert.deepEqual(notifications,mode==='normal'?[false,false]:mode==='rollback-uncertain'?[false,true]:[true]);
 window.dispatchEvent(new context.Event('officeauth:authenticated'));
 assert.equal(get('rag-index').disabled,mode!=='normal');
 assert(!Object.values(nodes).some(x=>x.textContent.includes('not displayed')));
 process.stdout.write('ok');
})().catch(e=>{process.stderr.write(e.stack);process.exitCode=1;});
"""


@pytest.mark.skipif(NODE is None, reason='Existing Node required for settings/RAG event contracts')
@pytest.mark.parametrize('mode',['normal','uncertain','rollback-uncertain','transport'])
def test_settings_retire_and_uncertain_events(tmp_path,mode):
    result = subprocess.run([NODE,'-e',SETTINGS_PROBE,str(ROOT/'office_agent/webui/static/rag.js'),
                             str(ROOT/'office_agent/webui/static/settings.js'),mode],
                            capture_output=True,text=True,encoding='utf-8',timeout=15)
    assert result.returncode == 0, result.stderr
    assert result.stdout == 'ok', 'Pending promises cannot prove a completed probe'
