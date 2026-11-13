"""Independent embedding parameter UI contracts, no native GUI/providers/secrets."""
import json
import subprocess
import threading
from pathlib import Path

import pytest
from test_web_settings import SettingsStructure
from test_web_simplify import NODE

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'
FIELDS = {'dimension': ('embedding_dimension', 1, 4096), 'timeout': ('embedding_timeout', 1, 20), 'budget': ('embedding_budget', 1, 60)}


def test_numeric_fields_mounted_in_authenticated_settings():
    dom = SettingsStructure((STATIC / 'index.html').read_text(encoding='utf-8'))
    for name, (_, minimum, maximum) in FIELDS.items():
        tag, attrs, parents = dom.nodes['settings-embedding-' + name]
        assert tag == 'input' and attrs['type'] == 'number'
        assert attrs['min'] == str(minimum) and attrs['max'] == str(maximum)
        assert attrs['step'] == ('1' if name == 'dimension' else 'any')
        assert 'disabled' in attrs and attrs['autocomplete'] == 'off'
        # 设置内容已移出任务栏，收进独立设置窗口的「模型与检索」分区。
        assert all(scope in parents for scope in ('settings-dialog', 'settings-pane-model'))


PROBE = r"""
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const mode=process.argv[2], nodes={}, events={}, calls=[], modals=[], notifications=[];
const id='a'.repeat(32),sid='b'.repeat(32),rid='c'.repeat(32);
let authenticated=true,release,hook=()=>{},approved=true;
const deferred=new Promise(resolve=>release=resolve);
const fields=['dimension','timeout','budget'];
function node(id){return nodes[id] ||= {value:'',checked:false,textContent:'',disabled:false,events:{},addEventListener(k,v){this.events[k]=v;}};}
const receipts=process.argv[3]?JSON.parse(fs.readFileSync(process.argv[3],'utf8')):null;
let position=0;
async function actual(method,args){const row=receipts[position++];assert.equal(row.method,method);assert.deepEqual(JSON.parse(JSON.stringify(args)),row.args);return row.response;}
const api={
 async get_settings(){calls.push(['get']);return receipts?actual('get_settings',[]):{ok:true,settings:{provider:'deepseek',model:'fixture',base_url:'https://example.invalid',context_window:8192,embedding_channel:'local',embedding_model:'C:/models/fixture',embedding_dimension:128,embedding_timeout:2.5,embedding_budget:30,api_key_configured:false}};},
 async plan_settings_update(changes){calls.push(['plan',JSON.parse(JSON.stringify(changes))]);if(mode.startsWith('late:'))await deferred;return receipts?actual('plan_settings_update',[changes]):{ok:true,plan:{plan_id:id,changed_fields:Object.keys(changes),settings:{embedding_dimension:1024,embedding_timeout:1.25,embedding_budget:60}}};},
 async confirm_settings_update(plan,allow){calls.push(['confirm',plan,allow]);return receipts?actual('confirm_settings_update',[plan,allow]):{ok:true,result:allow?{status:'saved',applied:true,save_id:sid}:{status:'rejected',applied:false,wrote_nothing:true}};},
 async plan_settings_rollback(save){calls.push(['rollback_plan',save]);return receipts?actual('plan_settings_rollback',[save]):{ok:true,plan:{rollback_id:rid}};},
 async confirm_settings_rollback(plan,allow){calls.push(['rollback_confirm',plan,allow]);return receipts?actual('confirm_settings_rollback',[plan,allow]):{ok:true,result:allow?{status:'rolled_back',restored:true}:{status:'rejected',restored:false,wrote_nothing:true}};}
};
const window={OfficeAuth:{isAuthenticated:()=>authenticated,wrapApi:a=>a},pywebview:{api},addEventListener:(k,v)=>events[k]=v,
 dispatchEvent:e=>notifications.push(e.ragUncertain),confirm:msg=>{modals.push(msg);hook();return approved;}};
vm.runInNewContext(fs.readFileSync(process.argv[1],'utf8'),{window,document:{getElementById:id=>id==='model-settings'?null:node(id),addEventListener(){}},Event:class{constructor(type){this.type=type;}}});
const click=id=>node(id).events.click(),tick=()=>new Promise(resolve=>setImmediate(resolve));
const fill=()=>{node('settings-embedding-dimension').value='1024';node('settings-embedding-timeout').value='1.25';node('settings-embedding-budget').value='60';};
(async()=>{
 for(const field of fields)assert(node('settings-embedding-'+field).disabled);
 await click('settings-load');
 for(const [field,value] of [['dimension',128],['timeout',2.5],['budget',30]])assert.equal(Number(node('settings-embedding-'+field).value),value);
 fill();
 if(mode.startsWith('invalid:')){
   const [,field,value]=mode.split(':');node('settings-embedding-'+field).value=value;node('settings-key').value='not-a-real-private-key';
   await click('settings-save');assert.equal(calls.length,1);assert.equal(modals.length,0);assert.equal(node('settings-key').value,'');assert(node('settings-status').textContent.includes('检索参数不对'));
 }else if(mode==='blank'){
   fields.forEach(field=>node('settings-embedding-'+field).value='');await click('settings-save');const change=calls.find(c=>c[0]==='plan')[1];fields.forEach(field=>assert(!(('embedding_'+field) in change)));
 }else if(mode.startsWith('late:')){
   const field=mode.split(':')[1];const pending=click('settings-save');await tick();node('settings-embedding-'+field).value='2';node('settings-embedding-'+field).events.input();release();await pending;
   assert.equal(modals.length,0);assert.strictEqual(calls.find(c=>c[0]==='confirm')[2],false);assert(node('settings-rollback').disabled);
 }else if(mode.startsWith('modal:')){
   const kind=mode.split(':')[1];hook=()=>{if(kind==='auth')authenticated=false;else node('settings-embedding-'+kind).events.input();};await click('settings-save');
   const confirm=calls.find(c=>c[0]==='confirm');if(kind==='auth')assert(!confirm);else assert.strictEqual(confirm[2],false);assert(node('settings-rollback').disabled);
 }else if(mode==='retire'){
   node('settings-embedding-dimension').events.input();authenticated=false;events['officeauth:unauthenticated']();fields.forEach(field=>{assert.equal(node('settings-embedding-'+field).value,'');assert(node('settings-embedding-'+field).disabled);});
 }else{
   await click('settings-save');const change=calls.find(c=>c[0]==='plan')[1];assert.strictEqual(change.embedding_dimension,1024);assert.strictEqual(change.embedding_timeout,1.25);assert.strictEqual(change.embedding_budget,60);
   assert(modals[0].includes('维度：1024'));assert(modals[0].includes('1.25秒'));assert(modals[0].includes('60秒'));assert.deepEqual(notifications,[false]);assert.strictEqual(calls.find(c=>c[0]==='confirm')[2],true);
   hook=()=>{if(mode==='rollback_modal')authenticated=false;};await click('settings-rollback');
   if(mode==='rollback_modal')assert(!calls.find(c=>c[0]==='rollback_confirm'));else{assert.strictEqual(calls.find(c=>c[0]==='rollback_confirm')[2],true);assert.deepEqual(notifications,[false,false]);}
   if(receipts)assert.equal(position,receipts.length);
 }
 assert(!Object.values(nodes).some(n=>n.textContent.includes('not-a-real-private-key')));
 process.stdout.write('EMBEDDING_UI_OK');
})().catch(error=>{process.stderr.write(error.stack);process.exitCode=1;});
"""

SCENARIOS = ['save', 'blank', 'retire', 'rollback_modal',
             *[f'late:{field}' for field in FIELDS],
             *[f'modal:{field}' for field in (*FIELDS, 'auth')],
             *[f'invalid:{field}:{value}' for field in FIELDS for value in ('0', '-1', 'NaN', 'Infinity', '1e0', 'true', ' ', '01', '1.' , '9' * 33)],
             'invalid:dimension:4097', 'invalid:dimension:1.5', 'invalid:timeout:20.1', 'invalid:budget:60.1']


def run_probe(mode, receipts=None):
    args = [NODE, '-e', PROBE, str(STATIC / 'settings.js'), mode]
    if receipts is not None:
        args.append(str(receipts))
    result = subprocess.run(args, capture_output=True, text=True, encoding='utf-8', timeout=15)
    assert result.returncode == 0, result.stderr
    assert result.stdout == 'EMBEDDING_UI_OK', 'Pending promises cannot count as completed probes'


@pytest.mark.skipif(NODE is None, reason='Existing Node required for embedding DOM contracts')
@pytest.mark.parametrize('mode', SCENARIOS)
def test_embedding_numeric_dom(mode):
    run_probe(mode)


@pytest.mark.skipif(NODE is None, reason='Existing Node required for actual settings receipt replay')
def test_actual_settings_receipts_numeric_save_rollback(tmp_path):
    from office_agent.webui.settings import Settings
    original = b'MODEL_PROVIDER=deepseek\nDEEPSEEK_MODEL=fixture\nDEEPSEEK_BASE_URL=https://example.invalid\nOFFICE_CONTEXT_WINDOW=8192\nOFFICE_EMBEDDING_CHANNEL=local\nOFFICE_EMBEDDING_MODEL=C:/models/fixture\nOFFICE_EMBEDDING_DIMENSION=128\nOFFICE_EMBEDDING_TIMEOUT=2.5\nOFFICE_EMBEDDING_BUDGET=30\n'
    (tmp_path / '.env').write_bytes(original)
    facade = Settings(threading.RLock(), lambda: (True, False), tmp_path)
    rows = []
    def record(method, *args):
        response = getattr(facade, method)(*args)
        assert response['ok'], response
        rows.append({'method': method, 'args': args, 'response': response})
        return response
    record('snapshot')
    rows[-1]['method'] = 'get_settings'
    changes = {'provider': 'deepseek', 'model': 'fixture', 'base_url': 'https://example.invalid', 'context_window': 8192, 'embedding_channel': 'local', 'embedding_model': 'C:/models/fixture', 'embedding_dimension': 1024, 'embedding_timeout': 1.25, 'embedding_budget': 60}
    plan = record('plan_update', changes)['plan']
    rows[-1]['method'] = 'plan_settings_update'
    result = record('confirm_update', plan['plan_id'], True)['result']
    rows[-1]['method'] = 'confirm_settings_update'
    rollback = record('plan_rollback', result['save_id'])['plan']
    rows[-1]['method'] = 'plan_settings_rollback'
    record('confirm_rollback', rollback['rollback_id'], True)
    rows[-1]['method'] = 'confirm_settings_rollback'
    assert (tmp_path / '.env').read_bytes() == original
    receipts = tmp_path / 'receipts.json'
    receipts.write_text(json.dumps(rows, ensure_ascii=False), encoding='utf-8')
    run_probe('actual', receipts)
