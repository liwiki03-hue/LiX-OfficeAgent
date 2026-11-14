"""DOM protocol stubs for the dynamically mounted export controls.

Save/rollback buttons now live inside the diagram message rather than a fixed card.
These checks keep the human-approval contract intact without native GUI proof.
"""
import subprocess

import pytest

from test_web_diagram_preview import NODE, STATIC


def test_export_controls_are_mounted_per_message_and_human_only():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    js = (STATIC / 'diagram_export.js').read_text(encoding='utf-8')
    app = (STATIC / 'app.js').read_text(encoding='utf-8')
    assert '<script src="diagram_export.js"></script>' in html
    # Buttons are created on demand, not declared in the document.
    assert 'id="diagram-save"' not in html and 'id="diagram-rollback"' not in html
    assert 'mount(host)' in js
    assert '保存这张图' in js and '撤销已保存的图' in js
    assert 'OfficeConfirm' in js
    assert 'window.confirm' not in js
    assert 'plan_diagram_save(expected.source_sha256, expected.svg_sha256)' in js
    assert 'window.OfficeDiagramExport?.mount?.(exportActions)' in app
    for forbidden in ('innerHTML', 'fetch(', 'localStorage', 'response.svg', 'eval('):
        assert forbidden not in js


@pytest.mark.parametrize('scenario', ['save_rollback', 'deny_save', 'deny_rollback', 'logout_pending',
                                      'double_click', 'unknown_after_approval', 'partial_rollback',
                                      'malformed_plan', 'unauthenticated', 'bad_metadata'])
def test_export_dom_approvals_and_late_results(tmp_path, scenario):
    assert NODE is not None, 'Node required for deterministic DOM checks'
    script = tmp_path / 'export_dom.js'
    script.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const events=new Map(); let active=true,resolvePlan,plans=0,confirms=[],undoConfirms=[],mode=process.argv[3];
const token='1'.repeat(32),rollback='2'.repeat(32),sha='3'.repeat(64),svg='4'.repeat(64);
function mk(){return {children:[],className:'',textContent:'',disabled:false,handlers:{},attrs:{},
 addEventListener(n,f){this.handlers[n]=f},fire(n){return this.handlers[n]?.()},
 setAttribute(n,v){this.attrs[n]=v},removeAttribute(n){delete this.attrs[n]},
 append(...n){this.children.push(...n)}}}
const window={OfficeAuth:{isAuthenticated:()=>active},addEventListener:(n,f)=>events.set(n,f),OfficeConfirm:{ask:async()=>mode!=='deny_save'}};
const plan={ok:true,plan:{plan_id:token,selected_document:'C:\\owned\\selected.docx',destination_basename:'office-diagram-'+token+'.svg'}};
const api={
plan_diagram_save:(...args)=>{assert.deepEqual(args,[sha,svg]);plans++;if(['logout_pending','double_click'].includes(mode))return new Promise(r=>resolvePlan=r);return Promise.resolve(mode==='malformed_plan'?{ok:true,plan:{...plan.plan,destination_basename:'../foreign.svg'}}:plan)},
confirm_diagram_save:async(...args)=>{confirms.push(args);if(mode==='unknown_after_approval')throw Error('PRIVATE FAILURE');return args[1]?{ok:true,result:{saved:true,save_id:token,path:'C:\\owned\\output.svg'}}:{ok:false,reason:'export_confirmation_required'}},
plan_diagram_rollback:async value=>{assert.equal(value,token);return {ok:true,plan:{rollback_id:rollback}}},
confirm_diagram_rollback:async(...args)=>{undoConfirms.push(args);return !args[1]?{ok:false,reason:'export_confirmation_required'}:{ok:true,result:mode==='partial_rollback'?{removed:true,audit_updated:null,code:'export_rollback_partial',service_disabled:true}:{removed:true,audit_updated:true}}}
};
vm.runInNewContext(fs.readFileSync(process.argv[2],'utf8'),{window,document:{createElement:mk,getElementById:()=>null},api});
const controls=window.OfficeDiagramExport,tick=()=>new Promise(r=>setImmediate(r));
const host=mk();
(async()=>{
controls.mount(host);const save=host.children[0],undo=host.children[1],status=host.children[2];
assert(save.disabled&&undo.disabled);
if(mode==='bad_metadata'){controls.ready({source_sha256:sha,svg_sha256:{}});await save.fire('click');assert.equal(plans,0);return}
controls.ready({source_sha256:sha,svg_sha256:svg});assert(!save.disabled);
if(mode==='unauthenticated'){active=false;await save.fire('click');assert.equal(plans,0);return}
const operation=save.fire('click');assert(save.disabled);
if(mode==='logout_pending'){active=false;events.get('officeauth:unauthenticated')();resolvePlan(plan)}
if(mode==='double_click'){await save.fire('click');assert.equal(plans,1);resolvePlan(plan)}
await operation;
if(['logout_pending','malformed_plan'].includes(mode)){assert.equal(confirms.length,0);return}
assert.equal(confirms.length,1);assert.equal(confirms[0][0],token);
if(mode==='deny_save'){assert.equal(confirms[0][1],false);assert(status.textContent.includes('已取消，没有新建文件'));return}
assert.equal(confirms[0][1],true);
if(mode==='unknown_after_approval'){assert(status.textContent.includes('结果没法确认'));assert(!status.textContent.includes('PRIVATE FAILURE'));assert(undo.disabled);return}
assert(!undo.disabled);
if(mode==='deny_rollback')window.OfficeConfirm.ask=async()=>false;
await undo.fire('click');assert.equal(undoConfirms.length,1);assert.equal(undoConfirms[0][0],rollback);
if(mode==='deny_rollback'){assert.equal(undoConfirms[0][1],false);assert(status.textContent.includes('已取消，文件没有删除'));return}
assert.equal(undoConfirms[0][1],true);
if(mode==='partial_rollback'){assert(status.textContent.includes('记录更新=null'));assert(save.disabled)}
else assert(status.textContent.includes('记录已更新并保留'));
assert(undo.disabled);
})().then(()=>console.log('export DOM passed')).catch(e=>{console.error(e);process.exitCode=1});
''', encoding='utf-8')
    result = subprocess.run([NODE, str(script), str(STATIC / 'diagram_export.js'), scenario],
                            capture_output=True, text=True, encoding='utf-8', timeout=10)
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'export DOM passed' in result.stdout


def test_new_diagram_retires_previous_export_authority(tmp_path):
    assert NODE is not None
    script = tmp_path / 'retire_dom.js'
    script.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
function mk(){return {children:[],className:'',textContent:'',disabled:false,handlers:{},attrs:{},
 addEventListener(n,f){this.handlers[n]=f},fire(n){return this.handlers[n]?.()},
 setAttribute(n,v){this.attrs[n]=v},removeAttribute(n){delete this.attrs[n]},
 append(...n){this.children.push(...n)}}}
const window={OfficeAuth:{isAuthenticated:()=>true},addEventListener(){},OfficeConfirm:{ask:async()=>true}};
const api={plan_diagram_save:async()=>({ok:true,plan:{plan_id:'1'.repeat(32),selected_document:'C:\\x.docx',destination_basename:'office-diagram-'+'1'.repeat(32)+'.svg'}}),
confirm_diagram_save:async()=>({ok:true,result:{saved:true,save_id:'1'.repeat(32),path:'C:\\out.svg'}}),
plan_diagram_rollback:async()=>({ok:true,plan:{rollback_id:'2'.repeat(32)}}),confirm_diagram_rollback:async()=>({ok:true,result:{removed:true,audit_updated:true}})};
vm.runInNewContext(fs.readFileSync(process.argv[2],'utf8'),{window,document:{createElement:mk,getElementById:()=>null},api});
const controls=window.OfficeDiagramExport,tick=()=>new Promise(r=>setImmediate(r));
const first=mk(),second=mk();
(async()=>{
controls.mount(first);controls.ready({source_sha256:'a'.repeat(64),svg_sha256:'b'.repeat(64)});
assert(!first.children[0].disabled);
// A newer diagram takes over: the first message can no longer save.
controls.mount(second);controls.ready({source_sha256:'c'.repeat(64),svg_sha256:'d'.repeat(64)});
assert(first.children[0].disabled,'旧图表不应再能导出');
assert(!second.children[0].disabled);
})().then(()=>console.log('retire passed')).catch(e=>{console.error(e);process.exitCode=1});
''', encoding='utf-8')
    result = subprocess.run([NODE, str(script), str(STATIC / 'diagram_export.js')],
                            capture_output=True, text=True, encoding='utf-8', timeout=10)
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'retire passed' in result.stdout
