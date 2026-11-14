"""Skills source and executable DOM contracts; no native UI/provider/Office."""
import os
import shutil
import subprocess
from pathlib import Path

import pytest
import playwright

STATIC = Path(__file__).resolve().parents[1] / "office_agent" / "webui" / "static"
BUNDLED_NODE = Path(playwright.__file__).resolve().parent / 'driver' / ('node.exe' if os.name == 'nt' else 'node')
NODE = str(BUNDLED_NODE) if BUNDLED_NODE.is_file() else shutil.which("node")


def test_html_mounts_skills_inside_authenticated_settings():
    from html.parser import HTMLParser
    class DOM(HTMLParser):
        def __init__(self):
            super().__init__(); self.stack=[]; self.nodes={}; self.scripts=[]
        def handle_starttag(self, tag, attrs):
            attrs=dict(attrs)
            if 'id' in attrs:
                assert attrs['id'] not in self.nodes
                self.nodes[attrs['id']]=(tag,attrs,tuple(self.stack))
            if tag=='script': self.scripts.append(attrs.get('src'))
            if tag not in {'input','meta','link','br','img','hr'}:
                self.stack.append(attrs.get('id',tag))
        def handle_startendtag(self, tag, attrs):
            self.handle_starttag(tag, attrs)
            if tag not in {'input','meta','link','br','img','hr'}:
                self.handle_endtag(tag)
        def handle_endtag(self, tag):
            if self.stack:self.stack.pop()
    dom=DOM();dom.feed((STATIC/'index.html').read_text(encoding='utf-8'))
    assert dom.nodes['skills-panel'][0]=='details'
    # 技能面板已从任务栏迁入独立设置窗口的「工作区技能」分区。
    assert {'settings-dialog', 'settings-pane-skills'} <= set(dom.nodes['skills-panel'][2])
    assert 'sidebar' not in dom.nodes['skills-panel'][2]
    assert 'hidden' in dom.nodes['workspace'][1] and 'inert' in dom.nodes['workspace'][1]
    for name in ('load','create','edit','delete','toggle','import','rollback'):
        tag,attrs,parents=dom.nodes['skills-'+name]
        assert tag=='button' and attrs['type']=='button' and 'skills-panel' in parents
        if name!='load': assert 'disabled' in attrs
    for name in ('slug','name','description','body','list','status'):
        assert 'skills-panel' in dom.nodes['skills-'+name][2]
    assert dom.nodes['skills-status'][1]['role']=='status'
    assert dom.scripts.count('skills.js')==1
    assert dom.scripts.index('auth.js')<dom.scripts.index('skills.js')


def test_controller_has_only_local_text_and_explicit_native_management():
    source = (STATIC / "skills.js").read_text(encoding="utf-8")
    for forbidden in ("innerHTML", "outerHTML", "insertAdjacentHTML", "localStorage",
                      "sessionStorage", "fetch(", "XMLHttpRequest", "eval(", "new Function",
                      "console.", "document.write", "execute_script", "run_script"):
        assert forbidden not in source
    for required in ("textContent", "createElement", "OfficeAuth.wrapApi", "OfficeConfirm",
                     "list_skills()", "get_skill(item.slug, item.source)", "plan_skill_change",
                     "plan_skill_import()", "confirm_skill_change", "plan_skill_rollback",
                     "confirm_skill_rollback", "officeauth:unauthenticated", "pywebviewready",
                     "officediagram:invalidate", "MutationObserver", "TextEncoder", "65536",
                     "JSON.stringify(name)", "JSON.stringify(description)"):
        assert required in source


PROBE = r"""
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[1], 'utf8'), scenario = process.argv[2];
class Element {
  constructor(id='') { this.id=id; this.value=''; this.textContent=''; this.children=[]; this.events={}; this.disabled=false; }
  setAttribute(name,value) { this[name]=value; }
  addEventListener(name,fn) { (this.events[name] ||= []).push(fn); }
  append(...nodes) { this.children.push(...nodes); }
  replaceChildren(...nodes) { this.children=nodes; }
  querySelectorAll() { return this.children; }
  fire(name) { return Promise.all((this.events[name] || []).map(fn => fn({type:name}))); }
}
const ids=['skills-panel','skills-list','skills-status','skills-slug','skills-name','skills-description','skills-body',
  ...['load','create','edit','delete','toggle','import','rollback'].map(x=>'skills-'+x),'file-name','file-dir'];
const nodes=Object.fromEntries(ids.map(id=>[id,new Element(id)])), events={}, observers=[];
let authenticated=true, raw={}, modal=true, modalHook=()=>{}, calls=[], modalTexts=[], ensured=0;
const id='a'.repeat(32), rid='b'.repeat(32), sid='c'.repeat(32);
const item={id:'user:alpha',source:'user',name:'<img onerror=x>',description:'human description',enabled:false};
const plan=(operation='create')=>({ok:true,plan:{plan_id:id,operation,id:'user:alpha',confirmation_required:true,wrote_nothing:true}});
const rollback=()=>({ok:true,plan:{rollback_id:rid,operation:'rollback',confirmation_required:true,wrote_nothing:true}});
const saved=()=>({ok:true,result:{status:'saved',applied:true,verified:true,save_id:sid}});
const methods={list_skills:async()=>({ok:true,skills:[item]}),get_skill:async()=>({ok:true,skill:{...item,body:'body <script>bad</script>'}}),
  plan_skill_change:async(op,changes)=>{const result=plan(op);if(op==='enabled')result.plan.enabled=changes.enabled;return result;},plan_skill_import:async()=>plan('import'),confirm_skill_change:async()=>saved(),
  plan_skill_rollback:async()=>rollback(),confirm_skill_rollback:async()=>({ok:true,result:{status:'rolled_back',restored:true,verified:true}})};
Object.keys(methods).forEach(key=>raw[key]=async(...args)=>{calls.push([key,...args]);return methods[key](...args);});
const window={pywebview:{api:raw},OfficeAuth:{isAuthenticated:()=>authenticated,wrapApi:api=>api,ensure:async()=>{ensured++;authenticated=true;}},
  addEventListener:(name,fn)=>(events[name] ||= []).push(fn),OfficeConfirm:{ask:async text=>{calls.push(['modal']);modalTexts.push(text);modalHook();return modal;}}};
const document={readyState:'complete',getElementById:id=>nodes[id]||null,createElement:()=>new Element(),addEventListener:()=>{}};
vm.runInNewContext(source,{window,document,TextEncoder,MutationObserver:class{constructor(fn){this.fn=fn;}observe(el){observers.push([el,this.fn]);}}});
const fire=async(name)=>{for(const fn of events[name]||[])await fn();};
const click=async(action)=>nodes['skills-'+action].fire('click');
const fill=()=>{nodes['skills-slug'].value='alpha';nodes['skills-name'].value='名 "\n字';nodes['skills-description'].value='描述: "value"';nodes['skills-body'].value='plain body';};
const load=async()=>{await click('load');assert.equal(nodes['skills-list'].children.length,1);};
const select=async()=>{await load();await nodes['skills-list'].children[0].fire('click');};
const deferred=()=>{let resolve;const promise=new Promise(r=>resolve=r);return {promise,resolve};};
const tick=async()=>{for(let i=0;i<8;i++)await Promise.resolve();};
(async()=>{
assert.equal(calls.length,0,'render must not issue native calls');
assert.equal(nodes['skills-panel'].open,true,'skills must remain visible after initialization');
nodes['skills-panel'].open=false;await fire('officeskills:invalidate');
assert.equal(nodes['skills-panel'].open,true,'reset must not blank the skills tab');
if(scenario==='create'){
  fill();await click('create');
  assert.deepEqual(calls.map(x=>x[0]),['plan_skill_change','modal','confirm_skill_change']);
  assert.equal(calls[0][1],'create');assert.deepEqual(Object.keys(calls[0][2]).sort(),['content','slug']);
  const text=calls[0][2].content;
  assert.equal(text,'---\nname: '+JSON.stringify('名 "\n字')+'\ndescription: '+JSON.stringify('描述: "value"')+'\n---\nplain body');
  assert.deepEqual(calls[2],['confirm_skill_change',id,true]);assert.equal(nodes['skills-rollback'].disabled,false);
  await click('rollback');assert.deepEqual(calls.slice(-3).map(x=>x[0]),['plan_skill_rollback','modal','confirm_skill_rollback']);
  assert.deepEqual(calls.at(-1),['confirm_skill_rollback',rid,true]);assert.equal(nodes['skills-rollback'].disabled,true);
}else if(scenario==='cancel'){
  modal=false;methods.confirm_skill_change=async()=>({ok:false,reason:'skills_denied'});fill();await click('create');
  assert.deepEqual(calls.at(-1),['confirm_skill_change',id,false]);assert.equal(nodes['skills-create'].disabled,false);
  assert(nodes['skills-status'].textContent.includes('取消'));
}else if(scenario==='metadata'){
  await select();assert.equal(nodes['skills-list'].children[0].textContent.includes('<img onerror=x>'),true);
  assert.equal(nodes['skills-body'].value,'body <script>bad</script>');assert.deepEqual(calls.at(-1),['get_skill','alpha','user']);
  await click('edit');const edit=calls.find(x=>x[0]==='plan_skill_change');assert.equal(edit[1],'edit');
  assert.deepEqual(Object.keys(edit[2]).sort(),['content','name','source']);assert.equal(edit[2].name,'alpha');
}else if(scenario==='builtin'){
  item.id='builtin:alpha';item.source='builtin';await select();
  assert(nodes['skills-edit'].disabled && nodes['skills-delete'].disabled);await click('edit');await click('delete');
  assert(!calls.some(x=>x[0]==='plan_skill_change'));await click('toggle');
  const enabled=calls.find(x=>x[0]==='plan_skill_change');assert.equal(enabled[1],'enabled');
  assert.deepEqual(enabled[2],{name:'alpha',source:'builtin',enabled:true});
}else if(scenario==='toggle_labels'){
  for(const enabled of [true,false]){
    item.enabled=enabled;await select();await click('toggle');
    assert(modalTexts.at(-1).startsWith(enabled?'确认停用技能？':'确认启用技能？'));
    assert(modalTexts.at(-1).includes('目标：user:alpha'));
    assert.deepEqual(calls.at(-1),['confirm_skill_change',id,true]);
  }
}else if(scenario==='invalid_toggle'){
  await select();methods.plan_skill_change=async()=>plan('enabled');await click('toggle');
  assert(!calls.some(x=>x[0]==='modal'||x[0]==='confirm_skill_change'));
  assert(nodes['skills-status'].textContent.includes('结果没能确认'));
  assert(nodes['skills-toggle'].disabled && nodes['skills-list'].children.length===0);
}else if(scenario==='delete'){
  await select();await click('delete');const entry=calls.find(x=>x[0]==='plan_skill_change');
  assert.equal(entry[1],'delete');assert.deepEqual(entry[2],{name:'alpha',source:'user'});
}else if(scenario==='import'){
  await click('import');assert.deepEqual(calls[0],['plan_skill_import']);assert.deepEqual(calls.at(-1),['confirm_skill_change',id,true]);
}else if(scenario==='import_cancel'){
  methods.plan_skill_import=async()=>({ok:true,cancelled:true});await click('import');assert.equal(calls.length,1);
}else if(scenario==='limits'){
  fill();for(const slug of ['../bad','A','a'.repeat(65),'a--b']){nodes['skills-slug'].value=slug;await click('create');}
  fill();nodes['skills-name'].value='字'.repeat(101);await click('create');
  fill();nodes['skills-description'].value='字'.repeat(501);await click('create');
  fill();nodes['skills-body'].value='字'.repeat(22000);await click('create');
  fill();nodes['skills-body'].value='\ud800';await click('create');assert.equal(calls.length,0);
}else if(['late_list','late_read'].includes(scenario)){
  const d=deferred();let pending;
  if(scenario==='late_list'){methods.list_skills=()=>d.promise;pending=click('load');}
  else {await load();methods.get_skill=()=>d.promise;pending=nodes['skills-list'].children[0].fire('click');}
  await tick();authenticated=false;await fire('officeauth:unauthenticated');
  d.resolve(scenario==='late_list'?{ok:true,skills:[item]}:{ok:true,skill:{...item,body:'late body'}});await pending;
  assert.equal(nodes['skills-list'].children.length,0);assert.equal(nodes['skills-body'].value,'');assert(nodes['skills-load'].disabled);
}else if(['late_plan','revision','scope','modal_scope','modal_auth','modal_revision'].includes(scenario)){
  fill();const d=deferred();methods.plan_skill_change=()=>d.promise;const pending=click('create');await tick();
  if(scenario==='late_plan')await fire('officediagram:invalidate');
  if(scenario==='revision'){nodes['skills-body'].value='changed';await nodes['skills-body'].fire('input');}
  if(scenario==='scope')nodes['file-dir'].textContent='changed workspace';
  if(scenario==='modal_scope')modalHook=()=>{nodes['file-name'].textContent='other file';};
  if(scenario==='modal_auth')modalHook=()=>{authenticated=false;};
  if(scenario==='modal_revision')modalHook=()=>{nodes['skills-body'].value='changed';nodes['skills-body'].fire('input');};
  d.resolve(plan());await pending;assert(!calls.some(x=>x[0]==='confirm_skill_change'));
  if(!scenario.startsWith('modal_'))assert(!calls.some(x=>x[0]==='modal'));
}else if(scenario==='late_save'){
  fill();const d=deferred();methods.confirm_skill_change=()=>d.promise;const pending=click('create');await tick();
  await fire('officediagram:invalidate');d.resolve(saved());await pending;
  assert(nodes['skills-rollback'].disabled);assert.equal(nodes['skills-body'].value,'');
}else if(scenario==='uncertain'){
  methods.confirm_skill_change=async()=>({ok:true,result:{status:'uncertain',applied:null}});fill();await click('create');
  assert(nodes['skills-status'].textContent.includes('没能确认'));assert(nodes['skills-rollback'].disabled);
  const count=calls.length;await click('create');await fire('officeauth:authenticated');await click('import');assert.equal(calls.length,count);
}else if(scenario==='nested_uncertain'){
  methods.confirm_skill_change=async()=>({ok:true,result:{status:'saved',applied:true,verified:true,save_id:sid,uncertain:true}});
  fill();await click('create');assert(nodes['skills-status'].textContent.includes('没能确认'));assert(nodes['skills-rollback'].disabled);
}else if(scenario==='rollback_scope'){
  fill();await click('create');modalHook=()=>{nodes['file-name'].textContent='different';};await click('rollback');
  assert(!calls.some(x=>x[0]==='confirm_skill_rollback'));
}else if(scenario==='readiness'){
  authenticated=false;delete window.pywebview;await fire('officeauth:unauthenticated');assert(nodes['skills-load'].disabled);
  window.pywebview={api:raw};await fire('pywebviewready');assert.equal(ensured,1);assert.equal(nodes['skills-load'].disabled,false);
}else if(scenario==='observer'){
  fill();await click('create');fill();nodes['file-name'].textContent='new';observers[0][1]();
  assert(nodes['skills-rollback'].disabled);assert.equal(nodes['skills-body'].value,'');
}else throw Error('unknown scenario');
process.stdout.write('ok');
})().catch(error=>{process.stderr.write(error.stack);process.exitCode=1;});
"""


@pytest.mark.skipif(NODE is None, reason="No existing Node runtime for executable DOM contract")
@pytest.mark.parametrize("scenario", [
    "create", "cancel", "metadata", "builtin", "toggle_labels", "invalid_toggle", "delete", "import", "import_cancel",
    "limits", "late_list", "late_read", "late_plan", "revision", "scope", "modal_scope",
    "modal_auth", "modal_revision", "late_save", "uncertain", "nested_uncertain", "rollback_scope", "readiness", "observer",
])
def test_dom_contract(scenario):
    result = subprocess.run([NODE, "-e", PROBE, str(STATIC / "skills.js"), scenario],
                            capture_output=True, text=True, encoding="utf-8", timeout=15)
    assert result.returncode == 0, result.stderr
    assert result.stdout == "ok", "Probe must finish; pending promises cannot count as passed"
