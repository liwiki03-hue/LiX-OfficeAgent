"""Configuration UI protocol: DOM stubs, temp local API fixtures, no native GUI."""
import json
import subprocess
from pathlib import Path

import pytest
from test_web_simplify import NODE, Structure

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'


class SettingsStructure(Structure):
    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in {'input', 'img', 'meta', 'link', 'br', 'hr'}:
            self.handle_endtag(tag)


def test_settings_form_mount_auth_scope_and_private_key_contract():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    dom = SettingsStructure(html)
    # 设置内容已从任务栏迁出：不再挂在侧栏里，而是收进独立设置窗口，按分区归档。
    assert dom.nodes['settings-dialog'][1]['role'] == 'dialog'
    assert dom.nodes['settings-dialog'][1]['aria-modal'] == 'true'
    assert 'hidden' in dom.nodes['settings-dialog'][1]
    assert 'hidden' in dom.nodes['settings-backdrop'][1]
    assert dom.nodes['nav-settings'][1]['aria-haspopup'] == 'dialog'
    assert dom.nodes['nav-settings'][1]['aria-controls'] == 'settings-dialog'
    # 旧的侧栏折叠容器必须彻底消失，否则内容还会被挤回任务栏。
    assert 'settings' not in dom.nodes, "侧栏 settings 折叠容器应已移除"
    # 分区导航与内容区一一对应。
    for tab, pane in (('tab-model', 'settings-pane-model'), ('tab-skills', 'settings-pane-skills'),
                      ('tab-rag', 'settings-pane-rag'), ('tab-machine', 'settings-pane-machine'),
                      ('tab-account', 'settings-pane-account')):
        assert tab in dom.nodes and pane in dom.nodes
        assert dom.nodes[tab][1]['aria-controls'] == pane
    # 每个设置控件都必须落在某个设置分区里，且不再挂在侧栏 nav 上。
    panes = ('settings-pane-model', 'settings-pane-skills', 'settings-pane-rag',
             'settings-pane-machine', 'settings-pane-account')
    for control in ('settings-key', 'settings-load', 'settings-save', 'settings-rollback',
                    'skills-load', 'rag-load', 'auth-logout'):
        ancestors = dom.nodes[control][2]
        assert 'settings-dialog' in ancestors, f"{control} 应位于设置窗口内"
        assert any(pane in ancestors for pane in panes), f"{control} 应归入某个设置分区"
        assert 'sidebar' not in ancestors, f"{control} 不应再留在任务栏里"
    assert dom.nodes['settings-key'][1]['type'] == 'password'
    assert dom.nodes['settings-key'][1]['autocomplete'] == 'off'
    assert 'disabled' in dom.nodes['settings-save'][1]
    assert 'disabled' in dom.nodes['settings-rollback'][1]
    # 折叠容器已由分区切换取代。
    assert 'model-settings' not in dom.nodes
    assert 'settings.js' in html and 'settings_dialog.js' in html
    source = (STATIC / 'settings.js').read_text(encoding='utf-8')
    for prohibited in ('innerHTML', 'localStorage', 'sessionStorage', 'fetch(', 'console.', 'eval('):
        assert prohibited not in source
    assert 'delete changes.api_key' in source
    assert 'window.confirm(' in source
    assert 'OfficeAuth.wrapApi' in source
    dialog = (STATIC / 'settings_dialog.js').read_text(encoding='utf-8')
    for prohibited in ('innerHTML', 'localStorage', 'sessionStorage', 'fetch(', 'console.', 'eval('):
        assert prohibited not in dialog
    assert "role=" not in dialog  # 语义写在 index.html 里，脚本不重建无障碍属性


@pytest.mark.skipif(NODE is None, reason='Node unavailable: deterministic DOM dependency')
@pytest.mark.parametrize('mode', ['save_rollback','deny_save','deny_rollback','nonliteral','logout_plan',
                                    'logout_approved','uncertain_save','partial_rollback','double_click',
                                    'edit_plan','malformed_id','clear_key','unauthenticated','login_enable','resumed_session',
                                     'auto_open','auto_closed'])
def test_settings_dom_confirmation_protocol(tmp_path, mode):
    probe = tmp_path / 'settings_probe.js'
    probe.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const mode=process.argv[3], nodes={}, events={}, documentEvents={}, calls=[], confirms=[];
const id='a'.repeat(32), saved='b'.repeat(32), rollback='c'.repeat(32);
let authenticated=!['unauthenticated','login_enable','resumed_session'].includes(mode), release;
const deferred=new Promise(resolve=>release=resolve);
function node(id) {return nodes[id] ||= {value:'',checked:false,disabled:false,open:false,textContent:'',events:{},addEventListener(name,fn){this.events[name]=fn;}};}
const api={
 async get_settings(){calls.push(['get']);return {ok:true,settings:{provider:'deepseek',model:'fixture',base_url:'https://example.invalid',context_window:8192,embedding_channel:'local',embedding_model:'',api_key_configured:true,api_key:'********'}};},
 async plan_settings_update(changes){calls.push(['plan',JSON.parse(JSON.stringify(changes))]);if(['logout_plan','edit_plan','double_click'].includes(mode)) await deferred;return {ok:true,plan:{plan_id:mode==='malformed_id'?'../foreign':id,changed_fields:Object.keys(changes)}};},
 async confirm_settings_update(plan,approved){calls.push(['confirm',plan,approved]);if(mode==='logout_approved') await deferred;return {ok:true,result:approved? mode==='uncertain_save'?{status:'uncertain',applied:null,wrote_nothing:false}:{status:'saved',applied:true,save_id:saved}:{status:'rejected',applied:false,wrote_nothing:true}};},
 async plan_settings_rollback(save){calls.push(['rollback_plan',save]);return {ok:true,plan:{rollback_id:rollback,restores_absent_source_as_empty_env:true}};},
 async confirm_settings_rollback(plan,approved){calls.push(['rollback_confirm',plan,approved]);return {ok:true,result:approved?mode==='partial_rollback'?{status:'uncertain',restored:null,wrote_nothing:false}:{status:'rolled_back',restored:true}:{status:'rejected',restored:false,wrote_nothing:true}};}
};
const window={OfficeAuth:{isAuthenticated:()=>authenticated,wrapApi:x=>x},pywebview:{api},confirm:text=>{assert(!text.includes('NewSecretFixture'));confirms.push(text);return mode==='nonliteral'?'true':mode==='deny_save'?false:mode==='deny_rollback'&&confirms.length===2?false:true;},addEventListener:(name,fn)=>events[name]=fn};
vm.runInNewContext(fs.readFileSync(process.argv[2],'utf8'),{window,document:{
  getElementById:id=>id==='model-settings'?null:node(id),
  addEventListener:(name,fn)=>documentEvents[name]=fn
}});
const click=id=>node(id).events.click();
const tick=()=>new Promise(resolve=>setImmediate(resolve));
(async()=>{
 if(mode==='auto_open'||mode==='auto_closed'){
  let opened=true, finish;
  window.OfficeSettings={isOpen:()=>opened};
  window.OfficeAuth.ensure=mode==='auto_open'?async()=>true:()=>new Promise(resolve=>{finish=resolve;});
  documentEvents['officesettings:opened']();
  if(mode==='auto_closed'){opened=false;finish(true);}
  await tick();
  assert.equal(calls.filter(c=>c[0]==='get').length,mode==='auto_open'?1:0);
  assert.equal(node('settings-model').disabled,mode==='auto_closed');
  if(mode==='auto_open')assert(node('settings-status').textContent.includes('已读取'));
  return;
 }
 if(mode==='unauthenticated'){await click('settings-load');await click('settings-save');assert.equal(calls.length,0);assert(node('settings-key').disabled);return;}
 if(mode==='login_enable'||mode==='resumed_session'){
  assert(node('settings-load').disabled, 'unauthenticated controls must start disabled');
  authenticated=true;
  if(mode==='login_enable')events['officeauth:authenticated']();
  else documentEvents['officesettings:opened']();
  assert(!node('settings-load').disabled, 'authenticated user can read configuration');
  assert(node('settings-save').disabled, 'saving requires an explicit read');
  await click('settings-load');
  assert(!node('settings-provider').disabled && !node('settings-model').disabled && !node('settings-key').disabled);
  node('settings-provider').value='dashscope';node('settings-provider').events.input();
  node('settings-model').value='fixture-model';node('settings-key').value='NewSecretFixture';
  await click('settings-save');
  const planned=calls.find(call=>call[0]==='plan')[1];
  assert.equal(planned.provider,'dashscope');assert.equal(planned.model,'fixture-model');
  assert.equal(planned.api_key,'NewSecretFixture');
  assert.equal(node('settings-key').value,'');
  assert(!node('settings-rollback').disabled);
  return;
 }
 await click('settings-load');assert.equal(node('settings-key').value,'');assert(!node('settings-key-state').textContent.includes('********'));assert(!node('settings-save').disabled);
 node('settings-key').value='NewSecretFixture';if(mode==='clear_key')node('settings-clear-key').checked=true;
 const pending=click('settings-save');await tick();
 if(mode==='double_click'){await click('settings-save');assert.equal(calls.filter(c=>c[0]==='plan').length,1);release();}
 if(mode==='logout_plan'){authenticated=false;events['officeauth:unauthenticated']();release();}
 if(mode==='edit_plan'){node('settings-model').events.input();release();}
 if(mode==='logout_approved'){assert.equal(calls.filter(c=>c[0]==='confirm').length,1);authenticated=false;events['officeauth:unauthenticated']();release();}
 await pending;
 assert.equal(node('settings-key').value,'');
 assert.equal(calls.find(c=>c[0]==='plan')[1].api_key,mode==='clear_key'?'':'NewSecretFixture');
 for(const el of Object.values(nodes))assert(!el.textContent.includes('NewSecretFixture'));
 const confirmation=calls.find(c=>c[0]==='confirm');
 if(['logout_plan','malformed_id'].includes(mode)){assert(!confirmation);assert.equal(confirms.length,0);return;}
 if(mode==='logout_approved'){assert(node('settings-rollback').disabled);assert.equal(node('settings-status').textContent,'尚未读取配置。');return;}
 if(['deny_save','nonliteral','edit_plan'].includes(mode)){assert.strictEqual(confirmation[2],false);assert(node('settings-rollback').disabled);return;}
 assert.strictEqual(confirmation[2],true);assert.equal(confirmation[1],id);
 if(mode==='uncertain_save'){assert(node('settings-rollback').disabled);assert(node('settings-status').textContent.includes('没能确认'));return;}
 assert(!node('settings-rollback').disabled);assert(node('settings-save').disabled);
 await click('settings-rollback');
 const rc=calls.find(c=>c[0]==='rollback_confirm');assert.equal(rc[1],rollback);assert(confirms[1].includes('恢复为空 .env'));
 if(mode==='deny_rollback'){assert.strictEqual(rc[2],false);assert(node('settings-status').textContent.includes('没有执行撤销'));}
 else if(mode==='partial_rollback'){assert(node('settings-status').textContent.includes('没能确认'));assert(node('settings-rollback').disabled);}
 else {assert.strictEqual(rc[2],true);assert(node('settings-status').textContent.includes('已恢复'));assert(node('settings-rollback').disabled);}
})().then(()=>console.log(JSON.stringify({ok:true,mode}))).catch(error=>{console.error(error);process.exitCode=1;});
''', encoding='utf-8')
    result = subprocess.run([NODE, str(probe), str(STATIC / 'settings.js'), mode], capture_output=True,
                            text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout) == {'ok': True, 'mode': mode}
