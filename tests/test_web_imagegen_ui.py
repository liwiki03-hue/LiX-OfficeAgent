"""P3 frontend wiring: static safety checks plus the node DOM-stub probe (no browser, no network)."""
import json
import os
import shutil
import subprocess
from pathlib import Path

import pytest
import playwright

STATIC = Path(__file__).resolve().parents[1] / "office_agent" / "webui" / "static"
NODE = Path(playwright.__file__).resolve().parent / "driver" / ("node.exe" if os.name == 'nt' else 'node')
PROBE = Path(__file__).with_name("probes") / "imagegen_dom.cjs"


def test_index_loads_imagegen_before_app_and_file_exists():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    assert '<script src="imagegen.js"></script>' in html
    assert html.index("imagegen.js") < html.index('src="app.js"')
    assert html.index('src="images.js"') < html.index("imagegen.js")
    # 生图已改为对话驱动：不再有独立卡片 / 状态药丸 / 表单控件，只留说明入口。
    assert 'id="imagegen-card"' not in html and 'id="imagegen-body"' not in html
    assert 'id="imagegen-state"' not in html
    assert "生成、保存、撤销各自单独确认" in html
    assert "直接在对话里说明你要什么图" in html


def test_app_delegates_the_six_image_events():
    app = (STATIC / "app.js").read_text(encoding="utf-8")
    for event in ("image_plan", "image_result", "image_save_plan", "image_saved",
                  "image_rollback_plan", "image_rolled_back"):
        assert f'case "{event}":' in app
    assert "window.OfficeImageGen?.handleEvent(ev," in app
    # The delegation must pass a caller bound to the authenticated api, never raw pywebview.
    assert "api[name](...args)" in app


def test_imagegen_module_has_no_network_storage_or_html_injection():
    source = (STATIC / "imagegen.js").read_text(encoding="utf-8")
    for forbidden in ("fetch(", "XMLHttpRequest", "localStorage", "sessionStorage",
                      "innerHTML", "eval(", "new Function", "http://", "https://",
                      "window.pywebview", "document.write"):
        assert forbidden not in source, forbidden
    for required in ("confirm_image_generation", "confirm_image_save", "confirm_image_rollback",
                     "plan_image_save", "plan_image_rollback", "officeauth:unauthenticated",
                     "showModal", "OfficeImages?.appendImages"):
        assert required in source, required
    # Only the user's click may submit: the module never calls without an explicit dialog.
    assert 'cancel.focus()' in source


@pytest.mark.skipif(not NODE.is_file(), reason="Playwright bundled Node unavailable")
def test_dom_stub_probe_confirms_no_automatic_authorization(tmp_path):
    assert PROBE.is_file()
    report = tmp_path / "probe.json"
    completed = subprocess.run([str(NODE), str(PROBE), str(STATIC), str(report)],
                               capture_output=True, text=True, timeout=60)
    assert completed.returncode == 0, completed.stderr[-2000:]
    data = json.loads(report.read_text(encoding="utf-8"))
    assert data["ok"] is True, data.get("error")
    assert data["checks"], data
    assert all(data["checks"].values()), data["checks"]
    assert data["automatic_calls_after_render"] == 0
    assert data["automatic_calls_after_dialog_open"] == 0
    assert data["submitted_operations"] == [
        "confirm_image_generation:true", "confirm_image_generation:false",
        "plan_image_save:plan-1", "confirm_image_save:true",
        "plan_image_rollback:save-1", "confirm_image_rollback:true"]
    assert data["unknown_event_ignored"] is True
    assert data["logout_disabled_buttons"] >= 2
    assert data["network_or_storage_used"] is False


@pytest.mark.skipif(shutil.which('node') is None and not NODE.is_file(), reason='Node unavailable: deterministic DOM-stub dependency')
def test_new_task_retires_old_image_actions_and_open_dialog(tmp_path):
    probe = tmp_path / 'image_new_task.js'
    probe.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
class El {
 constructor(tag){this.tagName=tag.toUpperCase();this.children=[];this.className='';this.disabled=false;this.listeners={};this.open=false;}
 append(...nodes){for(const child of nodes){child.parent=this;this.children.push(child);}}
 all(){return this.children.flatMap(child=>[child,...(child.all?.()||[])]);}
 querySelectorAll(selector){return this.all().filter(child=>child.tagName===selector.toUpperCase());}
 addEventListener(name,fn){this.listeners[name]=fn;}
 setAttribute(){} focus(){} showModal(){this.open=true;}
 close(){this.open=false;this.listeners.close?.();}
 remove(){if(this.parent)this.parent.children=this.parent.children.filter(child=>child!==this);}
}
const flow=new El('div'), body=new El('body'), events={}, calls=[];
const document={body,createElement:tag=>new El(tag),createTextNode:value=>Object.assign(new El('span'),{textContent:String(value)}),
 querySelectorAll:selector=>selector==='.imagegen-item .imagegen-actions button'
   ?flow.all().filter(child=>child.tagName==='BUTTON'&&child.parent?.className.includes('imagegen-actions')):[]};
const window={addEventListener:(name,fn)=>events[name]=fn};
vm.runInNewContext(fs.readFileSync(process.argv[2],'utf8'),{window,document});
const api=window.OfficeImageGen, call=async (...args)=>{calls.push(args);return {ok:true};};
const plan={type:'image_plan',plan:{plan_id:'offline',effective_prompt:'local',approval_ttl_seconds:30,configuration:{}}};
const buttons=()=>document.querySelectorAll('.imagegen-item .imagegen-actions button');
const tick=()=>new Promise(resolve=>setImmediate(resolve));
(async()=>{
 api.handleEvent(plan,call,card=>flow.append(card));
 assert.equal(buttons().length,2);assert(buttons().every(button=>!button.disabled));
 void buttons()[1].onclick();assert.equal(calls.length,0);
 assert.equal(body.children.filter(child=>child.open).length,1);
 api.retirePending('新任务已受理');
 assert(buttons().every(button=>button.disabled));
 assert.equal(body.children.filter(child=>child.open).length,0);
 void buttons()[0].onclick();void buttons()[1].onclick();await tick();
 assert.equal(calls.length,0,'stale controls and an open dialog cannot submit');
 api.handleEvent(plan,call,card=>flow.append(card));
 assert(buttons().slice(-2).every(button=>!button.disabled),'new plan remains actionable');
 assert(!flow.all().some(child=>String(child.textContent).includes('同意后 30 秒内有效')));
 assert(flow.all().some(child=>String(child.textContent).includes('计划创建后 30 秒内有效')));
})().catch(error=>{console.error(error);process.exitCode=1;});
''', encoding='utf-8')
    runtime = str(NODE) if NODE.is_file() else shutil.which('node')
    result = subprocess.run([runtime, str(probe), str(STATIC / 'imagegen.js')],
                            capture_output=True, text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stderr
    assert 'OfficeImageGen?.retirePending(' in (STATIC / 'app.js').read_text(encoding='utf-8')


@pytest.mark.skipif(shutil.which('node') is None and not NODE.is_file(), reason='Node unavailable: deterministic DOM-stub dependency')
def test_image_result_errors_give_safe_next_steps(tmp_path):
    probe = tmp_path / 'image_errors.js'
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[2], 'utf8');
const ctx = {tag: name => ({tag: name, textContent: '', setAttribute: () => {}}), text: value => String(value).slice(0, 4000)};
vm.createContext(ctx);
vm.runInContext(source.slice(source.indexOf('function showResult('), source.indexOf('function confirmDialog(')), ctx);
const show = (code, extra = {}) => { const card = {append: node => card.status = node}; ctx.showResult(card, {ok:false, code, ...extra}); return card.status.textContent; };
for (const code of ['image_configuration_missing', 'image_configuration_invalid']) {
  const notice = show(code);
  assert(notice.includes('本次没有请求生图服务'));
  assert(notice.includes('重启后重新规划'));
  assert(!notice.includes('IMAGE_API_KEY'));
}
assert(show('image_approval_expired').includes('旧计划不能继续'));
assert(show('image_generation_failed').includes('可能计费'));
assert(!show('image_generation_failed').includes('没有请求生图服务'));
for (const code of ['image_audit_changed', 'image_configuration_missing', 'image_preview_unavailable']) {
  const warning = show(code, {provider_request_state:'attempt_committed'});
  assert(warning.includes('可能') && warning.includes('计费'));
  assert(!warning.includes('本次没有请求生图服务'));
}
assert(show('image_generation_failed', {reason:'stale', provider_request_state:'attempt_committed'}).includes('不要直接重复确认'));
assert(!show('image_generation_failed', {reason:'stale'}).includes('可能计费'));
assert(show(undefined, {provider_request_state:'attempt_unknown'}).includes('可能计费'));
console.log('image error notices: passed');
''', encoding='utf-8')
    runtime = str(NODE) if NODE.is_file() else shutil.which('node')
    result = subprocess.run([runtime, str(probe), str(STATIC / 'imagegen.js')],
                            capture_output=True, text=True, encoding='utf-8', timeout=20)
    assert result.returncode == 0, result.stderr
    assert 'image error notices: passed' in result.stdout
