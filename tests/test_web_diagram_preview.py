"""Production entry and DOM protocol stubs for the conversation-driven diagram flow.

The diagram form is gone: a chart is produced by talking to the agent and is rendered
inline in the conversation. These checks pin that contract without providers or Office.
"""
import shutil
import subprocess
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / 'office_agent/webui/static'
BUNDLED = Path.home() / '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe'
NODE = shutil.which('node') or (str(BUNDLED) if BUNDLED.is_file() else None)


def test_diagram_is_conversation_driven_not_form_driven():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    js = (STATIC / 'diagram_preview.js').read_text(encoding='utf-8')
    app = (STATIC / 'app.js').read_text(encoding='utf-8')
    assert '<script src="diagram_preview.js"></script>' in html
    # No form controls remain; the only entry point is the conversation.
    for control in ('id="diagram-kind"', 'id="diagram-goal"', 'id="diagram-generate"',
                    'id="diagram-cancel"', 'id="diagram-render"', 'id="diagram-copy"',
                    'id="diagram-source"', 'id="diagram-image"', 'id="diagram-card"'):
        assert control not in html, f"{control} 应已移除（表单驱动）"
    # The conversation renderer must exist and stay safe.
    assert 'function appendDiagramMessage(' in app
    assert 'new Blob([svg], { type: "image/svg+xml" })' in app
    assert 'figure.src = url' in app
    assert 'innerHTML' not in js and 'fetch(' not in js and 'localStorage' not in js
    assert 'URL.revokeObjectURL' in app
    # Blobs are reclaimed on file change and on auth transitions.
    assert app.count('new Event("officediagram:invalidate")') >= 2
    assert 'releaseDiagramURLs()' in app


def test_inline_renderer_rejects_untrusted_svg():
    app = (STATIC / 'app.js').read_text(encoding='utf-8')
    # The inline renderer admits only a real <svg> root and a bounded size.
    assert '/^<svg\\s/.test(svg)' in app
    assert 'svg.length > 1_100_000' in app
    # A failed image load degrades to a sentence, not a broken frame.
    assert '图表图片未能显示，请重试。' in app


@pytest.mark.parametrize('scenario', ['success', 'logout', 'bad_svg', 'oversize'])
def test_dom_inline_diagram_lifecycle(tmp_path, scenario):
    assert NODE is not None, 'Node is required for deterministic production DOM checks'
    script = tmp_path / 'inline.js'
    script.write_text(r'''
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const nodes=new Map(),events=new Map(),created=[],revoked=[];
let active=true;
function mk(tag){return {tag,children:[],className:'',textContent:'',hidden:false,disabled:false,attrs:{},handlers:{},style:{},
 append(...n){this.children.push(...n)},addEventListener(n,f){this.handlers[n]=f},setAttribute(n,v){this.attrs[n]=v},
 fire(n){return this.handlers[n]?.({target:this})}}}
function node(id){if(!nodes.has(id))nodes.set(id,mk('div'));return nodes.get(id)}
const document={createElement:mk,getElementById:node,activeElement:null,body:{}};
const URL={createObjectURL:b=>{created.push(b);return 'blob:'+created.length},revokeObjectURL:u=>revoked.push(u)};
// Minimal slice of app.js: the inline diagram renderer plus its blob bookkeeping.
const src=fs.readFileSync(process.argv[2],'utf8');
const start=src.indexOf('const diagramObjectURLs');
const end=src.indexOf('function setLogFollowing(');
const slice=src.slice(start,end);
const ctx={document,URL,Blob,navigator:{clipboard:{writeText:async()=>{}}},
 window:{OfficeDiagramExport:{mount:h=>h.children.push(mk('btn')),ready(){},retire(){}},
         OfficeDiagramPreview:{ready(){},retire(){}}},
 logQueue:[],logQueueHead:0,logRecords:[],logLimit:2000,busy:false,requesting:false,modalOpen:false,currentRound:null,
 logFrame:null,logFollowing:true,logBatchSize:50,logScrollTarget:0,
 appendLog:function(t,x,n){ctx.__logged=n||null},refreshEmptyFlow(){},flushLogs(){},newFlowRound:()=>({number:1}),
 el:id=>node(id),setText(){},focusVisible(){},canFocus:()=>false};
vm.createContext(ctx);
vm.runInContext(slice,ctx);
const good={type:'diagram',kind:'flowchart',source:'flowchart TD\nA[Alpha]',text:'图表已生成（未保存文件）。',
 svg:'<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><text>Alpha</text></svg>',
 metadata:{source_sha256:'1'.repeat(64),svg_sha256:'2'.repeat(64)}};
const mode=process.argv[3];
if(mode==='logout'){
 ctx.appendDiagramMessage(good);assert.equal(created.length,1);
 ctx.releaseDiagramURLs();assert.deepEqual(revoked,['blob:1']);
}
else if(mode==='bad_svg'){
 const bad={...good,svg:'<script>alert(1)</script>'};
 ctx.appendDiagramMessage(bad);assert.equal(created.length,0);assert.equal(ctx.__logged,null);
}
else if(mode==='oversize'){
 const big={...good,svg:'<svg '+'a'.repeat(1_100_001)+'></svg>'};
 ctx.appendDiagramMessage(big);assert.equal(created.length,0);assert.equal(ctx.__logged,null);
}
else {
 ctx.appendDiagramMessage(good);
 assert.equal(created.length,1);
 const card=ctx.__logged;
 assert(card,'卡片应进入对话流');
 assert(card.children.some(c=>c.src==='blob:1'),'预览图应内联在消息内');
 const details=card.children.find(c=>c.tag==='details');
 assert(details,'源码应折起提供');
 assert(details.children[0].textContent.includes('查看图表源码'));
 assert(details.children[1].textContent.includes('flowchart TD'));
}
console.log('inline diagram passed');
''', encoding='utf-8')
    result = subprocess.run([NODE, str(script), str(STATIC / 'app.js'), scenario],
                            capture_output=True, text=True, encoding='utf-8', timeout=10)
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'inline diagram passed' in result.stdout


def test_preview_styles_are_scoped_and_mounted():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    css = (STATIC / 'theme.css').read_text(encoding='utf-8')
    assert '<link rel="stylesheet" href="diagram_preview.css">' in html
    assert '.diagram-message' in css and '.diagram-figure' in css
    assert '.diagram-source-text' in css
