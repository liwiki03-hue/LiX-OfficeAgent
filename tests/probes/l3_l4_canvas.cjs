const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');
const path=require('node:path');
const output=process.argv[2];
const fail=error=>{const value={passed:false,error:String(error.stack)};if(output)fs.writeFileSync(output,JSON.stringify(value));console.error(value.error);process.exitCode=1;};
process.on('uncaughtException',fail);

const nodes=new Map(),windowEvents=new Map();
let body,document;
class Element {
  constructor(tag){this.tagName=tag.toUpperCase();this.children=[];this.parentElement=null;this.dataset={};this.attrs={};this.style={};this.listeners=new Map();this.value='';this.hidden=false;this.disabled=false;this.open=false;this.textContent='';this.classes=new Set();this.classList={add:(...values)=>values.forEach(v=>this.classes.add(v)),remove:(...values)=>values.forEach(v=>this.classes.delete(v)),contains:value=>this.classes.has(value),toggle:(value,force)=>{const next=force??!this.classes.has(value);next?this.classes.add(value):this.classes.delete(value);return next;}};}
  set id(value){this._id=value;nodes.set(value,this);}
  get id(){return this._id||'';}
  set className(value){this.classes=new Set(value.split(/\s+/));}
  get className(){return [...this.classes].join(' ');}
  get isConnected(){let node=this;while(node){if(node===body)return true;node=node.parentElement;}return false;}
  append(...children){for(const child of children){child.remove();child.parentElement=this;this.children.push(child);}}
  remove(){if(this.parentElement){const parent=this.parentElement;parent.children=parent.children.filter(child=>child!==this);this.parentElement=null;}}
  replaceChildren(...children){for(const child of [...this.children])child.remove();this.append(...children);}
  setAttribute(name,value){this.attrs[name]=String(value);if(name==='id')this.id=value;}
  getAttribute(name){return this.attrs[name]??null;}
  addEventListener(name,callback){if(!this.listeners.has(name))this.listeners.set(name,[]);this.listeners.get(name).push(callback);}
  emit(name,event={}){for(const callback of this.listeners.get(name)||[])callback(event);}
  focus(){document.activeElement=this;}
  select(){this.allTextSelected=true;}
  scrollIntoView(){}
  showModal(){this.open=true;this.focus();}
  close(){this.open=false;this.emit('close');}
  matches(selector){
    if(selector==='.canvas-grid button'){let parent=this.parentElement;while(parent&&!parent.classes.has('canvas-grid'))parent=parent.parentElement;return this.tagName==='BUTTON'&&!!parent;}
    if(selector==='[data-target]')return !!this.dataset.target;
    if(selector==='dialog[data-canvas-document]')return this.tagName==='DIALOG'&&this.dataset.canvasDocument!==undefined;
    const match=/^button\[data-target="([A-Z]+\d+)"\]$/.exec(selector);
    if(match)return this.tagName==='BUTTON'&&this.dataset.target===match[1];
    const option=/^option\[value="([a-z]+)"\]$/.exec(selector);
    if(option)return this.tagName==='OPTION'&&this.value===option[1];
    return false;
  }
  querySelectorAll(selector){const values=[];const walk=node=>{for(const child of node.children){if(child.matches(selector))values.push(child);walk(child);}};walk(this);return values;}
  querySelector(selector){return this.querySelectorAll(selector)[0]||null;}
}
body=new Element('body');
document={body,activeElement:body,createElement:tag=>new Element(tag),getElementById:id=>nodes.get(id)?.isConnected?nodes.get(id):null,querySelectorAll:selector=>body.querySelectorAll(selector)};
const byId=id=>document.getElementById(id);
for(const id of ['floating-root','floating-body','floating-conversation','floating-controls','floating-scroll','floating-input']){const node=new Element(id==='floating-input'?'textarea':'section');node.id=id;body.append(node);}
byId('floating-root').append(byId('floating-body'),byId('floating-conversation'));
byId('floating-conversation').append(byId('floating-scroll'),byId('floating-controls'));
byId('floating-controls').append(byId('floating-input'));
const copyWrites=[],calls=[];
let revision=0,prepareResolve,commitResolve;
const baseCells=[{id:'A1',value:'甲',type:'inlineStr',editable:true},{id:'B1',value:'乙',type:'inlineStr',editable:true},{id:'A2',value:'丙',type:'inlineStr',editable:true},{id:'B2',value:'丁',type:'inlineStr',editable:true},{id:'C1',value:'45658',type:'n',editable:true,display:{text:'2025-01-01',format:'yyyy-mm-dd'}}];
const makeView=()=>({kind:'xlsx',item_id:'sheet-0',items:[{id:'sheet-0',name:'实验'}],row:1,col:1,rows:2,cols:3,merges:[],protected:false,formula_creation_allowed:true,scope:'code test',cells:baseCells.map(cell=>({...cell}))});
const makeState=()=>({ok:true,document_id:'doc',draft_revision:revision,operation_count:revision?4:0,undo_count:revision?1:0,redo_count:0,dirty:!!revision,recovery_available:false});
const api={
  canvas_open:async()=>({...makeState(),view:makeView()}),
  canvas_view:async()=>({...makeState(),view:makeView()}),
  canvas_prepare_range:async(...args)=>{calls.push(['prepare',...args]);return new Promise(resolve=>{prepareResolve=resolve;});},
  canvas_commit_range:async(...args)=>{calls.push(['commit',...args]);if(!args[3])return {...makeState(),cancelled:true};return new Promise(resolve=>{commitResolve=()=>{revision++;for(const cell of baseCells)cell.value='新'+cell.id;resolve({...makeState(),group_operations:4});};});},
  canvas_copy_range:async(...args)=>{calls.push(['copy',...args]);return {ok:true,tsv:'原文\t数据\n',copy_scope:'原始内容'};}
};
const window={OfficeAuth:{wrapApi:()=>api},pywebview:{api},OfficeAgentElements:{iconButton:(node,icon,label)=>{node.setAttribute('aria-label',label);return node;}},addEventListener:(name,callback)=>windowEvents.set(name,callback)};
const context=vm.createContext({window,document,navigator:{clipboard:{writeText:async text=>copyWrites.push(text)}},matchMedia:()=>({matches:false}),setTimeout,clearTimeout,console});
const staticRoot=path.resolve(__dirname,'../../office_agent/webui/static');
for(const file of ['canvas_grid_selection.js','office_canvas.js'])vm.runInContext(fs.readFileSync(path.join(staticRoot,file),'utf8'),context,{filename:file});
const Grid=window.OfficeCanvasGrid;
assert.equal(Grid.rectangle('C3','A1').label,'A1:C3');
assert.equal(Grid.rectangle('IV10000').count,1);
assert.equal(Grid.covered(['A1:C2'],'B2'),true);
assert.equal(Grid.covered(['A1:C2'],'A1'),false);
assert.equal(Grid.covered(['Z200:XFD300'],'B2'),false);
assert.throws(()=>Grid.rectangle('A1','K10'),/100/);
for(const invalid of ['A0','A01','IW1','A10001','<script>'])assert.throws(()=>Grid.position(invalid));
const tick=()=>new Promise(resolve=>setImmediate(resolve));
const click=(node,extra={})=>node.onclick({shiftKey:false,...extra});
const target=id=>byId('canvas-surface').querySelector('button[data-target="'+id+'"]');
const findButton=(node,text)=>{const all=[];const walk=value=>{for(const child of value.children){if(child.tagName==='BUTTON')all.push(child);walk(child);}};walk(node);return all.find(value=>value.textContent===text);};
const dialogs=()=>document.querySelectorAll('dialog[data-canvas-document]');
const descendants=(node,tag)=>{const values=[];const walk=current=>{for(const child of current.children){if(child.tagName===tag)values.push(child);walk(child);}};walk(node);return values;};
const plan=()=>({ok:true,plan:{id:'range-plan-'+revision,revision,start:'A1',end:'B2',changes:baseCells.map(cell=>({target:cell.id,action:'text',before:cell.value,after:'新'+cell.id}))},changed_cells:4,unchanged_cells:0});

(async()=>{
  await window.OfficeNative.update({document:{id:'doc',version:'v1',ready:true,name:'实验.xlsx'},busy:false});
  assert.equal(target('C1').textContent,'2025-01-01');
  await click(target('C1'));
  assert.equal(byId('canvas-value').value,'45658');
  assert.match(target('C1').title,/显示格式 yyyy-mm-dd/);
  assert.match(byId('canvas-feedback').textContent,/显示：2025-01-01/);
  assert.match(byId('canvas-feedback').textContent,/原始值：45658/);
  await click(target('A1'));
  await click(target('B2'),{shiftKey:true});
  assert.equal(byId('canvas-position').textContent,'A1:B2');
  assert.equal(byId('canvas-value').readOnly,true);
  assert.equal(byId('canvas-bold').disabled,true);
  await click(byId('canvas-paste-range'));
  let node=dialogs()[0],area=descendants(node,'TEXTAREA')[0];
  area.value='一\t二\n三\t四';
  const pending=click(findButton(node,'预览区域'));await tick();
  await click(findButton(node,'预览区域'));
  assert.equal(calls.filter(row=>row[0]==='prepare').length,1);
  assert.equal(calls[0][4],'A1');assert.equal(calls[0][6],'text');
  prepareResolve(plan());await pending;
  const accepting=click(findButton(node,'采用到草稿'));await tick();
  await click(findButton(node,'采用到草稿'));
  assert.equal(calls.filter(row=>row[0]==='commit'&&row[4]===true).length,1);
  commitResolve();await accepting;
  assert.equal(revision,1);assert.equal(dialogs().length,0);
  await click(target('A1'));
  byId('canvas-value').value='未采用文字';byId('canvas-value').oninput();
  await click(target('B2'));
  assert.equal(byId('canvas-value').value,'未采用文字');
  byId('canvas-value').onkeydown({key:'Escape',preventDefault(){}});
  assert.equal(byId('canvas-value').value,'新A1');
  await click(byId('canvas-copy-range'));
  node=dialogs()[0];assert.equal(copyWrites.length,0);
  await click(findButton(node,'复制'));assert.equal(copyWrites[0],'原文\t数据\n');
  await click(findButton(node,'关闭'));
  await click(byId('canvas-paste-range'));
  node=dialogs()[0];descendants(node,'TEXTAREA')[0].value='新的待处理内容';
  const late=click(findButton(node,'预览区域'));await tick();
  const commitCount=calls.filter(row=>row[0]==='commit').length;
  window.OfficeNative.close();prepareResolve(plan());await late;
  assert.equal(calls.filter(row=>row[0]==='commit').length,commitCount);
  assert.equal(dialogs().length,0);assert.equal(byId('canvas-work'),null);
  const result={passed:true,kind:'Node DOM code regression, not browser/native acceptance',duplicate_prepare_blocked:true,duplicate_commit_blocked:true,input_preserved:true,clipboard_explicit:true,late_response_isolated:true};
  if(output)fs.writeFileSync(output,JSON.stringify(result));
  console.log(JSON.stringify(result));
})().catch(fail);
