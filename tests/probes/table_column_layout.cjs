const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const output=process.argv[2];
try{
  const context=vm.createContext({structuredClone});vm.runInContext(fs.readFileSync(path.resolve(__dirname,'../../office_agent/webui/static/table_column_layout.js'),'utf8'),context);
  const policy=context.OfficeTableColumnLayout;
  const paragraph=text=>({type:'paragraph',attrs:{sdBlockRev:1},content:[{type:'run',attrs:{},content:[{type:'text',text}]}]});
  const model=()=>({type:'table',attrs:{sdBlockId:'table',tableLayout:'fixed',tableProperties:{tableWidth:{type:'auto',value:0}},grid:[{col:2160},{col:2160}],userEdited:false},content:Array.from({length:2},(_,row)=>({type:'tableRow',attrs:{tableRowProperties:{}},content:Array.from({length:2},(_,col)=>({type:'tableCell',attrs:{colspan:1,rowspan:1,colwidth:[144],widthUnit:'px',tableCellProperties:{cellWidth:{type:'dxa',value:2160},borders:{bottom:'retained'}}},content:[paragraph('数据'+row+col)]}))}))});
  function document(json,outer='非目标'){
    const node={type:{name:'table'},attrs:json.attrs,nodeSize:100,textContent:'数据',toJSON:()=>json,descendants:callback=>{const walk=current=>{for(const child of current.content||[]){callback({type:{name:child.type},attrs:child.attrs||{},marks:child.marks||[]});walk(child);}};walk(json);}};
    const outside={eq:value=>value?.text===outer,text:outer};
    return {content:{size:110},resolve:()=>({depth:0}),nodeAt:pos=>pos===10?node:outside,forEach:callback=>{callback(outside,0);callback(node,10);}};
  }
  const before=model(),doc=document(before),selected=policy.capture(doc,10,0,3),size=policy.validate({state:{doc}},selected,3,120);
  const snapshot=JSON.stringify(before);selected.json.attrs.grid[0].col=999;assert.equal(JSON.stringify(before),snapshot);selected.json.attrs.grid[0].col=2160;
  const changed=structuredClone(before);changed.attrs.grid[0].col=2400;changed.attrs.userEdited=true;changed.attrs.tableProperties.tableWidth={type:'dxa',value:4560};
  for(let row=0;row<2;row++){
    changed.content[row].content[0].attrs.colwidth=[160];
    if(row===0)changed.content[row].content[0].attrs.tableCellProperties.cellWidth.value=2400;
    for(const cell of changed.content[row].content)cell.content[0].attrs.sdBlockRev=2;
  }
  const engineSnapshot=JSON.stringify(changed),proof=policy.verify(selected,document(changed),size);
  assert.equal(proof.width_pt,120);assert.equal(proof.before_pt,108);assert.equal(proof.delta_pt,12);
  assert.equal(proof.total_before_pt,216);assert.equal(proof.total_after_pt,228);
  assert.match(proof.summary,/第1列 108→120磅/);assert.match(proof.summary,/216→228磅/);
  assert.equal(JSON.stringify(changed),engineSnapshot);assert.equal(JSON.stringify(before),snapshot);
  assert.throws(()=>policy.validate({state:{doc}},selected,4,120),/改变/);
  for(const value of [23,361,120.1,'120',NaN])assert.throws(()=>policy.width(value));
  for(const mutate of [x=>x.content[0].content[0].attrs.colspan=2,x=>x.attrs.tableLayout='autofit',x=>x.content[0].content[0].attrs.widthType='pct',x=>x.attrs.tableProperties.tableWidth.type='pct',x=>x.attrs.grid[0].col=2880]){
    const invalid=model();mutate(invalid);assert.throws(()=>policy.capture(document(invalid),10,0,3));
  }
  const corruption=structuredClone(changed);corruption.content[0].content[0].attrs.tableCellProperties.borders.bottom='changed';assert.throws(()=>policy.verify(selected,document(corruption),size),/未解释/);
  assert.throws(()=>policy.verify(selected,document(changed,'误改'),size),/以外/);
  fs.writeFileSync(output,JSON.stringify({passed:true,scope:'pure transaction policy and shared attribute regression; not native acceptance'}));
}catch(error){fs.writeFileSync(output,JSON.stringify({passed:false,error:String(error.stack)}));process.exitCode=1;}
