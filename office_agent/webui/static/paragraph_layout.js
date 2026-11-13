(() => {
  'use strict';
  const fields={before:[0,96],after:[0,96],left:[0,144],right:[0,144],firstLine:[-72,72]};
  function values(input){
    if(!input||typeof input!=='object'||Array.isArray(input)||Object.keys(input).some(key=>!Object.hasOwn(fields,key)))throw Error('段落布局参数无效');
    const result={};
    for(const [key,value] of Object.entries(input)){
      if(typeof value!=='number'||!Number.isFinite(value)||value<fields[key][0]||value>fields[key][1]||!Number.isInteger(value*20))throw Error('间距或缩进超出范围，最小精度为0.05磅');
      result[key]=Math.round(value*20);
    }
    if(!Object.keys(result).length)throw Error('请填写至少一项布局设置');
    return result;
  }
  function properties(current,input){
    const updates=values(input),result=structuredClone(current||{});
    for(const [key,value] of Object.entries(updates)){
      const spacing=['before','after'].includes(key),part=spacing?'spacing':'indent';
      const existing=result[part]||{};
      const competing=spacing?[key+'Lines']:['leftChars','rightChars','firstLineChars','hangingChars','start','end','startChars','endChars'];
      if(competing.some(name=>existing[name]!=null)||spacing&&[key+'Autospacing',key+'AutoSpacing'].some(name=>existing[name]!=null&&existing[name]!==false&&existing[name]!==0))throw Error('当前段落使用自动间距或字符缩进，请保留原设置');
      if(!spacing&&(result.numberingProperties||result.numPr||result.numbering||/list/i.test(result.styleId||'')))throw Error('列表缩进请使用编号工具；本入口仅调整普通段落缩进');
      result[part]={...existing};
      if(key==='firstLine'){
        delete result.indent.firstLine;delete result.indent.hanging;
        result.indent[value<0?'hanging':'firstLine']=Math.abs(value);
      }else result[part][key]=value;
    }
    return result;
  }
  function capture(doc,rows,ids,revision){
    const wanted=new Set(ids),targets=rows.filter(row=>wanted.has(row.id));
    if(!wanted.size||wanted.size>8||targets.length!==wanted.size)throw Error('请明确选择1至8个段落');
    return {doc,revision,targets:targets.map(row=>({id:row.id,position:row.position,node:doc.nodeAt(row.position)}))};
  }
  function prepare(editor,selection,revision,input){
    if(selection.revision!==revision||selection.doc!==editor.state.doc)throw Error('段落或选区已过期，请重新打开布局设置');
    const changes=selection.targets.map(target=>{
      const node=editor.state.doc.nodeAt(target.position);
      if(!node||node!==target.node||node.type.name!=='paragraph'||node.attrs.sdBlockId!==target.id)throw Error('段落身份已改变，请重新选择');
      return {target,node,properties:properties(node.attrs.paragraphProperties,input)};
    });
    if(changes.every(change=>JSON.stringify(change.properties)===JSON.stringify(change.node.attrs.paragraphProperties||{})))throw Error('布局没有实际变化');
    const transaction=editor.state.tr;
    for(const change of changes)transaction.setNodeMarkup(change.target.position,undefined,{...change.node.attrs,paragraphProperties:change.properties});
    return transaction;
  }
  globalThis.OfficeParagraphLayout=Object.freeze({values,properties,capture,prepare});
})();
