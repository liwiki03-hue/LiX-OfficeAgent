(() => {
  'use strict';
  const copy=value=>structuredClone(value);
  function width(value){
    if(typeof value!=='number'||!Number.isFinite(value)||value<24||value>360||!Number.isInteger(value*4/3))throw Error('列宽为24至360磅，以0.75磅为步长');
    return {pt:value,px:Math.round(value*4/3),twips:Math.round(value*20)};
  }
  function capture(doc,position,column,revision){
    const node=doc.nodeAt(position);
    if(!node||node.type.name!=='table'||!node.attrs.sdBlockId||doc.resolve(position).depth!==0)throw Error('请将光标放入普通正文表格的一列');
    const json=copy(node.toJSON()),rows=json.content||[],props=json.attrs.tableProperties||{},grid=json.attrs.grid;
    if((json.attrs.tableLayout||props.tableLayout)!=='fixed')throw Error('首版列宽调整仅支持固定布局表格，请保留自动布局表格');
    if(props.tableWidth?.type==='pct')throw Error('百分比表格请保留原宽度');
    if(!rows.length||rows.length>30||!Array.isArray(grid)||!grid.length||grid.length>16||rows.length*grid.length>100||node.textContent.length>20000)throw Error('列宽调整限30行、16列且总计100格、2万字以内表格');
    if(!Number.isInteger(column)||column<0||column>=grid.length)throw Error('表格列位置无效');
    if(grid.some(item=>!item||!Number.isFinite(item.col)||item.col<=0||Object.keys(item).some(key=>key!=='col')))throw Error('表格列网格不明确，请保留原布局');
    const supported=new Set(['tableRow','tableCell','paragraph','run','text','tab','lineBreak','hardBreak']);
    node.descendants(child=>{if(!supported.has(child.type.name)||child.attrs.trackChange||child.marks?.some(mark=>/track|deletion|insertion/i.test(mark.type.name)))throw Error('嵌套、修订或特殊单元格暂不支持调整列宽');});
    for(const row of rows){
      const cells=row.content||[],rowProps=row.attrs?.tableRowProperties||{};
      if(row.type!=='tableRow'||cells.length!==grid.length||rowProps.gridBefore||rowProps.gridAfter||row.attrs?.trackChange)throw Error('非矩形表格暂不支持调整列宽');
      for(const [columnIndex,cell] of cells.entries()){
        const attrs=cell.attrs||{},properties=attrs.tableCellProperties||{};
        if(cell.type!=='tableCell'||(attrs.colspan||1)!==1||(attrs.rowspan||1)!==1||properties.gridSpan||properties.vMerge||attrs.cellWidthType==='pct'||attrs.widthType==='pct'||properties.cellWidth?.type==='pct'||attrs.widthUnit&&attrs.widthUnit!=='px')throw Error('合并或百分比单元格请保留原列宽');
        if(!Array.isArray(attrs.colwidth)||attrs.colwidth.length!==1||!Number.isFinite(attrs.colwidth[0])||attrs.colwidth[0]<=0)throw Error('单元格宽度不明确，请保留原布局');
        if(Math.round(attrs.colwidth[0]*15)!==grid[columnIndex].col||properties.cellWidth?.type==='dxa'&&Number(properties.cellWidth.value)!==grid[columnIndex].col)throw Error('列网格与单元格宽度不一致，请保留原布局');
      }
    }
    if(props.tableWidth?.type==='dxa'&&Number(props.tableWidth.value)!==grid.reduce((sum,item)=>sum+item.col,0))throw Error('表格总宽与列网格不一致，请保留原布局');
    const widths=new Set(rows.map(row=>row.content[column].attrs.colwidth[0]));
    return {doc,node,json,position,column,revision,rows:rows.length,cols:grid.length,current:widths.size===1?[...widths][0]*.75:null};
  }
  function validate(editor,selected,revision,value){
    const size=width(value);
    if(selected.revision!==revision||selected.doc!==editor.state.doc||editor.state.doc.nodeAt(selected.position)!==selected.node)throw Error('表格或草稿已改变，请重新打开列宽设置');
    if(selected.json.attrs.grid[selected.column].col===size.twips&&selected.json.content.every(row=>row.content[selected.column].attrs.colwidth[0]===size.px))throw Error('列宽没有实际变化');
    return size;
  }
  function verify(selected,doc,size){
    const node=doc.nodeAt(selected.position);
    if(!node||node.type.name!=='table'||node.nodeSize!==selected.node.nodeSize||doc.content.size!==selected.doc.content.size)throw Error('列宽操作产生非预期结构变化');
    selected.doc.forEach((old,pos)=>{if(pos!==selected.position&&!old.eq(doc.nodeAt(pos)))throw Error('列宽操作改变了表格以外的内容');});
    const before=copy(selected.json),after=copy(node.toJSON()),index=selected.column;
    if(after.attrs.grid?.length!==before.attrs.grid.length||after.attrs.grid[index]?.col!==size.twips)throw Error('列网格与目标宽度不匹配');
    const total=after.attrs.grid.reduce((sum,col)=>sum+col.col,0),totalWidth=after.attrs.tableProperties?.tableWidth;
    if(totalWidth?.type!=='dxa'||totalWidth.value!==total||after.attrs.userEdited!==true)throw Error('表格总宽派生结果不明确');
    after.attrs.grid[index]=before.attrs.grid[index];
    after.attrs.userEdited=before.attrs.userEdited;
    if(before.attrs.tableProperties?.tableWidth===undefined)delete after.attrs.tableProperties.tableWidth;
    else after.attrs.tableProperties.tableWidth=before.attrs.tableProperties.tableWidth;
    if(after.content?.length!==before.content.length)throw Error('列宽操作改变了表格行数');
    for(let row=0;row<before.content.length;row++){
      if(after.content[row].content?.length!==before.content[row].content.length)throw Error('列宽操作改变了表格列数');
      const old=before.content[row].content[index],current=after.content[row].content[index];
      if(current.attrs.colwidth?.length!==1||current.attrs.colwidth[0]!==size.px)throw Error('目标列存在漏改');
      current.attrs.colwidth=old.attrs.colwidth;
      if(row===0){
        if(current.attrs.tableCellProperties?.cellWidth?.type!=='dxa'||current.attrs.tableCellProperties.cellWidth.value!==size.twips)throw Error('单元格宽度与网格不一致');
        if(old.attrs.tableCellProperties?.cellWidth===undefined)delete current.attrs.tableCellProperties.cellWidth;
        else current.attrs.tableCellProperties.cellWidth=old.attrs.tableCellProperties.cellWidth;
      }
    }
    const strip=value=>{
      if(value?.type==='paragraph'&&value.attrs)delete value.attrs.sdBlockRev;
      for(const child of value?.content||[])strip(child);
    };
    strip(before);strip(after);
    if(JSON.stringify(before)!==JSON.stringify(after))throw Error('列宽操作包含未解释的内容或格式变化');
    const beforeWidth=selected.json.attrs.grid[index].col/20,beforeTotal=selected.json.attrs.grid.reduce((sum,col)=>sum+col.col,0)/20,afterTotal=total/20;
    return {column:index+1,rows:selected.rows,width_pt:size.pt,table_width_pt:afterTotal,
      before_pt:beforeWidth,delta_pt:size.pt-beforeWidth,total_before_pt:beforeTotal,total_after_pt:afterTotal,
      summary:'第'+(index+1)+'列 '+beforeWidth+'→'+size.pt+'磅；表格总宽 '+beforeTotal+'→'+afterTotal+'磅'};
  }
  globalThis.OfficeTableColumnLayout=Object.freeze({width,capture,validate,verify});
})();
