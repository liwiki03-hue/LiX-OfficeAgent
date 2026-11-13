(() => {
  'use strict';
  function column(number) {
    if (!Number.isInteger(number)||number<1||number>256) throw Error('列位置超出A—IV范围');
    let value='';
    while(number){const digit=(number-1)%26;value=String.fromCharCode(65+digit)+value;number=Math.floor((number-1)/26);}
    return value;
  }
  function parsePosition(value,maxRows,maxCols) {
    const match=/^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(value||'');
    if(!match) throw Error('请输入A1形式的单元格位置');
    let col=0;for(const char of match[1])col=col*26+char.charCodeAt(0)-64;
    const row=Number(match[2]);
    if(row>maxRows||col>maxCols)throw Error('单元格位置超出支持范围');
    return [row,col];
  }
  function position(value){return parsePosition(value,10000,256);}
  function rectangle(anchor,focus=anchor){
    const [a,b]=position(anchor),[c,d]=position(focus);
    const top=Math.min(a,c),left=Math.min(b,d),bottom=Math.max(a,c),right=Math.max(b,d);
    const count=(bottom-top+1)*(right-left+1);
    if(count>100)throw Error('区域超过100格，请缩小选择后分组处理');
    const start=column(left)+top,end=column(right)+bottom;
    return {top,left,bottom,right,count,start,end,label:start===end?start:start+':'+end};
  }
  function includes(range,target){
    const [row,col]=position(target);
    return row>=range.top&&row<=range.bottom&&col>=range.left&&col<=range.right;
  }
  function covered(merges,target){
    const [row,col]=position(target);
    try{
      return merges.some(ref=>{
        const [start,end=start]=ref.split(':'),[r1,c1]=parsePosition(start,1048576,16384),[r2,c2]=parsePosition(end,1048576,16384);
        return row>=r1&&row<=r2&&col>=c1&&col<=c2&&(row!==r1||col!==c1);
      });
    }catch{return true;}
  }
  function indices(view){
    const read=(explicit,start,count,limit,budget)=>{
      const values=explicit||Array.from({length:count},(_,index)=>start+index);
      if(!Array.isArray(values)||!values.length||values.length>budget||values.some((value,index)=>!Number.isInteger(value)||value<1||value>limit||index&&value<=values[index-1]))throw Error('工作表视窗索引无效，请重新读取');
      return values;
    };
    return {rows:read(view.row_indices,view.row,view.rows,10000,30),cols:read(view.col_indices,view.col,view.cols,256,16)};
  }
  function adjacent(view,direction,visible=null){
    const current=indices(view),rows=view.panes?.rows||0,cols=view.panes?.cols||0;
    let row=view.row,col=view.col;
    const limits={down:10000-(30-rows)+1,right:256-(16-cols)+1};
    if(direction==='down'&&current.rows.at(-1)<10000)row=Math.min(limits.down,Math.max(rows+1,row+10));
    if(direction==='up'&&row>rows+1)row=Math.max(rows+1,row-10);
    if(direction==='right'&&current.cols.at(-1)<256)col=Math.min(limits.right,Math.max(cols+1,col+10));
    if(direction==='left'&&col>cols+1)col=Math.max(cols+1,col-10);
    if(visible){
      if(direction==='down')row=Math.min(row,Math.max(rows+1,visible.top-1));
      if(direction==='up')row=Math.max(row,visible.bottom+2-(30-rows)+1);
      if(direction==='right')col=Math.min(col,Math.max(cols+1,visible.left-2));
      if(direction==='left')col=Math.max(col,visible.right+2-(16-cols)+1);
      if(direction==='down'&&row<=Math.max(view.row,rows+1)||direction==='up'&&row>=view.row||direction==='right'&&col<=Math.max(view.col,cols+1)||direction==='left'&&col>=view.col)return null;
    }
    return row===view.row&&col===view.col?null:{row,col};
  }
  window.OfficeCanvasGrid={column,position,rectangle,includes,covered,indices,adjacent};
})();
