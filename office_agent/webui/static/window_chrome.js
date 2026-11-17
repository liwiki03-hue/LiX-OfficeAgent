(() => {
  'use strict';
  let header, controls, bar, maximum, locked = false, secondaryOpen = false, sequence = 0, resizeTimer, regionFrame;
  let regionBusy=false,regionQueued=false,lastRegionKey='';
  let pointerDrag=null,pointerBusy=false,pointerPending=null,pointerFrame=null;
  async function sendPointer(packet){
    if(pointerBusy){pointerPending=packet;return;}
    pointerBusy=true;
    try{
      const result=await api().floating_window_pointer(packet);
      if(!result?.ok)throw Error(result?.error||'窗口移动未完成');
      update(result);
    }catch(_){pointerDrag=null;pointerPending=null;fallbackButton();}
    finally{pointerBusy=false;if(pointerPending){const next=pointerPending;pointerPending=null;sendPointer(next);}}
  }
  function pointerPacket(phase,event){
    const drag=pointerDrag;
    if(!drag)return null;
    drag.screen=[event.screenX,event.screenY];
    return {id:drag.id,phase,seq:++drag.seq,screen:drag.screen,client:drag.client,viewport:drag.viewport};
  }
  function startPointer(event){
    if(pointerDrag||pointerBusy||pointerPending||event.button!==0||!event.isTrusted)return;
    const target=event.target.closest('#floating-title,#desktop-window-bar .desktop-drag-region,#settings-dialog .settings-head');
    if(!target||event.target.closest('button,input,select,textarea,a,label,summary,[contenteditable]'))return;
    event.preventDefault();
    pointerDrag={id:crypto.randomUUID(),seq:0,pointer:event.pointerId,target,client:[event.clientX,event.clientY],viewport:[innerWidth,innerHeight],screen:[event.screenX,event.screenY]};
    try{target.setPointerCapture(event.pointerId);}catch(_){}
    sendPointer(pointerPacket('start',event));
  }
  window.addEventListener('pointermove',event=>{
    if(!pointerDrag||event.pointerId!==pointerDrag.pointer)return;
    pointerDrag.screen=[event.screenX,event.screenY];
    if(pointerFrame)return;
    pointerFrame=requestAnimationFrame(()=>{pointerFrame=null;if(pointerDrag)sendPointer(pointerPacket('move',{screenX:pointerDrag.screen[0],screenY:pointerDrag.screen[1]}));});
  });
  function finishPointer(event){
    if(!pointerDrag||event.pointerId!==pointerDrag.pointer)return;
    if(pointerFrame){cancelAnimationFrame(pointerFrame);pointerFrame=null;}
    const drag=pointerDrag,packet=pointerPacket(event.type==='pointercancel'?'cancel':'end',event);
    pointerDrag=null;
    try{drag.target.releasePointerCapture(drag.pointer);}catch(_){}
    sendPointer(packet);
  }
  window.addEventListener('pointerup',finishPointer);
  window.addEventListener('pointercancel',finishPointer);
  function cancelPointer(){
    if(pointerDrag)finishPointer({type:'pointercancel',pointerId:pointerDrag.pointer,screenX:pointerDrag.screen[0],screenY:pointerDrag.screen[1]});
  }
  window.addEventListener('blur',cancelPointer);
  window.addEventListener('lostpointercapture',event=>{if(pointerDrag?.pointer===event.pointerId)cancelPointer();});
  function fallbackButton() {
    const status=document.getElementById('floating-status');
    if(!status||document.getElementById('desktop-system-frame'))return;
    const recovery=document.createElement('p');recovery.id='desktop-chrome-recovery';recovery.textContent='拖动未就绪，可改用系统窗框。';
    const button=document.createElement('button');button.type='button';button.id='desktop-system-frame';button.textContent='使用系统窗框';
    button.onclick=()=>action('standard-frame');recovery.append(button);status.after(recovery);
  }
  async function sendRegions() {
    if(regionBusy){regionQueued=true;return;}
    if(document.body.dataset.desktopWindow!=='true')return;
    const width=innerWidth,height=innerHeight;
    if(width<1||height<1)return;
    const rectangle=node=>{
      const r=node.getBoundingClientRect(),left=Math.max(0,r.left),top=Math.max(0,r.top),right=Math.min(width,r.right),bottom=Math.min(height,r.bottom);
      return left<right&&top<bottom?[left,top,right,bottom]:null;
    };
    const targets=[...document.querySelectorAll('#floating-title,#desktop-window-bar .desktop-drag-region,#settings-dialog .settings-head')];
    const visible=targets.filter(node=>node.getClientRects().length&&getComputedStyle(node).visibility==='visible'&&!node.closest('[hidden]'));
    const regions=visible.map(rectangle).filter(Boolean);
    const exclude=[...new Set(visible.flatMap(node=>[...node.querySelectorAll('button,input,select,textarea,a,label,summary,[contenteditable],.desktop-window-controls')]))].map(rectangle).filter(Boolean);
    const blocking=[...document.querySelectorAll('dialog[open]')].some(node=>node.id!=='settings-dialog'&&node.matches(':modal'));
    const payload={viewport:[width,height],regions:blocking||exclude.length>64||regions.length>16?[]:regions,exclude:exclude.slice(0,64)};
    const key=JSON.stringify(payload);if(key===lastRegionKey)return;
    regionBusy=true;
    try{
      const result=await api().floating_window_regions({version:performance.timeOrigin+performance.now(),...payload});
      if(result?.ok)lastRegionKey=key;
    }catch(_){}
    finally{regionBusy=false;if(regionQueued){regionQueued=false;regionsChanged();}}
  }
  function regionsChanged() {
    if(regionFrame)return;
    regionFrame=requestAnimationFrame(()=>{regionFrame=null;sendRegions();});
  }
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  async function action(name) {
    if (locked) return;
    locked = true;
    const mine = ++sequence;
    try {
      const result = await api().floating_window_action(name);
      if (!result?.ok) throw Error(result?.error || '窗口操作未完成，请使用悬浮球恢复窗口');
      if (mine === sequence) update(result);
      if(result.drag_started===false||result.caption_failed)fallbackButton();
      return result;
    } catch (error) {
      const status = document.getElementById('floating-status');
      if (status) status.textContent = error.message;
      return {ok:false};
    } finally { locked = false; }
  }
  function update(state) {
    if (!maximum) return;
    document.body.dataset.windowMaximized = String(!!state.maximized);
    document.body.dataset.nativeCaption=String(!!state.native_caption);
    document.body.dataset.dragDriver=state.drag_driver||'';
    document.body.dataset.standardFrame=String(!!state.standard_frame);
    if(state.standard_frame){
      document.getElementById('desktop-chrome-recovery')?.remove();
    }
    window.OfficeAgentElements.iconButton(maximum, state.maximized ? 'minimize-2' : 'maximize-2', state.maximized ? '还原窗口' : '最大化窗口');
    maximum.setAttribute('aria-pressed', String(!!state.maximized));
  }
  function attach(parent, collapse) {
    if (header) return;
    header = parent;
    const title = document.getElementById('floating-title');
    title.classList.add('desktop-drag-region');
    controls = document.createElement('div'); controls.className = 'desktop-window-controls';
    for (const [name, label, icon] of [['minimize', '最小化窗口', 'minus'], ['maximize', '最大化窗口', 'maximize-2']]) {
      const node = document.createElement('button'); node.type = 'button'; node.hidden = true;
      node.dataset.windowAction = name;
      window.OfficeAgentElements.iconButton(node, icon, label);
      node.addEventListener('click', () => action(name));
      controls.append(node);
      if (name === 'maximize') maximum = node;
    }
    window.OfficeAgentElements.iconButton(collapse, 'x', '收起', '收起到悬浮球');
    collapse.classList.add('desktop-window-close');
    controls.append(collapse); parent.append(controls);
    bar = document.createElement('div'); bar.id = 'desktop-window-bar'; bar.hidden = true;
    const name = document.createElement('span'); name.textContent = 'LiX'; name.className = 'desktop-drag-region';
    const logo = document.createElement('img'); logo.src = 'brand/lix.png'; logo.alt = ''; logo.className = 'lix-brand-icon'; logo.width = logo.height = 24;
    name.prepend(logo); bar.append(name); document.body.prepend(bar);
    let lastDown;
    for (const node of [title, name]) node.addEventListener('pointerdown', event => {
      if (event.button !== 0 || document.body.dataset.desktopWindow !== 'true'||document.body.dataset.nativeCaption==='true') return;
      const now = performance.now(), prior = lastDown;
      lastDown = {time:now,x:event.screenX,y:event.screenY};
      const doubleClick = prior && now-prior.time < 360 && Math.hypot(event.screenX-prior.x,event.screenY-prior.y) < 4;
      if(document.body.dataset.dragDriver==='pointer'){
        if(doubleClick){event.preventDefault();action('maximize');}
        else startPointer(event);
        return;
      }
      event.preventDefault();
      if (doubleClick) action('maximize');
      else action('drag');
    });
    for(const node of [title,name])node.addEventListener('dblclick',event=>{
      if(document.body.dataset.dragDriver==='pywebview'&&event.button===0){event.preventDefault();action('maximize');}
    });
    let tracing=false,traceMoved=false;
    for(const phase of ['pointerdown','mousedown','pointermove','pointerup','mouseup'])window.addEventListener(phase,event=>{
      if(document.body.dataset.desktopWindow!=='true')return;
      const region=!!event.target.closest?.('#floating-title,#desktop-window-bar .desktop-drag-region,#settings-dialog .settings-head');
      if(phase==='pointerdown'){tracing=region;traceMoved=false;}
      if(!tracing||phase==='pointermove'&&traceMoved)return;
      if(phase==='pointermove')traceMoved=true;
      api().floating_window_trace({phase,driver:document.body.dataset.dragDriver||'unset',region:!!event.target.closest?.('.desktop-drag-region,#settings-dialog .settings-head'),prevented:event.defaultPrevented,buttons:event.buttons}).catch(()=>{});
      if(phase==='mouseup')tracing=false;
    });
    document.addEventListener('pointerdown', event => {
      if(event.button!==0||document.body.dataset.desktopWindow!=='true'||document.body.dataset.nativeCaption==='true'||!event.target.closest('#settings-dialog .settings-head')||event.target.closest('button,input,select,textarea,a,label,summary,[contenteditable]'))return;
      if(document.body.dataset.dragDriver==='pointer'){startPointer(event);return;}
      event.preventDefault();action('drag');
    });
    for (const edge of ['left','right','top','bottom','top-left','top-right','bottom-left','bottom-right']) {
      const handle = document.createElement('div'); handle.className = 'desktop-resize-handle'; handle.dataset.edge = edge;
      handle.setAttribute('aria-hidden','true');
      handle.addEventListener('pointerdown', event => { if(event.button===0){event.preventDefault();action('resize-'+edge);} });
      document.body.append(handle);
    }
    const observer=new MutationObserver(regionsChanged);
    observer.observe(document.body,{subtree:true,childList:true,attributes:true,attributeFilter:['hidden','open','class','style']});
    new ResizeObserver(regionsChanged).observe(document.body);
    window.addEventListener('scroll',regionsChanged,true);
    refresh();
  }
  async function refresh() {
    const mine = ++sequence;
    try {
      const state = await api().floating_window_state();
      if (mine !== sequence || !state?.ok || !state.available || !controls) return;
      document.body.dataset.desktopWindow = 'true';
      controls.querySelectorAll('[data-window-action]').forEach(node => { node.hidden = false; });
      update(state);
      secondary(secondaryOpen);
      regionsChanged();
      if(state.caption_failed)fallbackButton();
    } catch (_) {}
  }
  function secondary(open) {
    secondaryOpen = open;
    if (!controls) return;
    const native = document.body.dataset.desktopWindow === 'true';
    document.body.classList.toggle('desktop-secondary', open && native);
    bar.hidden = !open || !native;
    (open && native ? bar : header).append(controls);
    regionsChanged();
  }
  window.addEventListener('focus', () => { if (header) refresh(); });
  window.addEventListener('resize', () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (header) refresh(); }, 160); });
  window.OfficeWindowChrome = {attach, secondary, maximize: () => action('maximize')};
})();
