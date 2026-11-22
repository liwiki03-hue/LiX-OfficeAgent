(() => {
  'use strict';
  function mount({composer, footer, input, state, api, button, make, reference}) {
    const selected = {};
    let returnPoint, generation = 0, serial = 0, picking = false, page = 'root', pendingDrop;
    const chips = make('div', '', null, 'composer-context-chips');
    composer.prepend(chips);
    const error = make('div', '', chips, 'composer-context-feedback');
    const note = make('span', '', error);note.setAttribute('role','status');
    const retry = button('重新选择', error, () => chooseFile());
    error.hidden=true;retry.hidden=true;
    const chipList=make('div','',chips,'composer-context-list');
    const add=window.OfficeAgentElements.iconButton(button('添加本轮材料与工具', footer, toggle, 'composer-context-add'),'plus');
    const popup=make('div','',footer,'composer-context-menu');
    popup.setAttribute('popover','manual');
    add.setAttribute('aria-haspopup','menu');add.setAttribute('aria-expanded','false');add.setAttribute('aria-controls',popup.id);
    const isOpen=()=>popup.matches(':popover-open');
    const close=focus=>{
      generation++;
      if(isOpen())popup.hidePopover();
      add.setAttribute('aria-expanded','false');
      if(focus&&!add.disabled)add.focus({preventScroll:true});
    };
    function remember() {
      return {text:input.value, start:input.selectionStart, end:input.selectionEnd, direction:input.selectionDirection};
    }
    function resume(point=returnPoint) {
      input.focus({preventScroll:true});
      if(point?.text===input.value)input.setSelectionRange(point.start,point.end,point.direction);
    }
    function owner() {
      const data=state();
      return data?.document?.ready?JSON.stringify([data.epoch,data.document.id,data.document.version]):'';
    }
    function noteFile(text,canRetry=false) {
      note.textContent=text;error.hidden=!text;retry.hidden=!canRetry;
    }
    function paint() {
      chipList.replaceChildren();
      for(const [kind,value] of Object.entries(selected)){
        const word={file:'参考',skill:'技能',plugin:'插件'}[kind];
        const chip=make('span','',chipList);chip.className='composer-context-chip';
        const text=make('span',word+' · '+value.name,chip);text.title=value.name;
        const remove=button('移除'+word+'：'+value.name,chip,()=>{
          delete selected[kind];paint();resume(remember());
          if(kind==='file'){noteFile('');api()?.conversation_remove_file(value.token).catch(()=>{});}
        });
        window.OfficeAgentElements.iconButton(remove,'x');
      }
      chipList.hidden=!chipList.childElementCount;
    }
    function update(data) {
      add.disabled=!data?.document?.ready||picking||!!pendingDrop;
      if(!data?.document?.ready)close(false);
    }
    function clear() {
      if(pendingDrop)clearTimeout(pendingDrop.timer);
      pendingDrop=null;
      update(state());
      if(selected.file)api()?.conversation_remove_file(selected.file.token).catch(()=>{});
      for(const kind of Object.keys(selected))delete selected[kind];
      close(false);noteFile('');paint();
    }
    function position() {
      if(!isOpen())return;
      if(!add.getClientRects().length){close(false);return;}
      const anchor=add.getBoundingClientRect(), box=composer.getBoundingClientRect();
      const width=Math.min(300,box.width,innerWidth-16);
      popup.style.width=width+'px';
      popup.style.left=Math.max(8,Math.min(box.left,innerWidth-width-8))+'px';
      popup.style.bottom=innerHeight-anchor.top+5+'px';
      popup.style.maxHeight=Math.min(340,Math.max(84,anchor.top-16))+'px';
    }
    function toggle() {
      if(isOpen()){close(true);return;}
      returnPoint=remember();
      popup.showPopover();add.setAttribute('aria-expanded','true');
      renderRoot();
    }
    function item(text,icon,action) {
      const trigger=button(text,popup,action);trigger.setAttribute('role','menuitem');trigger.tabIndex=-1;
      const svg=make('span','');svg.className='agent-icon composer-icon-'+icon;svg.setAttribute('aria-hidden','true');
      trigger.prepend(svg);
    }
    function renderRoot() {
      page='root';
      popup.replaceChildren();popup.setAttribute('role','menu');popup.setAttribute('aria-label','添加本轮材料与工具');
      item(selected.file?'替换参考文件…':'引用文件…','paperclip',()=>chooseFile(returnPoint));
      if(reference())item('引用当前选区','quote',async()=>{
        const point=returnPoint;close(false);
        try{await reference()?.();}catch(e){noteFile(e.message);}
        resume(point);
      });
      item('技能','skill',()=>showChoices('skill'));
      item('插件','plugin',()=>showChoices('plugin'));
      position();popup.querySelector('button')?.focus();
    }
    async function chooseFile(point=remember()) {
      const document=state()?.document;
      if(!document?.ready||picking)return;
      const before=owner(), priorGeneration=generation;
      close(false);const expectedGeneration=priorGeneration+1;
      picking=true;update(state());
      try {
        const result=await api().conversation_choose_file(document.id,document.version);
        if(owner()!==before||generation!==expectedGeneration){
          if(result?.file)api().conversation_remove_file(result.file.token).catch(()=>{});
          return;
        }
        if(!result?.ok)throw Error(result?.error||'无法选择参考文件。');
        if(!result.cancelled){selected.file=result.file;noteFile('');paint();}
      } catch(e) {
        if(owner()===before&&generation===expectedGeneration)noteFile(e.message,true);
      } finally {
        picking=false;update(state());
        if(owner()===before&&generation===expectedGeneration)resume(point);
      }
    }
    function dropFile(files, transfer) {
      const doc=state()?.document;
      if(!doc?.ready){noteFile('先打开一份主文档，再拖入参考文件。');return;}
      if(picking||pendingDrop){noteFile('正在接收参考文件，请稍等。');return;}
      if(files.length!==1){noteFile('一次只能拖入一份参考文件。');return;}
      const request=crypto.randomUUID();
      const point=remember();
      close(false);
      noteFile('正在添加参考文件…');
      pendingDrop={request, owner:owner(), point, timer:setTimeout(()=>{
        pendingDrop=null;update(state());noteFile('未收到参考文件，请重试。',true);
      },10000)};
      update(state());
      try{transfer({route:'reference',request,owner:[state().epoch,doc.id,doc.version]},files);}
      catch(_){finishDrop({request,ok:false,error:'无法接收拖放文件，请重新选择。'});}
    }
    function finishDrop(result) {
      if(pendingDrop?.request!==result.request||pendingDrop.owner!==owner()){
        if(result.file)api()?.conversation_remove_file(result.file.token).catch(()=>{});
        return;
      }
      const point=pendingDrop.point;
      clearTimeout(pendingDrop.timer);
      pendingDrop=null;update(state());
      if(result.ok){selected.file=result.file;noteFile('');paint();}
      else noteFile(result.error||'无法添加参考文件。',true);
      const active=document.activeElement;
      if(active===document.body||composer.contains(active))resume(point);
    }
    function choose(kind,entry) {
      selected[kind]={...entry,serial:++serial};close(false);noteFile('');paint();resume();
    }
    async function showChoices(kind) {
      const doc=state()?.document;
      if(!doc?.ready)return;
      page=kind;const mine=++generation, before=owner();
      popup.replaceChildren();popup.setAttribute('role','dialog');popup.setAttribute('aria-label',kind==='skill'?'选择本轮技能':'选择本轮插件工具');
      button('返回',popup,()=>{generation++;renderRoot();});
      const search=make('input','',popup);search.type='search';search.placeholder='查找名称';search.setAttribute('aria-label','查找名称');
      const rows=make('div','',popup);rows.className='composer-context-options';
      button(kind==='skill'?'管理技能':'配置插件',popup,()=>{
        close(false);
        const point=returnPoint;
        document.addEventListener('officesettings:closed',()=>{if(owner()===before)resume(point);},{once:true});
        window.OfficeSettings?.show();window.OfficeSettings?.selectTab(kind==='skill'?'tab-skills':'tab-plugins');
      });
      rows.textContent='正在读取本地列表…';position();search.focus();
      try{
        let entries;
        if(kind==='skill'){
          const result=await api().conversation_skill_choices(doc.id,doc.version);
          if(!result?.ok)throw Error(result?.error||'无法读取技能。');
          entries=result.items;
        } else entries=await window.OfficePlugins.available();
        if(mine!==generation||!isOpen()||owner()!==before)return;
        const paintChoices=()=>{
          const term=search.value.trim().toLocaleLowerCase();
          const matched=entries.filter(entry=>entry.name.toLocaleLowerCase().includes(term));
          rows.replaceChildren();
          if(!matched.length)rows.textContent=entries.length?'未找到匹配项。':'当前没有可用项。';
          for(const entry of matched)button(entry.name,rows,()=>choose(kind,entry)).title=entry.description||entry.name;
        };
        search.oninput=paintChoices;paintChoices();position();
      }catch(e){if(mine===generation)rows.textContent=e.message;}
    }
    add.addEventListener('keydown',event=>{
      if(!event.isComposing&&event.key==='ArrowDown'&&!isOpen()){event.preventDefault();toggle();}
    });
    popup.addEventListener('keydown',event=>{
      if(event.isComposing||event.keyCode===229)return;
      if(event.key==='Escape'){event.preventDefault();event.stopPropagation();close(true);return;}
      if(event.key==='Enter'&&event.target.type==='search'){event.preventDefault();return;}
      if(!['ArrowUp','ArrowDown','Home','End'].includes(event.key)||event.target.type==='search'&&['Home','End'].includes(event.key))return;
      const items=[...popup.querySelectorAll(page==='root'?'button':'.composer-context-options button')];
      if(!items.length)return;
      event.preventDefault();
      const index=items.indexOf(document.activeElement), count=items.length;
      items[event.key==='Home'?0:event.key==='End'?count-1:event.key==='ArrowDown'?(index+1)%count:index<0?count-1:(index+count-1)%count].focus();
    });
    document.addEventListener('pointerdown',event=>{
      if(isOpen()&&!popup.contains(event.target)&&!add.contains(event.target))close(false);
    });
    document.addEventListener('focusin',event=>{
      if(isOpen()&&!popup.contains(event.target)&&!add.contains(event.target))close(false);
    });
    window.addEventListener('resize',position);
    const observer=new ResizeObserver(position);observer.observe(add);
    paint();
    return {
      update,clear,dropFile,finishDrop,waitingForFile:()=>!!pendingDrop,
      choices(){
        const {file,skill,plugin}=selected;
        return {resources:{...(file?{file:file.token}:{}),...(skill?{skill:{id:skill.id,sha256:skill.sha256}}:{})},
          plugin:plugin?.ticket||null, skillPick:skill?.serial, pluginPick:plugin?.serial};
      },
      consume(sent,turn){
        if(turn?.resources?.file&&selected.file?.token===sent.resources?.file){delete selected.file;noteFile('');}
        if(turn?.resources?.skill&&selected.skill&&selected.skill.serial===sent.skillPick)delete selected.skill;
        if(turn?.plugin&&sent.plugin&&selected.plugin&&selected.plugin.serial===sent.pluginPick)delete selected.plugin;
        paint();
      },
      rejected(text,sent) {
        if(sent.resources?.file&&selected.file?.token===sent.resources.file)noteFile(text,/重新选择/.test(text));
      },
      restorePlugin(entry){selected.plugin={...entry,serial:++serial};paint();},
      restoreResources(row) {
        if(input.value.trim()&&input.value.trim()!==row.question)return false;
        input.value=row.question;
        resume(remember());
        if(row.resources?.file)chooseFile(remember());
        else if(row.resources?.skill){if(!isOpen())toggle();showChoices('skill');}
        return true;
      }
    };
  }
  window.OfficeComposerContext={mount};
})();
