(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const api = () => window.pywebview?.api;
  const status = text => { $('plugin-status').textContent = text; };
  const make = (tag, text, parent) => {
    const element = document.createElement(tag);
    element.textContent = text || '';
    if (parent) parent.append(element);
    return element;
  };
  function button(text, parent, action) {
    const element = make('button', text, parent);
    element.type = 'button'; element.className = 'btn secondary'; element.onclick = action;
    return element;
  }
  async function call(name, ...args) {
    if (typeof api()?.[name] !== 'function') throw Error('插件工作区尚未连接。');
    const result = await api()[name](...args);
    if (!result?.ok) throw Error(result?.error || '操作未完成，请重试。');
    return result;
  }
  let rows = [], picker, listGeneration = 0, catalogGeneration = 0, timer, cursor = '', term = '';
  const catalog = new Map();
  function clear() { if (picker) picker.value = ''; }
  function choice() { return picker?.selectedOptions[0]?.ticket || null; }
  function updatePicker() {
    if (!picker) return;
    const previous = picker.value;
    picker.replaceChildren(new Option('本次插件', ''));
    for (const row of rows.filter(value => value.enabled)) {
      for (const tool of row.allowed) {
        const option = new Option(row.name + ' · ' + tool, row.id + ':' + tool);
        option.ticket = {id: row.id, revision: row.revision, tool};
        picker.add(option);
      }
    }
    if ([...picker.options].some(opt => opt.value === previous)) picker.value = previous;
    picker.hidden = picker.options.length === 1;
  }
  async function refresh() {
    const mine = ++listGeneration;
    const result = await call('plugin_list');
    if (mine !== listGeneration) return;
    rows = result.items; updatePicker();
    const list = $('plugin-installed');
    list.replaceChildren();
    if (!rows.length) make('p', '还没有添加插件。', list);
    for (const row of rows) {
      const line = make('div', '', list); line.className = 'plugin-row';
      const info = make('div', '', line);
      make('strong', row.name, info);
      make('small', row.state + ' · ' + (row.source || (row.transport === 'stdio' ? '本机程序' : '手动添加')), info);
      button('管理', line, event => detail(row, event.currentTarget)).dataset.pluginId=row.id;
    }
  }
  function attach(footer) {
    picker = make('select', '', footer); picker.id = 'floating-plugin'; picker.title = '本次使用的插件';
    picker.setAttribute('aria-label', '本次使用的插件'); updatePicker();
    if (api()?.plugin_list) refresh().catch(() => {});
  }
  function field(parent, label, value, type = 'text', options) {
    const wrap = make('label', '', parent); make('span', label, wrap);
    const input = make(options ? 'select' : type === 'textarea' ? 'textarea' : 'input', '', wrap);
    if (options) for (const [id, title] of options) input.add(new Option(title, id));
    else if (type !== 'textarea') input.type = type;
    if (type === 'checkbox') input.checked = Boolean(value); else input.value = value || '';
    return input;
  }
  function detail(initial, trigger, entry = null) {
    let row = initial, busy = false, returned = false, lastTrigger = trigger;
    const dialog = make('dialog', '', document.body); dialog.className = 'plugin-dialog';
    dialog.setAttribute('aria-label','插件连接详情');
    const head = make('header','',dialog), heading = make('h3','',head);
    button('关闭', head, () => dialog.close());
    const content = make('div','',dialog), message = make('p','',dialog);
    message.setAttribute('role','status');
    dialog.addEventListener('close', () => {
      returned = true;
      if (busy && row?.id) call('plugin_cancel',row.id).catch(()=>{});
      dialog.remove();
      const manager=rows.some(value=>value.id===row?.id)&&[...$('plugin-installed').querySelectorAll('button')].find(b=>b.dataset.pluginId===row.id);
      (lastTrigger?.isConnected?lastTrigger:manager||$('plugin-manual'))?.focus();
    },{once:true});
    async function perform(name,...args) {
      if (busy) return;
      busy = true; content.inert = true; message.textContent = '处理中…';
      try {
        const result = await call(name,...args);
        if (returned) return;
        await refresh();
        if (result.removed) { dialog.close(); return; }
        if (result.item) row = result.item;
        paint();
        message.textContent = result.duplicate ? '已打开相同配置的连接。' : row?.error || row?.state || '';
      } catch(error) { if (!returned) message.textContent = error.message; }
      finally { busy = false; content.inert = false; }
    }
    function paint() {
      content.replaceChildren(); message.textContent = '';
      heading.textContent = row?.id ? row.name : '添加插件';
      const data = row || {};
      if (entry) {
        make('p',entry.source+' · '+entry.version,content);
        if (entry.website) make('small',entry.website,content);
      }
      const name = field(content,'名称',data.name);
      const transport = field(content,'连接方式',data.transport || 'http','',[['http','远程 HTTPS'],['stdio','本地 stdio']]);
      const remote = make('div','',content), local = make('div','',content);
      const url = field(remote,'服务地址',data.url,'url');
      const priv = field(remote,'允许连接这个私有服务',data.private,'checkbox');
      const auth = field(remote,'认证方式',data.auth || 'none','',[['none','无需凭据'],['bearer','密钥引用'],['oauth','OAuth 登录']]);
      const key = field(remote,'本机 .env 凭据名称',data.credential_ref);
      const header = field(remote,'凭据标头',data.api_header || 'Authorization','',[['Authorization','Authorization (Bearer)'],['X-Api-Key','X-Api-Key']]);
      const issuer = field(remote,'授权服务器（同源时留空）',data.oauth_issuer,'url');
      const program = field(local,'本机程序',data.program);
      button('选择程序…',local,async()=>{
        try{const result=await call('plugin_pick_program');if(result.path){program.value=result.path;program.dispatchEvent(new Event('input'));}}
        catch(error){message.textContent=error.message;}
      });
      const args = field(local,'启动参数（每行一项）',(data.args || []).join('\n'),'textarea');
      const refs = field(local,'环境变量引用（JSON 对象）',JSON.stringify(data.env_refs || {}),'textarea');
      make('small','程序将以当前 Windows 用户权限运行。请核对路径和参数。',local);
      const importer = make('details','',local); make('summary','从 JSON 填入',importer);
      const raw = field(importer,'本地插件配置','','textarea');
      button('预览配置',importer,()=>{
        try {
          const draft=JSON.parse(raw.value);
          if (!draft || typeof (draft.program || draft.command)!=='string' ||
              !Array.isArray(draft.args || []) || !(draft.args || []).every(v=>typeof v==='string'&&!/[\r\n]/.test(v)) ||
              (draft.env_refs && (typeof draft.env_refs!=='object'||Array.isArray(draft.env_refs))) ||
              Object.keys(draft).some(k=>!['program','command','name','args','env_refs','transport'].includes(k))) throw Error();
          program.value=draft.program || draft.command; args.value=(draft.args || []).join('\n');
          refs.value=JSON.stringify(draft.env_refs || {});if(draft.name)name.value=draft.name;
          program.dispatchEvent(new Event('input'));
          message.textContent='配置已填入，请核对后保存。';
        } catch {message.textContent='JSON 无效，使用 program、args 和 env_refs 填写配置。';}
      });
      function fields() {
        remote.hidden=transport.value!=='http';local.hidden=!remote.hidden;
        key.parentElement.hidden=header.parentElement.hidden=auth.value!=='bearer';
        issuer.parentElement.hidden=auth.value!=='oauth';
      }
      transport.onchange=auth.onchange=fields;fields();
      const separate = !data.id ? field(content,'另建独立连接',false,'checkbox') : null;
      const actions=make('div','',content);actions.className='plugin-actions';
      button('保存配置',actions,()=>{
        try{
          const config={name:name.value,transport:transport.value};
          if(transport.value==='http'){
            Object.assign(config,{url:url.value.trim(),private:priv.checked,auth:auth.value,
              credential_ref:key.value.trim(),api_header:header.value});
            if(auth.value==='oauth' && issuer.value.trim())config.oauth_issuer=issuer.value.trim();
            if(data.source && data.url===config.url && !config.private)config.source=data.source;
          }else Object.assign(config,{program:program.value.trim(),args:args.value ? args.value.split('\n') : [],env_refs:JSON.parse(refs.value || '{}')});
          perform('plugin_save',config,data.id || null,data.revision || null,Boolean(separate?.checked));
        }catch{message.textContent='请核对环境变量引用的 JSON 格式。';}
      });
      if(!data.id)return;
      make('p',data.state + ' · 本机用户 · '+(data.source || '手动添加'),content);
      if(data.auth==='bearer'||data.auth==='oauth'){
        const credential=make('small','正在核对凭据…',content);
        call('plugin_credentials',data.id,data.revision).then(r=>{
          credential.textContent=r.available?'凭据已配置':'尚无可用凭据';
        }).catch(()=>{credential.textContent='凭据状态未知';});
      }
      if(data.error)make('p',data.error,content);
      const probeButton=button('测试连接',actions,()=>perform('plugin_probe',data.id,data.revision));
      if(data.auth==='oauth')button('登录并连接',actions,()=>perform('plugin_login',data.id,data.revision));
      const tools=make('fieldset','',content);make('legend','可用工具',tools);
      const found=data.tools || [],checks=[];
      make('small','发现 '+found.length+' 个工具，'+found.filter(t=>t.usable).length+' 个可选。',tools);
      for(const t of found){
        const checkbox=field(tools,t.name,data.allowed.includes(t.name),'checkbox');
        checkbox.disabled=!t.usable;checks.push([checkbox,t.name]);
        make('small',t.usable?t.description:t.reason,tools);
      }
      if(found.some(t=>t.usable))button('启用所选工具',tools,()=>perform('plugin_enable',data.id,data.revision,checks.filter(([cb])=>cb.checked).map(([,name])=>name)));
      const manage=make('div','',content);manage.className='plugin-actions';
      if(data.enabled)button('停用',manage,()=>perform('plugin_disable',data.id,data.revision));
      if(data.auth==='oauth')button('断开登录',manage,()=>perform('plugin_disconnect',data.id,data.revision));
      const remove=button('移除',manage,()=>{
        remove.hidden=true;
        button('确认移除 '+data.name,manage,()=>perform('plugin_remove',data.id,data.revision));
        button('保留',manage,paint);
      });
      const savedOnly=[...actions.querySelectorAll('button')].slice(1).concat([...tools.querySelectorAll('button')]);
      function edited(){
        for(const action of savedOnly)action.disabled=true;
        probeButton.title='请先保存更改';
        message.textContent='请先保存配置，再连接或启用工具。';
      }
      for(const control of [name,transport,url,priv,auth,key,header,issuer,program,args,refs])control.addEventListener('input',edited);
    }
    paint();dialog.showModal();
  }


  function renderCatalog() {
    const list=$('plugin-catalog');list.replaceChildren();
    if(catalog.size)make('h4','发现',list);
    for(const item of catalog.values()){
      const line=make('div','',list);line.className='plugin-row';
      const info=make('div','',line);
      make('strong',item.name,info);make('small',item.description,info);
      make('small',item.source+' · '+(item.url?'远程 HTTPS':'需本地安装')+' · '+item.version,info);
      if(item.url)button('添加',line,event=>detail({name:item.name,url:item.url,source:item.source},event.currentTarget,item));
      else for(const pkg of item.packages)make('small',[pkg.name,pkg.version,pkg.registry,pkg.runtime].filter(Boolean).join(' · '),info);
    }
    $('plugin-more').hidden=!cursor;
  }
  async function search(next=false) {
    clearTimeout(timer);
    const mine=++catalogGeneration;
    if(!next){term=$('plugin-search').value.trim();cursor='';catalog.clear();renderCatalog();}
    if(!term){status('');return;}
    status('正在搜索公开目录…');$('plugin-more').disabled=true;
    try{
      const result=await call('plugin_search',term,next?cursor:'');
      if(mine!==catalogGeneration)return;
      for(const item of result.items)catalog.set(item.source,item);
      cursor=result.cursor;renderCatalog();
      status(result.offline?'离线结果 · '+new Date(result.cached_at*1000).toLocaleString():catalog.size?'':'没有匹配的插件。');
    }catch(error){if(mine===catalogGeneration)status(error.message);}
    finally{if(mine===catalogGeneration)$('plugin-more').disabled=false;}
  }
  function renderTurn(node,row,handlers) {
    node.pluginArea?.remove();node.pluginArea=null;
    if(!row.plugin)return;
    const recover=['failed','interrupted','uncertain'].includes(row.status);
    if(!row.plugin_preview && !row.plugin_sources?.length && !recover)return;
    const area=make('div','',node.assistant);area.className='conversation-plugin';node.pluginArea=area;
    if(recover){
      const note=make('small','',area);
      button('重新提问',area,async()=>{
        const input=$('floating-input');
        if(!input)return;
        if(input.value.trim() && input.value.trim()!==row.question){
          note.textContent='输入区已有草稿，请先处理后再重提。';input.focus();return;
        }
        try{await refresh();}catch(error){note.textContent=error.message;return;}
        if(!area.isConnected)return;
        const option=[...(picker?.options || [])].find(o=>o.ticket&&o.ticket.id===row.plugin.id&&o.ticket.tool===row.plugin.tool);
        if(!option){note.textContent='请先在插件设置中核对并启用这个工具。';return;}
        picker.value=option.value;input.value=row.question;input.dispatchEvent(new Event('input',{bubbles:true}));input.focus();
        note.textContent='已填回问题，请核对后发送。';
      });
      button('插件设置',area,()=>{window.OfficeSettings?.show();window.OfficeSettings?.selectTab('tab-plugins');});
    }else if(row.plugin_preview){
      make('small',row.plugin.name+' · '+row.plugin.tool,area);make('small',row.plugin.destination,area);
      make('pre',row.plugin_preview.query,area);
      for(const approved of [true,false]){
        button(approved?'允许本次发送':'取消',area,async()=>{
          const buttons=[...area.querySelectorAll('button')];buttons.forEach(b=>b.disabled=true);
          try{await handlers.pluginConsent?.(row.id,row.plugin_preview.token,approved);}
          finally{buttons.forEach(b=>b.disabled=false);}
        });
      }
    }else{
      const details=make('details','',area);make('summary','来源 · '+row.plugin.name,details);
      row.plugin_sources.forEach((source,index)=>{
        const link=make('a',source.title,details);link.href=source.url;link.title=source.url;
        link.onclick=event=>{event.preventDefault();handlers.pluginSource?.(row.id,index);};
      });
    }
  }
  $('plugin-search').addEventListener('input',()=>{
    catalogGeneration++;clearTimeout(timer);timer=setTimeout(()=>search(),350);
  });
  $('plugin-search').addEventListener('keydown',event=>{if(event.key==='Enter'&&!event.isComposing){event.preventDefault();search();}});
  $('plugin-more').onclick=()=>search(true);
  $('plugin-manual').onclick=event=>detail(null,event.currentTarget);
  document.addEventListener('officesettings:tab',event=>{
    if(event.detail.tab==='tab-plugins')refresh().catch(e=>status(e.message));
    else {catalogGeneration++;clearTimeout(timer);}
  });
  document.addEventListener('officesettings:closed',()=>{catalogGeneration++;clearTimeout(timer);});
  window.addEventListener('pywebviewready',()=>{if(picker)refresh().catch(()=>{});});
  window.OfficePlugins={attach,choice,clear,renderTurn,reload:()=>refresh().catch(()=>{})};
})();

