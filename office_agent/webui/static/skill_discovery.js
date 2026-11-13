(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const root = $('skills-discovery');
  if (!root || root.dataset.bound) return;
  root.dataset.bound = 'true';
  let epoch = 0, pending = false, items = [], result = null;
  let featured = null;
  const input = $('skills-search'), status = $('skills-public-status');
  const active = () => !root.hidden && window.OfficeSettings?.isOpen();
  const make = (tag, text, parent, className) => {
    const node = document.createElement(tag); node.textContent = text;
    if (className) node.className = className;
    parent.append(node); return node;
  };
  const sourceLabel=make('label','搜索来源',root.querySelector('.actions'),'skills-source-choice');
  const sourceChoice=make('select','',sourceLabel);sourceChoice.id='skills-public-provider';sourceChoice.setAttribute('aria-label','公开技能搜索来源');
  for(const [value,label] of [['skills.sh','skills.sh'],['openai/plugins','OpenAI 插件目录'],['lix','LiX 工作流']]){const option=make('option',label,sourceChoice);option.value=value;}
  const iconFor=(parent,item)=>{
    const icon=make('span','',parent,'skill-entry-icon');icon.setAttribute('aria-hidden','true');
    icon.dataset.category=['writing','style','review','document','table','presentation'].includes(item.category)?item.category:'document';
    if(item.icon?.data_url&&/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(item.icon.data_url)&&item.icon.data_url.length<150000){
      const image=document.createElement('img');image.src=item.icon.data_url;image.alt='';image.width=28;image.height=28;icon.append(image);icon.classList.add('has-source-icon');
      image.onerror=()=>{icon.classList.remove('has-source-icon');image.remove();};
    }
  };
  const errors = {invalid_query:'请输入 2–120 个字符。', busy:'查询尚未结束，请稍后重试。', timeout:'查询超时，请重试。', network_error:'网络不可用，请重试。', service_error:'目录服务暂不可用，请重试。', rate_limited:'查询过于频繁，请稍后重试。', invalid_response:'目录格式发生变化，暂时无法读取。', response_too_large:'目录响应过大，已停止读取。', not_authenticated:'本机服务未就绪，请重新打开设置。'};
  const controls = () => {
    $('skills-public-search').disabled = pending;
    $('skills-public-cancel').hidden = !pending;
    $('skills-public-results').setAttribute('aria-busy', String(pending));
    $('skills-github-check').disabled=pending;
  };
  function detail(item) {
    const panel=$('skills-public-detail');panel.replaceChildren();panel.hidden=false;
    const back=make('button','返回结果',panel,'btn secondary');back.type='button';back.onclick=()=>{panel.hidden=true;$('skills-public-results').hidden=false;Array.from($('skills-public-results').children).find(node=>node.dataset.candidate===item.candidate_id)?.focus();};
    iconFor(panel,item);make('h3',item.display_name||item.name,panel);
    make('p',item.source+' · '+(item.provider==='lix'?'LiX 提供 · 应用内置工作流':item.provider==='featured'?'固定来源精选 · '+item.name:item.provider==='openai/plugins'?'插件清单中声明的 Skill':'skills.sh 收录'),panel);
    make('p',item.description||'安装前检查固定版本、许可与兼容性。',panel);
    if(item.provider==='featured'){make('p','快照版本：'+item.commit+' · '+item.license,panel);make('p',item.reason,panel);}
    else if(item.provider==='lix'){make('p','工作流版本：'+item.version+' · '+item.license,panel);make('p',item.reason,panel);}
    else if(item.provider==='openai/plugins'){make('p','固定版本：'+item.commit,panel);make('p',item.plugin_context.scope,panel);if(item.plugin_context.requirements.length)make('p','附带依赖未接入：'+item.plugin_context.requirements.join('、'),panel);}
    else make('p','尚未核对许可、固定版本与 LiX 兼容性；当前只可浏览。',panel);
    const label=make('label','来源地址',panel);
    const source=make('input','',label);source.readOnly=true;source.value=item.url;source.setAttribute('aria-label','技能来源地址');
    const copy=make('button','复制来源地址',panel,'btn secondary');copy.type='button';copy.onclick=async()=>{
      try {await navigator.clipboard.writeText(item.url);copy.textContent='已复制';}
      catch (_) {source.focus();source.select();copy.textContent='请复制选中地址';}
    };
    make('p','安装状态以当前工作区“已安装”为准；未验证本轮 Agent 使用。',panel);
    const inspect=make('button','检查来源与兼容性',panel,'btn secondary');inspect.type='button';inspect.onclick=()=>inspectSource('resolve_public_skill',item.candidate_id);
    $('skills-public-results').hidden=true;back.focus();
  }
  async function inspectSource(method,value) {
    if(pending||!active())return;
    const mine=++epoch;pending=true;controls();status.textContent='正在固定版本并核对资源、许可…';
    try{
      const response=await window.OfficeAuth.wrapApi(window.pywebview.api)[method](value);
      if(mine!==epoch||!active())return;
      if(!response?.ok){status.textContent=({fixed_commit_required:'请使用含40位提交摘要的 GitHub 链接。',candidate_expired:'候选已过期，请重新查询。',no_file:'打开文档后才能安装到当前工作区。',body_budget:'技能主指令超过当前加载上限，暂不能安装；可查看来源原文。',package_budget:'技能包超过当前保存上限，暂不能安装。',plugin_skill_not_declared:'插件未声明这个 Skill 路径，请核对来源。',plugin_manifest_unsupported:'插件声明格式暂不支持，当前未安装。',source_rate_limited:'来源限流，请稍后主动重试。',source_timeout:'来源读取超时，可主动重试。'})[response?.reason]||'来源未通过检查（'+String(response?.reason||'unknown')+'），可以重新核对或重试。';return;}
      const panel=$('skills-public-detail');panel.replaceChildren();panel.hidden=false;$('skills-public-results').hidden=true;
      const back=make('button','返回结果',panel,'btn secondary');back.type='button';back.onclick=()=>{panel.hidden=true;$('skills-public-results').hidden=false;};
      iconFor(panel,response);make('h3',response.display?.display_name||response.name,panel);make('p',response.description,panel);
      if(response.icon_status==='fallback')make('p','来源图标未通过检查，使用分类图标。',panel);
      make('p',response.origin.repository+' · '+response.origin.license,panel);
      if(response.origin.plugin_context){const dependency=make('details','',panel);make('summary','插件声明与独立 Skill 范围',dependency);make('pre',JSON.stringify(response.origin.plugin_context,null,2),dependency);}
      make('p',(response.origin.provider==='lix'?'工作流版本：'+response.origin.version:'固定版本：'+response.origin.commit),panel);
      make('p',response.compatibility_note,panel);
      const details=make('details','',panel);make('summary','原文、参考与许可',details);make('pre',response.body+'\n\n'+response.license_text,details);
      if(response.origin.declarations&&Object.keys(response.origin.declarations).length)make('pre','来源声明（不授予权限）\n'+JSON.stringify(response.origin.declarations,null,2),details);
      for(const [name,item] of Object.entries(response.resources||{})){make('strong',name,details);make('pre',item.text,details);}
      if(response.compatible){
        make('p','受支持的指令、资料、展示资源与许可可完整封装；尚未安装或验证 Agent 使用。',panel);
        const install=make('button','预览安装到当前工作区',panel,'btn');install.type='button';install.onclick=async()=>{
          if(pending)return;pending=true;controls();install.disabled=true;
          const own=epoch;
          try {const outcome=await window.OfficeSkills.installPublic(response.resolution_id);if(own===epoch&&active()&&outcome?.text)status.textContent=outcome.text;}
          finally{pending=false;controls();if(install.isConnected)install.disabled=false;}
        };
      }else{make('p','当前无法完整安装：',panel);const list=make('ul','',panel);for(const reason of response.blockers)make('li',reason,list);
        if(response.progressive_available){const progressive=make('button','改用按需参考模式',panel,'btn secondary');progressive.type='button';progressive.onclick=()=>inspectSource('prepare_progressive_skill',response.resolution_id);}
      }
      status.textContent='已核对固定来源；安装仍需当前工作区和独立确认。';back.focus();
    }catch(_){if(mine===epoch&&active())status.textContent='来源检查未完成，可重新发起；尚未安装。';}
    finally{pending=false;controls();}
  }
  function render() {
    const list=$('skills-public-results');list.replaceChildren();list.hidden=false;$('skills-public-detail').hidden=true;
    const query=input.value.trim().toLocaleLowerCase();
    const matches=result?.query===input.value.trim()?items:items.filter(item=>(item.name+' '+item.source+' '+(item.display_name||'')+' '+(item.description||'')).toLocaleLowerCase().includes(query));
    for(const item of matches){
      const entry=make('button','',list,'skill-entry');entry.type='button';entry.dataset.candidate=item.candidate_id;
      iconFor(entry,item);
      const text=make('span','',entry,'skill-entry-text');make('strong',item.display_name||item.name,text);make('span',item.description||item.source,text,'skill-entry-description');make('span',item.provider==='featured'?item.source+' · 安装前检查兼容性':'兼容性待核对',text,'skill-entry-description');
      entry.onclick=()=>detail(item);
    }
    if(result&&!matches.length)make('p',result.ok===false?'本次查询未完成；可主动重试或改用其他来源。':result.query===input.value.trim()?'没有匹配的已收录技能。':'当前结果没有匹配项，可搜索公开目录。',list);
  }
  async function showFeatured(){
    if(!active()||pending)return;
    const mine=epoch;
    try{
      const response=await window.OfficeAuth.wrapApi(window.pywebview.api).list_featured_skills();
      if(mine!==epoch||!active())return;
      if(!response?.ok){status.textContent='精选目录不可读；仍可搜索公开技能或检查 GitHub 来源。';return;}
      featured=response;items=response.items;result={featured:true};render();
      status.textContent='精选 '+items.length+' 项 · 离线来源快照 · '+response.reviewed_at.slice(0,10)+' · 安装前检查兼容性';
    }catch(_){if(mine===epoch&&active())status.textContent='精选目录未加载；仍可主动搜索公开技能。';}
  }
  function invalidate(message) {
    epoch++;
    if(pending)status.textContent=message || '已停止等待；服务端查询可能仍在进行。';
    else if(message)status.textContent=message;
  }
  async function search() {
    if(pending||!active())return;
    const query=input.value.trim();
    const provider=sourceChoice.value;
    if((provider!=='lix'&&query.length<2)||query.length>120){status.textContent=errors.invalid_query;return;}
    const mine=++epoch;pending=true;controls();status.textContent=provider==='lix'?'正在读取本机工作流…':provider==='openai/plugins'?'正在读取 OpenAI 插件清单…':'正在查询 skills.sh…';
    $('skills-public-detail').hidden=true;
    try{
      const api=window.OfficeAuth.wrapApi(window.pywebview.api);
      const response=await (provider==='lix'?api.list_office_workflows(query):provider==='openai/plugins'?api.search_plugin_skills(query):api.search_public_skills(query));
      if(mine!==epoch||!active()||query!==input.value.trim()||provider!==sourceChoice.value)return;
      result=response;items=Array.isArray(response?.items)?response.items:[];
      if(items.some(item=>!item||!/^[a-f0-9]{32}$/.test(item.candidate_id)||(provider==='lix'?item.provider!=='lix'||item.source!=='LiX':!(item.provider==='openai/plugins'?/^https:\/\/github\.com\/openai\/plugins\/blob\/[a-f0-9]{40}\/plugins\/[A-Za-z0-9_.\/-]+\/SKILL\.md$/:/^https:\/\/skills\.sh\/[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/).test(item.url))))throw Error();
      render();
      status.textContent=response.ok?'已获取 '+items.length+' 项 · '+provider+' · 最多 24 项':({plugin_partial:'部分插件未能检查；当前仅显示可读取结果。',source_rate_limited:'来源限流，请稍后主动重试。',source_timeout:'来源超时，请主动重试。',plugin_manifest_unsupported:'插件清单格式暂不支持，仍可检查固定 GitHub Skill 链接。'})[response.reason]||(errors[response.reason]||'查询失败，请重试。');
      if(provider==='openai/plugins')status.textContent+=' 按插件名/类别查询'+(response.limited?'；已达本次检查上限，请缩小查询。':'。');
      if(provider==='lix'&&response.ok)status.textContent='LiX 提供 · '+items.length+' 项本机工作流 · 未联网 · 安装仍需独立确认';
      if(response.errors?.some(item=>item.reason==='plugin_source_unsupported'))status.textContent+=' 外部仓库/子目录插件来源暂不支持，可单独检查固定 Skill 链接。';
      if(response.cached)status.textContent+=' 显示本次运行的缓存（'+response.fetched_at+'），不是最新结果。';
    }catch(_){if(mine===epoch&&active())status.textContent='查询未完成，请重试。';}
    finally{pending=false;controls();}
  }
  $('skills-public-search').onclick=search;
  sourceChoice.onchange=()=>{invalidate(sourceChoice.value==='lix'?'输入已保留；点击浏览工作流，本机读取不联网。':'搜索来源已切换；输入已保留，按 Enter 查询。');$('skills-public-search').textContent=sourceChoice.value==='lix'?'浏览工作流':'搜索公开技能';items=[];result=null;render();};
  const featuredButton=make('button','精选目录',root.querySelector('.actions'),'btn secondary');featuredButton.type='button';featuredButton.onclick=()=>{invalidate();showFeatured();};
  $('skills-github-check').onclick=()=>inspectSource('resolve_github_skill',$('skills-github-url').value.trim());
  $('skills-public-cancel').onclick=()=>invalidate();
  input.addEventListener('input',()=>{if(active()){invalidate('按 Enter 搜索公开目录；当前仅筛选已获取结果。');render();}});
  input.addEventListener('keydown',event=>{if(event.key==='Enter'&&!event.isComposing&&active()){event.preventDefault();search();}});
  document.addEventListener('officeskills:view',event=>{invalidate();if(event.detail?.view==='discover'&&!featured)showFeatured();});
  document.addEventListener('officesettings:closed',()=>invalidate());
  document.addEventListener('officesettings:tab',event=>{if(event.detail?.tab!=='tab-skills')invalidate();});
  for(const event of ['officeauth:unauthenticated','pagehide'])window.addEventListener(event,()=>{invalidate();items=[];result=null;$('skills-public-results').replaceChildren();$('skills-public-detail').replaceChildren();});
})();
