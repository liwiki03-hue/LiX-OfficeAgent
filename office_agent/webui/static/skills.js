/* Human-only Skills management. Skill text is untrusted reference material. */
(() => {
  'use strict';
  const node = id => document.getElementById(id);
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  const fields = ['skills-slug', 'skills-name', 'skills-description', 'skills-body'];
  const actions = ['load', 'create', 'edit', 'delete', 'toggle', 'import', 'rollback'];
  let epoch = 0, revision = 0, inflight = false, selected = null, saveId = null;
  let uncertain = false, initialized = false;
  let rows = [], baseline = '', editing = false, view = 'installed', loaded = false;
  let reloadAfter = false, detached = null, lastScope = '', workspaceReady = null;
  const values = () => fields.map(id => node(id)?.value || '');
  const dirty = () => initialized && JSON.stringify(values()) !== baseline;
  const browser = () => Boolean(node('skills-search'));
  async function guard() {
    if (inflight) { status('操作尚未结束，请等待后再离开。'); return false; }
    if (dirty() && !await ask('放弃未保存的技能改动？已保存内容不会改变。')) return false;
    if (dirty()) {
      const saved = JSON.parse(baseline);
      fields.forEach((id, index) => { node(id).value = saved[index]; });
      revision++;
    }
    editing = false;
    refresh();
    return true;
  }
  const live = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const ready = () => live() && Boolean(window.pywebview?.api);
  const opaque = value => typeof value === 'string' && /^[a-f0-9]{32}$/.test(value);
  const slugValid = value => typeof value === 'string' && value.length <= 64 && /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(value);
  const scope = () => JSON.stringify(['file-name', 'file-dir'].map(id => node(id)?.textContent || ''));
  const identity = () => selected?.id || '';
  function status(message) { if (node('skills-status')) node('skills-status').textContent = message; }
  function refresh() {
    if (!initialized) return;
    const blocked = inflight || !ready() || uncertain;
    actions.forEach(action => { node('skills-' + action).disabled = blocked; });
    node('skills-edit').disabled ||= selected?.source !== 'user';
    node('skills-delete').disabled ||= selected?.source !== 'user';
    node('skills-toggle').disabled ||= !selected;
    node('skills-rollback').disabled ||= !opaque(saveId);
    if (workspaceReady === false) actions.filter(action => action !== 'load').forEach(action => { node('skills-' + action).disabled = true; });
    node('skills-toggle').textContent = selected?.enabled === false ? '启用技能' : '停用技能';
    fields.forEach(id => { node(id).disabled = blocked || (selected?.source === 'builtin'); });
    // Slug identifies the selected file and cannot rename it through edit.
    node('skills-slug').disabled ||= Boolean(selected);
    node('skills-list').querySelectorAll('button').forEach(button => { button.disabled = blocked; });
    node('skills-panel').setAttribute('aria-busy', String(inflight));
    actions.forEach(action => {
      const button = node('skills-' + action);
      button.title = !button.disabled ? '' : inflight ? '正在处理，请等待' : !ready() ? '本机服务尚未就绪' :
        uncertain ? '结果不确定，需核对记录' : workspaceReady === false ? '打开文档后管理工作区技能' : action === 'rollback' ? '当前没有可撤销的修改' : '先选择用户技能；内置技能只读';
    });
    if (browser()) {
      node('skills-editor').hidden = !editing;
      node('skills-detail').hidden = editing || !selected || view !== 'installed';
      node('skills-list').hidden = editing || view !== 'installed';
      node('skills-panel').dataset.view = editing ? 'editor' : selected ? 'detail' : view;
      node('skills-discovery').hidden = editing || view !== 'discover';
      node('skills-create').hidden = Boolean(selected);
      node('skills-edit').hidden = !selected;
      node('skills-open-editor').disabled = blocked || selected?.source !== 'user';
      node('skills-recover-draft').hidden = !detached;
      for (const id of ['skills-new', 'skills-installed', 'skills-discover', 'skills-detail-back', 'skills-editor-cancel', 'skills-recover-draft']) node(id).disabled = blocked;
      node('skills-rollback').hidden = !opaque(saveId);
      node('skills-installed').setAttribute('aria-pressed', String(view === 'installed'));
      node('skills-discover').setAttribute('aria-pressed', String(view === 'discover'));
      node('skills-status').hidden = view === 'discover';
      node('skills-scope').hidden = view === 'discover';
      node('skills-load').hidden = view === 'discover';
      node('skills-search').previousSibling.textContent = view === 'discover' ? '搜索公开技能' : '搜索已安装技能';
    }
  }
  function clearEditor() { selected = null; editing = false; fields.forEach(id => { node(id).value = ''; }); baseline = JSON.stringify(values()); }
  function retire() {
    const changed = lastScope && lastScope !== scope();
    if (browser() && changed && dirty()) detached = values();
    lastScope = scope(); loaded = false; rows = []; reloadAfter = false; workspaceReady = null;
    if (browser()) node('skills-scope').textContent = '打开文档后，查看和管理该工作区的技能。';
    epoch++; revision++; saveId = null;
    if (!initialized) return;
    clearEditor(); node('skills-list').replaceChildren();
    node('skills-panel').open = true;
    status(uncertain ? '技能操作结果没能确认；请核对本地文件、备份和记录，不要自动重试。' : detached ? '工作区已变化；未保存内容可恢复为新草稿，旧操作已失效。' : '打开技能页后读取当前工作区。');
    refresh();
  }
  function doubt() {
    uncertain = true; saveId = null; clearEditor(); node('skills-list').replaceChildren();
    status('技能操作结果没能确认；请核对本地文件、备份和记录，不要自动重试。');
  }
  function rejected(response) {
    if (response?.reason === 'no_file') workspaceReady = false;
    if (response?.outcome_uncertain === true || response?.reason === 'skills_uncertain' || response?.result?.status === 'uncertain') doubt();
    else status(({no_file:'尚未选择文档。打开文档后即可查看和管理工作区技能。', busy:'任务运行中，结束后可刷新和管理技能。', not_authenticated:'本机会话尚未就绪，请重新打开工作区。', skills_source_changed:'技能已变化或有本地修改，请刷新后重新核对。', skills_scope_changed:'工作区已变化，请重新检查更新。', candidate_expired:'检查结果已过期，请重新检查更新。', update_source_unavailable:'更新来源无法核对，可能不兼容或网络不可用；当前版本未改变。'})[response?.reason] || '技能操作被拒绝，未获得执行批准；请核对工作区及运行状态。');
  }
  async function run(operation) {
    if (inflight || !ready() || uncertain) return;
    inflight = true; status('正在处理工作区技能…'); refresh();
    const mine = epoch, version = revision, target = identity(), fileScope = scope();
    const current = () => mine === epoch && version === revision && target === identity() && fileScope === scope() && ready();
    try {
      const api = window.OfficeAuth.wrapApi(window.pywebview.api);
      await operation(api, current);
    } catch (_) { if (current()) doubt(); }
    finally {
      inflight = false; refresh();
      if (browser() && reloadAfter && mine === epoch && fileScope === scope() && ready()) {
        reloadAfter = false;
        await loadCatalog(false);
      }
    }
  }
  function metadata(item) {
    if (!item || !['user', 'builtin'].includes(item.source) || typeof item.id !== 'string') throw Error();
    const parts = item.id.split(':');
    if (parts.length !== 2 || parts[0] !== item.source || !slugValid(parts[1]) ||
        typeof item.name !== 'string' || typeof item.description !== 'string' || typeof item.enabled !== 'boolean') throw Error();
    return {...item, name:item.display_name||item.name, slug:parts[1]};
  }
  function showList(items) {
    if (!Array.isArray(items)) throw Error();
    rows = items.map(metadata); const seen = new Set();
    if (rows.some(item => seen.has(item.id) || !seen.add(item.id))) throw Error();
    clearEditor(); revision++; node('skills-list').replaceChildren();
    loaded = true; workspaceReady = true; renderList();
    if (browser()) node('skills-scope').textContent = '';
    status(rows.length ? '' : '当前工作区还没有技能。');
  }
  function renderList() {
    node('skills-list').replaceChildren();
    const query = (node('skills-search')?.value || '').trim().toLocaleLowerCase();
    const matches = rows.filter(item => (item.name + ' ' + item.description + ' ' + item.slug).toLocaleLowerCase().includes(query));
    matches.forEach(item => {
      const button = document.createElement('button'); button.type = 'button';
      button.className = 'skill-entry';
      if (browser()) {
        const icon = document.createElement('span'); icon.className = 'skill-entry-icon'; icon.setAttribute('aria-hidden', 'true');
        const text = document.createElement('span'); text.className = 'skill-entry-text';
        const title = document.createElement('strong'); title.textContent = item.name;
        const description = document.createElement('span'); description.textContent = item.description; description.className = 'skill-entry-description';
        const info = document.createElement('small'); info.className = 'skill-entry-state'; info.textContent = item.enabled ? '已启用' : '已停用';
        text.append(title, description); button.append(icon, text, info); button.title = item.name;
      } else button.textContent = item.name + ' — ' + item.description + ' [' + item.source + ' · ' + (item.enabled ? '已启用' : '已停用') + ']';
      button.setAttribute('aria-label', '读取技能 ' + item.name + ' (' + item.id + ')');
      button.addEventListener('click', () => select(item)); node('skills-list').append(button);
    });
    if (browser() && loaded && !matches.length) {
      const empty = document.createElement('p'); empty.className = 'skills-empty';
      empty.textContent = query ? '没有匹配的已安装技能。' : '还没有安装技能。通过“添加”创建自己的工作方法。';
      node('skills-list').append(empty);
    }
  }
  async function select(item) {
    if (inflight || !ready() || uncertain) return;
    if (browser() && !await guard()) return;
    clearEditor(); selected = item; revision++; refresh();
    return run(async (api, current) => {
      const response = await api.get_skill(item.slug, item.source);
      if (!current()) return;
      if (response?.ok !== true) { clearEditor(); rejected(response); return; }
      const skill = metadata(response.skill);
      if (skill.id !== item.id || typeof skill.body !== 'string') throw Error();
      selected = skill;
      node('skills-slug').value = item.slug; node('skills-name').value = skill.name;
      node('skills-description').value = skill.description; node('skills-body').value = skill.body;
      baseline = JSON.stringify(values());
      if (browser()) {
        const presentation=skill.origin?.presentation;
        node('skills-detail-title').textContent = presentation?.display?.display_name||skill.origin?.display_name||skill.name;
        node('skills-detail-icon')?.remove();
        if(presentation?.icon?.data_url&&/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(presentation.icon.data_url)&&presentation.icon.data_url.length<150000){
          const icon=document.createElement('span');icon.id='skills-detail-icon';icon.className='skill-entry-icon has-source-icon';icon.setAttribute('aria-hidden','true');
          const image=document.createElement('img');image.src=presentation.icon.data_url;image.alt='';image.width=28;image.height=28;icon.append(image);node('skills-detail-title').before(icon);
          image.onerror=()=>icon.remove();
        }
        node('skills-detail-description').textContent = skill.description;
        node('skills-detail-meta').textContent = (skill.source === 'builtin' ? '内置技能' : '当前工作区') + ' · ' + (skill.enabled ? '已启用' : '已停用') + ' · ' + skill.id + (skill.origin?' · '+(skill.origin.provider==='lix'?'LiX 提供 @ '+skill.origin.version:skill.origin.repository+' @ '+skill.origin.commit)+' · '+skill.origin.license+(skill.origin.local_modified?' · 本地已修改':' · 指令与原文一致'):' · 未登记公开来源与版本');
        let license=node('skills-detail-license');
        if(!license){license=document.createElement('details');license.id='skills-detail-license';const summary=document.createElement('summary');summary.textContent='来源与许可';license.append(summary,document.createElement('pre'));node('skills-detail-meta').after(license);}
        license.hidden=!skill.origin;
        license.querySelector('pre').textContent=skill.origin?skill.origin.url+'\n'+skill.origin.format+'\n'+skill.origin.license_text+'\n原始描述：'+(skill.origin.original_description||skill.description)+'\n来源声明（不授予权限）\n'+JSON.stringify(skill.origin.declarations||{},null,2):'';
        node('skills-detail-body').textContent = skill.body;
        publicDetails(skill);
        node('skills-detail').hidden = false;
        node('skills-detail-back').focus();
      }
      status('已读取。编辑、启停与删除均需单独确认。');
    });
  }
  function publicDetails(skill) {
    node('skills-public-management')?.remove();
    if (!skill.origin) return;
    const panel=document.createElement('div');panel.id='skills-public-management';node('skills-detail-body').closest('details').after(panel);
    const add=(tag,text,parent=panel)=>{const element=document.createElement(tag);element.textContent=text;parent.append(element);return element;};
    const resources=Object.entries(skill.origin.resources||{});
    if(skill.origin.presentation){
      const details=add('details','');add('summary','展示信息与图标',details);
      add('p','仅用于界面；不发送给模型，不授予权限。',details);
      add('pre',JSON.stringify(skill.origin.presentation.display,null,2),details);
      for(const file of skill.origin.presentation.files)add('p',file.path+' · '+file.size+' 字节',details);
    }
    if(resources.length){
      const details=add('details','');add('summary','包内参考 · '+resources.length+' 项',details);
      for(const [name,item] of resources){add('strong',name,details);add('pre',item.text,details);}
      add('p',skill.origin.reference_loading==='task-matched'?'按需参考：按本次要求匹配文件名和首行；未加载项会列出，不能当作已读。':'Agent 选中此技能时一并读取这些文字；不增加文件权限。',details);
    }
    if(skill.origin.provider==='lix'){add('p','LiX 工作流随应用更新；本地编辑与撤销仍需单独确认。');return;}
    if(skill.source!=='user')return;
    const check=add('button','检查更新');check.type='button';check.className='btn secondary';
    check.onclick=()=>run(async(api,current)=>{
      check.disabled=true;
      let response;
      try{response=await api.check_public_skill_update(skill.slug,skill.source);}
      catch(_){if(current())status('更新检查未完成，当前版本未改变。');return;}
      finally{if(check.isConnected)check.disabled=false;}
      if(!current())return;
      if(!response?.ok){rejected(response);return;}
      panel.querySelector('[data-update]')?.remove();
      const comparison=add('section','');comparison.dataset.update='true';
      add('h4','版本比较',comparison);
      add('p',response.update.current.commit+' → '+response.origin.commit,comparison);
      add('p',response.update.permissions,comparison);
      const files=add('ul','',comparison);
      for(const item of response.update.files)add('li',item.status+' · '+item.path,files);
      if(!response.update.files.length)add('p','文件内容一致；仅固定版本信息可能变化。',comparison);
      if(response.blockers.length){for(const reason of response.blockers)add('p',reason,comparison);}
      else{
        const preview=add('button','预览更新',comparison);preview.type='button';preview.className='btn';
        preview.onclick=()=>{if(!current())return;return run(async(nextApi,nextCurrent)=>{
          const plan=await nextApi.plan_public_skill_update(response.resolution_id);
          await approve(nextApi,nextCurrent,plan,'edit');
        });};
      }
      status('更新检查完成；尚未写入。');
    });
  }
  function content() {
    const name = node('skills-name').value.trim(), description = node('skills-description').value.trim();
    if (!name || Array.from(name).length > 100 || !description || Array.from(description).length > 500) {
      status('名称需为 1–100 字符，描述需为 1–500 字符。'); return null;
    }
    const text = '---\nname: ' + JSON.stringify(name) + '\ndescription: ' + JSON.stringify(description) + '\n---\n' + node('skills-body').value;
    if (new TextEncoder().encode(text).length > 65536 || /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(text)) {
      status('SKILL.md 需为有效 UTF-8，且不得超过 65536 字节。'); return null;
    }
    return text;
  }
  async function approve(api, current, response, operation) {
    if (!current()) return;
    if (response?.ok !== true) { rejected(response); return; }
    if (response.cancelled === true) { status('已取消导入；未批准写入。'); return; }
    const plan = response.plan;
    if (!opaque(plan?.plan_id) || plan.operation !== operation || plan.confirmation_required !== true || plan.wrote_nothing !== true) { doubt(); return; }
    if (operation === 'enabled' && typeof plan.enabled !== 'boolean') { doubt(); return; }
    const action = operation === 'enabled' ? (plan.enabled ? '启用技能' : '停用技能') : ({create:'新增技能',edit:'保存技能修改',delete:'删除技能',import:'导入技能',update:'更新技能'})[operation];
    if (!action) { doubt(); return; }
    const confirmed = await ask('确认' + action + '？\n目标：' + String(plan.id || '') +
      (plan.origin?'\n工作区：'+plan.workspace+'\n来源：'+plan.origin.repository+'\n固定版本：'+(plan.origin.provider==='lix'?plan.origin.version:plan.origin.commit)+'\n许可：'+plan.origin.license+'\n'+plan.format+'\n仅作文字参考，安装不代表 Agent 已使用。':'')+
      (plan.previous_origin?'\n原版本：'+plan.previous_origin.commit+'\n文件变化：'+plan.changes.map(item=>item.status+' '+item.path).join('；'):'')+
      '\n仅改动当前工作区的技能或启用状态；公开技能保留已核验的受支持资料。文档不会因此被修改。\n不会新增工具权限；操作留有备份与记录，可另行确认撤销。') === true;
    if (!current()) return;
    const result = await api.confirm_skill_change(plan.plan_id, confirmed === true ? true : false);
    if (!current()) return;
    const outcome = result?.result;
    if (result?.outcome_uncertain === true || outcome?.outcome_uncertain === true || outcome?.uncertain === true) { doubt(); return; }
    if (result?.ok === true && outcome?.status === 'saved' && outcome.applied === true && outcome.verified === true && opaque(outcome.save_id)) {
      saveId = outcome.save_id; clearEditor(); revision++; node('skills-list').replaceChildren();
      reloadAfter = true;
      status('技能操作已保存并通过校验；可以重新读取目录，或撤销这次操作。');
    } else if (result?.ok === false && ['skills_denied', 'confirmation_required'].includes(result.reason) && result.outcome_uncertain !== true) {
      status('已取消，没有写入；再次操作需要重新发起并确认一次。');
    } else if (result?.ok === true && outcome?.status === 'rejected' && outcome.wrote_nothing === true && outcome.applied === false) {
      status('没有写入；再次操作需要重新发起并确认一次。');
    } else doubt();
  }
  function change(operation) {
    return run(async (api, current) => {
      if (['edit', 'delete'].includes(operation) && selected?.source !== 'user') return;
      if (operation !== 'create' && !selected) return;
      let changes;
      if (operation === 'create') {
        const slug = node('skills-slug').value;
        if (!slugValid(slug)) { status('slug 需为最多 64 字符的小写字母、数字及单个连接短横线。'); return; }
        const text = content(); if (text === null) return; changes = {slug, content: text};
      } else {
        changes = {name: selected.slug, source: selected.source};
        if (operation === 'edit') { const text = content(); if (text === null) return; changes.content = text; }
        if (operation === 'enabled') changes.enabled = !selected.enabled;
      }
      const response = await api.plan_skill_change(operation, changes);
      if (!current()) return;
      await approve(api, current, response, operation);
    });
  }
  function init() {
    if (initialized || !node('skills-panel')) return;
    if ([...fields, ...actions.map(action => 'skills-' + action), 'skills-list', 'skills-status'].some(id => !node(id))) return;
    initialized = true;
    node('skills-load').addEventListener('click', () => loadCatalog());
    fields.forEach(id => node(id).addEventListener('input', () => { revision++; refresh(); }));
    ['create', 'edit', 'delete'].forEach(operation => node('skills-' + operation).addEventListener('click', () => change(operation)));
    node('skills-toggle').addEventListener('click', () => change('enabled'));
    node('skills-import').addEventListener('click', async () => {
      if (browser() && !await guard()) return;
      return run(async (api, current) => {
      const response = await api.plan_skill_import(); if (!current()) return;
      await approve(api, current, response, 'import');
      });
    });
    node('skills-rollback').addEventListener('click', () => run(async (api, current) => {
      const target = saveId; if (!opaque(target)) return;
      const response = await api.plan_skill_rollback(target); if (!current()) return;
      if (response?.ok !== true) { saveId = null; rejected(response); return; }
      const plan = response.plan;
      if (!opaque(plan?.rollback_id) || plan.operation !== 'rollback' || plan.confirmation_required !== true || plan.wrote_nothing !== true) { doubt(); return; }
      const confirmed = await ask('确认撤销这次的技能操作？\n会从已校验的备份还原，记录和资源都保留；技能不会因此获得执行脚本的权限。') === true;
      if (!current() || saveId !== target) return;
      const result = await api.confirm_skill_rollback(plan.rollback_id, confirmed === true ? true : false); if (!current()) return;
      const outcome = result?.result;
      if (result?.outcome_uncertain === true || outcome?.outcome_uncertain === true || outcome?.uncertain === true) { doubt(); return; }
      if (result?.ok === true && outcome?.status === 'rolled_back' && outcome.restored === true && outcome.verified === true) {
        saveId = null; clearEditor(); revision++; node('skills-list').replaceChildren(); status('已撤销并通过校验；请重新读取技能目录。');
        reloadAfter = true;
      } else if (result?.ok === true && outcome?.status === 'rejected' && outcome.wrote_nothing === true && outcome.restored === false) status('没有执行撤销；再次撤销需要重新发起并确认。');
      else doubt();
    }));
    ['file-name', 'file-dir'].forEach(id => {
      if (node(id)) new MutationObserver(retire).observe(node(id), {childList: true, characterData: true, subtree: true});
    });
    retire();
    if (browser()) initBrowser();
  }
  async function loadCatalog(check = true) {
    if (browser() && check && !await guard()) return;
    return run(async (api, current) => {
      const response = await api.list_skills(); if (!current()) return;
      if (response?.ok !== true) { rejected(response); return; }
      showList(response.skills);
      if (browser() && response.agent_catalog) {
        const report=response.agent_catalog;
        node('skills-scope').textContent='当前工作区 · Agent 目录 '+report.available+' 项'+(report.excluded.length?'，'+report.excluded.length+' 项停用或未纳入':'')+' · 模型使用描述摘要';
      }
    });
  }
  function initBrowser() {
    const activate = () => {
      if (window.OfficeSettings?.isOpen() && !node('settings-pane-skills').hidden && !loaded && !dirty() && !inflight && ready()) loadCatalog(false);
    };
    const setView = async next => {
      if (!await guard()) return;
      clearEditor(); revision++; view = next; refresh();
      document.dispatchEvent(new CustomEvent('officeskills:view', {detail:{view: next}}));
      if (next === 'installed') activate();
    };
    const begin = async recovered => {
      if (!await guard()) return;
      clearEditor(); revision++; view = 'installed'; editing = true;
      node('skills-slug').value = 'skill-' + Array.from(window.crypto.getRandomValues(new Uint8Array(8)), value => value.toString(16).padStart(2, '0')).join('');
      baseline = JSON.stringify(values());
      if (recovered) {
        fields.slice(1).forEach((id, index) => { node(id).value = recovered[index + 1]; });
        detached = null;
      }
      node('skills-editor-title').textContent = '新建技能';
      node('skills-add-menu').open = false; refresh(); node('skills-name').focus();
    };
    node('skills-new').addEventListener('click', () => begin());
    node('skills-recover-draft').addEventListener('click', () => begin(detached));
    node('skills-installed').addEventListener('click', () => setView('installed'));
    node('skills-back-installed').addEventListener('click', () => setView('installed'));
    node('skills-discover').addEventListener('click', () => setView('discover'));
    node('skills-search').addEventListener('input', () => { renderList(); refresh(); });
    node('skills-search').addEventListener('keydown', event => { if (view === 'installed' && event.key === 'Enter' && !event.isComposing) { event.preventDefault(); renderList(); } });
    node('skills-open-editor').addEventListener('click', () => { editing = true; node('skills-editor-title').textContent = '编辑技能'; refresh(); node('skills-name').focus(); });
    node('skills-detail-back').addEventListener('click', async () => { if (await guard()) { clearEditor(); revision++; refresh(); } });
    node('skills-editor-cancel').addEventListener('click', async () => { if (await guard()) refresh(); });
    document.addEventListener('officesettings:opened', activate);
    document.addEventListener('officesettings:tab', event => {
      if (event.detail?.tab === 'tab-skills' && !dirty()) { loaded = false; activate(); }
    });
    document.addEventListener('officesettings:closed', () => { epoch++; revision++; });
    window.addEventListener('officeapp:state', activate);
    window.OfficeSkills = {requestClose: guard, dirty, refresh: activate, installPublic: async resolution => {
      if (!await guard()) return {installed:false,text:node('skills-status').textContent};
      const previousSave=saveId;
      await run(async (api,current) => {
        const response=await api.plan_public_skill_install(resolution);
        await approve(api,current,response,'create');
      });
      if(saveId&&saveId!==previousSave){view='installed';refresh();document.dispatchEvent(new CustomEvent('officeskills:view',{detail:{view:'installed'}}));}
      return {installed:Boolean(saveId&&saveId!==previousSave),text:node('skills-status').textContent};
    }};
    window.addEventListener('pagehide', () => { detached = null; });
  }
  ['officeauth:unauthenticated', 'officeauth:authenticated', 'officediagram:invalidate', 'officeskills:invalidate', 'pagehide'].forEach(event => window.addEventListener(event, retire));
  window.addEventListener('pywebviewready', async () => {
    const mine = epoch;
    try { await window.OfficeAuth?.ensure(); } catch (_) { /* Keep the authentication gate closed. */ }
    if (mine === epoch) refresh();
  });
  document.addEventListener('DOMContentLoaded', init);
  if (document.readyState !== 'loading') init();
})();
