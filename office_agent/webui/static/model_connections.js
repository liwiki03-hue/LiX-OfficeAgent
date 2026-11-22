(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  if (!$('workspace-document-bar') || !$('settings-pane-model')) return;
  const labels = {chat: '普通问答', agent: '文档 Agent', image: '生图'};
  const states = {verified: '实际调用已验证', unverified: '未验证', failed: '检测失败', changed: '配置已变更'};
  const errors = {not_authenticated: '登录已失效，请重新登录。', busy: '有任务或检测正在进行，请结束后再保存或检测。',
    settings_source_changed: '配置已被修改，请重新读取后再操作。', settings_plan_expired: '保存计划已过期，请重新预览。',
    settings_denied: '已取消，没有写入；新密钥已清空。', settings_no_changes: '没有实际改动。',
    settings_invalid: '配置无效。请核对地址、模型、用途和密钥；云端使用 HTTPS，Ollama HTTP 仅限回环地址。',
    model_selection_required: '此用途尚未选择模型，请配置接口或选择已保存模型。',
    model_selection_invalid: '当前选择已失效，请重新选择或修复连接。', model_configuration_changed: '当前连接配置已变更，请重新核对并选择。',
    settings_path_changed: '配置路径安全检查未通过。', settings_file_invalid: '配置格式无效，请检查文件。',
    settings_uncertain: '保存结果不确定，停止自动操作；请核对配置与审计记录后重启。',
    settings_io_failed: '配置读写失败，请核对权限和文件占用。'};
  const probeLabels = {reachable: '模型列表已获取；不代表模型调用或工具能力已验证。', authentication_failed: '认证失败，请检查密钥。',
    models_unsupported: '无法通过模型列表验证；可以保留并手动配置模型。', timeout: '检测超时，没有重试。',
    redirect_refused: '服务返回重定向，未携带凭据跟随。', service_failed: '服务返回失败状态。', models_unverified: '服务未返回可验证的模型列表。',
    response_too_large: '模型列表超过大小上限。', connection_failed: '连接失败，请核对 API 地址和证书。'};
  let snapshot = null, draft = null, baseline = '', revision = '', mode = 'agent', epoch = 0, sequence = 0;
  let pending = false, selecting = false, editorId = null, saveId = null, uncertain = false;
  let selectionBlocked = false;
  let modelProbeSequence=0, modelProbePending=false;
  let chosenCandidate=null;
  const live = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const copy = object => JSON.parse(JSON.stringify(object));
  const value = selection => JSON.stringify([selection.connection_id, selection.model_id, selection.capability]);
  function make(tag, id, text, parent) {
    const node = document.createElement(tag);
    if (id) node.id = id;
    if (text) node.textContent = text;
    if (parent) parent.append(node);
    return node;
  }
  function button(id, text, parent, fn) {
    const node = make('button', id, text, parent);
    node.type = 'button'; node.className = 'btn secondary'; node.onclick = fn;
    return node;
  }
  const toolbar = make('section', 'model-selector-bar', '', $('workspace-document-bar'));
  toolbar.setAttribute('aria-label', '当前请求模型');
  const modeLabel = make('label', '', '用途', toolbar);
  const modeInput = make('select', 'model-mode', '', modeLabel);
  Object.entries(labels).forEach(([id, name]) => { const option = make('option', '', name, modeInput); option.value = id; });
  modeInput.value = mode;
  const searchLabel = make('label', '', '查找模型或连接', toolbar);
  const search = make('input', 'model-search', '', searchLabel); search.type = 'search'; search.placeholder = '搜索连接或模型';
  const selectLabel = make('label', '', '模型与接口', toolbar);
  const selector = make('select', 'model-selector', '', selectLabel);
  const status = make('p', 'model-selection-status', '登录后读取已保存模型；不会自动联网检测。', toolbar);
  status.setAttribute('role', 'status');
  const actions = make('div', '', '', toolbar); actions.className = 'actions';
  button('model-configure', '配置接口', actions, configure);
  button('model-refresh', '刷新已保存配置', actions, () => refresh());
  button('model-default', '设为本用途默认', actions, setDefault);
  button('model-open-input', '打开对话输入', actions, () => window.dispatchEvent(new Event('officeworkspace:compose')));
  const manager = make('details', 'connection-manager', '', $('settings-pane-model'));
  $('settings-pane-model').prepend(manager);
  manager.prepend(make('summary', '', '接口与模型 · 支持多个命名连接'));
  make('p', '', '本机共享配置。仅已保存且启用的模型进入主界面；密钥不进入浏览器存储。工具用途是配置声明，需实际任务验证。', manager);
  const list = make('div', 'connection-list', '', manager);
  const edit = make('form', 'connection-editor', '', manager);
  edit.onsubmit = event => event.preventDefault(); edit.hidden = true;
  function field(id, label, type = 'text') {
    const node = make('label', '', label, edit);
    const input = make(type === 'textarea' ? 'textarea' : type === 'select' ? 'select' : 'input', id, '', node);
    if (input.tagName === 'INPUT') input.type = type;
    return input;
  }
  field('connection-name', '连接自定义名称');
  const providerChoice=field('connection-provider','模型提供方','select');
  for(const [id,title] of [['custom','自定义兼容服务'],['deepseek','DeepSeek'],['dashscope','百炼'],['ollama','Ollama']]){const item=make('option','',title,providerChoice);item.value=id;}
  const protocol = field('connection-protocol', '接口协议', 'select');
  [['openai', 'OpenAI 兼容接口'], ['ollama', 'Ollama']].forEach(([id, name]) => { const option = make('option', '', name, protocol); option.value = id; });
  field('connection-url', 'API 地址');
  const context = field('connection-context', '上下文窗口', 'number'); context.min = '1024'; context.max = '2097152';
  const enabled = field('connection-enabled', '启用连接', 'checkbox');
  const keyAction = field('connection-key-action', '密钥处理', 'select');
  [['keep', '留空保持旧密钥'], ['replace', '显式替换'], ['clear', '显式清除']].forEach(([id, name]) => { const option = make('option', '', name, keyAction); option.value = id; });
  const key = field('connection-key', 'API 密钥（仅本次保存使用）', 'password'); key.autocomplete = 'new-password'; key.disabled = true;
  keyAction.onchange = () => { key.value = ''; key.disabled = keyAction.value !== 'replace'; };
  make('p', 'connection-key-state', '', edit);
  const models = field('connection-models', '手动模型：每行 模型ID | 用途（chat,agent,image）', 'textarea'); models.rows = 5;
  const modelName=field('connection-model-name','模型名称');modelName.setAttribute('list','connection-model-candidates');modelName.autocomplete='off';
  const candidates=make('datalist','connection-model-candidates','',edit);
  const purpose=field('connection-model-purpose','用途（配置声明，仍需实际验证）','select');
  for(const [id,title] of [['chat','普通问答'],['chat,agent','问答与文档 Agent'],['image','图片生成']]){const item=make('option','',title,purpose);item.value=id;}
  const probeState=make('p','connection-probe-state','可直接填写模型名称。',edit);probeState.setAttribute('role','status');
  button('connection-probe-draft','刷新模型列表',edit,()=>probeDraft(true));
  modelName.addEventListener('change',()=>{
    const id=modelName.value.trim();if(!/^[A-Za-z0-9][A-Za-z0-9._:/@+-]{0,199}$/.test(id)||!editorId)return;
    const existing=models.value.split('\n').filter(Boolean),index=existing.findIndex(line=>line.split('|')[0].trim()===id);
    const line=id+' | '+purpose.value;if(index<0)existing.push(line);else existing[index]=line;models.value=existing.join('\n');
    for(const capability of purpose.value.split(','))draft.defaults[capability]={connection_id:editorId,model_id:id,capability};
    chosenCandidate={connection_id:editorId,model_id:id,capability:purpose.value.split(',').includes(mode)?mode:purpose.value.split(',')[0]};
    probeState.textContent='已填入配置草稿；保存后使用，能力尚未验证。';controls();
  });
  providerChoice.onchange=()=>{
    const provider=providerChoice.value;
    if(provider==='custom')return;
    protocol.value=provider==='ollama'?'ollama':'openai';
    $('connection-url').value=snapshot?.provider_defaults?.[provider]?.base_url||'';
    if(!$('connection-name').value)$('connection-name').value=providerChoice.selectedOptions[0].textContent;
    keyAction.value=provider==='ollama'?'clear':'replace';key.value='';key.disabled=provider==='ollama';
    modelProbeSequence++;candidates.replaceChildren();controls();
    if(provider==='ollama')probeDraft(false);
  };
  for(const node of [$('connection-url'),protocol,key,keyAction])node.addEventListener('input',()=>{modelProbeSequence++;candidates.replaceChildren();});
  make('p', '', '例如使用你的实际模型ID，加上 | chat,agent。生图单独使用 image；不会根据名称猜测能力。可追加第三列输出限制字段：auto（当前适配器默认）、max_tokens、max_completion_tokens；Ollama 使用 num_predict；unsupported 表示不支持受限回答。须按接口文档选择，不会失败后自动换参数。刷新列表不覆盖手动条目。', edit);
  button('connection-apply-draft', '保留到本次草稿', edit, applyEditor);
  const footer = make('div', '', '', manager); footer.className = 'actions';
  button('connection-add', '新增接口', footer, addConnection);
  button('connection-reload', '重新读取', footer, () => reloadDraft());
  button('connection-save', '预览并保存接口', footer, save);
  button('connection-discard', '放弃改动', footer, () => discard());
  button('connection-undo', '撤销本次保存', footer, undo);
  const message = make('p', 'connection-status', '尚未读取。', manager); message.setAttribute('role', 'status');
  const feedback = make('div', 'connection-feedback', '', manager);
  const changes = make('pre', 'connection-differences', '', manager); changes.hidden = true;
  let closeReview = null;
  function reviewConfiguration(plan) {
    return new Promise(resolve => {
      const previousFocus = document.activeElement;
      const dialog = make('dialog', 'connection-save-review', '', document.body);
      dialog.className = 'connection-review';
      dialog.setAttribute('aria-labelledby', 'connection-review-title');
      make('h2', 'connection-review-title', '确认接口配置改动', dialog);
      make('p', '', '请审阅全部差异。所有本机账户共享配置；确认后备份、写入、校验并审计，不会调用模型。新配置用于后续请求，运行中的任务保持原配置。', dialog);
      const content = make('div', 'connection-review-content', '', dialog); content.tabIndex = 0;
      content.setAttribute('aria-label', '完整接口配置差异');
      const labels = {name:'连接名称', protocol:'接口协议', base_url:'API 地址', enabled:'启用状态', context_window:'上下文窗口', models:'模型、用途及输出字段'};
      const display = value => value === undefined ? '未配置' : typeof value === 'string' ? value : JSON.stringify(value, null, 2);
      for (const item of plan.differences) {
        const section = make('section', '', '', content);
        make('h3', '', item.connection || (item.field === 'defaults' ? '各用途默认模型' : '配置变更'), section);
        if (item.operation === 'delete') make('p', '', '删除此连接；当前选择可能失效，需要重新选择。', section);
        else {
          const table = make('table', '', '', section), head = make('tr', '', '', make('thead', '', '', table));
          ['字段', '修改前', '修改后'].forEach(text => make('th', '', text, head));
          const body = make('tbody', '', '', table);
          const fields = item.field === 'defaults' ? ['chat','agent','image'] : Object.keys(labels);
          for (const field of fields) {
            if (JSON.stringify(item.before?.[field]) === JSON.stringify(item.after?.[field])) continue;
            const row = make('tr', '', '', body);
            make('th', '', labels[field] || ({chat:'普通问答',agent:'文档 Agent',image:'生图'})[field], row);
            make('td', '', display(item.before?.[field]), row);
            make('td', '', display(item.after?.[field]), row);
          }
        }
        if (item.key_action) make('p', '', '密钥处理：' + ({keep:'保持',replace:'替换',clear:'清除'})[item.key_action] + '（不显示原文）', section);
      }
      const details = make('details', '', '', content);
      make('summary', '', '完整脱敏计划详情', details);
      make('pre', '', JSON.stringify(plan.differences, null, 2), details);
      const actions = make('div', '', '', dialog); actions.className = 'actions';
      let settled = false;
      const finish = approved => {
        if (settled) return;
        settled = true; closeReview = null;
        dialog.close(); dialog.remove(); resolve(approved);
        if (previousFocus?.isConnected && !previousFocus.disabled) previousFocus.focus();
      };
      closeReview = finish;
      const cancel = button('connection-review-cancel', '取消，不保存', actions, () => finish(false));
      button('connection-review-confirm', '确认保存接口', actions, () => finish(true));
      dialog.addEventListener('keydown', event => {
        event.stopPropagation();
        if (event.key === 'Escape') { event.preventDefault(); finish(false); }
        if (event.key === 'Tab') {
          const targets = [content, ...dialog.querySelectorAll('summary,button')];
          const first = targets[0], last = targets.at(-1);
          if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
          else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
        }
      });
      dialog.addEventListener('cancel', event => { event.preventDefault(); finish(false); });
      dialog.addEventListener('click', event => { if (event.target === dialog) { const box = dialog.getBoundingClientRect(); if (event.clientX < box.left || event.clientX > box.right || event.clientY < box.top || event.clientY > box.bottom) finish(false); } });
      dialog.addEventListener('close', () => finish(false), {once:true});
      dialog.showModal(); cancel.focus();
    });
  }
  const editorValue = () => JSON.stringify(['connection-name', 'connection-protocol', 'connection-url', 'connection-context', 'connection-enabled', 'connection-models'].map(id => $(id).type === 'checkbox' ? $(id).checked : $(id).value));
  let editorBaseline = '';
  const dirty = () => Boolean(draft && (JSON.stringify(draft) !== baseline || editorId && editorValue() !== editorBaseline || key.value || keyAction.value !== 'keep'));
  function error(reply) {
    const uncertainResult = reply?.outcome_uncertain || reply?.result?.status === 'uncertain';
    if (uncertainResult) uncertain = true;
    message.textContent = errors[uncertainResult ? 'settings_uncertain' : reply?.reason || reply?.result?.code] || '操作未完成，请核对当前状态；未知结果不自动重试。';
    controls();
  }
  function controls() {
    const blocked = pending || !live() || uncertain;
    manager.querySelectorAll('button,input,select,textarea').forEach(node => { node.disabled = blocked; });
    key.disabled = blocked || keyAction.value !== 'replace';
    $('connection-save').disabled = blocked || !dirty();
    $('connection-discard').disabled = blocked || !dirty();
    $('connection-undo').disabled = blocked || !saveId;
    selector.disabled = selecting || !snapshot || !live() || uncertain;
    modeInput.disabled = selecting || !live() || uncertain;
    const current = snapshot?.catalog.current[mode], saved = snapshot?.catalog.defaults[mode];
    const isDefault = current?.valid && saved?.connection_id === current.connection_id && saved?.model_id === current.model_id;
    $('model-default').textContent = isDefault ? '当前为本用途默认' : '设为本用途默认';
    $('model-default').disabled = pending || selecting || selectionBlocked || !current?.valid || !live() || uncertain || isDefault;
    $('model-refresh').disabled = selecting || !live();
    list.querySelectorAll('[data-requires-saved="no"]').forEach(node => { node.disabled = true; });
  }
  function rowsForMode() {
    return snapshot?.catalog.connections.filter(row => row.enabled).flatMap(row => row.models.filter(model => model.capabilities.includes(mode)).map(model => ({row, model}))) || [];
  }
  function renderSelector() {
    selector.replaceChildren();
    const current = snapshot?.catalog.current[mode];
    const hint = make('option', '', current?.valid ? '选择另一个模型' : '当前选择无效或尚未配置，请选择', selector); hint.value = '';
    const query = search.value.trim().toLowerCase();
    const groups = new Map();
    rowsForMode().forEach(({row, model}) => {
      if (query && !`${row.name} ${model.id}`.toLowerCase().includes(query)) return;
      if (!groups.has(row.id)) { const group = make('optgroup', '', '', selector); group.label = row.name; groups.set(row.id, group); }
      const option = make('option', '', `${model.id} · ${row.name} · ${labels[mode]} · ${states[model.states[mode]]}`, groups.get(row.id));
      option.value = value({connection_id: row.id, model_id: model.id, capability: mode});
      option.disabled = !row.api_key_configured && row.protocol !== 'ollama';
    });
    if (current?.valid && Array.from(selector.options).some(o => o.value === value(current))) selector.value = value(current);
    status.textContent = current?.valid ? `当前：${current.model_id} · ${current.connection_name} · ${labels[mode]}。新选择用于下一次请求；运行中的任务保持原选择。` : errors[current?.reason] || '没有可用模型，请配置接口；输入不会清空。';
    if (current?.valid) {
      const item = rowsForMode().find(({row, model}) => row.id === current.connection_id && model.id === current.model_id);
      if (item) {
        status.textContent += ' ' + states[item.model.states[mode]] + '。';
        if (item.row.model_list_status !== 'not_requested') status.textContent += ' 列表检测：' + (probeLabels[item.row.model_list_status] || '未验证');
      }
    }
    if (mode === 'chat') status.textContent += ' 普通问答不读取文档或调用编辑工具；仅保留有界的本机对话上下文。';
    if (mode === 'image') status.textContent += ' 此入口每次一张，尺寸与风格在独立生图确认中核对；不会自动重试。';
    controls();
  }
  function cleanDraft(result) {
    draft = {connections: result.catalog.connections.map(row => ({...Object.fromEntries(Object.entries(row).filter(([k]) => !['revision', 'api_key_configured', 'model_list_status'].includes(k))),
      models: row.models.map(model => ({id: model.id, capabilities: [...model.capabilities], ...('output_parameter' in model ? {output_parameter:model.output_parameter} : {})})), key_action: 'keep'})), defaults: copy(result.catalog.defaults)};
    baseline = JSON.stringify(draft); editorId = null; edit.hidden = true; key.value = ''; keyAction.value = 'keep';
    revision = result.revision; renderConnections();
  }
  async function refresh(reset = false, duringSelection = false) {
    if (!live() || selecting && !duringSelection) return;
    const mine = epoch, request = ++sequence;
    try {
      const result = await api().get_model_connections();
      if (mine !== epoch || request !== sequence || !live()) return;
      if (!result?.ok) { selectionBlocked = true; status.textContent = errors[result?.reason] || '无法读取模型配置，请刷新核对后再发送。'; controls(); return; }
      selectionBlocked = false;
      snapshot = result; mode = result.catalog.mode; modeInput.value = mode;
      if (reset || !draft || !dirty() && !editorId) cleanDraft(result);
      renderSelector();
      window.dispatchEvent(new CustomEvent('officemodels:mode', {detail: {mode}}));
    } catch (_) { if (mine === epoch) { selectionBlocked = true; status.textContent = '无法读取配置；请确认宿主连接，刷新核对后再发送。'; controls(); } }
  }
  function configure() {
    window.OfficeSettings?.show(); window.OfficeSettings?.selectTab('tab-model'); manager.open = true;
    if (!draft) refresh(true); manager.scrollIntoView({block: 'start'});
  }
  async function openEditor(row) {
    if (editorId && dirtyEditor() && !await ask('放弃当前接口表单中尚未保留的改动并清空新密钥？')) return;
    editorId = row.id; edit.hidden = false;
    $('connection-name').value = row.name; protocol.value = row.protocol; $('connection-url').value = row.base_url;
    context.value = row.context_window; enabled.checked = row.enabled;
    models.value = row.models.map(model => `${model.id} | ${model.capabilities.join(',')}${model.output_parameter ? ' | ' + model.output_parameter : ''}`).join('\n');
    key.value = ''; keyAction.value = 'keep'; key.disabled = true;
    providerChoice.value=row.protocol==='ollama'?'ollama':Object.entries(snapshot.provider_defaults||{}).find(([id,item])=>id!=='ollama'&&item.base_url===row.base_url)?.[0]||'custom';
    purpose.value=mode==='agent'?'chat,agent':mode==='image'?'image':'chat';modelName.value=row.models[0]?.id||'';candidates.replaceChildren();modelProbeSequence++;
    const saved = snapshot.catalog.connections.find(r => r.id === row.id);
    $('connection-key-state').textContent = saved?.api_key_configured ? '旧密钥已配置；不返回原文。' : '未配置密钥；Ollama 通常无需密钥。';
    editorBaseline = editorValue(); controls();
    probeDraft(false);
  }
  async function probeDraft(refresh=false){
    if(modelProbePending||!editorId||!live()||pending)return;
    const previous=draft?.connections.find(row=>row.id===editorId),saved=snapshot?.catalog.connections.find(row=>row.id===editorId);
    const action=keyAction.value==='keep'&&previous?.key_action==='replace'?'replace':keyAction.value;
    const target={id:editorId,protocol:protocol.value,base_url:$('connection-url').value.trim(),key_action:action};
    if(action==='replace')target.api_key=key.value||previous?.api_key||'';
    if(!target.base_url||target.protocol!=='ollama'&&!(target.api_key||action==='keep'&&saved?.api_key_configured))return;
    const own=++modelProbeSequence,owner=epoch,identity=JSON.stringify([editorId,protocol.value,$('connection-url').value,keyAction.value,key.value]);
    modelProbePending=true;probeState.textContent='正在读取模型列表…';
    try{
      const result=await api().test_model_connection_draft(target,refresh);
      if(owner!==epoch||own!==modelProbeSequence||identity!==JSON.stringify([editorId,protocol.value,$('connection-url').value,keyAction.value,key.value]))return;
      if(!result?.ok){probeState.textContent=result?.field==='api_key'?'当前地址需要匹配的 API 密钥；模型名称可手填。':'未能读取列表，手动模型保留。';return;}
      candidates.replaceChildren();
      for(const id of result.result.models||[]){const option=make('option','',id,candidates);option.value=id;}
      probeState.textContent=result.result.status==='reachable'?(result.result.models?.length?'模型列表已读取；选中即可填入草稿。':'接口可达，未返回候选；可手动填写。'):(probeLabels[result.result.status]||'模型列表不可用，保留手动名称。');
    }catch(_){if(owner===epoch&&own===modelProbeSequence)probeState.textContent='列表暂不可用，可手动填写模型名称。';}
    finally{delete target.api_key;modelProbePending=false;}
  }
  const dirtyEditor = () => editorValue() !== editorBaseline || key.value || keyAction.value !== 'keep';
  function applyEditor() {
    if (!editorId || pending || !draft) return false;
    const parsed = models.value.split('\n').filter(line => line.trim()).map(line => {
      const parts = line.split('|'); return {id: parts[0].trim(), capabilities: parts.length >= 2 && parts.length <= 3 ? parts[1].split(',').map(s => s.trim()) : [], ...(parts.length === 3 ? {output_parameter:parts[2].trim()} : {})};
    });
    if (!$('connection-name').value.trim() || !parsed.length || parsed.some(m => !/^[A-Za-z0-9][A-Za-z0-9._:/@+-]{0,199}$/.test(m.id) || !m.capabilities.length || m.capabilities.some(c => !labels[c])) || keyAction.value === 'replace' && !key.value) {
      message.textContent = '请填写连接名称与有效模型ID，每行明确用途；替换密钥时必须输入新密钥。'; return false;
    }
    const index = draft.connections.findIndex(r => r.id === editorId);
    const previous = draft.connections[index];
    const row = {id: editorId, name: $('connection-name').value.trim(), protocol: protocol.value, base_url: $('connection-url').value.trim(),
      context_window: Number(context.value), enabled: enabled.checked, models: parsed, key_action: keyAction.value === 'keep' && previous?.key_action === 'replace' ? 'replace' : keyAction.value};
    if (keyAction.value === 'replace') row.api_key = key.value;
    else if (row.key_action === 'replace') row.api_key = previous.api_key;
    if (index < 0) draft.connections.push(row); else draft.connections[index] = row;
    Object.keys(draft.defaults).forEach(cap => {
      const selected = draft.defaults[cap];
      if (selected.connection_id === row.id && (!row.enabled || !row.models.some(m => m.id === selected.model_id && m.capabilities.includes(cap)))) delete draft.defaults[cap];
    });
    key.value = ''; keyAction.value = 'keep'; editorBaseline = editorValue();
    message.textContent = '已保留到本次内存草稿，尚未保存；主界面不会使用草稿。'; renderConnections(); return true;
  }
  function addConnection() {
    if (!draft || pending) return;
    const id = Array.from(crypto.getRandomValues(new Uint8Array(16)), n => n.toString(16).padStart(2, '0')).join('');
    openEditor({id: 'c-' + id, name: '', protocol: 'openai', base_url: '', context_window: 8192, enabled: true, models: []});
  }
  function renderConnections() {
    list.replaceChildren();
    if (!draft?.connections.length) make('p', '', '暂无已配置连接。新增接口后手动填写模型，保存后才能使用。', list);
    draft?.connections.forEach(row => {
      const article = make('article', '', '', list); article.className = 'connection-row';
      make('strong', '', `${row.name} · ${row.enabled ? '启用' : '禁用'}`, article);
      make('p', '', row.models.map(m => `${m.id}（${m.capabilities.map(c => labels[c]).join('、')}）`).join('；'), article);
      button('', '编辑', article, () => openEditor(row));
      button('', row.enabled ? '禁用到草稿' : '启用到草稿', article, () => {
        row.enabled = !row.enabled;
        if (!row.enabled) Object.keys(draft.defaults).forEach(cap => { if (draft.defaults[cap].connection_id === row.id) delete draft.defaults[cap]; });
        renderConnections();
      });
      button('', '删除到草稿', article, async () => {
        if (!await ask('从草稿移除此连接？保存确认后才从配置移除，当前会话选择会失效。')) return;
        draft.connections = draft.connections.filter(r => r.id !== row.id);
        Object.keys(draft.defaults).forEach(cap => { if (draft.defaults[cap].connection_id === row.id) delete draft.defaults[cap]; });
        if (editorId === row.id) { editorId = null; key.value = ''; edit.hidden = true; }
        renderConnections();
      });
      const saved = snapshot?.catalog.connections.find(r => r.id === row.id);
      const detect = button('', '获取已保存接口模型列表', article, () => probe(saved));
      detect.dataset.requiresSaved = saved ? 'yes' : 'no';
    });
    controls();
    list.querySelectorAll('[data-requires-saved="no"]').forEach(node => { node.disabled = true; node.title = '请先保存接口'; });
  }
  async function discard() {
    if (!snapshot || pending) return;
    if (dirty() && !await ask('放弃全部接口草稿并清空新密钥？不会写入磁盘。')) return;
    cleanDraft(snapshot); controls(); message.textContent = '已恢复最近读取的配置，没有写入。';
  }
  async function reloadDraft() {
    if (dirty() && !await ask('重新读取将丢弃接口草稿及新密钥，继续？')) return;
    key.value = ''; draft = null; await refresh(true);
  }
  async function savePayload(payload, expected) {
    if (pending || uncertain || !live()) return;
    pending = true; const mine = epoch, entry = document.activeElement, chosen=chosenCandidate?copy(chosenCandidate):null; let writeRequested = false; controls();
    try {
      const planned = await api().plan_model_connections(payload, expected);
      payload.connections.forEach(row => { delete row.api_key; row.key_action = 'keep'; });
      key.value = '';
      if (mine !== epoch || !live()) return;
      if (!planned?.ok) { error(planned); return; }
      const plan = planned.plan;
      changes.textContent = JSON.stringify(plan.differences, null, 2); changes.hidden = false;
      const approved = await reviewConfiguration(plan);
      if (mine !== epoch || !live()) return;
      writeRequested = approved;
      const result = await api().confirm_settings_update(plan.plan_id, approved);
      if (mine !== epoch || !live()) return;
      if (!result?.ok || result.result?.status !== 'saved') { error(result); return; }
      saveId = result.result.save_id; draft = null; await refresh(true);
      if (mine !== epoch || !live()) return;
      chosenCandidate=null;
      const chosenRow=chosen&&snapshot.catalog.connections.find(row=>row.id===chosen.connection_id&&row.enabled&&row.models.some(model=>model.id===chosen.model_id&&model.capabilities.includes(chosen.capability)));
      if(chosenRow)await selectModel(chosenRow,chosen.model_id,chosen.capability);
      message.textContent = '接口已保存。Web主界面后续请求可使用新配置；既有任务保持原配置。其他启动入口及检索配置仍按原规则，可能需重启。';
      showFeedback();
      window.dispatchEvent(new Event('officemodels:saved'));
      return true;
    } catch (_) {
      if (mine === epoch) {
        if (writeRequested) error({outcome_uncertain:true});
        else message.textContent = '配置预览或取消未完成；没有提交保存，不自动重试。请重新读取核对。';
      }
    } finally {
      if (mine === epoch) { pending = false; controls(); if (document.activeElement === document.body && entry?.isConnected && !entry.disabled) entry.focus(); }
    }
  }
  async function save() {
    await probeDraft(false);
    if (editorId && dirtyEditor() && !applyEditor()) return;
    if (!draft || !dirty()) return;
    return savePayload(draft, revision);
  }
  function showFeedback() {
    feedback.replaceChildren();
    snapshot.catalog.connections.filter(r => r.enabled).forEach(row => {
      const article = make('article', '', '', feedback);
      make('h4', '', row.name, article);
      make('p', '', '已保存 · ' + (probeLabels[row.model_list_status] || '模型列表尚未获取；不自动联网。'), article);
      button('', '获取模型列表', article, () => probe(row));
      row.models.forEach(model => model.capabilities.forEach(cap => button('', cap === 'chat' ? `用于当前对话：${model.id}` : `用于${labels[cap]}：${model.id}`, article, async () => {
        if (await selectModel(row, model.id, cap)) window.OfficeSettings?.close();
      })));
    });
  }
  async function probe(row) {
    if (pending || !row || !live()) return;
    pending = true; const mine = epoch; controls(); message.textContent = '正在获取已保存接口的模型列表；不生成内容、不自动重试…';
    try {
      const result = await api().test_model_connection(row.id, row.revision);
      if (mine !== epoch || !live()) return;
      if (!result?.ok) { error(result); return; }
      const outcome = result.result;
      message.textContent = probeLabels[outcome.status] || '模型列表未验证。';
      await refresh();
      if (mine !== epoch || !live()) return;
      showFeedback();
      if (outcome.models?.length) {
        const box = make('div', '', '', feedback); make('p', '', '检测返回的候选（采用到草稿并保存后才可选；用途需人工声明）：', box);
        outcome.models.forEach(id => button('', '添加候选到草稿：' + id, box, () => {
          const target = draft.connections.find(r => r.id === row.id);
          if (!target || target.models.some(m => m.id === id)) return;
          if (target.models.length >= 64) { message.textContent = '每个连接最多64个模型，请先整理条目。'; return; }
          target.models.push({id, capabilities: ['chat']}); renderConnections();
          message.textContent = '候选已加入普通问答草稿；可编辑用途，尚未保存或验证实际调用。';
        }));
      }
    } catch (_) { if (mine === epoch) message.textContent = '检测结果未确认；没有自动重试。'; }
    finally { if (mine === epoch) { pending = false; controls(); } }
  }
  async function selectModel(row, modelId, capability) {
    if (selecting || !live()) return false;
    selecting = true; selectionBlocked = true; sequence++; const mine = epoch; controls();
    try {
      const result = await api().select_session_model({connection_id: row.id, model_id: modelId, capability, revision: row.revision});
      if (mine !== epoch || !live()) return false;
      if (!result?.ok) { selectionFailure(errors[result?.reason] || '选择失败。'); return false; }
      mode = capability; await refresh(false, true);
      if (selectionBlocked) return false;
      return true;
    } catch (_) { if (mine === epoch) selectionFailure('宿主响应未确认。'); return false; }
    finally { if (mine === epoch) { selecting = false; controls(); } }
  }
  function selectionFailure(text) {
    selectionBlocked = true;
    modeInput.value = mode;
    const current = snapshot?.catalog.current[mode];
    selector.value = current?.valid ? value(current) : '';
    status.textContent = '切换未确认：' + text + ' 已恢复上次已知显示，请重新选择或刷新核对后再发送；输入已保留。';
  }
  async function setDefault() {
    const current = snapshot?.catalog.current[mode];
    if (!current?.valid || pending) return;
    if (dirty()) { configure(); message.textContent = '请先保存或放弃接口草稿，再单独设置默认模型。'; return; }
    if (!draft) cleanDraft(snapshot);
    const payload = copy(draft);
    payload.defaults[mode] = {connection_id: current.connection_id, model_id: current.model_id, capability: mode};
    const saved = await savePayload(payload, revision);
    if (saved) status.textContent += ' 本用途默认模型已保存；重启后按此默认恢复。';
    else if (live()) status.textContent = message.textContent;
  }
  async function undo() {
    if (!saveId || pending || uncertain) return;
    if (dirty() && !await ask('撤销将丢弃接口草稿和新密钥，继续？')) return;
    key.value = ''; const mine = epoch; pending = true; controls();
    try {
      const result = await api().plan_settings_rollback(saveId);
      if (mine !== epoch || !live()) return;
      if (!result?.ok) { error(result); return; }
      const approved = await ask('确认恢复本次保存前的完整配置备份？现有选择可能失效；不会调用模型。') === true;
      const receipt = await api().confirm_settings_rollback(result.plan.rollback_id, approved);
      if (mine !== epoch || !live()) return;
      if (!receipt?.ok || receipt.result?.status !== 'rolled_back') { error(receipt); return; }
      saveId = null; draft = null; await refresh(true); message.textContent = '本次保存已撤销，请重新核对模型选择。';
    } catch (_) { if (mine === epoch) error({outcome_uncertain: true}); }
    finally { if (mine === epoch) { pending = false; controls(); } }
  }
  selector.onchange = () => {
    if (!selector.value) return;
    const [cid, mid, cap] = JSON.parse(selector.value);
    const row = snapshot.catalog.connections.find(r => r.id === cid);
    if (row) selectModel(row, mid, cap);
  };
  modeInput.onchange = async () => {
    if (selecting || !live()) return;
    const mine = epoch; selecting = true; selectionBlocked = true; sequence++; controls();
    try {
      const result = await api().set_model_mode(modeInput.value);
      if (mine !== epoch || !live()) return;
      if (result?.ok) await refresh(false, true);
      else selectionFailure(errors[result?.reason] || '用途切换失败。');
    } catch (_) { if (mine === epoch) selectionFailure('用途切换响应未确认。'); }
    finally { if (mine === epoch) { selecting = false; controls(); } }
  };
  search.oninput = renderSelector;
  edit.addEventListener('input', controls);
  const clear = () => {
    closeReview?.(false);
    epoch++; sequence++; snapshot = draft = null; baseline = revision = ''; editorId = saveId = null;
    pending = selecting = uncertain = selectionBlocked = false; key.value = ''; keyAction.value = 'keep'; models.value = '';
    edit.reset(); edit.hidden = true; list.replaceChildren(); feedback.replaceChildren(); changes.replaceChildren();
    changes.hidden = true; mode = 'agent'; modeInput.value = mode; renderSelector();
  };
  document.addEventListener('officesettings:opened', () => { refresh(); });
  document.addEventListener('officesettings:closed', () => {
    key.value = ''; draft = null; editorId = null; edit.hidden = true; changes.textContent = ''; changes.hidden = true;
  });
  window.addEventListener('officeauth:unauthenticated', clear);
  window.addEventListener('officeauth:authenticated', () => refresh(true));
  window.addEventListener('officesettings:saved', () => refresh());
  window.addEventListener('officemodels:request-finished', () => refresh());
  window.addEventListener('officemodels:host-ready', () => refresh());
  window.OfficeModelSelector = {
    documentRecipient: () => {
      const selected=snapshot?.catalog.current.agent;
      if(mode!=='agent'||!selected?.valid||selectionBlocked||selecting||uncertain)return false;
      return {connection_id:selected.connection_id,model_id:selected.model_id,revision:selected.revision};
    },
    mode: () => mode,
    selection: () => {
      const current = snapshot?.catalog.current[mode];
      const item = rowsForMode().find(({row, model}) => row.id === current?.connection_id && model.id === current?.model_id);
      return current?.valid && !selectionBlocked ? {valid:true, model:current.model_id, connection:current.connection_name, state:states[item?.model.states[mode]] || '未验证'} : {valid:false,reason:rowsForMode().length?'当前模型不可用，请检查连接':'尚未配置模型'};
    },
    refresh,
    readyForRequest: (notify = true) => {
      if (!live() || pending || uncertain || selecting || selectionBlocked || !snapshot?.catalog.current[mode]?.valid) {
        if (notify) status.textContent = '请先选择有效的已保存模型并核对配置状态；输入已保留。';
        return false;
      }
      return true;
    },
    requestClose: async () => {
      if (pending) { message.textContent = '操作尚未结束，请先等待结果；没有自动重试。'; return false; }
      if (dirty() && !await ask('关闭将放弃未保存的接口草稿并清空新密钥，是否继续？')) return false;
      if (draft) draft.connections.forEach(row => { delete row.api_key; row.key_action = 'keep'; });
      key.value = ''; return true;
    }
  };
  controls();
})();
