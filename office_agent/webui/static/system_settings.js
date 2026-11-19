(() => {
  'use strict';
  const parent = document.querySelector('#settings-pane-machine .settings-form');
  if (!parent) return;
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const make = (tag, text, host) => { const node = document.createElement(tag); if (text) node.textContent = text; if(tag==='button'){node.type='button';node.className='btn secondary';}host?.append(node); return node; };
  const section = make('section', '', parent); section.className = 'settings-group';
  make('h4', '网络', section);
  const label = make('label', '网络方式', section), mode = make('select', '', label); mode.id = 'network-mode';
  for (const [value, title] of [['direct', '直连'], ['auto', '自动（显式系统／环境代理）'], ['manual', '手动代理']]) { const option = make('option', title, mode); option.value = value; }
  const proxyLabel = make('label', '代理地址', section), proxy = make('input', '', proxyLabel); proxy.id = 'network-proxy'; proxy.type = 'password'; proxy.autocomplete = 'off'; proxy.placeholder = 'http://127.0.0.1:7890；留空保持已存代理';
  const proxyRule = make('p', '使用显式代理，跳过 PAC。本机模型和 NO_PROXY 地址直连。', section);
  const renderMode = () => {
    proxyLabel.hidden = mode.value !== 'manual';
    proxyRule.hidden = mode.value !== 'auto';
  };
  mode.addEventListener('change', renderMode);
  renderMode();
  const status = make('p', '', section); status.setAttribute('role', 'status');
  const actions = make('div', '', section); actions.className = 'actions';
  const save = make('button', '保存网络设置', actions); save.type = 'button';
  const extra = make('details', '', section); extra.className = 'settings-explanation';
  make('summary', '诊断与缓存', extra);
  const utilities = make('div', '', extra); utilities.className = 'actions';
  const diagnostics = make('button', '查看诊断', utilities); diagnostics.type = 'button';
  const repair = make('button', '检查下载缓存', utilities); repair.type = 'button';
  let revision = '', pending = false;
  const load = async () => {
    if (pending) return;
    const result = await api().get_settings();
    if (!result?.ok) return;
    const snapshot = result.settings; revision = snapshot.revision || '';
    const value = snapshot.values || snapshot;
    mode.value = value.network_mode || 'direct'; proxy.value = '';
    proxy.placeholder = value.network_proxy_configured ? '已保存代理；留空保持' : 'http://127.0.0.1:7890';
    renderMode();
  };
  save.onclick = async () => {
    if (pending) return;
    pending = true; save.disabled = true;
    try {
      const changes = {network_mode: mode.value}; if (mode.value === 'manual' && proxy.value) changes.network_proxy = proxy.value;
      const result = await api().plan_settings_edit(changes, revision);
      proxy.value = '';
      if (!result?.ok) { status.textContent = '设置未保存，请重新读取或检查代理地址。'; return; }
      const yes = await window.OfficeConfirm.ask('保存网络方式？代理凭据保存在本机配置中，不写入诊断日志。');
      const receipt = await api().confirm_settings_update(result.plan.plan_id, yes);
      status.textContent = receipt?.result?.status === 'saved' ? '已保存。新请求使用新配置。' : '未保存。';
    } finally { pending = false; save.disabled = false; await load(); }
  };
  diagnostics.onclick = async () => {
    if (pending) return;
    const result = await api().floating_diagnostics();
    if (!result?.ok) { status.textContent = '诊断暂不可用。'; return; }
    const previous = document.activeElement, dialog = make('dialog', '', document.body); dialog.className = 'paged-dialog'; dialog.setAttribute('aria-label', '诊断摘要');
    make('h2', '诊断摘要', dialog); make('p', '仅运维状态，不包含配置、会话或文档正文。', dialog);
    const content = make('pre', JSON.stringify(result.diagnostics, null, 2), dialog); content.style.maxHeight = '50vh'; content.style.overflow = 'auto';
    const buttons = make('div', '', dialog); buttons.className = 'paged-dialog-actions';
    const close = make('button', '关闭', buttons); close.onclick = () => dialog.close();
    const exportButton = make('button', '导出摘要…', buttons); exportButton.onclick = async () => {
      exportButton.disabled = true;
      const saved = await api().floating_export_diagnostics(result.preview_id);
      status.textContent = saved?.saved ? '诊断摘要已导出。' : saved?.cancelled ? '已取消导出。' : '未能导出摘要，原文件未改动。';
      if (saved?.saved) dialog.close(); else exportButton.disabled = false;
    };
    dialog.addEventListener('close', () => { dialog.remove(); previous?.focus(); }); dialog.showModal(); close.focus();
  };
  document.addEventListener('officesettings:opened', load);
  repair.onclick=async()=>{
    if(pending)return;pending=true;repair.disabled=true;
    try{
      for(const kind of ['frontend','renderer']){
        const result=await api().floating_download_cache(kind);
        if(!result?.ok){status.textContent='缓存检查未完成，已保留原文件。';return;}
        const plan=result.plan;
        if(!plan.repairable)continue;
        const yes=await window.OfficeConfirm.ask((kind==='frontend'?'编辑组件':'渲染组件')+'缓存未验证。将此组件缓存移入同目录隔离文件，保留原字节；不清理其他文件。继续？');
        const saved=await api().floating_quarantine_download(plan.id,yes);
        if(!saved?.ok){status.textContent='缓存身份已变化，未继续处理。';return;}
        status.textContent=yes?'已隔离。请在安装器选择重新准备；不会自动重放安装。':'未更改缓存。';
        if(!yes)return;
      }
      if(!status.textContent.includes('已隔离'))status.textContent='缓存检查完成；无需隔离。';
    }finally{pending=false;repair.disabled=false;}
  };
  document.addEventListener('officesettings:closed', () => { proxy.value = ''; });
})();
