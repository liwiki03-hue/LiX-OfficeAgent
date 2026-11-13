(() => {
  'use strict';
  const $ = id => document.getElementById('updates-' + id);
  if (!$('version')) return;
  let epoch = 0, busy = false, configured = false, available = false, plan = null, loaded = false, automatic = false, queued = false, loading = null;
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const errors = {not_authenticated: '请先登录。', busy: '当前有任务或更新操作正在进行，请稍后重试。', update_not_configured: '发布仓库和签名公钥尚未配置，当前版本可正常使用。', update_timeout: '检查或下载超时，未自动重试。', update_network_failed: '无法连接更新服务，请检查网络后主动重试。', update_release_missing: '仓库尚无公开稳定版本。', update_asset_missing: '该版本没有适用于本软件的更新包，请联系发布者。', update_rate_limited: 'GitHub 请求受到限流，请稍后重试。', update_signature_invalid: '签名校验失败，更新包已拒绝。', update_package_invalid: '更新包摘要或内容校验失败，未安装。', update_dependencies_changed: '此版本需要更新运行依赖，当前自动更新通道无法安装，请使用发布者完整发行包。', update_local_changes: '本地程序文件或依赖已有修改，为保留改动已拒绝覆盖。', update_plan_expired: '下载或安装计划已过期，请重新检查并下载。', update_install_locked: '存在未完成的安装记录，请先核对安装日志，不自动覆盖。', update_platform_unsupported: '当前窗口无法执行 Windows 重启安装。', update_session_changed: '登录状态已改变，迟到结果已丢弃。', update_url_refused: '更新下载地址不符合发布源限制，已拒绝。', update_capacity: '更新内容超过容量上限，已拒绝。', update_failed: '更新操作未完成，请核对本机安装记录后重试。'};
  const say = text => { $('status').textContent = text; };
  errors.update_handoff_failed = '安装组件未准备就绪，软件保持打开。请重新检查并下载后再试。';
  function refresh() {
    const blocked = busy || queued || !authenticated();
    $('check').disabled = blocked || !configured;
    $('download').disabled = blocked || !available || Boolean(plan);
    $('install').disabled = blocked || !plan;
    $('auto-check').disabled = blocked || !configured;
    $('auto-download').disabled = blocked || !configured || !$('auto-check').checked;
    $('status').setAttribute('aria-busy', String(busy));
  }
  function readPreference(id) { try { return localStorage.getItem('office-agent-' + id) === 'true'; } catch (_) { return false; } }
  function writePreference(id, value) { try { localStorage.setItem('office-agent-' + id, String(value)); } catch (_) { say('浏览器未允许保存偏好，本次选项只在当前窗口有效。'); } }
  async function run(action) {
    if (busy || queued || !authenticated()) return false;
    const current = epoch; busy = true; refresh();
    try { await action(window.OfficeAuth.wrapApi(window.pywebview.api), () => current === epoch && authenticated()); return true; }
    catch (_) { if (current === epoch) say('宿主请求中断，未自动重试。安装结果请以本机记录为准。'); return false; }
    finally { if (current === epoch) { busy = false; refresh(); } }
  }
  function failure(reply) { say(errors[reply?.reason] || '更新内容或来源无效，已停止。'); }
  function load() {
    if (loading) return loading;
    const pending = run(async (api, current) => {
      const reply = await api.get_release_info(); if (!current()) return; if (!reply?.ok) { failure(reply); return; }
      const info = reply.result; configured = info.configured; loaded = true;
      $('version').textContent = 'v' + info.version;
      if (document.getElementById('app-version')) document.getElementById('app-version').textContent = 'v' + info.version;
      $('source').textContent = configured ? '稳定版发布源：' + info.repository : '发布源尚未配置；创建 GitHub 仓库后可由发布者接入签名更新。';
      $('changelog').replaceChildren();
      for (const release of info.history) {
        const article = document.createElement('article'), title = document.createElement('h4'), list = document.createElement('ul');
        title.textContent = 'v' + release.version + ' · ' + release.date + ' · ' + release.title;
        for (const entry of release.items) { const item = document.createElement('li'); item.textContent = entry; list.append(item); }
        article.append(title, list); $('changelog').append(article);
      }
      const statuses = {queued: '已安排，等待软件退出', refused: '安装前被拒绝，程序未替换', installed: '安装成功', rolled_back: '安装失败，已恢复原版', uncertain: '结果不确定，请核对备份和审计', prepared: '已备份，尚未完成安装', installing: '安装被中断，需核对备份和审计'};
      $('install-history').replaceChildren();
      for (const item of info.install_history) { const row = document.createElement('p'); row.textContent = (statuses[item.status] || '未知状态，需核对') + '：' + item.from + ' → ' + item.to; $('install-history').append(row); }
      if (!info.install_history.length) $('install-history').textContent = '暂无安装记录。';
      if (!configured) say(errors.update_not_configured);
    });
    loading = pending;
    pending.finally(() => { if (loading === pending) loading = null; });
    return pending;
  }
  async function check() {
    available = false; plan = null;
    const done = await run(async (api, current) => {
      say('正在检查稳定版本…'); const reply = await api.check_app_updates(); if (!current()) return; if (!reply?.ok) { failure(reply); return; }
      available = reply.result.status === 'available'; $('release').hidden = false;
      $('release-title').textContent = (available ? '发现新版本 v' : '当前已是最新稳定版 · v') + reply.result.version;
      $('release-notes').textContent = reply.result.notes || '发布者未提供说明。';
      say(available ? '可下载并验证更新包；尚未修改程序。' : '检查完成，无需更新。');
    });
    if (done && available && $('auto-download').checked) await download();
  }
  async function download() {
    plan = null;
    await run(async (api, current) => {
      say('正在下载并验证签名…'); const reply = await api.download_app_update(); if (!current()) return; if (!reply?.ok) { failure(reply); return; }
      plan = reply.result; say('v' + plan.version + ' 已通过签名与内容校验，涉及 ' + plan.changed_files + ' 个程序文件；5分钟内可确认安装。当前软件未被替换。');
    });
  }
  $('check').addEventListener('click', check); $('download').addEventListener('click', download);
  $('install').addEventListener('click', () => {
    if (!plan || busy) return;
    const approved = window.confirm('确认安装 v' + plan.version + ' 并关闭、重启软件？仅替换受管程序文件；配置、文档、账户和模型目录保留。本地代码有改动或依赖不同会拒绝；安装失败会尝试恢复原版。');
    run(async (api, current) => { const token = plan.plan_id; plan = null; const reply = await api.confirm_app_update(token, approved); if (!current()) return; if (!reply?.ok) { failure(reply); return; } queued = reply.result.status === 'queued'; say(queued ? '安装已安排，正在关闭软件；退出后安装并从原启动入口重启。' : '已取消安装，程序文件未改动。'); });
  });
  $('auto-check').checked = readPreference('updates-auto-check'); $('auto-download').checked = readPreference('updates-auto-download');
  $('auto-check').addEventListener('change', () => {
    if ($('auto-check').checked && !window.confirm('启用后，每次登录会向配置的 GitHub 仓库检查一次稳定版本，不发送文档和密钥。是否启用？')) $('auto-check').checked = false;
    if (!$('auto-check').checked) { $('auto-download').checked = false; writePreference('updates-auto-download', false); }
    writePreference('updates-auto-check', $('auto-check').checked); refresh();
  });
  $('auto-download').addEventListener('change', () => {
    if ($('auto-download').checked && !window.confirm('发现新版后自动下载并验证更新包，最多10MB；仍不会自动批准安装。是否启用？')) $('auto-download').checked = false;
    writePreference('updates-auto-download', $('auto-download').checked);
  });
  document.getElementById('tab-updates')?.addEventListener('click', () => { if (!loaded) load(); });
  const activate = async () => {
    await load(); if (configured && $('auto-check').checked && !automatic) { automatic = true; await check(); }
  };
  window.addEventListener('officeauth:authenticated', activate);
  const reset = () => { epoch++; busy = configured = available = loaded = automatic = queued = false; plan = loading = null; $('version').textContent = ''; $('source').textContent = ''; $('changelog').replaceChildren(); $('install-history').replaceChildren(); $('release').hidden = true; $('release-notes').textContent = ''; say('登录后可查看版本与更新日志。'); refresh(); };
  window.addEventListener('officeauth:unauthenticated', reset); window.addEventListener('pagehide', reset);
  refresh(); if (authenticated()) activate();
})();
