/* Human-only local configuration; never put keys in output, storage, history or logs. */
(() => {
  'use strict';
  if (window.OfficeSettingsEditor?.mounted) return;
  const node = id => document.getElementById(id);
  let epoch = 0, inflight = false, revision = 0, loaded = null, saveId = null;
  const embeddingNumbers = [
    ['settings-embedding-dimension', 'embedding_dimension', 1, 4096, true],
    ['settings-embedding-timeout', 'embedding_timeout', 1, 20, false],
    ['settings-embedding-budget', 'embedding_budget', 1, 60, false]
  ];
  const ids = ['settings-provider', 'settings-model', 'settings-url', 'settings-key',
    'settings-clear-key', 'settings-context', 'settings-embedding-channel', 'settings-embedding-model',
    ...embeddingNumbers.map(field => field[0])];
  const live = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const bridge = () => live() ? window.OfficeAuth.wrapApi(window.pywebview.api) : null;
  const opaque = id => typeof id === 'string' && /^[a-f0-9]{32}$/.test(id);
  function status(text) { node('settings-status').textContent = text; }
  function retireRag(uncertain = false) {
    if (typeof window.dispatchEvent !== 'function') return;
    const event = new Event('officerag:invalidate');
    event.ragUncertain = uncertain; window.dispatchEvent(event);
  }
  function refresh() {
    node('settings-load').disabled = inflight || !live();
    node('settings-save').disabled = inflight || !live() || !loaded;
    node('settings-rollback').disabled = inflight || !live() || !saveId;
    ids.forEach(id => { node(id).disabled = inflight || !live() || !loaded; });
    node('settings-key').disabled ||= node('settings-provider').value === 'ollama';
    node('settings-clear-key').disabled ||= node('settings-provider').value === 'ollama';
  }
  function retire() {
    epoch++; revision++; loaded = null; saveId = null;
    ids.forEach(id => { node(id).value = ''; });
    node('settings-clear-key').checked = false;
    node('settings-key-state').textContent = '';
    status('尚未读取配置。'); refresh();
  }
  function show(value) {
    if (!value || !['deepseek', 'dashscope', 'ollama'].includes(value.provider)) throw Error();
    loaded = value;
    node('settings-provider').value = value.provider;
    node('settings-model').value = value.model || '';
    node('settings-url').value = value.base_url || '';
    node('settings-context').value = value.context_window ?? '';
    node('settings-embedding-channel').value = value.embedding_channel || 'local';
    node('settings-embedding-model').value = value.embedding_model || '';
    embeddingNumbers.forEach(([id, key]) => { node(id).value = String(value[key] ?? ''); });
    node('settings-key').value = '';
    node('settings-clear-key').checked = false;
    node('settings-key-state').textContent = value.api_key_configured ? '当前提供方已配置 Key（不回显）。' : '当前提供方未配置 Key。';
  }
  async function run(operation) {
    if (inflight || !live()) return;
    inflight = true; refresh();
    const mine = epoch;
    try { await operation(bridge(), () => mine === epoch && live()); }
    catch (_) {
      if (mine === epoch && live()) { saveId = null; retireRag(true); status('配置操作没法确认结果；请核对本机配置和记录，不要自动重试。'); }
    } finally { inflight = false; node('settings-key').value = ''; refresh(); }
  }
  function loadSettings() {
    return run(async (api, current) => {
      status('正在读取本机配置…');
      const result = await api.get_settings();
      if (!current()) return;
      if (!result?.ok) { status('配置读取被拒绝；请核对登录、运行状态与本地配置。'); return; }
      show(result.settings);
      status(result.settings.restart_required ? '配置已改变，请重启应用后使用。' : '已读取；密钥不会显示出来，保存前会再让你确认一次。');
    });
  }
  node('settings-load').addEventListener('click', loadSettings);
  ids.forEach(id => node(id).addEventListener('input', () => {
    revision++;
    if (id === 'settings-provider') {
      node('settings-model').value = ''; node('settings-url').value = '';
      node('settings-key').value = ''; node('settings-clear-key').checked = false;
      node('settings-key-state').textContent = '已切换提供方；填写其配置，空 Key 保留该提供方原值。';
    }
    refresh();
  }));
  node('settings-save').addEventListener('click', () => run(async (api, current) => {
    if (!loaded) return;
    const version = revision;
    const changes = {provider: node('settings-provider').value};
    for (const [id, key] of [['settings-model','model'],['settings-url','base_url'],['settings-embedding-model','embedding_model']]) {
      const value = node(id).value.trim(); if (value) changes[key] = value;
    }
    if (node('settings-context').value) changes.context_window = Number(node('settings-context').value);
    changes.embedding_channel = node('settings-embedding-channel').value;
    for (const [id, key, min, max, integral] of embeddingNumbers) {
      const raw = node(id).value;
      if (raw === '') continue;
      const syntax = integral ? /^(?:0|[1-9][0-9]*)$/ : /^(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$/;
      const value = Number(raw);
      if (typeof raw !== 'string' || raw.length > 32 || !syntax.test(raw) || !Number.isFinite(value) ||
          value < min || value > max || (integral && !Number.isSafeInteger(value))) {
        status('检索参数不对：维度要填 1 到 4096 的整数，单次超时 1 到 20 秒，总时长上限 1 到 60 秒；不支持科学计数法或无穷大的数。'); return;
      }
      changes[key] = value;
    }
    if (changes.provider !== 'ollama') {
      if (node('settings-clear-key').checked) changes.api_key = '';
      else if (node('settings-key').value) changes.api_key = node('settings-key').value;
    }
    let response;
    try { response = await api.plan_settings_update(changes); }
    finally { delete changes.api_key; node('settings-key').value = ''; }
    if (!current()) return;
    if (!response?.ok || !opaque(response.plan?.plan_id)) { status('配置计划被拒绝，尚未批准写入。'); return; }
    const plan = response.plan;
    let confirmed = false;
    if (version === revision && current()) {
      const view = plan.settings || {};
      confirmed = window.confirm('确认保存配置？\n改动的项：' + plan.changed_fields.join('、') +
        '\n提供方：' + (view.provider || '') + '\n模型：' + (view.model || '') +
        '\n地址：' + (view.base_url || '') + '\n上下文窗口：' + (view.context_window ?? '') +
        '\n检索通道：' + (view.embedding_channel || '') + ' / ' + (view.embedding_model || '') +
        '\n维度：' + (view.embedding_dimension ?? '') + '；单次超时：' + (view.embedding_timeout ?? '') + '秒；总时长上限：' + (view.embedding_budget ?? '') + '秒' +
        '\n密钥：' + (view.api_key_configured ? '已配置（不显示）' : '未配置') +
        '\n本机的 .env 会被替换；备份里含旧密钥，记录里不含。所有账户共用这份配置。保存后需重启；不会自动调用外部服务。') === true;
    }
    if (!current()) return;
    if (version !== revision) confirmed = false;
    const result = await api.confirm_settings_update(plan.plan_id, confirmed);
    if (!current()) return;
    const outcome = result?.result;
    if (result?.ok && outcome?.status === 'saved' && outcome.applied === true && opaque(outcome.save_id)) {
      saveId = outcome.save_id; loaded = null;
      retireRag();
      status('配置已保存并通过校验；请重启应用。如需还原，可以撤销这次改动。');
    } else if (result?.ok && outcome?.status === 'rejected' && outcome.wrote_nothing === true) {
      status('本次批准未执行；需要新计划才能再次保存。');
    } else { saveId = null; loaded = null; retireRag(true); status('配置结果没能确认；请核对本机配置、备份和记录，不要自动重试。'); }
  }));
  node('settings-rollback').addEventListener('click', () => run(async (api, current) => {
    const target = saveId;
    if (!opaque(target)) return;
    const response = await api.plan_settings_rollback(target);
    if (!current()) return;
    if (!response?.ok || !opaque(response.plan?.rollback_id)) { saveId = null; status('撤销没能发起；配置可能已经变了，没有执行还原。'); return; }
    const note = response.plan.restores_absent_source_as_empty_env ? '\n原先没有 .env，将恢复为空 .env 文件。' : '';
    const confirmed = window.confirm('确认撤销这次配置改动？\n会从备份还原配置，重启后生效；备份和记录都会保留。' + note) === true;
    if (!current() || target !== saveId) return;
    const result = await api.confirm_settings_rollback(response.plan.rollback_id, confirmed);
    if (!current()) return;
    if (result?.ok && result.result?.status === 'rolled_back' && result.result.restored === true) {
      saveId = null; loaded = null; status('已恢复配置并校验；请重启应用。');
      retireRag();
    } else if (result?.ok && result.result?.status === 'rejected' && result.result.wrote_nothing === true) {
      status('没有执行撤销；需要重新发起并确认一次。');
    } else { saveId = null; loaded = null; retireRag(true); status('撤销结果没能确认；请核对配置和记录。'); }
  }));
  window.addEventListener('officeauth:unauthenticated', retire);
  window.addEventListener('officeauth:authenticated', retire);
  // A restored session may not emit a login event; resolve it before loading local settings.
  document.addEventListener('officesettings:opened', () => {
    refresh();
    if (!window.OfficeAuth?.ensure) return;
    const mine = epoch;
    window.OfficeAuth.ensure().then(authenticated => {
      if (mine !== epoch || !window.OfficeSettings?.isOpen?.()) return;
      refresh();
      if (authenticated && live() && !loaded && !inflight) loadSettings();
    }, () => {
      if (mine === epoch && window.OfficeSettings?.isOpen?.()) status('无法确认登录状态；请重新登录后读取配置。');
    });
  });
  refresh();
})();
