(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  if (!$('settings-editor-version')) return;
  const fields = {
    provider: 'settings-provider', model: 'settings-model', base_url: 'settings-url',
    context_window: 'settings-context', embedding_channel: 'settings-embedding-channel',
    embedding_model: 'settings-embedding-model', embedding_dimension: 'settings-embedding-dimension',
    embedding_timeout: 'settings-embedding-timeout', embedding_budget: 'settings-embedding-budget'
  };
  const labels = {provider: '模型提供方', model: '模型名称', base_url: 'API 地址', api_key: 'API 密钥',
    context_window: '上下文窗口', embedding_channel: '检索通道', embedding_model: '检索模型',
    embedding_dimension: '检索维度', embedding_timeout: '单次超时', embedding_budget: '总时长上限'};
  const reasons = {
    busy: '当前有任务或检测正在进行，请结束后再操作。', not_authenticated: '登录已失效，请重新登录。',
    settings_invalid: '配置校验失败，请按字段提示修正。', settings_source_changed: '磁盘配置已被外部修改，请重新读取后再编辑。',
    settings_plan_expired: '确认计划已过期，请重新查看改动。', settings_plan_invalid: '确认计划已失效，请重新查看改动。',
    settings_path_changed: '配置路径安全检查未通过，请检查目录或链接。', settings_file_invalid: '配置文件格式无效，请检查重复字段或非法值。',
    settings_io_failed: '配置文件读写失败，请检查本机权限和文件占用。', settings_verification_failed: '写入校验失败，请核对备份和记录。',
    settings_capacity: '本次会话的设置操作已达上限，请重启后重新读取。', settings_no_changes: '没有实际改动，无需保存。',
    settings_denied: '已取消，没有执行写入；新密钥输入已清空，再次替换需重新输入。', confirmation_required: '必须单独确认才能保存。',
    settings_session_invalid: '设置会话已失效，请重新读取。', settings_service_disabled: '设置服务已锁定，请核对记录并重启。'
  };
  let snapshot = null, baseline = null, drafts = {}, provider = '', saveId = null;
  let epoch = 0, revision = 0, pending = false, state = 'clean', uncertain = false;
  let providerChosen = false, guideExpanded = null, connection = null;
  const guideActions = ['provider', 'defaults', 'key', 'test', 'save', 'context', 'retrieval'];
  const live = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const opaque = value => typeof value === 'string' && /^[a-f0-9]{32}$/.test(value);
  const key = () => $('settings-key').value;
  const clearKey = () => $('settings-clear-key').checked;
  Object.values({...fields, api_key: 'settings-key'}).forEach(id => {
    const error = document.createElement('span');
    error.id = id + '-error'; error.className = 'settings-field-error'; error.setAttribute('aria-live', 'polite');
    $(id).insertAdjacentElement('afterend', error);
    $(id).setAttribute('aria-describedby', error.id);
  });
  function read() {
    return Object.fromEntries(Object.entries(fields).map(([name, id]) => [name, String($(id).value)]));
  }
  function dirty() {
    if (!snapshot) return false;
    if (!snapshot.disk_provider_explicit && providerChosen) return true;
    if (key() || (clearKey() && snapshot.profiles[provider]?.api_key_configured)) return true;
    const values = read();
    if (Object.keys(values).some(name => (name in snapshot.limits && values[name] === '' ? baseline[name] : values[name]) !== baseline[name])) return true;
    return Object.entries(drafts).some(([name, draft]) => name !== provider && (
      draft.model !== (snapshot.profiles[name]?.model || '') || draft.base_url !== (snapshot.profiles[name]?.base_url || '')));
  }
  function status(next, text) {
    state = next;
    $('settings-status').dataset.state = next;
    $('settings-status').textContent = text;
    refresh();
  }
  function refresh() {
    const blocked = pending || !live() || uncertain;
    $('settings-load').disabled = blocked;
    $('settings-save').disabled = blocked || !snapshot || !dirty();
    $('settings-discard').disabled = blocked || !snapshot || !dirty();
    $('settings-test').disabled = blocked || !snapshot;
    $('settings-rollback').disabled = blocked || !saveId;
    [...Object.values(fields), 'settings-key', 'settings-clear-key'].forEach(id => {
      $(id).disabled = blocked || !snapshot;
    });
    $('settings-key').disabled ||= provider === 'ollama' || clearKey();
    $('settings-clear-key').disabled ||= provider === 'ollama';
    const local = $('settings-embedding-channel').value === 'local';
    $('settings-embedding-api').hidden = local;
    $('settings-embedding-local').hidden = !local;
    document.querySelector('label[for="settings-embedding-timeout"]')?.toggleAttribute('hidden', local);
    $('settings-embedding-timeout').hidden = local;
    $('settings-embedding-model-label').textContent = local ? '离线检索模型目录（绝对路径）' : '百炼检索模型名称';
    refreshGuide(blocked);
  }
  function refreshGuide(blocked) {
    const visible = Boolean(snapshot) && live();
    $('settings-setup').hidden = !visible;
    if (!visible) {
      ['summary', 'provider-state', 'model-state', 'url-state', 'key-state', 'test-state', 'save-state', 'retrieval-state'].forEach(name => {
        $('settings-setup-' + name).textContent = '';
      });
      return;
    }
    const values = read(), saved = snapshot.profiles[provider];
    const hasKey = provider === 'ollama' || (!clearKey() && Boolean(key() || saved.api_key_configured));
    const incomplete = !providerChosen || !values.model || !values.base_url || !hasKey;
    if (guideExpanded === null) guideExpanded = incomplete;
    $('settings-setup-body').hidden = !guideExpanded;
    $('settings-setup-toggle').setAttribute('aria-expanded', String(guideExpanded));
    $('settings-setup-toggle').textContent = guideExpanded ? '收起引导' : '展开引导';
    $('settings-setup-summary').textContent = uncertain ? '操作结果不确定，请先核对配置和记录。' :
      state === 'invalid' ? '字段格式或范围需要修正，请查看对应字段提示。' :
      incomplete ? '当前配置仍有待明确的项目，可按步骤补齐草稿。' :
      dirty() ? '当前配置项已填写；仍有未保存改动，服务能力尚未确认。' :
      snapshot.restart_required ? '磁盘配置已保存，重启后生效；服务能力仍需另行核对。' : '磁盘配置项已填写；运行状态与模型能力仍需另行核对。';
    $('settings-setup-provider-state').textContent = '当前选择：' + $('settings-provider').selectedOptions[0].textContent + '。' + (!providerChosen ?
      '磁盘未指定提供方。本页初始选择不代表运行配置；先选择提供方，再确认选用。' :
      !snapshot.disk_provider_explicit || provider !== snapshot.provider ? '已选择当前提供方，仅在草稿中；保存后才写入。' : '磁盘已明确指定当前提供方。');
    for (const [name, label, id] of [['model', '模型', 'model-state'], ['base_url', '地址', 'url-state']]) {
      $('settings-setup-' + id).textContent = values[name] ?
        (values[name] === (saved[name] || '') ? label + '已在磁盘填写，尚未证明可用。' : label + '已填写在草稿中，尚未保存。') :
        saved[name + '_present'] ? label + '在磁盘显式为空，请填写替换值。' : label + '未在磁盘配置，启动可能采用进程值或默认值；建议明确填写。';
    }
    $('settings-setup-key-state').textContent = provider === 'ollama' ? 'Ollama 无需此密钥；本页不会启动服务或下载模型。' :
      clearKey() ? '当前选择清除密钥；保存后将不再使用磁盘旧密钥。' :
      key() ? '已输入新密钥，仅在当前表单中；尚未保存。' :
      saved.api_key_configured ? '该提供方磁盘密钥已配置，留空保持；尚未证明认证有效。' : '该提供方尚未配置密钥，请在下方输入；其他提供方的密钥不会代用。';
    $('settings-setup-test-state').textContent = connection ?
      (probeMessages[connection.status] || '检测没有完成，请核对检测提示。') + ' 不能证明生成任务或工具调用成功。' :
      '当前草稿尚未检测。只有点击原“测试连接”按钮才联网，不发送生成请求；此步骤可跳过。';
    $('settings-setup-save-state').textContent = uncertain ? '不能确定是否已写入，禁止自动重试。' :
      dirty() ? '草稿未保存；前往查看实际差异，再单独确认。' :
      snapshot.restart_required ? '已保存，重启后生效；原有本次撤销入口继续保留。' : '没有未保存改动；引导不会写入磁盘。';
    $('settings-setup-retrieval-state').textContent = values.embedding_channel === 'api' ?
      '联网检索独立使用百炼配置及密钥，还需核对检索模型、维度和地址；切换模型或维度须重建索引。' :
      snapshot.local_embedding_available ? '本地依赖可发现，仍需准备离线模型目录和维度；不会自动下载。' :
      '缺少 sentence-transformers 或完整的 research-local ONNX 依赖。本地检索暂不可用，需要另行准备依赖和离线目录；引导不会安装或下载。';
    $('settings-setup-toggle').disabled = blocked;
    guideActions.forEach(name => { $('settings-setup-' + name).disabled = blocked; });
    $('settings-setup-provider').disabled ||= providerChosen;
    $('settings-setup-defaults').disabled ||= Boolean(values.model && values.base_url);
    $('settings-setup-key').hidden = provider === 'ollama';
    $('settings-setup-key').disabled ||= clearKey();
    $('settings-setup-save').disabled ||= !dirty();
  }
  function focusSetting(id) {
    const target = $(id);
    for (let parent = target.parentElement; parent; parent = parent.parentElement) {
      if (parent.tagName === 'DETAILS') parent.open = true;
    }
    target.scrollIntoView({block: 'center'});
    target.focus();
  }
  function errors(values) {
    const issues = {};
    if (!snapshot.providers.includes(values.provider)) issues.provider = '请选择已支持的提供方。';
    for (const name of ['model', 'embedding_model']) {
      if (values[name] && !(new RegExp('^' + snapshot.model_pattern + '$')).test(values[name])) {
        const localPath = name === 'embedding_model' && values.embedding_channel === 'local' &&
          /^(?:[A-Za-z]:\/|\/)/.test(values[name]) && values[name].length <= 200 && !/[\x00-\x1f\x7f'"$#=\\]/.test(values[name]);
        if (!localPath) issues[name] = '请输入有效名称；离线目录使用绝对路径和正斜杠。';
      }
    }
    for (const name of ['model', 'base_url', 'embedding_model']) {
      const previous = name === 'embedding_model' ? baseline?.[name] : snapshot.profiles[values.provider]?.[name];
      if (!values[name] && previous) issues[name] = '已有值不能在此留空清除，请填写替换值；放弃改动可恢复原值。';
      if (name !== 'embedding_model' && !values[name] && snapshot.profiles[values.provider]?.[name + '_present']) issues[name] = '磁盘此字段为空；启动不会自动采用默认值，请填写替换值。';
    }
    if (values.base_url) {
      try {
        const url = new URL(values.base_url);
        if (!/^https?:\/\//.test(values.base_url) || values.base_url.length > 2048 || !['http:', 'https:'].includes(url.protocol) ||
            !url.hostname || url.username || url.password || url.search || url.hash || url.port === '0' ||
            /[\s\\\x00-\x1f\x7f]/.test(values.base_url) ||
            (values.provider !== 'ollama' && url.protocol !== 'https:')) throw Error();
      } catch (_) { issues.base_url = '填写完整 API 地址；云端须为 HTTPS，不含用户信息、查询参数或片段。'; }
    }
    if (!['local', 'api'].includes(values.embedding_channel)) issues.embedding_channel = '请选择本地离线或联网接口。';
    for (const [name, [min, max, integral]] of Object.entries(snapshot.limits)) {
      const raw = values[name];
      if (!raw) continue;
      const value = Number(raw), syntax = integral ? /^(?:0|[1-9][0-9]*)$/ : /^(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$/;
      if (raw.length > 32 || !syntax.test(raw) || !Number.isFinite(value) || value < min || value > max ||
          (integral && !Number.isSafeInteger(value))) issues[name] = `请输入 ${min} 到 ${max} 的${integral ? '整数' : '数值'}，不要使用科学计数法。`;
    }
    if (key() && (key().length > 4096 || /[\s'"\\$#=\x00-\x1f\x7f]/.test(key()))) issues.api_key = '密钥不能含空白、引号或配置控制字符。';
    Object.entries({...fields, api_key: 'settings-key'}).forEach(([name, id]) => {
      $(id).setAttribute('aria-invalid', String(Boolean(issues[name])));
      $(id + '-error').textContent = issues[name] || '';
    });
    return Object.keys(issues).length === 0;
  }
  function clearSensitive() {
    $('settings-key').value = '';
    $('settings-clear-key').checked = false;
  }
  function profile(name) {
    const value = drafts[name] || snapshot.profiles[name];
    $('settings-provider').value = name;
    $('settings-model').value = value?.model || '';
    $('settings-url').value = value?.base_url || '';
    $('settings-model').placeholder = snapshot.defaults[name].model;
    $('settings-url').placeholder = snapshot.defaults[name].base_url;
    $('settings-key-state').textContent = name === 'ollama' ? 'Ollama 不使用此密钥。' :
      (snapshot.profiles[name].api_key_configured ? '该提供方已配置密钥。' : '该提供方未配置密钥。') + ' 留空保持，输入替换，勾选清除。';
    provider = name;
  }
  function show(value) {
    snapshot = value;
    providerChosen = Boolean(value.disk_provider_explicit);
    connection = null;
    drafts = {};
    Object.entries(fields).forEach(([name, id]) => { $(id).value = value[name] ?? ''; });
    clearSensitive(); profile(value.provider);
    baseline = read();
    $('settings-runtime').textContent = '磁盘：本机 .env。运行配置快照不可得，不能据此确认当前模型；已构建的模型不会被本页替换。保存后请重启。';
    if (!value.disk_provider_explicit) $('settings-runtime').textContent += ' 磁盘未指定提供方：本页初始展示 DeepSeek；主模型入口另有命令行、进程环境和百炼默认值，不能据此推断运行提供方。';
    $('settings-embedding-local').textContent = value.local_embedding_available ?
      '本地依赖可发现；仍需准备已有离线模型目录。不会自动安装或下载。' :
      '缺少可选依赖 sentence-transformers 或完整的 research-local ONNX 依赖，本地检索暂不可用。需另行准备依赖与离线模型；本页不会安装或下载。';
    errors(read()); refresh();
  }
  function failure(result) {
    const reason = result?.reason || result?.error;
    if (result?.field && fields[result.field]) {
      $(fields[result.field] + '-error').textContent = '宿主校验未通过，请按字段格式和范围修正。';
      $(fields[result.field]).setAttribute('aria-invalid', 'true');
    }
    if (result?.outcome_uncertain || reason === 'settings_uncertain') {
      uncertain = true; saveId = null;
      invalidateRag(true);
      status('uncertain', '操作结果没能确认；可能已经写入，请核对配置、备份和记录，不要自动重试。');
    } else status(reason === 'settings_invalid' ? 'invalid' : 'denied', reasons[reason] || '操作被拒绝，请重新读取或检查本机设置服务。');
  }
  async function run(kind, operation) {
    if (pending || !live() || uncertain) return;
    pending = true; const mine = epoch;
    const current = () => mine === epoch && live();
    refresh();
    try { await operation(api(), current); }
    catch (_) {
      if (current()) {
        if (kind === 'write') failure({outcome_uncertain: true});
        else status('denied', kind === 'probe' ? '检测请求失败，配置没有写入。' : '本机配置读取失败，请检查设置服务。');
      }
    } finally {
      if (mine === epoch) { pending = false; refresh(); }
    }
  }
  async function guard() {
    if (pending) { status(state, '操作正在进行，请等待返回后再关闭。'); return false; }
    const discarded = dirty();
    if (discarded && await ask('放弃全部未保存改动并清空新密钥？各提供方草稿也会丢弃；磁盘配置不变。') !== true) return false;
    epoch++; revision++; clearSensitive();
    if (snapshot && discarded) show(snapshot);
    drafts = {};
    if (!uncertain && discarded) status(snapshot?.restart_required ? 'saved_restart' : 'clean',
      snapshot?.restart_required ? '已放弃未保存改动；磁盘配置已保存，重启后生效。' : '已放弃未保存改动，磁盘配置不变。');
    refresh(); return true;
  }
  window.OfficeSettingsEditor = {mounted: true, requestClose: guard};
  $('settings-setup-toggle').addEventListener('click', () => {
    if (pending || !snapshot || !live() || uncertain) return;
    guideExpanded = !guideExpanded; refresh();
  });
  $('settings-setup-provider').addEventListener('click', () => {
    if (pending || !snapshot || !live() || uncertain || providerChosen) return;
    providerChosen = true; revision++;
    status('dirty', '已在草稿中选用当前提供方；尚未写入磁盘，保存仍需单独确认。');
  });
  $('settings-setup-defaults').addEventListener('click', () => {
    if (pending || !snapshot || !live() || uncertain) return;
    const missing = ['model', 'base_url'].filter(name => !$(fields[name]).value);
    if (!missing.length) return;
    missing.forEach(name => { $(fields[name]).value = snapshot.defaults[provider][name]; });
    providerChosen = true; connection = null; revision++;
    const valid = errors(read());
    status(valid ? 'dirty' : 'invalid', valid ? '已补齐空白地址与模型草稿；自定义值和密钥保持，尚未保存或联网。' : '已补齐空白草稿；请继续修正字段旁提示。');
    focusSetting(fields[missing[0]]);
  });
  for (const [name, id] of [['key', 'settings-key'], ['test', 'settings-test'], ['save', 'settings-save'],
    ['context', 'settings-context'], ['retrieval', 'settings-embedding-channel']]) {
    $('settings-setup-' + name).addEventListener('click', () => {
      if (!pending && snapshot && live() && !uncertain) focusSetting(id);
    });
  }
  async function load() {
    if (pending || !live()) return;
    if (dirty() && await ask('重新读取会放弃全部草稿并清空新密钥，是否继续？') !== true) return;
    clearSensitive();
    return run('read', async (host, current) => {
      status('loading', '正在读取磁盘配置…');
      const result = await host.get_settings();
      if (!current()) return;
      if (!result?.ok) { failure(result); return; }
      show(result.settings);
      status(result.settings.restart_required ? 'saved_restart' : 'clean', result.settings.restart_required ? '已读取磁盘配置，已保存，重启后生效。' : '已读取，没有未保存改动。');
    });
  }
  $('settings-load').addEventListener('click', load);
  $('settings-discard').addEventListener('click', () => {
    if (pending || !snapshot || !live()) return;
    revision++; show(snapshot); status('clean', '已放弃改动，恢复最近读取的快照；没有写入磁盘。');
  });
  [...Object.values(fields), 'settings-key', 'settings-clear-key'].forEach(id => $(id).addEventListener('input', async () => {
    if (!snapshot || !live()) return;
    revision++; connection = null;
    if (id === 'settings-provider') {
      const target = $('settings-provider').value;
      if (!snapshot.providers.includes(target)) { $('settings-provider').value = provider; return; }
      if ((key() || clearKey()) && await ask('切换提供方会清空新密钥和清除密钥选项；模型与地址草稿会保留，是否继续？') !== true) {
        $('settings-provider').value = provider; refresh(); return;
      }
      drafts[provider] = {model: $('settings-model').value, base_url: $('settings-url').value};
      clearSensitive(); profile(target);
      providerChosen = true;
    }
    if (id === 'settings-clear-key' && clearKey()) $('settings-key').value = '';
    const valid = errors(read());
    status(valid ? (dirty() ? 'dirty' : 'clean') : 'invalid', valid ?
      (dirty() ? '存在未保存改动。提供方草稿仅保留在当前窗口内存；保存仅提交当前提供方。' : '没有未保存改动。') : '校验失败，请修正字段旁的提示。');
  }));
  function changes(values) {
    const result = {provider: values.provider};
    for (const [name, value] of Object.entries(values)) {
      if (name === 'provider' || value === '') continue;
      const previous = ['model', 'base_url'].includes(name) ? snapshot.profiles[values.provider][name] || '' : baseline[name];
      if (value !== previous) result[name] = name in snapshot.limits ? Number(value) : value;
    }
    if (provider !== 'ollama') {
      if (clearKey()) result.api_key = '';
      else if (key()) result.api_key = key();
    }
    return result;
  }
  function invalidateRag(uncertain = false) {
    const event = new Event('officerag:invalidate'); event.ragUncertain = uncertain;
    window.dispatchEvent(event);
  }
  async function reread(host, current, text) {
    let result;
    try { result = await host.get_settings(); }
    catch (_) { if (current()) status('saved_restart', text + ' 磁盘复读失败；撤销入口保留。'); return; }
    if (!current()) return;
    if (!result?.ok) { status('saved_restart', text + ' 磁盘复读失败；撤销入口保留，请检查本机配置。'); return; }
    show(result.settings); status('saved_restart', text);
  }
  $('settings-save').addEventListener('click', () => {
    if (!snapshot || !dirty() || pending || !live()) return;
    const values = read();
    if (!errors(values)) { status('invalid', '校验失败，请修正对应字段。'); return; }
    const change = changes(values), version = revision;
    const otherDrafts = Object.entries(drafts).some(([name, value]) => name !== provider &&
      (value.model !== (snapshot.profiles[name].model || '') || value.base_url !== (snapshot.profiles[name].base_url || '')));
    if (Object.keys(change).length === 1 && change.provider === snapshot.provider && snapshot.disk_provider_explicit) {
      status('dirty', '当前提供方没有改动；其他提供方草稿尚未保存，请切换后保存。'); return;
    }
    return run('write', async (host, current) => {
      status('saving', '正在校验并生成保存计划…');
      let result;
      try { result = await host.plan_settings_edit(change, snapshot.revision); }
      finally { delete change.api_key; if (current()) clearSensitive(); }
      if (!current()) return;
      if (!result?.ok || !opaque(result.plan?.plan_id)) { failure(result); return; }
      const plan = result.plan;
      const lines = (plan.differences || []).map(item => `${labels[item.field] || item.field}：${item.before ?? '未设置'} → ${item.after ?? '未设置'}`);
      const approved = version === revision && await ask('确认保存以下实际改动？\n' + lines.join('\n') +
        '\n密钥：' + (plan.differences?.find(item => item.field === 'api_key')?.after || '保持') +
        (otherDrafts ? '\n其他提供方还有草稿；本次保存成功后重新读取会丢弃这些草稿。' : '') +
        '\n所有本机账户共享此配置；确认后备份、写入并校验。已保存，重启后生效。') === true;
      if (!current()) return;
      const confirmed = approved && version === revision;
      const receipt = await host.confirm_settings_update(plan.plan_id, confirmed);
      if (!current()) return;
      if (!receipt?.ok) { failure(receipt); return; }
      const outcome = receipt.result;
      if (outcome?.status === 'saved' && outcome.applied === true && opaque(outcome.save_id)) {
        window.dispatchEvent(new Event('officesettings:saved'));
        saveId = outcome.save_id; invalidateRag();
        await reread(host, current, '已保存，重启后生效；可撤销本次改动。');
      } else if (outcome?.status === 'rejected' && outcome.wrote_nothing === true) failure({reason: outcome.code || 'settings_denied'});
      else failure({outcome_uncertain: true});
    });
  });
  $('settings-rollback').addEventListener('click', async () => {
    if (!saveId || pending || !live()) return;
    if (dirty() && await ask('撤销将丢弃未保存草稿并清空新密钥，是否继续？') !== true) return;
    clearSensitive();
    return run('write', async (host, current) => {
      status('saving', '正在准备撤销…');
      const target = saveId, result = await host.plan_settings_rollback(target);
      if (!current()) return;
      if (!result?.ok || !opaque(result.plan?.rollback_id)) { failure(result); return; }
      const approved = await ask('从本次备份恢复配置？确认后写入、校验，重启后生效。' +
        (result.plan.restores_absent_source_as_empty_env ? ' 原来没有配置文件，将恢复为空 .env。' : '')) === true;
      if (!current() || target !== saveId) return;
      const receipt = await host.confirm_settings_rollback(result.plan.rollback_id, approved);
      if (!current()) return;
      if (receipt?.ok && receipt.result?.status === 'rolled_back' && receipt.result.restored === true) {
        window.dispatchEvent(new Event('officesettings:saved'));
        saveId = null; invalidateRag(); await reread(host, current, '已恢复并校验配置，重启后生效。');
      } else if (receipt?.result?.wrote_nothing === true) failure({reason: receipt.result.code || 'settings_denied'});
      else failure(receipt?.ok ? {outcome_uncertain: true} : receipt);
    });
  });
  const probeMessages = {
    reachable: '服务可达，模型列表读取成功。', authentication_failed: '认证失败，请检查当前提供方的密钥或访问权限。',
    timeout: '检测超时，请检查地址和网络。', models_unsupported: '服务响应，但无法通过模型列表验证；可继续保留手动模型名称。',
    redirect_refused: '服务要求重定向，已拒绝转发凭据；请手动核对最终 API 地址。',
    models_unverified: '服务响应，但无法通过模型列表验证；可继续保留手动模型名称。',
    service_failed: '服务返回错误，请检查服务状态；未发送生成请求。',
    response_too_large: '模型列表超过检测大小上限，未继续读取。', connection_failed: '连接失败，请检查 API 地址、证书或本机服务状态。'
  };
  $('settings-test').addEventListener('click', () => {
    if (!snapshot || pending || !live() || !errors(read())) return;
    const values = read(), draft = {provider, model: values.model || snapshot.defaults[provider].model,
      base_url: values.base_url || snapshot.defaults[provider].base_url, clear_key: clearKey()};
    if (key()) draft.api_key = key();
    return run('probe', async (host, current) => {
      $('settings-test-status').textContent = '正在检测当前草稿；只读取模型列表，不生成内容…';
      let result;
      try { result = await host.test_settings_connection(draft); }
      finally { delete draft.api_key; }
      if (!current()) return;
      if (!result?.ok) { $('settings-test-status').textContent = reasons[result?.reason || result?.error] || '检测被拒绝。'; return; }
      const outcome = result.result;
      connection = {status: outcome?.status};
      $('settings-test-status').textContent = (probeMessages[outcome?.status] || '检测未能完成。') +
        (outcome?.status === 'reachable' && !outcome.model_listed ? ' 当前手动模型未在列表中，仍可保留。' : '') +
        ' 此检测不能证明生成任务或工具调用成功；未改动配置。';
    });
  });
  function retire() {
    epoch++; revision++; pending = false; snapshot = null; baseline = null; drafts = {}; saveId = null;
    providerChosen = false; guideExpanded = null; connection = null;
    clearSensitive(); Object.values(fields).forEach(id => { $(id).value = ''; });
    $('settings-test-status').textContent = ''; $('settings-key-state').textContent = '';
    status('clean', '尚未读取配置。');
  }
  window.addEventListener('officeauth:unauthenticated', retire);
  window.addEventListener('officeauth:authenticated', retire);
  window.addEventListener('pagehide', retire);
  document.addEventListener('officesettings:closed', () => { epoch++; clearSensitive(); refresh(); });
  document.addEventListener('officesettings:opened', async () => {
    refresh(); const mine = epoch;
    try {
      const authenticated = await window.OfficeAuth.ensure();
      if (mine !== epoch || !window.OfficeSettings.isOpen()) return;
      refresh(); if (authenticated && !snapshot && !pending) await load();
    } catch (_) { if (mine === epoch) status('denied', '无法确认登录状态，请重新登录。'); }
  });
  refresh();
})();
