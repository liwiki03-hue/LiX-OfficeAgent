"use strict";
(() => {
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  const $ = id => document.getElementById(id);
  const host = () => window.pywebview?.api && window.OfficeAuth?.wrapApi(window.pywebview.api);
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated?.());
  let mode = 'agent', epoch = 0, readSequence = 0, state = null, selected = null, changing = false;
  let sending = null, streamed = '', failedPrompt = '', transfer = null, preferencePending = false;
  let renameDirty = false, renameFor = null;
  let submitting = null, contextPreview = null, previewSequence = 0, previewPending = false;
  let noticeFeedback = null;
  let budgetFeedbackToken = null;
  let continuation = null, replySaving = false;
  let exporting = false;
  let imageRead = 0, imagePreferencePending = false;
  let liveStream = null;
  const retiredStreams = new Set();
  function retireStream(id) {
    if (!id) return;
    retiredStreams.add(id);
    if (retiredStreams.size > 128) retiredStreams.delete(retiredStreams.values().next().value);
  }
  let inputRevision = 0;
  const exportSelection = new Set();
  const drafts = {chat: '', agent: '', image: ''}, sessionDrafts = new Map();
  const input = $('prompt'), panes = $('content-panes');
  const make = (tag, id, text, parent) => {
    const node = document.createElement(tag);
    if (id) node.id = id;
    if (text) node.textContent = text;
    parent.append(node);
    return node;
  };
  const button = (id, text, parent, action) => {
    const node = make('button', id, text, parent);
    node.type = 'button'; node.className = 'btn secondary'; node.onclick = action;
    return node;
  };
  const panel = make('section', 'daily-conversations', '', panes);
  panel.setAttribute('aria-label', '普通问答会话');
  const controls = make('div', '', '', panel); controls.className = 'daily-actions';
  button('chat-new', '新建问答会话', controls, () => change('new_conversation'));
  const searchLabel = make('label', '', '查找 ', controls);
  const search = make('input', 'chat-search', '', searchLabel); search.type = 'search'; search.maxLength = 100;
  search.placeholder = '会话标题或模型'; search.setAttribute('aria-label', '搜索会话标题或模型'); search.oninput = () => renderChoices();
  const label = make('label', '', '会话 ', controls);
  const picker = make('select', 'chat-session', '', label); picker.setAttribute('aria-label', '普通问答会话列表');
  picker.onchange = () => change('select_conversation', picker.value);
  button('chat-refresh', '重新读取历史', controls, async () => { if (await ask('重新读取将替换窗口内的会话记录；未保存消息可能不再显示，输入草稿保留。不会调用模型或恢复批准。继续？')) change('reload_conversations'); });
  button('chat-delete', '删除本会话', controls, async () => {
    const session = selected, born = epoch;
    if (session && await ask('删除此问答会话的本地文字记录？不会撤销模型费用，也不删除文档或图片。') && born === epoch && selected === session) await change('delete_conversation', session, true);
  });
  const previewButton = button('chat-preview-context', '预览本次发送范围', controls, previewContext);
  const nameControls = make('div', '', '', panel); nameControls.className = 'daily-actions';
  const nameLabel = make('label', '', '会话名称 ', nameControls);
  const titleInput = make('input', 'chat-title', '', nameLabel); titleInput.maxLength = 80;
  titleInput.setAttribute('aria-label', '当前会话名称');
  titleInput.oninput = () => { renameDirty = true; renameFor = selected; titleHint.textContent = '名称草稿尚未采用；请点击“重命名”。'; };
  button('chat-rename', '重命名', nameControls, () => change('rename_conversation', selected, titleInput.value));
  button('chat-save-now', '保存当前历史', nameControls, () => change('save_conversation_history'));
  const titleHint = make('small', 'chat-title-hint', '', nameControls); titleHint.setAttribute('role', 'status');
  const capacity = make('p', 'chat-capacity', '', panel); capacity.setAttribute('role', 'status');
  const saveStatus = make('p', 'chat-save-status', '', panel); saveStatus.setAttribute('role', 'status');
  const pref = make('label', '', ' ', controls);
  const enabled = make('input', 'chat-save-history', '', pref); enabled.type = 'checkbox';
  pref.append(document.createTextNode('保存普通问答到本机'));
  enabled.onchange = async () => {
    const desired = enabled.checked, born = epoch;
    enabled.disabled = true; preferencePending = true;
    try {
      const plan = await host().plan_settings_update({chat_history: desired});
      if (born !== epoch) return;
      if (!plan?.ok) { setNotice(error(plan)); return; }
      const yes = await ask((desired ? '启用本机问答保存：当前窗口的会话及后续消息会保存文字、时间和模型来源。' : '停止保存新消息；已有历史仍可查看或主动删除。') + '不保存密钥，不恢复任何文件或写入批准。是否保存此偏好？');
      if (born !== epoch) return;
      const saved = await host().confirm_settings_update(plan.plan.plan_id, yes);
      if (born !== epoch) return;
      if (saved?.ok && yes && desired) await host().save_conversation_history();
      if (born !== epoch) return;
      setNotice(saved?.ok ? (yes ? '保存偏好已更新。' : '已取消，偏好未改变。') : error(saved));
      await refresh();
    } catch (_) { if (born === epoch) setNotice('保存偏好结果不确定，请重新读取设置核对；没有自动重试。'); }
    finally { if (born === epoch) { preferencePending = false; enabled.disabled = false; render(); } }
  };
  const streamLabel = make('label', '', ' ', controls);
  const stream = make('input', 'chat-stream', '', streamLabel); stream.type = 'checkbox'; stream.checked = true;
  streamLabel.append(document.createTextNode('流式回答'));
  const lengthLabel = make('label', '', '回答长度 ', controls);
  const length = make('select', 'chat-reply-length', '', lengthLabel);
  for (const [value,title] of [['512','简短（兼容默认）'],['2048','标准'],['4096','详细'],['custom','高级数值']]) { const item = make('option','',title,length); item.value = value; }
  const outputTokens = make('input','chat-reply-tokens','',lengthLabel); outputTokens.type = 'number'; outputTokens.min = '64'; outputTokens.max = '16384'; outputTokens.value = '512'; outputTokens.hidden = true; outputTokens.setAttribute('aria-label','最大输出 tokens');
  length.onchange = () => { outputTokens.hidden = length.value !== 'custom'; if (length.value !== 'custom') outputTokens.value = length.value; clearPreview(); };
  outputTokens.oninput = clearPreview;
  button('chat-reply-save','保存回答预算',controls,async () => {
    if (replySaving) return;
    const born = epoch, tokens = Number(outputTokens.value);
    if (!Number.isInteger(tokens) || tokens < 64 || tokens > 16384) { setNotice('最大输出预算须为64至16384的整数。'); return; }
    replySaving = true;
    try {
      const plan = await host().plan_settings_update({chat_reply_tokens:tokens});
      if (born !== epoch) return;
      if (!plan?.ok) { setNotice(error(plan)); return; }
      const yes = await ask('保存普通问答最大输出预算为 ' + tokens + ' tokens？这是预算倾向，不保证字数；较大预算可能增加等待与费用，已有本地32000字符、2048分块、90秒限制保持。');
      if (born !== epoch) return;
      const result = await host().confirm_settings_update(plan.plan.plan_id, yes);
      if (born !== epoch) return;
      clearPreview(); setNotice(result?.ok ? (yes ? '回答预算已保存，用于后续普通问答。' : '已取消预算保存。') : error(result)); await refresh();
    } catch (_) { if (born === epoch) setNotice('预算保存结果未确认，请重读设置；未自动重试。'); }
    finally { if (born === epoch) replySaving = false; }
  });
  make('small','','长度是输出预算；与上下文窗口、输入估算和服务商用量分别核对。未保存的选择不用于发送。',panel);
  const notice = make('p', 'chat-notice', '', panel); notice.setAttribute('role', 'status');
  const location = make('details', '', '', panel); make('summary', '', '本地保存范围与位置', location);
  const storage = make('p', 'chat-storage', '', location);
  button('chat-clear-saved', '清除本账户的已保存问答', location, async () => {
    if (await ask('清除当前账户已保存的问答，并清空本窗口未保存消息？输入草稿保留；不删除文档或图片，不撤销费用。')) await change('clear_saved_conversations', true);
  });
  const contextPanel = make('details', 'chat-context-preview', '', panel); contextPanel.hidden = true;
  make('summary', '', '本次发送范围 · 只读预览', contextPanel);
  const contextContent = make('div', 'chat-context-content', '', contextPanel); contextContent.setAttribute('aria-live', 'polite');
  const budgetFeedback = make('section', 'chat-budget-feedback', '', panel); budgetFeedback.hidden = true;
  budgetFeedback.setAttribute('aria-label', '上下文预算处理');
  const budgetMessage = make('p', 'chat-budget-message', '', budgetFeedback); budgetMessage.setAttribute('role', 'status');
  const budgetActions = make('div', '', '', budgetFeedback); budgetActions.className = 'daily-actions';
  button('chat-budget-shorten', '回到输入缩短内容', budgetActions, () => input.focus());
  button('chat-budget-configure', '调整上下文配置', budgetActions, () => $('model-configure').click());
  button('chat-budget-preview', '重新预览', budgetActions, previewContext);
  const transcript = make('div', 'chat-transcript', '', panel); transcript.setAttribute('aria-label', '当前问答消息');
  const exportControls = make('div','chat-export-controls','',panel); exportControls.className = 'daily-actions';
  const exportFormatLabel = make('label','','导出格式 ',exportControls);
  const exportFormat = make('select','chat-export-format','',exportFormatLabel);
  for (const [value,title] of [['md','Markdown 原文'],['txt','纯文本（保留原文语法）']]) { const node = make('option','',title,exportFormat); node.value = value; }
  const exportMetaLabel = make('label','','',exportControls);
  const exportMeta = make('input','chat-export-metadata','',exportMetaLabel); exportMeta.type = 'checkbox'; exportMetaLabel.append(document.createTextNode('包含接口、模型与时间'));
  button('chat-export-selected','导出所选轮次',exportControls,() => exportTurns([...exportSelection]));
  button('chat-export-session','导出当前会话',exportControls,() => exportTurns((state?.sessions.find(item=>item.id===selected)?.turns || []).map(turn=>turn.id)));
  const exportResult = make('section','chat-export-result','',panel); exportResult.setAttribute('role','status');
  async function exportTurns(ids, content='turns', allowPartial=false) {
    if (exporting || !authenticated()) return;
    if (!ids.length) { setNotice('请先选择要导出的轮次。'); return; }
    const born = epoch, session = selected; exporting = true;
    try {
      const result = await host().plan_chat_export({session_id:session,turn_ids:ids,format:exportFormat.value,metadata:exportMeta.checked,content,allow_partial:allowPartial});
      if (born !== epoch || selected !== session) return;
      if (!result?.ok) { setNotice(error(result)); return; }
      const plan = result.plan;
      const yes = await ask('导出' + (plan.partial ? '当前部分内容的固定快照（未完成）' : content === 'answer' ? '单条回答' : plan.turns + '轮对话') + '，格式 ' + plan.format + '，' + (plan.metadata ? '含模型和时间' : '不含模型和时间') + '。\n新文件：' + plan.path + '\n' + plan.bytes + ' 字节；不覆盖已有文件，不调用模型。确认写入？');
      if (born !== epoch || selected !== session) return;
      const saved = await host().confirm_chat_export(plan.plan_id,yes);
      if (born !== epoch || selected !== session) return;
      exportResult.replaceChildren();
      if (!saved?.ok) { make('p','',error(saved),exportResult); return; }
      const value = saved.result;
      if (value.status !== 'saved') { make('p','',value.status === 'uncertain' ? '写入结果不确定，请核对目标目录；不要重复导出。' : !yes ? '已取消；没有写入导出文件。' : value.code === 'chat_export_path_conflict' ? '导出路径已存在，没有覆盖；请重新选择并生成新计划。' : '写入被拒绝，请重新核对范围、路径和有效期。',exportResult); return; }
      make('p','','已导出 ' + value.path.split(/[\\/]/).pop(),exportResult);
      const savedDetails = make('details','','',exportResult); make('summary','','路径与校验详情',savedDetails);
      make('p','','已重开校验：' + value.path,savedDetails);
      button('','打开导出位置',exportResult,async()=> { const opened = await host().open_chat_export_folder(value.save_id); if (born === epoch && !opened?.ok) setNotice(error(opened)); });
      let rollbackPending = false;
      button('','撤销本次导出',exportResult,async()=> {
        if (rollbackPending || born !== epoch || selected !== session) return;
        rollbackPending = true;
        try {
          const rollback = await host().plan_chat_export_rollback(value.save_id);
          if (born !== epoch || selected !== session) return;
          if (!rollback?.ok) { setNotice(error(rollback)); return; }
          const approved = await ask('删除本次仍未变化的导出文件？审计保留。') === true;
          if (born !== epoch || selected !== session) return;
          const removed = await host().confirm_chat_export_rollback(value.save_id,approved);
          if (born === epoch) setNotice(removed?.result?.status === 'rolled_back' ? '本次导出文件已撤销，审计保留。' : '撤销未完成，请核对结果；没有自动重试。');
        } finally { rollbackPending = false; }
      });
    } catch (_) { if (born === epoch) setNotice('导出结果未确认，请核对目录和记录；未自动重试。'); }
    finally { if (born === epoch) exporting = false; }
  }
  const live = make('pre', 'chat-live-answer', '', panel); live.hidden = true; live.setAttribute('aria-label', '正在生成，尚未完成的回答');
  button('chat-latest','回到最新',controls,()=> (live.hidden ? transcript.lastElementChild || transcript : live).scrollIntoView({block:'end'}));
  const retry = button('chat-retry', '填回失败输入，再主动发送', panel, () => { input.value = failedPrompt; keepDraft(); input.focus(); }); retry.hidden = true;
  button('chat-cancel-continuation','改为普通提问（取消续写关联）',controls,()=> { continuation=null; clearPreview(); setNotice('已取消续写关联，输入保留；后续作为普通问答重新核对。'); input.focus(); });
  const imageParams = make('div', 'daily-image-parameters', '', $('model-selector-bar'));
  const sizeLabel = make('label', '', '图片尺寸 ', imageParams);
  const size = make('select', 'daily-image-size', '', sizeLabel);
  for (const value of ['1024x1024', '1024x1536', '1536x1024']) { const option = make('option', '', value, size); option.value = value; }
  const styleLabel = make('label', '', '风格提示 ', imageParams);
  const style = make('select', 'daily-image-style', '', styleLabel);
  for (const [value, title] of [['natural','自然'], ['vivid','鲜明']]) { const option = make('option', '', title, style); option.value = value; }
  make('span', '', '单图；风格加入提示词，生成前确认有效参数。', imageParams);
  const historyActions = make('div', 'daily-history-actions', '', $('model-selector-bar')); historyActions.className = 'daily-actions';
  button('daily-open-chats', '问答会话', historyActions, async () => {
    const result = await host()?.set_model_mode('chat');
    if (result?.ok) { await window.OfficeModelSelector.refresh(); window.dispatchEvent(new Event('officeworkspace:compose')); }
  });
  const imageHistory = make('section', 'daily-image-history', '', panes); imageHistory.hidden = true;
  make('h3', '', '本机图片记录 · 只读', imageHistory);
  const imageFilters = make('div','','',imageHistory); imageFilters.className = 'daily-actions';
  const imageFilter = make('input','image-history-filter','',imageFilters); imageFilter.type = 'search'; imageFilter.placeholder = '按模型或日期筛选'; imageFilter.setAttribute('aria-label','按模型或日期筛选历史图片');
  const imageStatus = make('select','image-history-status','',imageFilters); imageStatus.setAttribute('aria-label','图片记录状态');
  for (const [value,title] of [['','全部状态'],['generated','已生成'],['generating','请求已记录'],['failed','失败或待核对'],['saved','已记录保存']]) { const item=make('option','',title,imageStatus); item.value=value; }
  const filterImages = ()=> { for (const row of imageRecords.children) if (row.dataset.search !== undefined) row.hidden = !row.dataset.search.includes(imageFilter.value.toLocaleLowerCase()) || Boolean(imageStatus.value && (imageStatus.value === 'saved' ? row.dataset.saved !== 'true' : row.dataset.status !== imageStatus.value)); };
  imageFilter.oninput = imageStatus.onchange = filterImages;
  const reuseLabel = make('label','','',imageHistory);
  const reusePreference = make('input','image-reuse-preference','',reuseLabel); reusePreference.type='checkbox'; reuseLabel.append(document.createTextNode('为后续生图保存提示词与有效参数到本机审计（未加密，默认关闭）'));
  reusePreference.onchange = async()=> {
    if (imagePreferencePending) return;
    imagePreferencePending = true; const desired = reusePreference.checked, born=epoch;
    try {
      const plan = await host().plan_settings_update({image_reuse_metadata:desired});
      if (born !== epoch) return;
      if (!plan?.ok) { setNotice(error(plan)); return; }
      const yes = await ask(desired ? '为后续新生图记录保存原提示词、尺寸和风格？按本机账户保存，未加密；不保存密钥、完整响应或图片base64，不补回旧记录。' : '关闭后续提示词保存？已存记录仍保留，可在详情中查看。');
      if (born !== epoch) return;
      const result=await host().confirm_settings_update(plan.plan.plan_id,yes);
      if (born === epoch) setNotice(result?.ok ? yes ? '生图记录偏好已保存，仅用于后续任务。' : '已取消偏好修改。' : error(result));
    } catch (_) { if (born === epoch) setNotice('偏好保存结果未确认，请重读设置；未自动重试。'); }
    finally { if (born === epoch) imagePreferencePending=false; }
  };
  const imageRecords = make('div', 'daily-image-records', '', imageHistory);
  button('', '收起图片记录', imageHistory, () => { imageHistory.hidden = true; });
  button('daily-open-images', '图片记录', historyActions, async () => {
    const born = epoch, request=++imageRead;
    try {
      const result = await host()?.list_independent_image_history();
      if (born !== epoch || request !== imageRead || !authenticated()) return;
      imageRecords.replaceChildren(); imageHistory.hidden = false;
      if (!window.OfficeWorkbench) window.dispatchEvent(new Event('officeworkspace:compose'));
      if (!result?.ok) { make('p', '', '图片记录无法安全读取；不会读取图片文件或恢复旧操作。', imageRecords); return; }
      reusePreference.checked = Boolean(result.reuse_metadata_enabled);
      make('p', '', result.scope || '当前没有保存过的独立生图记录。', imageRecords);
      if (result.limited || result.unreadable) make('p', '', '记录读取有界，部分记录未覆盖或未通过检查。', imageRecords);
      for (const item of result.items) {
        const row = make('article', '', '', imageRecords); row.className = 'daily-turn';
        row.dataset.search = (item.model+' '+new Date(item.created*1000).toLocaleString()).toLocaleLowerCase(); row.dataset.status=item.status; row.dataset.saved=String(item.saved);
        const states = {generating:'生成尝试已记录，最终结果未知',generated:'已生成',failed:'失败或结果待核对',rolled_back:'已撤销输出'};
        make('p', '', new Date(item.created * 1000).toLocaleString() + ' · 生图 · ' + item.model + ' · ' + (states[item.status] || '状态需核对'), row);
        make('p', '', item.connection, row);
        if (item.path) make('p', '', (item.saved ? '记录的保存位置：' : '历史输出位置，当前不具备删除批准：') + item.path, row);
        make('small', '', '不会自动加载图片；历史不能恢复生成、保存或撤销权限。', row);
        const detail = make('details','','',row); make('summary','','记录详情',detail);
        make('p','','请求状态：' + item.request_state + '；摘要：' + item.sha256 + '；记录ID：' + item.id,detail);
        const feedback = make('p','','',row); feedback.setAttribute('role','status');
        const preview = make('div','','',row);
        let pending = false;
        const act = async(operation) => {
          if (pending || born !== epoch || request !== imageRead) return;
          pending = true; feedback.textContent='正在核验记录与文件…';
          const draftRevision = inputRevision, draftMode = mode, draftSession = selected;
          if (operation === 'preview_historical_image') preview.replaceChildren();
          try {
            const value = await host()[operation](item.id);
            if (born !== epoch || request !== imageRead || !authenticated()) return;
            if (!value?.ok) { feedback.textContent=imageErrors[value?.reason] || '无法安全读取或复用，未请求生图；请重新选择文件并授权读取。'; return; }
            if (operation === 'preview_historical_image') { preview.replaceChildren(); window.OfficeImages?.appendImages(preview,[value.image]); feedback.textContent='已核验并预览；不恢复旧保存或撤销批准。'; }
            else if (operation === 'open_historical_image_folder') feedback.textContent='已核验并打开保存位置。';
            else {
              if (draftRevision !== inputRevision || draftMode !== mode || draftSession !== selected) { feedback.textContent='输入、模式或会话已改变，未覆盖新草稿；可重新主动复用。'; return; }
              const switched = await host().set_model_mode('image');
              if (born !== epoch || request !== imageRead || !switched?.ok) return;
              await window.OfficeModelSelector?.refresh();
              if (born !== epoch || request !== imageRead || inputRevision !== draftRevision || selected !== draftSession) return;
              updateMode('image'); input.value=value.draft.prompt; size.value=value.draft.size; style.value=value.draft.style; keepDraft(); input.focus();
              if (window.OfficeWorkbench) window.dispatchEvent(new Event('officeworkspace:compose'));
              feedback.textContent='已填入可编辑草稿，当前接口：'+value.route.connection_name+' · '+value.route.model_id+'；请核对兼容参数，生成仍需新方案和独立确认。没有发送请求。';
            }
          } catch (_) { if (born === epoch && request === imageRead) feedback.textContent='操作未完成，没有重新生图；请核对文件和模型配置。'; }
          finally { pending=false; }
        };
        if (item.saved) { button('','预览已保存图片',row,()=>act('preview_historical_image')); button('','打开保存位置',row,()=>act('open_historical_image_folder')); }
        if (item.reusable) button('','复用提示词和参数到草稿',row,()=>act('reuse_historical_image'));
        else make('p','','旧记录没有保存提示词，无法恢复；不会根据图片或文件名推测。',row);
      }
      filterImages();
    } catch (_) { if (born === epoch && request === imageRead) { imageHistory.hidden = false; imageRecords.replaceChildren(); make('p', '', '无法读取图片记录，请重新打开记录；没有重新生图。', imageRecords); } }
  });
  const summary = make('p', 'daily-next-action', '', $('model-selector-bar')); summary.setAttribute('role', 'status');
  const errors = {not_authenticated:'请先登录。', chat_sensitive_input:'输入包含已配置的凭据，请移除后再发送；没有请求模型或保存该条消息。', chat_account_required:'当前账户身份不可用，不能保存或恢复问答。', chat_history_corrupt:'本地历史损坏或安全检查失败；仍可临时问答。请自行备份后主动清除记录，不会自动覆盖。', chat_history_write_failed:'回答已保留在本窗口，但本地历史写入失败。请检查保存目录和磁盘空间。', chat_history_conflict:'历史被其他窗口改动；未覆盖，请重新读取核对。', chat_history_capacity:'本机会话容量已满，请主动删除不需要的会话或新建分段任务。', busy:'任务仍在处理，请等待当前请求释放；浏览历史不受影响。', model_context_budget:'当前输入超出模型上下文预算，请缩短输入或选择合适模型。', chat_session_missing:'会话已不存在，请刷新列表。', settings_no_changes:'偏好没有变化。', settings_source_changed:'配置已被其他操作修改，请重新读取。'};
  Object.assign(errors, {chat_title_invalid:'会话名称需为1–80个字符，不含换行或控制字符。', chat_history_disabled:'本机保存尚未启用，请先通过保存偏好启用。', chat_history_uncertain:'本机历史保存结果不确定，磁盘可能已更新；不自动重试，请重新读取核对。'});
  Object.assign(errors, {chat_context_changed:'预览后输入、会话或模型配置已改变；没有发送，请重新预览核对。', chat_submission_uncertain:'未确认宿主是否受理；输入已保留，请先核对会话记录，不要重复发送。', model_selection_required:'请先选择普通问答模型；没有模型时可点击“配置接口”。', model_selection_invalid:'所选接口或模型不可用，请重新选择或修复配置。', model_configuration_changed:'接口配置已改变，请重新选择模型后再预览。'});
  Object.assign(errors,{chat_continuation_changed:'原回答或会话已改变，请重新选择续写并预览。',chat_continuation_unavailable:'此回复不能续写，请核对结束原因。',chat_continuation_limit:'续写链已达到16段，请调整任务范围。',chat_output_parameter_unsupported:'此模型未配置可用的输出限制字段，请在连接中按接口文档配置。',chat_export_in_progress:'当前回答仍在生成；请停止并核对已收到内容或等待完成后导出，不会把运行中内容当作完整答案。',chat_export_invalid:'请选择当前会话中有效的消息。',chat_export_capacity:'导出超过512 KiB或本窗口8项事务上限，请缩小范围或在新窗口重试。',chat_export_expired:'导出批准已失效，请重新核对范围并创建计划。',chat_export_io_failed:'导出文件或目录无法安全访问，请核对位置与权限。',evidence_path_changed:'目标路径不安全或已改变，请重新选择目录。',evidence_uncertain:'写入结果不确定，请核对目标文件与审计；不要重复写入。',cancelled:'已取消，没有创建文件。'});
  const error = result => errors[result?.reason] || '操作未完成，请重新读取状态核对；没有自动重试。';
  const imageErrors = {not_authenticated:'请先登录。',image_history_expired:'记录读取身份已失效，请重新打开图片记录。',image_history_changed:'与原生成记录不一致；请重新选择文件并授权读取，未预览。',image_history_missing:'记录的图片不存在；没有重新生成。',image_history_too_large:'图片大小超限，未读取或预览。',image_history_not_saved:'此记录没有仍有效的保存结果。',image_history_prompt_unavailable:'原记录没有保存提示词，不能恢复。',image_history_sensitive:'提示词包含当前敏感配置，拒绝复用。',model_selection_required:'没有有效的当前生图模型，请配置或重新选择。'};
  function setNotice(text, source = null) {
    noticeFeedback = {source};
    notice.textContent = text;
  }
  function clearPreview() {
    contextPreview = null; previewSequence++; previewPending = false;
    previewButton.disabled = false; contextPanel.hidden = true; contextContent.replaceChildren();
    budgetFeedback.hidden = true; budgetMessage.textContent = '';
    budgetFeedbackToken = null;
  }
  function showBudget(value) {
    budgetFeedbackToken = {};
    budgetFeedback.hidden = false;
    const usable = value && ['estimated_input_tokens','reply_token_reserve','context_window','estimated_excess_tokens'].every(key => Number.isFinite(value[key]));
    budgetMessage.textContent = usable ? '输入估算 ' + value.estimated_input_tokens + ' tokens，回答预留 ' + value.reply_token_reserve + ' tokens；配置窗口 ' + value.context_window + ' tokens，预计超出 ' + value.estimated_excess_tokens + ' tokens。输入已保留，没有请求模型；请先缩短输入或核对配置，调整配置不会自动发送。' : '输入已超出上下文预算，没有请求模型；请缩短输入或核对配置后重新预览。';
  }
  async function previewContext() {
    if (!authenticated() || mode !== 'chat' || previewPending || submitting) return;
    const born = epoch, request = ++previewSequence, question = input.value, session = selected;
    const previousNotice = noticeFeedback, previousBudget = budgetFeedbackToken;
    const feedback = window.OfficeRequestFeedback, previousStatus = feedback?.snapshot();
    previewPending = true; previewButton.disabled = true;
    try {
      const result = await host().preview_chat_context(question, continuation ? {continuation} : undefined);
      if (born !== epoch || request !== previewSequence || session !== selected || mode !== 'chat' || question !== input.value) return;
      contextContent.replaceChildren(); contextPanel.hidden = false; contextPanel.open = true;
      if (!result?.ok) {
        contextPreview = null; make('p', '', error(result), contextContent);
        setNotice(error(result), result?.reason === 'model_context_budget' ? 'chat-budget' : null);
        if (result?.reason === 'model_context_budget') showBudget(result.budget);
        return;
      }
      const value = result.result;
      if (previousBudget === budgetFeedbackToken) {
        budgetFeedback.hidden = true; budgetMessage.textContent = ''; budgetFeedbackToken = null;
      }
      if (previousNotice && previousNotice === noticeFeedback && previousNotice.source === 'chat-budget') setNotice('当前输入的发送范围已重新核对；尚未调用模型。');
      feedback?.resolveBudget(previousStatus);
      contextPreview = {question, session, revision:value.context_revision};
      if (continuation) make('p','','本次是关联续写：保留原问题和原回答全文及其续写链；无法装入预算时拒绝发送，不丢弃原回答。',contextContent);
      make('p', '', value.route.connection_name + ' · ' + value.route.model_id + '；发送时重新核验此预览。', contextContent);
      make('p', '', '预计发送 ' + value.context_turns + ' 组历史；较早 ' + value.omitted_turns + ' 组不发送；未完成/失败 ' + value.excluded.length + ' 组不发送。', contextContent);
      make('p', '', '输入估算 ' + value.estimated_input_tokens + ' tokens；预留回答 ' + value.reply_token_reserve + ' tokens；配置窗口 ' + value.context_window + ' tokens。', contextContent);
      if (value.output_parameter) make('p', '', '本次输出限制参数：' + value.output_parameter + '=' + value.reply_token_reserve + '。参数由宿主按已保存策略和接口身份解析；服务端实际遵守情况仍需调用核对，不自动更换参数或重试。', contextContent);
      if (Number.isFinite(value.ollama_num_ctx)) make('p', '', '本次将请求Ollama上下文 ' + value.ollama_num_ctx + ' tokens（num_ctx=' + value.ollama_num_ctx + '）；未验证服务端实际窗口。较大窗口可能增加本机内存占用，实际由Ollama决定。', contextContent);
      if (value.transfer_required) make('p', '', '历史将发往不同接口或配置版本；发送时仍需单独确认。', contextContent);
      for (const [key, title] of [['included','将发送的历史'], ['omitted','超出预算，未发送'], ['excluded','未完成，未发送']]) {
        if (!value[key].length) continue;
        make('strong', '', title, contextContent);
        const list = make('ol', '', '', contextContent);
        for (const turn of value[key]) make('li', '', turn.question + ' · ' + turn.route.connection_name + ' · ' + turn.route.model_id + (turn.finish_reason !== 'stop' && key === 'included' ? '（原回复完整性待核对）' : ''), list);
      }
      make('small', '', value.scope, contextContent);
    } catch (_) {
      if (born === epoch && request === previewSequence) { contextPreview = null; setNotice('发送范围预览未完成；没有调用模型，可主动重新预览。'); }
    } finally {
      if (born === epoch && request === previewSequence) { previewPending = false; previewButton.disabled = false; }
    }
  }
  function keepDraft() {
    drafts[mode] = input.value;
    if (mode === 'chat' && selected) sessionDrafts.set(selected, input.value);
  }
  input.addEventListener('input', () => { inputRevision++; keepDraft(); clearPreview(); });
  $('model-selector').addEventListener('change', clearPreview);
  window.addEventListener('officesettings:saved', clearPreview);
  function updateMode(next) {
    if (!['chat','agent','image'].includes(next)) return;
    if (next !== mode) { keepDraft(); clearPreview(); continuation = null; retireStream(liveStream); liveStream=null; live.hidden=true; mode = next; input.value = mode === 'chat' && selected ? sessionDrafts.get(selected) || '' : drafts[mode]; }
    panes.classList.toggle('daily-chat', mode === 'chat');
    panel.hidden = mode !== 'chat'; imageParams.hidden = mode !== 'image';
    input.setAttribute('aria-label', mode === 'chat' ? '普通问答输入' : mode === 'image' ? '图片生成提示词' : '自然语言任务指令');
    summary.textContent = mode === 'chat' ? '普通问答：无需文件；输入后发送。新建会话、切换模式和选择文件是独立操作。' : mode === 'image' ? '生图：无需文件；先查看方案并独立确认，生成后再选择保存目录。' : '文档 Agent：先选择文件并同意读取；建议采用不等于写入批准。';
    if (mode === 'chat') refresh();
  }
  function renderChoices() {
    const query = search.value.trim().toLocaleLowerCase();
    picker.replaceChildren();
    for (const item of state?.sessions || []) {
      const route = item.turns.at(-1)?.route;
      const origin = route ? route.connection_name + ' · ' + route.model_id : '尚未调用模型';
      const matches = (item.title + ' ' + origin).toLocaleLowerCase().includes(query);
      if (!matches && item.id !== state.current) continue;
      const option = make('option', '', new Date(item.created * 1000).toLocaleString() + ' · 问答 · ' + item.title + ' · ' + origin + (!matches ? '（当前，未匹配筛选）' : ''), picker);
      option.value = item.id;
    }
    picker.value = state?.current || '';
  }
  function viewSnapshot() {
    return {current:selected, title:state?.sessions.find(item => item.id === selected)?.title || '新会话',
      save_state:state?.save_state, sessions:(state?.sessions || []).map(item => ({id:item.id,title:item.title,created:item.created}))};
  }
  function render() {
    const row = state?.sessions.find(item => item.id === state.current);
    const source = noticeFeedback?.source;
    const finished = !source?.terminal && source?.session_id === selected && row?.turns.find(turn => turn.id === source.turn_id && turn.status !== 'interrupted');
    if (finished) {
      retireStream(finished.id);
      const completed = {stop:'本轮回答已返回；可复制、导出或继续提问。', length:'回复达到输出预算，尚未完整；可核对后主动继续回答。', unknown:'服务未报告结束原因，回复完整性未确认；请核对后决定是否继续。', content_filter:'服务过滤了回复；不提供续写操作，请核对服务规则。', local_limit:'达到本地字符、分块或等待限制；已保留收到内容，可核对后主动续写。'};
      setNotice(finished.status === 'complete' ? completed[finished.finish_reason] : finished.status === 'stopped' ? '已停止本地显示；远端是否结束及费用仍需核对，未自动重试。' : '本轮请求失败；已保留收到内容，请核对错误后主动处理。', {turn_id:finished.id,session_id:selected,terminal:true});
    }
    if (!replySaving && document.activeElement !== length && document.activeElement !== outputTokens) {
      const tokens = state?.reply_policy?.reply_token_reserve || 512;
      length.value = ['512','2048','4096'].includes(String(tokens)) ? String(tokens) : 'custom'; outputTokens.value = String(tokens); outputTokens.hidden = length.value !== 'custom';
    }
    renderChoices();
    if (renameFor !== selected || !renameDirty) { titleInput.value = row?.title || ''; renameFor = selected; renameDirty = false; }
    titleHint.textContent = renameDirty && titleInput.value.trim() !== row?.title ? '名称草稿尚未采用；请点击“重命名”。' : '';
    const limits = state?.capacity;
    capacity.textContent = limits ? '会话 ' + limits.sessions + '/' + limits.max_sessions + '；当前轮次 ' + limits.current_turns + '/' + limits.max_turns + '；历史约 ' + Math.ceil(limits.bytes / 1024) + '/' + Math.floor(limits.max_bytes / 1024) + ' KiB。达到上限后请主动删除不需要的会话或分段处理，不自动清理。' : '';
    const states = {saved:'当前历史已保存到本机。',unsaved:'当前历史有未保存改动。',disabled:'本机保存已关闭；当前新改动仅在本窗口，已有磁盘历史保留。',failed:'本机保存失败，窗口内容保留；请核对错误后主动处理。',uncertain:'保存结果不确定；请重新读取磁盘核对，不要重复保存。'};
    saveStatus.textContent = states[state?.save_state] || '';
    const fused = ['chat_history_uncertain','chat_history_corrupt','chat_history_conflict'].includes(state?.warning);
    $('chat-save-now').disabled = changing || !state?.saving_enabled || fused || state?.save_state === 'saved';
    $('chat-save-now').textContent = state?.warning === 'chat_history_write_failed' ? '重试保存本机历史' : '保存当前历史';
    $('chat-save-now').title = fused ? '请重新读取核对；不会覆盖冲突、损坏或结果不确定的记录。' : !state?.saving_enabled ? '先启用本机保存偏好；本按钮不会改变设置。' : '仅保存当前账户普通问答，不发送模型请求。';
    $('chat-rename').disabled = changing || !selected || fused;
    titleInput.disabled = changing || !selected;
    if (!preferencePending) enabled.checked = Boolean(state?.saving_enabled);
    storage.textContent = (state?.location || '') + '。' + (state?.scope || '') + ' 未加密；保存偏好为本机共享，文字按账户记录；关闭保存不删除已有历史。';
    transcript.replaceChildren();
    for (const turn of row?.turns || []) {
      const article = make('article', '', '', transcript); article.className = 'daily-turn';
      article.dataset.turnId = turn.id;
      const turnNotice = text => { const node = make('p', '', text, article); node.dataset.turnNotice = turn.id; };
      const selectionLabel = make('label','','',article);
      const selectionBox = make('input','','',selectionLabel); selectionBox.type = 'checkbox'; selectionBox.checked = exportSelection.has(turn.id);
      selectionLabel.append(document.createTextNode('选择此轮导出')); selectionBox.onchange = ()=>selectionBox.checked ? exportSelection.add(turn.id) : exportSelection.delete(turn.id);
      if (turn.answer) button('','导出此回答',article,()=>exportTurns([turn.id],'answer'));
      if (turn.status === 'interrupted') button('','导出当前已收到部分（未完成）',article,()=>exportTurns([turn.id],'answer',true));
      make('strong', '', '你', article); make('p', '', turn.question, article).className = 'daily-question';
      const info = make('details', '', '', article); make('summary', '', '本轮实际接口与上下文来源', info);
      make('p', '', turn.route.connection_name + ' · ' + turn.route.model_id + ' · 普通问答 · ' + new Date(turn.created * 1000).toLocaleString() + '；实际发送' + turn.context_turns + '组历史，较早' + turn.omitted_turns + '组不发送。', info);
      if (turn.status === 'complete') {
        if (turn.finish_reason === 'unknown') turnNotice('服务未报告结束原因，回复完整性未确认；请核对后决定是否继续。');
        if (turn.finish_reason === 'length') turnNotice('服务报告输出达到预算，回答未完整；可主动续写。');
        if (turn.finish_reason === 'content_filter') turnNotice('服务过滤了回答；不提供续写操作，请核对服务规则。');
        if (turn.finish_reason === 'local_limit') turnNotice('达到本地字符、分块或等待限制；已保留收到的部分内容。');
        if (turn.continuation_of) make('p','','关联续写 · 原回答 ' + turn.continuation_of.turn_id.slice(0,8) + '（原回答保留，未覆盖）',article);
        if (turn.budget) make('p','','本轮输出预算 ' + turn.budget.reply_token_reserve + ' tokens · ' + turn.budget.output_parameter,info);
        if (turn.usage) make('p','','服务商报告用量：' + JSON.stringify(turn.usage) + '；不等于计费记录。',info);
        make('strong', '', '回答', article); const answer = make('div', '', '', article); answer.className = 'daily-answer';
        if (window.OfficeMarkdown) window.OfficeMarkdown.render(answer,turn.answer); else answer.textContent = turn.answer;
        button('', '复制回答', article, async () => {
          try { await navigator.clipboard.writeText(turn.answer); setNotice('回答已复制。'); }
          catch (_) { setNotice('系统剪贴板不可用，请选择回答文字复制。'); answer.tabIndex = 0; answer.focus(); }
        });
        button('', '以原问题再次提问', article, () => { input.value = turn.question; keepDraft(); input.focus(); setNotice('已填回原问题；发送前可编辑，尚未调用模型。'); });
        const source = state?.continuations?.find(item => item.turn_id === turn.id);
        if (source) button('','继续回答',article,async () => {
          continuation = {session_id:selected,...source};
          input.value = '请继续上一条回答，从中断处接续，保留原问题的要求；不要重复已经完成的内容。'; keepDraft();
          await previewContext(); input.focus();
          setNotice('续写草稿已准备；原问题：' + turn.question.slice(0,120) + '；原回答末段：' + turn.answer.slice(-240) + '。请核对发送范围，再主动发送。');
        });
      } else {
        if (turn.answer) { const partial = make('div','','',article); partial.className = 'daily-answer'; if (window.OfficeMarkdown) window.OfficeMarkdown.render(partial,turn.answer); else partial.textContent = turn.answer; make('small','','以上为未完成的已接收内容。',article); }
        make('p', '', {failed:'请求失败，未自动重试。',stopped:'已停止本地显示，远端状态和费用需核对。',interrupted:'该轮未完成；历史恢复不会继续请求。'}[turn.status] || '等待结果。', article);
        button('', '填回原问题', article, () => { input.value = turn.question; keepDraft(); input.focus(); });
      }
    }
    if (!row?.turns.length) make('p', '', '这是普通问答会话。可以直接发送问题，无需选择文件。', transcript);
    if (state?.warning) setNotice(errors[state.warning] || '本地历史不可用，请核对保存状态。');
    $('chat-new').disabled = changing || limits?.sessions >= limits?.max_sessions;
    $('chat-new').title = limits?.sessions >= limits?.max_sessions ? '已达到32个会话；请主动删除不需要的会话后新建。' : '';
    picker.disabled = changing;
    window.dispatchEvent(new CustomEvent('officedaily:state', {detail:viewSnapshot()}));
  }
  async function refresh() {
    if (changing || !host()?.conversation_snapshot || !authenticated()) return;
    const born = epoch, request = ++readSequence;
    try {
      const result = await host().conversation_snapshot();
      if (born !== epoch || request !== readSequence || !authenticated()) return;
      if (!result?.ok) { setNotice(error(result)); return; }
      const previous = selected;
      state = result.result; selected = state.current;
      if (previous !== selected) { clearPreview(); continuation = null; retireStream(liveStream); liveStream=null; live.hidden=true; exportSelection.clear(); exportResult.replaceChildren(); }
      if (previous !== selected && mode === 'chat') input.value = sessionDrafts.get(selected) || (previous ? '' : drafts.chat);
      render();
    } catch (_) { if (born === epoch) setNotice('无法读取本地会话；没有重新调用模型，请重试读取。'); }
  }
  async function change(method, ...args) {
    if (changing || !authenticated()) return;
    keepDraft(); const retainedDraft = input.value; changing = true; readSequence++; const born = epoch;
    render();
    try {
      const result = await host()[method](...args);
      if (born !== epoch) return;
      if (!result?.ok) {
        if (result?.result) { state = result.result; selected = state.current; sessionDrafts.set(selected, retainedDraft); input.value = retainedDraft; }
        setNotice(error(result)); return;
      }
      state = result.result; selected = state.current; live.hidden = true; streamed = '';
      clearPreview(); continuation = null; retireStream(liveStream); liveStream=null; exportSelection.clear(); exportResult.replaceChildren();
      if (method === 'rename_conversation') renameDirty = false;
      if (['reload_conversations','clear_saved_conversations','delete_conversation'].includes(method)) sessionDrafts.set(selected, retainedDraft);
      input.value = sessionDrafts.get(selected) || ''; render();
      const messages = {rename_conversation:'会话名称已更新。', delete_conversation:'会话已删除，尚未发送的输入保留。', save_conversation_history:'当前历史已保存。'};
      setNotice(state.warning ? errors[state.warning] || '历史保存状态需核对。' : (messages[method] || '已切换问答会话；没有发送请求，也没有恢复文件读取或任何批准。'));
    } catch (_) { if (born === epoch) setNotice('会话操作未确认，请重新读取核对；没有自动重试。'); }
    finally { if (born === epoch) { changing = false; render(); } }
  }
  window.addEventListener('officemodels:mode', event => updateMode(event.detail.mode));
  window.addEventListener('officeauth:authenticated', () => { updateMode(window.OfficeModelSelector?.mode() || 'agent'); refresh(); });
  window.addEventListener('officemodels:host-ready', refresh);
  window.addEventListener('officeauth:unauthenticated', () => {
    epoch++; readSequence++; imageRead++; state = null; selected = null; sending = null; changing = false; transfer = null; preferencePending = false; enabled.disabled = false;
    for (const key of Object.keys(drafts)) drafts[key] = '';
    sessionDrafts.clear(); imageRecords.replaceChildren(); imageHistory.hidden = true; failedPrompt = ''; streamed = ''; input.value = ''; transcript.replaceChildren(); live.hidden = true; retry.hidden = true;
    search.value = ''; titleInput.value = ''; renameDirty = false; renameFor = null; titleHint.textContent = ''; capacity.textContent = ''; saveStatus.textContent = '';
    submitting = null; continuation = null; retireStream(liveStream); liveStream=null; replySaving=exporting=imagePreferencePending=false; exportSelection.clear(); exportResult.replaceChildren(); clearPreview();
  });
  window.addEventListener('officedaily:event', event => {
    const ev = event.detail;
    if (!authenticated() || ev.session_id && ev.session_id !== selected) return;
    if (ev.type === 'chat_changed') refresh();
    if (ev.type === 'answer_draft_start' && mode === 'chat' && !retiredStreams.has(ev.stream_id)) { liveStream=ev.stream_id; streamed = ''; live.textContent = ''; live.hidden = false; setNotice('正在接收流式回答；尚未完成。', {turn_id:ev.stream_id,session_id:selected}); }
    if (ev.type === 'answer_draft_delta' && mode === 'chat' && liveStream===ev.stream_id) { streamed += ev.text; live.append(document.createTextNode(ev.text)); }
    if (ev.type === 'answer_draft_cancel') { retireStream(ev.stream_id); if (liveStream===ev.stream_id) { liveStream=null; live.hidden = true; streamed = ''; } }
    if (ev.type === 'chat_error' && mode === 'chat') {
      if (ev.turn_id && retiredStreams.has(ev.turn_id)) return;
      const detail = {authentication_failed:'鉴权失败，请修复当前 API 密钥。',model_not_found:'模型或接口路径不存在，请核对模型名称和 API 地址。',rate_limited:'接口限流，请稍后主动重试。',redirect_refused:'接口返回重定向，为保护凭据未跟随；请填写最终 API 地址。',service_failed:'模型服务异常，请核对服务状态。',timeout:'等待超时；服务端可能已受理，请先核对再重试。',address_unreachable:'连接地址不可达，请核对地址和网络。',response_incompatible:'响应格式不兼容；可关闭流式选项后主动重试。'};
      setNotice(detail[ev.code] || '请求失败，未自动重试。', {turn_id:ev.turn_id,session_id:selected,terminal:true});
      failedPrompt = ev.prompt; retry.hidden = false;
      if (sending?.question === ev.prompt) sending.failed = true;
      if (!input.value) { input.value = failedPrompt; keepDraft(); }
    }
    if (ev.type === 'busy' && !ev.busy) refresh();
  });
  window.OfficeDaily = {
    viewSnapshot,
    feedbackSnapshot: () => noticeFeedback,
    send: async question => {
      if (submitting) return {ok:false,reason:'busy'};
      const born = epoch, usedMode = mode, usedSession = selected;
      const options = usedMode === 'chat' ? {stream: stream.checked} : usedMode === 'image' ? {size: size.value, style: style.value} : undefined;
      if (usedMode === 'chat' && continuation) {
        if (!contextPreview || contextPreview.question !== question || contextPreview.session !== usedSession) return {ok:false,reason:'chat_context_changed'};
        options.continuation = {...continuation};
      }
      if (usedMode === 'chat' && contextPreview?.question === question && contextPreview.session === usedSession) options.context_revision = contextPreview.revision;
      const submission = {}; submitting = submission;
      const previousNotice = noticeFeedback;
      sending = {mode: usedMode, session: usedSession, question};
      const current = () => born === epoch && usedMode === mode && (usedMode !== 'chat' || usedSession === selected);
      try {
        let result = await host().ask(question, options);
        if (!current()) return {ok:false,reason:'stale'};
        if (result?.reason === 'chat_transfer_required') {
          const destination = result.route;
          if (!await ask('继续发送会把本会话预算内的历史文字发送到：' + destination.connection_name + ' · ' + destination.model_id + '。是否允许本次发送？')) return {ok:false,reason:'cancelled'};
          if (!current()) return {ok:false,reason:'stale'};
          result = await host().ask(question, {...options, transfer_to: destination});
          if (!current()) return {ok:false,reason:'stale'};
        }
        if (result?.ok && usedMode === 'chat') {
          continuation = null;
          budgetFeedback.hidden = true; budgetMessage.textContent = ''; budgetFeedbackToken = null;
          retry.hidden = true;
          if (noticeFeedback === previousNotice || result.turn_id && noticeFeedback?.source?.turn_id === result.turn_id && noticeFeedback.source.session_id === usedSession) setNotice('请求已受理；实际发送' + result.context_turns + '组历史，较早' + result.omitted_turns + '组未发送。', result.turn_id ? {turn_id:result.turn_id,session_id:usedSession} : null);
          await refresh();
          if (!current()) return {ok:false,reason:'stale'};
        } else if (!result?.ok && usedMode === 'chat') {
          setNotice(error(result), result?.reason === 'model_context_budget' ? 'chat-budget' : null);
          if (result?.reason === 'model_context_budget') showBudget(result.budget);
        }
        return result;
      } catch (_) {
        if (current()) setNotice(errors.chat_submission_uncertain);
        return {ok:false,reason:current() ? 'chat_submission_uncertain' : 'stale'};
      } finally {
        if (submitting === submission) submitting = null;
      }
    },
    mayClear: question => sending?.question === question && !sending.failed && sending.mode === mode && (mode !== 'chat' || sending.session === selected),
    sent: question => {
      if (!sending || sending.question !== question) return;
      if (sending.failed) return;
      drafts[sending.mode] = '';
      if (sending.mode === 'chat' && sending.session) sessionDrafts.set(sending.session, '');
      sending = null;
    },
    refresh
  };
  updateMode(window.OfficeModelSelector?.mode() || 'agent');
})();
