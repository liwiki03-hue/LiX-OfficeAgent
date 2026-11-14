(() => {
  'use strict';
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  const $ = id => document.getElementById('doc-task-' + id);
  if (!$('panel')) return;
  let epoch = 0, busy = false, uncertain = false, task = null, plan = null, next = null, navigationOffset = null, returnOffset = null;
  let comparison = null, handoff = null, phase = '读取或任务准备';
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const stateNames = {pending: '待处理', completed: '已执行并验证', no_change: '无需修改', excluded: '用户排除（不计成功）', restricted: '受限未完成', failed: '失败', uncertain: '结果不确定'};
  const errors = {professional_task_filename_budget: '当前文件名无法容纳全部分组副本后缀；请另存为较短名称后重新选择。尚未批准或执行本次写入。', professional_task_needs_specification: '未能无歧义地编译完整要求。精确术语使用示例句式；章节润色先选择章节；组合任务请由Agent生成明确规格。', professional_source_changed: '源或当前副本已改变，旧计划失效；重新选择并核验版本。', professional_task_expired: '任务会话已失效；使用存档恢复，不重放批准。', professional_task_reselect_current: '请选择存档记录中准确的当前DOCX版本，再恢复。', professional_task_scope: '章节或目标不明确，请重新读取大纲并选择。', professional_task_capacity: '任务超过有界容量，请缩小明确范围；没有提高单事务上限。', professional_task_no_preview: '尚无本会话Agent任务方案；先在对话中启用专业工具并请求生成任务规格。', professional_task_no_pending_group: '没有待执行组；请检查完成账本中的受限项或排除项。', professional_task_unresolved: '存在歧义、待确认计划或结果不确定，当前操作已拒绝。', professional_task_changed: '任务状态已变化，请重新预览。', professional_plan_invalid: '计划已失效，请重新预览本组。', professional_task_model_unavailable: '模型方案生成不可用，未写入；精确术语与明确规格仍可使用。', not_authenticated: '请先登录。', busy: '宿主正在处理其他任务。'};
  const status = text => { $('status').textContent = text; };
  Object.assign(errors, {professional_return_expired: '回流核对已过期或被新操作替代；重新同意读取三个版本后再建立候选。', professional_return_target: '该文段不在本次安全候选中；冲突或未知位置不能自动继续。', professional_return_unique_text: '原文字在当前文段中必须恰好出现一次；请补充相邻文字以消除歧义。', professional_return_acknowledgement: '请先核对当前原文及未知项，再明确勾选。', professional_protected_boundary: '选中文字跨越保护标记或既有修订；请选择独立普通文字范围。', professional_return_lineage: '三个版本与回执链不匹配；请核对准确文件和执行顺序。', professional_return_three_versions: '请选择用户另存的当前版本；原稿、交付稿、当前稿必须是三个不同文件。'});
  Object.assign(errors, {professional_profile_version: '保护规则版本不匹配，重新建立任务并审阅规则清单。', professional_profile_capacity: '识别的保护项超过30种，请缩小明确范围；不会静默丢弃保护项。', professional_profile_literal_missing: '额外保留文字须逐项出现在本次目标原文中，且每项不超过200字符。', professional_profile_changed_quantity: '候选改变了已识别数值及单位组合；保留原量值或回到独立事实核验。', professional_profile_changed_condition: '候选改变了已识别适用条件；请保留原条件重新提出候选。', professional_profile_changed_qualifier: '候选改变了否定、限定或义务字面项；请保留后重新提出候选。', professional_protected_literal: '候选改变了明确保留的文字，请核对本任务保护清单。'});
  errors.professional_profile_operation = '技术报告方案首版只支持局部文字、格式和批注；结构操作请独立规划，不继承此方案。';
  const protectionInput = () => ({editing_profile: $('profile').value, protected_literals: $('profile-literals').value.split(/\r?\n/).map(value => value.trim()).filter(Boolean)});
  function clearComparison() {
    comparison = null; $('return-target').replaceChildren();
    ['return-find', 'return-replacement', 'return-reason'].forEach(id => { $(id).value = ''; });
    $('return-ack').checked = false; $('return-mode').value = 'clean';
  }
  function refresh() {
    $('panel').setAttribute('aria-busy', String(busy));
    document.querySelectorAll('[id^="doc-task-"], #doc-task-ledger button, #doc-task-deliveries button').forEach(node => { if (node.matches('button,input,select,textarea')) node.disabled = busy || uncertain || !authenticated(); });
    $('plan').disabled ||= !task || task.group_index >= task.groups_total;
    $('route').disabled ||= !task || task.group_index >= task.groups_total;
    $('confirm').disabled ||= !plan;
    $('undo').disabled ||= !task?.receipts?.length || task?.historical_receipts_only;
    $('save').disabled ||= !task || Boolean(plan);
    $('next').disabled ||= next === null;
    $('return-next').disabled ||= returnOffset === null;
    $('navigate-next').disabled ||= navigationOffset === null;
    const selectedProtection = protectionInput(), dedicatedProtection = selectedProtection.editing_profile !== 'general' || selectedProtection.protected_literals.length > 0;
    ['polish', 'block-create', 'import'].forEach(id => { $(id).disabled ||= dedicatedProtection; $(id).title = dedicatedProtection ? '当前保护选择用于精确术语或回流候选；该入口不继承此选择，请使用显式任务规格。' : ''; });
    $('prepare-return').disabled ||= !task?.return_context || Boolean(plan);
    $('return-target').disabled ||= !comparison;
    $('return-create').disabled ||= !comparison || !$('return-target').value || !$('return-find').value || !$('return-replacement').value || !$('return-reason').value.trim() || !$('return-ack').checked;
    $('workflow').textContent = uncertain ? '已停止：结果不确定。先核对审计，不自动重写。' : busy ? '正在处理当前阶段，请等待；不会重复提交。' : plan ? '下一步：审阅本组完整差异；确认才会写副本，取消不写入。' : task ? task.deterministic_complete ? `已生成候选：${task.current_name}。下一步：审阅事务页图；需要Word继续修改时先准备回流交付链。语义与真实Word仍待核对。` : `当前候选：${task.current_name}。下一步：核对待处理和受限项，预览下一组；已执行${task.receipts.length}个独立事务。` : comparison ? '下一步：核对当前原文，选择安全文段建立新候选；冲突项保留双方，不自动合并。' : handoff ? '下一步：选择用户另存的当前DOCX，再主动同意三版本读取；草稿不恢复旧批准。' : '下一步：选择准确的DOCX，输入要求并建立任务；也可重新准入已有回流版本。';
  }
  function reset(clearHandoff = true) {
    window.dispatchEvent(new CustomEvent('officeworkspace:clear'));
    epoch++; busy = uncertain = false; task = plan = null; next = returnOffset = navigationOffset = null;
    ['request', 'spec', 'receipt', 'query', 'block-texts', 'block-end'].forEach(id => { $(id).value = ''; });
    ['summary', 'deliveries', 'navigation', 'ledger', 'preview', 'visual', 'section', 'block-first', 'block-last', 'block-destination', 'block-bookmark'].forEach(id => $(id).replaceChildren());
    $('block-ack').checked = false; $('block-start').value = '0'; $('block-inheritance').value = 'reject_mixed';
    $('profile').value = 'general'; $('profile-literals').value = ''; $('protection').textContent = ''; $('return-tables').checked = false;
    ['return-baseline', 'return-delivered', 'return-receipts'].forEach(id => { $(id).value = ''; }); $('return-output').textContent = '';
    clearComparison();
    if (clearHandoff) handoff = null;
    if (handoff) fillHandoff();
    status('会话已清理；旧任务批准不恢复。'); refresh();
  }
  function fillHandoff() {
    $('return-baseline').value = handoff.baseline_name; $('return-delivered').value = handoff.delivered_name;
    $('return-receipts').value = handoff.receipts.join(', '); $('return-panel').open = true;
  }
  $('prepare-return').onclick = async () => {
    if (busy || uncertain || !authenticated() || !task?.return_context) return;
    if (['return-baseline', 'return-delivered', 'return-receipts', 'return-find', 'return-replacement', 'return-reason'].some(id => $(id).value) && !await ask('用当前已知交付链替换回流草稿？未保存的回流输入将清除，不写磁盘。')) return;
    const directory = document.getElementById('file-dir')?.textContent.trim();
    handoff = {...task.return_context, receipts: [...task.return_context.receipts], directory};
    clearComparison(); returnOffset = null; $('return-output').textContent = ''; fillHandoff();
    $('return-panel').scrollIntoView({block: 'start'}); $('return-baseline').focus();
    status('交付链已填入，没有读取或写入。请用文件选择器重新选择用户另存版本；同目录草稿可保留，旧批准不保留。'); refresh();
  };
  function fail(reply) { status('阶段：' + phase + '。' + (errors[reply?.reason] || '操作被拒绝：' + (reply?.reason || '宿主不可用') + '。未自动重试。')); }
  async function run(write, operation, stage = '读取或任务准备') {
    if (busy || uncertain || !authenticated()) return;
    const mine = epoch, current = () => epoch === mine && authenticated(); phase = stage; busy = true; refresh();
    try { await operation(window.OfficeAuth.wrapApi(window.pywebview.api), current); }
    catch (_) { if (current()) { if (write) { uncertain = true; plan = null; status('阶段：' + phase + '。结果不确定，可能已有副本或存档；请核对审计，不自动重复写入。'); } else status('阶段：' + phase + '。读取或计划失败，没有新的写入批准。重新读取或重新规划不等于重新写入。'); } }
    finally { if (epoch === mine) { busy = false; refresh(); } }
  }
  function show(value, keepPlan = false) {
    window.dispatchEvent(new CustomEvent('officeworkspace:task', {detail: {source_name: value.source_name, current_name: value.current_name, receipts: value.receipts}}));
    task = value; next = value.next_offset; if (!keepPlan) { plan = null; $('preview').textContent = ''; }
    $('summary').textContent = `当前版本：${value.current_name}\n组进度：${value.group_index}/${value.groups_total}；各组项数：${value.group_sizes.join('、')}\n状态：${Object.entries(value.states).map(([key, count]) => `${stateNames[key] || key} ${count}`).join('；')}\n确定性检查：${value.deterministic_complete ? '支持范围已执行且检查通过；语义/Word仍需审阅' : '尚未完整完成，请核对待处理、受限、排除或检查残留'}\n歧义：${value.ambiguities.join('；') || '未登记'}\n扫描：${value.coverage.map(row => `${row.requirement}：${row.scanned}个目标，${row.complete ? '该登记范围扫描完成' : '未完成'}`).join('；')}\n漏改检查：${JSON.stringify(value.verification_checks)}\n每组独立确认和回滚；中间副本不等于最终稿。恢复不恢复旧批准。`;
    $('summary').textContent += `\n实际已完成独立事务：${value.receipts.length}。以上组进度为原选区回退分组；推荐路线可能跨组完成，请以逐项账本和残留检查为准。`;
    $('summary').textContent += `\n交付模式：${value.mode === 'tracked' ? '原生文字修订；仅本轮文字范围，无关历史修订保持' : value.mode === 'clean' ? '净稿局部修改；不接受或拒绝其他作者修订' : '未报告，请核对计划'}。`;
    const protection = value.protection_profile || {profile: 'general', items: [], manual_literals: []};
    $('protection').textContent = `本任务保护方案：${protection.profile === 'technical_report_v1' ? '技术报告字面保护v1' : '通用保护'}\n规则摘要：${protection.rule_sha256 || '原有规则'}\n已识别保护：${(protection.items || []).map(row => `${row.kind} · ${row.literal} × ${row.count}`).join('；') || '未登记；不表示没有风险'}\n额外保留：${(protection.manual_literals || []).join('；') || '未指定'}\n仅字面保护；语义、单位等价及事实仍需核对。`;
    $('deliveries').replaceChildren();
    value.receipts.forEach((receipt, index) => {
      const button = document.createElement('button'); button.type = 'button'; button.className = 'btn secondary'; button.textContent = `审阅第${index + 1}次事务`; button.dataset.receipt = receipt;
      button.onclick = () => {
        if (!authenticated() || busy || uncertain) return;
        if (document.getElementById('range-panel').getAttribute('aria-busy') === 'true') { status('页图正在读取，请完成本次读取后再切换事务。'); return; }
        const input = document.getElementById('range-delivery-id'); input.value = receipt; input.dispatchEvent(new Event('input', {bubbles: true}));
        window.dispatchEvent(new CustomEvent('officeworkspace:review'));
        const panel = document.getElementById('range-panel'); panel.open = true; panel.scrollIntoView({block: 'start'}); input.focus();
      };
      $('deliveries').append(button);
    });
    $('ledger').replaceChildren();
    for (const item of value.items) {
      const row = document.createElement('div'), text = document.createElement('pre');
      row.dataset.reviewState = {pending: 'pending', completed: 'completed', restricted: 'unsupported', uncertain: 'needs_review', failed: 'needs_review'}[item.state] || item.state;
      text.textContent = `${stateNames[item.state] || item.state} · ${item.requirement} · ${item.location || '无目标'}\n${item.before || ''} → ${item.after || ''}\n${item.reason || ''}`; row.append(text);
      if (['pending', 'restricted'].includes(item.state)) {
        const button = document.createElement('button'); button.type = 'button'; button.className = 'btn secondary'; button.textContent = '排除并记录原因';
        button.onclick = () => run(false, async (api, current) => { const reason = window.prompt('排除原因；排除项不计为完成'); if (!reason?.trim() || !current()) return; const reply = await api.exclude_document_task_item(task.task_id, item.id, reason); if (!current()) return; if (reply?.ok) show(reply.result); else fail(reply); });
        row.append(button);
      }
      $('ledger').append(row);
    }
    refresh();
  }
  const replaceTask = async () => !task || await ask('建立新任务会结束当前会话任务及其撤销批准；已有副本和审计保留。需要继续恢复时请先保存任务。是否继续？');
  function navigate(offset) { return run(false, async (api, current) => {
    const reply = await api.navigate_document_task({query: $('query').value, offset, limit: 20});
    if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    const data = reply.result; navigationOffset = data.next_offset; $('navigation').textContent = `登记目标 ${data.registered_targets}；匹配 ${data.matching_targets}；本页 ${data.items.length}；后续偏移 ${data.next_offset ?? '无'}\n${data.items.map(row => `${row.location} · ${row.excerpt}`).join('\n')}\n${data.coverage}`;
    $('section').replaceChildren(); const empty = document.createElement('option'); empty.value = ''; empty.textContent = '请选择明确章节'; $('section').append(empty);
    for (const row of data.outline) { const option = document.createElement('option'); option.value = row.locator_id; option.textContent = `${row.level}级 · ${row.title}`; $('section').append(option); }
    $('block-destination').replaceChildren();
    for (const row of data.outline) { const option = document.createElement('option'); option.value = row.locator_id; option.textContent = row.title; $('block-destination').append(option); }
    for (const id of ['block-first', 'block-last']) { $(id).replaceChildren(); for (const row of data.items) { const option = document.createElement('option'); option.value = row.locator_id; option.textContent = row.location + ' · ' + row.excerpt.slice(0, 80); $(id).append(option); } }
    $('block-bookmark').replaceChildren(); for (const row of data.reference_targets || []) { const option = document.createElement('option'); option.value = JSON.stringify(row); option.textContent = row.name + ' · ' + row.location; $('block-bookmark').append(option); }
    status(data.outline_truncated ? '大纲已截断，不能将首批视作全部章节；使用Agent有界分页读取。' : '已读取有界大纲和首批精确匹配；扫描覆盖不代表所有结构可写。');
  }); }
  $('navigate').onclick = () => navigate(0);
  $('navigate-next').onclick = () => navigate(navigationOffset);
  $('query').addEventListener('input', () => { navigationOffset = null; refresh(); });
  $('block-create').onclick = () => run(false, async (api, current) => {
    if (!await replaceTask() || !current()) return;
    const action = $('block-action').value;
    let bookmark = null; if (action === 'bookmark_text') { try { bookmark = JSON.parse($('block-bookmark').value); } catch (_) { status('未找到已登记的受限书签目标。'); return; } }
    const payload = {action, first: bookmark?.locator_id || (action === 'chapter_move' ? $('section').value : $('block-first').value), last: $('block-last').value,
      destination: action === 'chapter_move' ? $('block-destination').value : null, placement: $('block-placement').value,
      new_texts: action === 'chapter_move' ? [] : $('block-texts').value.split(/\n---\n/),
      start: Number($('block-start').value || 0), end: $('block-end').value ? Number($('block-end').value) : null,
      inheritance: $('block-inheritance').value, bookmark_name: bookmark?.name || null, acknowledge_stale_fields: $('block-ack').checked};
    const reply = await api.create_docx_block_task(payload); if (!current()) return; if (reply?.ok) { show(reply.result); status('结构任务已建立；先预览并独立确认，不提供结构原生修订。'); } else fail(reply);
  });
  $('create').onclick = () => run(false, async (api, current) => { if (!await replaceTask() || !current()) return; const protection = protectionInput(); const payload = protection.editing_profile === 'general' && !protection.protected_literals.length ? $('request').value : {request: $('request').value, ...protection}; const reply = await api.create_document_task(payload); if (!current()) return; if (reply?.ok) { show(reply.result); status('整体范围与分组已生成，尚未批准写入。'); } else fail(reply); });
  $('create-spec').onclick = () => run(false, async (api, current) => { if (!await replaceTask() || !current()) return; let spec; try { spec = JSON.parse($('spec').value); } catch (_) { status('任务规格不是有效JSON。'); return; } const reply = await api.create_document_task(spec); if (!current()) return; if (reply?.ok) show(reply.result); else fail(reply); });
  $('polish').onclick = () => run(false, async (api, current) => { if (!$('section').value) { status('先读取大纲并明确选择章节。'); return; } if (!await replaceTask() || !await ask('主动调用已配置模型生成所选章节候选？最多24段、4次调用；可能产生费用。内容与条件仍需人工核对，不自动写入。') || !current()) return; const reply = await api.propose_document_task_polish($('request').value, $('section').value); if (!current()) return; if (reply?.ok) { show(reply.result); status('模型候选已生成；遗漏、受限和未处理范围见账本，不作为语义正确证明。'); } else fail(reply); });
  $('import').onclick = () => run(false, async (api, current) => { if (!await replaceTask() || !current()) return; const reply = await api.import_document_task_preview(); if (!current()) return; if (reply?.ok) show(reply.result); else fail(reply); });
  $('next').onclick = () => run(false, async (api, current) => { const reply = await api.get_document_task(task.task_id, next); if (!current()) return; if (reply?.ok) show(reply.result, true); else fail(reply); });
  function previewRoute(route) { return run(false, async (api, current) => {
    const reply = await api.plan_document_task_group(task.task_id, route); if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    plan = reply.result; const before = JSON.parse(plan.before).items, after = JSON.parse(plan.after).items;
    window.dispatchEvent(new CustomEvent('officeworkspace:plan', {detail: {before, after, group: plan.group, groups: plan.groups_total, input: plan.input_name, mode: task.mode}}));
    $('preview').textContent = `第${plan.group}/${plan.groups_total}组；输入${plan.input_name}\n${after.map((item, i) => `${i + 1}. ${item.action} · ${before[i]?.location}\n原：${before[i]?.text ?? before[i]?.texts?.join('\n¶\n')}\n新：${item.text ?? item.texts?.join('\n¶\n')}\n移动：${item.destination_location || item.destination_locator || '无'} · ${item.placement || '无'}\n格式：${JSON.stringify(item.character_format || item.paragraph_format || {})}\n原因：${item.reason}`).join('\n')}\n引用影响：${JSON.stringify(after[0]?.verification?.reference_impact || {})}\n只确认本组；后续组绑定实际产物版本，可能部分完成。`;
    const routing = plan.routing || {};
    $('preview').textContent = `执行方案：${routing.selected || plan.kind}\n原因：${routing.reason || '原有有界路径'}\n本事务覆盖账本${plan.affected_ledger_items}项；原有选区分组仅供回退进度对照，批量原语可以覆盖多组。\n候选比较：${JSON.stringify(routing.candidates || [])}\n计划页图上下文：${routing.page_alignment_items ?? after.length}/${routing.page_alignment_total ?? after.length}；实际定位以渲染后审阅为准，尚未证明全部唯一。文字/直接格式与非目标部件仍整组核验。\n` + $('preview').textContent;
    status('本组实际变化已预览；接受任务不等于批准本组写入。');
  }, '实际输入版本的计划预览'); }
  $('plan').onclick = () => previewRoute('ranges');
  $('route').onclick = () => previewRoute('auto');
  $('confirm').onclick = () => run(true, async (api, current) => { phase = "本组副本执行"; const target = plan; if (!target) return; const approved = await ask('独立确认本组DOCX副本？取消不写入。\n' + $('preview').textContent) === true; if (!current()) return; plan = null; const reply = await api.confirm_document_task_group(task.task_id, target.id, approved); if (!current()) return; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'uncertain') uncertain = true; show(reply.result.task); if (reply.result.status === 'saved') { const layout = reply.result.layout_observations || {}; $('visual').textContent = '审计回执：' + reply.result.save_id + '\n字符越页/原生空文字页观察：原稿 ' + JSON.stringify(layout.source_counts || {}) + '；副本 ' + JSON.stringify(layout.output_counts || {}) + '\n新增观察：' + (layout.increased_observations?.join('、') || '未登记') + '\n仅原生字符与页面边界检查；图片页、表格越格、遮挡、孤立标题、页眉页脚和页码仍需人工核对。合理分页不自动判失败。'; } status(reply.result.status === 'saved' ? '本组已重开/渲染/审计验证；当前副本：' + reply.result.task.current_name + '。是否完整完成以账本和检查为准。' : reply.result.status === 'cancelled' ? '已取消本组写入。' : '执行结果不确定，请核对审计。'); });
  $('undo').onclick = () => run(true, async (api, current) => { phase = "独立逆序撤销"; if (!await ask('独立撤销最后一组？只能逆序逐个事务撤销，不是整体原子回滚。') || !current()) return; const reply = await api.undo_document_task_group(task.task_id, true); if (!current()) return; if (reply?.ok) { show(reply.result.task); status('最后一组已撤销或保守拒绝；原文件、备份和审计保留。'); } else fail(reply); });
  $('save').onclick = () => run(true, async (api, current) => { phase = "任务存档"; const planned = await api.plan_document_task_archive(task.task_id); if (!current()) return; if (!planned?.ok) { fail(planned); return; } const approved = await ask('独立保存任务规格、版本摘要、计划与完成账本？包含待改文字；不保存写入或撤销批准。') === true; if (!current()) return; const reply = await api.confirm_document_task_archive(planned.result.plan_id, approved); if (!current()) return; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'saved') { $('receipt').value = reply.result.save_id; status('任务已存档；恢复前重新选择当前版本：' + task.current_name); } else if (reply.result.status === 'uncertain') { uncertain = true; status('存档结果不确定，不自动重试。'); } else status('已取消任务存档。'); });
  $('restore').onclick = () => run(false, async (api, current) => { phase = "重新准入恢复"; if (!await replaceTask() || !current()) return; const reply = await api.restore_document_task($('receipt').value.trim()); if (!current()) return; if (reply?.ok) { show(reply.result); status('已重新核验当前版本并恢复账本；旧写入和撤销批准未恢复，下一组必须新预览确认。'); } else fail(reply); });
  const returnLabels = {unchanged: '保持', system_suggestion_still_present: '系统建议仍存在', system_revision_still_present: '本系统文字修订仍存在', system_revision_accepted: '本系统文字修订已接受', system_revision_rejected: '本系统文字修订已拒绝', system_revision_partially_resolved: '部分修订已处理', revision_projection_ambiguous: '修订处理含歧义', structure_changed_unmapped: '结构变化无法可靠定位', ambiguous_position: '位置含歧义', stored_field_result_changed_unverified: '域缓存变化，尚未证明刷新正确', baseline_restored_resolution_unknown: '恢复基线内容，处理方式未知', format_or_structure_changed: '格式或结构变化', user_content_change: '用户新增文字修改', same_position_conflict: '同位置冲突'};
  const continuationLabels = {comparison_requires_review: '变化或位置尚需人工核对', unmapped_or_unsupported: '未映射或未主动纳入的结构', sensitive_or_oversize: '内容敏感或超过范围上限', table_structure_changed: '三版本表格位置或网格属性变化', table_structure_unsupported: '表格含合并、嵌套、多段或特殊标记，或超过预算'};
  function compareReturn(offset) { return run(false, async (api, current) => {
    if (['return-find', 'return-replacement', 'return-reason'].some(id => $(id).value) && !await ask('重新读取比较会放弃尚未建立候选的文字输入；不写磁盘。是否继续？')) return;
    if (offset === 0 && !await ask('同意读取当前已选版本及下列两个同目录DOCX进行只读比较？不合并、不恢复旧批准。') || !current()) return;
    clearComparison(); returnOffset = null; $('return-output').textContent = '';
    const reply = await api.compare_document_return({baseline_name: $('return-baseline').value.trim(), delivered_name: $('return-delivered').value.trim(), receipts: $('return-receipts').value.split(/[,，\s]+/).filter(Boolean), approved_read: true, offset, include_plain_tables: $('return-tables').checked});
    if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    const data = reply.result; returnOffset = data.next_offset;
    comparison = data.comparison_id;
    const placeholder = document.createElement('option'); placeholder.value = ''; placeholder.textContent = '请选择本页已核验文段；下方仍需填写明确文字'; $('return-target').append(placeholder);
    for (const row of data.items) {
      const option = document.createElement('option'); option.value = row.continuation_id || ''; option.disabled = !row.continuation_id;
      option.textContent = `${row.continuation_family === 'ordinary_table_cell' ? '普通单元格 · ' : ''}${returnLabels[row.classification] || row.classification} · ${row.current_location || '未映射'} · ${row.current}${row.continuation_requires_current_review ? '（须先核对用户当前文字）' : ''}${row.continuation_blocked ? '（受限：' + (continuationLabels[row.continuation_blocked] || '需人工核对') + '）' : ''}`;
      $('return-target').append(option);
    }
    $('return-output').textContent = data.versions.map(row => `${row.role} · ${row.name} · ${row.sha256}`).join('\n') + '\n' + Object.entries(data.classifications).map(([key, count]) => `${returnLabels[key] || key}：${count}`).join('；') + '\n' + data.items.map(row => `${returnLabels[row.classification] || row.classification} · ${row.delivered_location} → ${row.current_location || '未映射'}\n原：${row.baseline}\n交付：${row.delivered}\n当前：${row.current}\n新增修订作者：${row.additional_revision_authors?.join('、') || '未识别'}\n${row.requires_review ? '需人工核对' : '仅支持范围的确定性匹配'}`).join('\n\n') + '\n部件变化：' + JSON.stringify(data.package_changes) + '\n仅覆盖已登记正文及简单直接文字修订；未知结构保留，Word分页和完整语义未验证。继续编辑时对当前选中文件建立新任务。';
    status('三个版本只读比较完成；未恢复旧定位、写入或撤销批准。');
  }, '三版本只读核对'); }
  $('return-compare').onclick = () => compareReturn(0);
  $('return-next').onclick = () => compareReturn(returnOffset);
  $('return-create').onclick = () => run(false, async (api, current) => {
    if (!comparison || !await replaceTask() || !current()) return;
    const reply = await api.create_document_return_continuation({comparison_id: comparison, row_id: $('return-target').value,
      find: $('return-find').value, replacement: $('return-replacement').value, reason: $('return-reason').value.trim(),
      mode: $('return-mode').value, acknowledge_current: $('return-ack').checked, ...protectionInput()});
    if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    clearComparison(); returnOffset = null; show(reply.result);
    status('当前版本的新候选已建立，尚未写入。下一步审阅推荐方案，再独立确认副本；旧批准和撤销权限均未恢复。');
    $('route').scrollIntoView({block: 'center'}); $('route').focus();
  }, '当前版本继续修改候选');
  ['return-target', 'return-find', 'return-replacement', 'return-reason', 'return-mode', 'return-ack'].forEach(id => $(id).addEventListener('input', refresh));
  ['profile', 'profile-literals'].forEach(id => $(id).addEventListener('input', refresh));
  ['return-baseline', 'return-delivered', 'return-receipts'].forEach(id => $(id).addEventListener('input', () => { handoff = null; clearComparison(); returnOffset = null; $('return-output').textContent = ''; refresh(); }));
  $('return-tables').addEventListener('change', () => { clearComparison(); returnOffset = null; $('return-output').textContent = ''; refresh(); });
  ['officeauth:authenticated', 'officeauth:unauthenticated', 'pagehide'].forEach(event => window.addEventListener(event, reset));
  ['file-name', 'file-dir'].forEach(id => { const node = document.getElementById(id); if (node) new MutationObserver(() => {
    const directory = document.getElementById('file-dir')?.textContent.trim();
    reset(!authenticated() || !handoff?.directory || directory !== handoff.directory);
  }).observe(node, {childList: true, characterData: true, subtree: true}); });
  refresh();
})();
