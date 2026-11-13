(() => {
  'use strict';
  const $ = id => document.getElementById('doc-task-' + id);
  if (!$('panel')) return;
  let epoch = 0, busy = false, uncertain = false, task = null, plan = null, next = null, navigationOffset = null, returnOffset = null;
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const stateNames = {pending: '待处理', completed: '已执行并验证', no_change: '无需修改', excluded: '用户排除（不计成功）', restricted: '受限未完成', failed: '失败', uncertain: '结果不确定'};
  const errors = {professional_task_filename_budget: '当前文件名无法容纳全部分组副本后缀；请另存为较短名称后重新选择。尚未批准或执行本次写入。', professional_task_needs_specification: '未能无歧义地编译完整要求。精确术语使用示例句式；章节润色先选择章节；组合任务请由Agent生成明确规格。', professional_source_changed: '源或当前副本已改变，旧计划失效；重新选择并核验版本。', professional_task_expired: '任务会话已失效；使用存档恢复，不重放批准。', professional_task_reselect_current: '请选择存档记录中准确的当前DOCX版本，再恢复。', professional_task_scope: '章节或目标不明确，请重新读取大纲并选择。', professional_task_capacity: '任务超过有界容量，请缩小明确范围；没有提高单事务上限。', professional_task_no_preview: '尚无本会话Agent任务方案；先在对话中启用专业工具并请求生成任务规格。', professional_task_no_pending_group: '没有待执行组；请检查完成账本中的受限项或排除项。', professional_task_unresolved: '存在歧义、待确认计划或结果不确定，当前操作已拒绝。', professional_task_changed: '任务状态已变化，请重新预览。', professional_plan_invalid: '计划已失效，请重新预览本组。', professional_task_model_unavailable: '模型方案生成不可用，未写入；精确术语与明确规格仍可使用。', not_authenticated: '请先登录。', busy: '宿主正在处理其他任务。'};
  const status = text => { $('status').textContent = text; };
  function refresh() {
    $('panel').setAttribute('aria-busy', String(busy));
    $('panel').querySelectorAll('button,input,select,textarea').forEach(node => { node.disabled = busy || uncertain || !authenticated(); });
    $('plan').disabled ||= !task || task.group_index >= task.groups_total;
    $('route').disabled ||= !task || task.group_index >= task.groups_total;
    $('confirm').disabled ||= !plan;
    $('undo').disabled ||= !task?.receipts?.length || task?.historical_receipts_only;
    $('save').disabled ||= !task || Boolean(plan);
    $('next').disabled ||= next === null;
    $('return-next').disabled ||= returnOffset === null;
    $('navigate-next').disabled ||= navigationOffset === null;
  }
  function reset() {
    epoch++; busy = uncertain = false; task = plan = null; next = returnOffset = navigationOffset = null;
    ['request', 'spec', 'receipt', 'query', 'block-texts', 'block-end'].forEach(id => { $(id).value = ''; });
    ['summary', 'deliveries', 'navigation', 'ledger', 'preview', 'visual', 'section', 'block-first', 'block-last', 'block-destination', 'block-bookmark'].forEach(id => $(id).replaceChildren());
    $('block-ack').checked = false; $('block-start').value = '0'; $('block-inheritance').value = 'reject_mixed';
    ['return-baseline', 'return-delivered', 'return-receipts'].forEach(id => { $(id).value = ''; }); $('return-output').textContent = '';
    status('会话已清理；旧任务批准不恢复。'); refresh();
  }
  function fail(reply) { status(errors[reply?.reason] || '操作被拒绝：' + (reply?.reason || '宿主不可用') + '。未自动重试。'); }
  async function run(write, operation) {
    if (busy || uncertain || !authenticated()) return;
    const mine = epoch, current = () => epoch === mine && authenticated(); busy = true; refresh();
    try { await operation(window.OfficeAuth.wrapApi(window.pywebview.api), current); }
    catch (_) { if (current()) { if (write) { uncertain = true; plan = null; status('结果不确定，可能已有副本或存档；请核对审计，不自动重复写入。'); } else status('读取或计划失败，没有新的写入批准。'); } }
    finally { if (epoch === mine) { busy = false; refresh(); } }
  }
  function show(value, keepPlan = false) {
    task = value; next = value.next_offset; if (!keepPlan) { plan = null; $('preview').textContent = ''; }
    $('summary').textContent = `当前版本：${value.current_name}\n组进度：${value.group_index}/${value.groups_total}；各组项数：${value.group_sizes.join('、')}\n状态：${Object.entries(value.states).map(([key, count]) => `${stateNames[key] || key} ${count}`).join('；')}\n确定性检查：${value.deterministic_complete ? '支持范围已执行且检查通过；语义/Word仍需审阅' : '尚未完整完成，请核对待处理、受限、排除或检查残留'}\n歧义：${value.ambiguities.join('；') || '未登记'}\n扫描：${value.coverage.map(row => `${row.requirement}：${row.scanned}个目标，${row.complete ? '该登记范围扫描完成' : '未完成'}`).join('；')}\n漏改检查：${JSON.stringify(value.verification_checks)}\n每组独立确认和回滚；中间副本不等于最终稿。恢复不恢复旧批准。`;
    $('summary').textContent += `\n实际已完成独立事务：${value.receipts.length}。以上组进度为原选区回退分组；推荐路线可能跨组完成，请以逐项账本和残留检查为准。`;
    $('deliveries').replaceChildren();
    value.receipts.forEach((receipt, index) => {
      const button = document.createElement('button'); button.type = 'button'; button.className = 'btn secondary'; button.textContent = `审阅第${index + 1}次事务`; button.dataset.receipt = receipt;
      button.onclick = () => {
        if (!authenticated() || busy || uncertain) return;
        if (document.getElementById('range-panel').getAttribute('aria-busy') === 'true') { status('页图正在读取，请完成本次读取后再切换事务。'); return; }
        const input = document.getElementById('range-delivery-id'); input.value = receipt; input.dispatchEvent(new Event('input', {bubbles: true}));
        const panel = document.getElementById('range-panel'); panel.open = true; panel.scrollIntoView({block: 'start'}); input.focus();
      };
      $('deliveries').append(button);
    });
    $('ledger').replaceChildren();
    for (const item of value.items) {
      const row = document.createElement('div'), text = document.createElement('pre');
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
  const replaceTask = () => !task || window.confirm('建立新任务会结束当前会话任务及其撤销批准；已有副本和审计保留。需要继续恢复时请先保存任务。是否继续？');
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
    if (!replaceTask() || !current()) return;
    const action = $('block-action').value;
    let bookmark = null; if (action === 'bookmark_text') { try { bookmark = JSON.parse($('block-bookmark').value); } catch (_) { status('未找到已登记的受限书签目标。'); return; } }
    const payload = {action, first: bookmark?.locator_id || (action === 'chapter_move' ? $('section').value : $('block-first').value), last: $('block-last').value,
      destination: action === 'chapter_move' ? $('block-destination').value : null, placement: $('block-placement').value,
      new_texts: action === 'chapter_move' ? [] : $('block-texts').value.split(/\n---\n/),
      start: Number($('block-start').value || 0), end: $('block-end').value ? Number($('block-end').value) : null,
      inheritance: $('block-inheritance').value, bookmark_name: bookmark?.name || null, acknowledge_stale_fields: $('block-ack').checked};
    const reply = await api.create_docx_block_task(payload); if (!current()) return; if (reply?.ok) { show(reply.result); status('结构任务已建立；先预览并独立确认，不提供结构原生修订。'); } else fail(reply);
  });
  $('create').onclick = () => run(false, async (api, current) => { if (!replaceTask() || !current()) return; const reply = await api.create_document_task($('request').value); if (!current()) return; if (reply?.ok) { show(reply.result); status('整体范围与分组已生成，尚未批准写入。'); } else fail(reply); });
  $('create-spec').onclick = () => run(false, async (api, current) => { if (!replaceTask() || !current()) return; let spec; try { spec = JSON.parse($('spec').value); } catch (_) { status('任务规格不是有效JSON。'); return; } const reply = await api.create_document_task(spec); if (!current()) return; if (reply?.ok) show(reply.result); else fail(reply); });
  $('polish').onclick = () => run(false, async (api, current) => { if (!$('section').value) { status('先读取大纲并明确选择章节。'); return; } if (!replaceTask() || !window.confirm('主动调用已配置模型生成所选章节候选？最多24段、4次调用；可能产生费用。内容与条件仍需人工核对，不自动写入。') || !current()) return; const reply = await api.propose_document_task_polish($('request').value, $('section').value); if (!current()) return; if (reply?.ok) { show(reply.result); status('模型候选已生成；遗漏、受限和未处理范围见账本，不作为语义正确证明。'); } else fail(reply); });
  $('import').onclick = () => run(false, async (api, current) => { if (!replaceTask() || !current()) return; const reply = await api.import_document_task_preview(); if (!current()) return; if (reply?.ok) show(reply.result); else fail(reply); });
  $('next').onclick = () => run(false, async (api, current) => { const reply = await api.get_document_task(task.task_id, next); if (!current()) return; if (reply?.ok) show(reply.result, true); else fail(reply); });
  function previewRoute(route) { return run(false, async (api, current) => {
    const reply = await api.plan_document_task_group(task.task_id, route); if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    plan = reply.result; const before = JSON.parse(plan.before).items, after = JSON.parse(plan.after).items;
    $('preview').textContent = `第${plan.group}/${plan.groups_total}组；输入${plan.input_name}\n${after.map((item, i) => `${i + 1}. ${item.action} · ${before[i]?.location}\n原：${before[i]?.text ?? before[i]?.texts?.join('\n¶\n')}\n新：${item.text ?? item.texts?.join('\n¶\n')}\n移动：${item.destination_location || item.destination_locator || '无'} · ${item.placement || '无'}\n格式：${JSON.stringify(item.character_format || item.paragraph_format || {})}\n原因：${item.reason}`).join('\n')}\n引用影响：${JSON.stringify(after[0]?.verification?.reference_impact || {})}\n只确认本组；后续组绑定实际产物版本，可能部分完成。`;
    const routing = plan.routing || {};
    $('preview').textContent = `执行方案：${routing.selected || plan.kind}\n原因：${routing.reason || '原有有界路径'}\n本事务覆盖账本${plan.affected_ledger_items}项；原有选区分组仅供回退进度对照，批量原语可以覆盖多组。\n候选比较：${JSON.stringify(routing.candidates || [])}\n计划页图上下文：${routing.page_alignment_items ?? after.length}/${routing.page_alignment_total ?? after.length}；实际定位以渲染后审阅为准，尚未证明全部唯一。文字/直接格式与非目标部件仍整组核验。\n` + $('preview').textContent;
    status('本组实际变化已预览；接受任务不等于批准本组写入。');
  }); }
  $('plan').onclick = () => previewRoute('ranges');
  $('route').onclick = () => previewRoute('auto');
  $('confirm').onclick = () => run(true, async (api, current) => { const target = plan; if (!target) return; const approved = window.confirm('独立确认本组DOCX副本？取消不写入。\n' + $('preview').textContent) === true; if (!current()) return; plan = null; const reply = await api.confirm_document_task_group(task.task_id, target.id, approved); if (!current()) return; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'uncertain') uncertain = true; show(reply.result.task); if (reply.result.status === 'saved') { const layout = reply.result.layout_observations || {}; $('visual').textContent = '审计回执：' + reply.result.save_id + '\n字符越页/原生空文字页观察：原稿 ' + JSON.stringify(layout.source_counts || {}) + '；副本 ' + JSON.stringify(layout.output_counts || {}) + '\n新增观察：' + (layout.increased_observations?.join('、') || '未登记') + '\n仅原生字符与页面边界检查；图片页、表格越格、遮挡、孤立标题、页眉页脚和页码仍需人工核对。合理分页不自动判失败。'; } status(reply.result.status === 'saved' ? '本组已重开/渲染/审计验证；当前副本：' + reply.result.task.current_name + '。是否完整完成以账本和检查为准。' : reply.result.status === 'cancelled' ? '已取消本组写入。' : '执行结果不确定，请核对审计。'); });
  $('undo').onclick = () => run(true, async (api, current) => { if (!window.confirm('独立撤销最后一组？只能逆序逐个事务撤销，不是整体原子回滚。') || !current()) return; const reply = await api.undo_document_task_group(task.task_id, true); if (!current()) return; if (reply?.ok) { show(reply.result.task); status('最后一组已撤销或保守拒绝；原文件、备份和审计保留。'); } else fail(reply); });
  $('save').onclick = () => run(true, async (api, current) => { const planned = await api.plan_document_task_archive(task.task_id); if (!current()) return; if (!planned?.ok) { fail(planned); return; } const approved = window.confirm('独立保存任务规格、版本摘要、计划与完成账本？包含待改文字；不保存写入或撤销批准。') === true; if (!current()) return; const reply = await api.confirm_document_task_archive(planned.result.plan_id, approved); if (!current()) return; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'saved') { $('receipt').value = reply.result.save_id; status('任务已存档；恢复前重新选择当前版本：' + task.current_name); } else if (reply.result.status === 'uncertain') { uncertain = true; status('存档结果不确定，不自动重试。'); } else status('已取消任务存档。'); });
  $('restore').onclick = () => run(false, async (api, current) => { if (!replaceTask() || !current()) return; const reply = await api.restore_document_task($('receipt').value.trim()); if (!current()) return; if (reply?.ok) { show(reply.result); status('已重新核验当前版本并恢复账本；旧写入和撤销批准未恢复，下一组必须新预览确认。'); } else fail(reply); });
  const returnLabels = {unchanged: '保持', system_suggestion_still_present: '系统建议仍存在', system_revision_still_present: '本系统文字修订仍存在', system_revision_accepted: '本系统文字修订已接受', system_revision_rejected: '本系统文字修订已拒绝', system_revision_partially_resolved: '部分修订已处理', revision_projection_ambiguous: '修订处理含歧义', structure_changed_unmapped: '结构变化无法可靠定位', ambiguous_position: '位置含歧义', stored_field_result_changed_unverified: '域缓存变化，尚未证明刷新正确', baseline_restored_resolution_unknown: '恢复基线内容，处理方式未知', format_or_structure_changed: '格式或结构变化', user_content_change: '用户新增文字修改', same_position_conflict: '同位置冲突'};
  function compareReturn(offset) { return run(false, async (api, current) => {
    if (offset === 0 && !window.confirm('同意读取当前已选版本及下列两个同目录DOCX进行只读比较？不合并、不恢复旧批准。') || !current()) return;
    const reply = await api.compare_document_return({baseline_name: $('return-baseline').value.trim(), delivered_name: $('return-delivered').value.trim(), receipts: $('return-receipts').value.split(/[,，\s]+/).filter(Boolean), approved_read: true, offset});
    if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    const data = reply.result; returnOffset = data.next_offset;
    $('return-output').textContent = data.versions.map(row => `${row.role} · ${row.name} · ${row.sha256}`).join('\n') + '\n' + Object.entries(data.classifications).map(([key, count]) => `${returnLabels[key] || key}：${count}`).join('；') + '\n' + data.items.map(row => `${returnLabels[row.classification] || row.classification} · ${row.delivered_location} → ${row.current_location || '未映射'}\n原：${row.baseline}\n交付：${row.delivered}\n当前：${row.current}\n新增修订作者：${row.additional_revision_authors?.join('、') || '未识别'}\n${row.requires_review ? '需人工核对' : '仅支持范围的确定性匹配'}`).join('\n\n') + '\n部件变化：' + JSON.stringify(data.package_changes) + '\n仅覆盖已登记正文及简单直接文字修订；未知结构保留，Word分页和完整语义未验证。继续编辑时对当前选中文件建立新任务。';
    status('三个版本只读比较完成；未恢复旧定位、写入或撤销批准。');
  }); }
  $('return-compare').onclick = () => compareReturn(0);
  $('return-next').onclick = () => compareReturn(returnOffset);
  ['return-baseline', 'return-delivered', 'return-receipts'].forEach(id => $(id).addEventListener('input', () => { returnOffset = null; $('return-output').textContent = ''; refresh(); }));
  ['officeauth:authenticated', 'officeauth:unauthenticated', 'pagehide'].forEach(event => window.addEventListener(event, reset));
  ['file-name', 'file-dir'].forEach(id => { const node = document.getElementById(id); if (node) new MutationObserver(reset).observe(node, {childList: true, characterData: true, subtree: true}); });
  refresh();
})();
