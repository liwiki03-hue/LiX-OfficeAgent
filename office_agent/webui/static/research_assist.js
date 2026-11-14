(() => {
  'use strict';
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  const $ = id => document.getElementById('assist-' + id);
  if (!$('panel')) return;
  let epoch = 0, busy = false, uncertain = false, taskDirty = false, version = null, offset = 0, total = 0, nextOffset = null, history = [], task = null, plan = null, saved = null, taskSaved = null, timer = null;
  const targets = new Map(), auth = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const say = text => { $('status').textContent = text; };
  const stages = {materials_validated: '材料已核对', question_and_query_planning: '规划问题与查询', retrieving_selected_sources: '检索所选来源', original_contexts_read: '已补读原文上下文', model_review: '整理证据与建议', awaiting_human_review: '等待逐项人工审阅', regenerating_one_suggestion: '仅重新生成本项', stopped: '已停止', restored_requires_new_human_review: '已恢复，等待重新人工审阅', comparison_only_requires_fresh_admission: '历史材料仅供比较'};
  const revisionErrors = {professional_revision_capacity: '词级修订超过3000个分词单元或100个插入/删除标记，请缩小段落或关闭词级模式。', professional_revision_run_unsupported: '词级模式仅支持普通文本运行，不支持其中的制表符、换行或域。请改用原有段落修订模式或缩小目标范围。', professional_revision_projection_failed: '修订的接受/拒绝投影不一致，已拒绝整单；请保留原格式并缩小目标范围。'};
  const retrievalPaths = {bounded_capture_reading_with_lexical_priority_not_vector_retrieval: '所选页段按关键词读取，未使用向量模型', approved_saved_retrieval: '已保存且获准的检索证据', lexical_only_no_vector_model: '本地向量模型不可用，已降级为关键词检索', real_local_dense_lexical_rrf: '本机真实嵌入与关键词融合，仅读取所选来源', local_embedding_failed_lexical_fallback: '本地嵌入失败，已降级为关键词检索'};
  const messages = {not_authenticated: '请先登录。', busy: '已有任务进行中，请稍后操作。', professional_source_changed: '源文档或会话已改变，请重新读取。', professional_evidence_changed: '证据来源已变化，请重新取证。', professional_evidence_invalid: '证据已失效，请重新核对纳入状态。', professional_assist_citation_review_required: '新增证据主张须在多引文面板核对身份、支持关系和编号，当前仅允许润色副本。', professional_assist_review_required: '请逐项核对并明确勾选人工审阅。', research_context_budget: '内容超出当前模型上下文预算，请减少段落和证据。', research_model_configuration: '当前模型配置不完整，请检查设置。', research_model_timeout: '模型请求超时，未自动重试。', research_model_service: '模型服务不可用，请检查连接。', research_model_response: '模型结果不符合结构或证据约束，未作为可写建议接受。', professional_protected_numbers: '数字保护校验失败，请恢复原数字及顺序。', professional_plan_expired: '确认计划已过期，请重新预览。'};
  Object.assign(messages, revisionErrors);
  Object.assign(messages, {professional_assist_response: '模型引用了未知来源或目标，结果已拒绝；请缩小任务后主动重做。', research_model_context_budget: '所选材料超出上下文预算，请减少段落或来源。', research_model_output_budget: '模型输出超出预算，未自动重试；请缩小任务。', professional_assist_budget_exhausted: '本次模型调用预算已用完，请保留草稿后开始新任务。', professional_assist_budget_expired: '本次计算时间预算已用完，未继续调用模型。', professional_ocr_review_required: 'OCR 文本尚未对照原页核对，请先完成人工校正。', professional_existing_revisions: '文档已有修订，当前原生修订输出已拒绝；请先在 Word 处理既有修订。', insufficient_materials: '所选证据不足，已停止，未编造主张。', user_cancelled: '已取消，迟到结果不会恢复本次建议。', research_retrieval_cleanup_failed: '本地推理进程未能确认清理，已停止；请核对本机进程后重启。'});
  function refresh() {
    const blocked = busy || uncertain || !auth(), running = task?.status === 'running';
    $('panel').setAttribute('aria-busy', String(busy || running));
    $('panel').querySelectorAll('input,select,textarea,button').forEach(node => { node.disabled = blocked || running; });
    $('panel').querySelectorAll('[data-exhausted=true]').forEach(node => { node.disabled = true; });
    if ($('granular')) $('granular').disabled ||= !$('tracked').checked;
    $('cancel').disabled = blocked || !task || task.status === 'cancelled';
    $('prev').disabled = blocked || running || offset === 0;
    $('next').disabled = blocked || running || nextOffset === null;
    $('start').disabled = blocked || running || !version;
    $('plan').disabled = blocked || running || taskDirty || task?.status !== 'completed' || task.request.mode !== 'polish';
    $('confirm').disabled = blocked || running || !plan;
    $('undo').disabled = blocked || running || !saved;
    $('save-task').disabled = blocked || running || taskDirty || task?.status !== 'completed';
    $('undo-task').disabled = blocked || running || !taskSaved;
  }
  function changed(context = false) {
    plan = null; $('preview').textContent = '';
    if (context === true && task) { taskDirty = true; $('suggestions').querySelectorAll('input[value=review]').forEach(node => { node.checked = false; }); say('问题、材料或保护要求已改变。旧建议仅供比较，请重新生成后再审阅。'); }
    refresh();
  }
  function reset() {
    epoch++; clearTimeout(timer); busy = uncertain = taskDirty = false; version = task = plan = saved = taskSaved = null; offset = total = 0; nextOffset = null; history = []; targets.clear(); preserved = null; regenerating = null;
    ['targets', 'captures', 'cards', 'suggestions', 'preview', 'output', 'task-output'].forEach(id => $(id).replaceChildren());
    ['question', 'terms', 'receipt', 'task-receipt'].forEach(id => { $(id).value = ''; });
    $('tracked').checked = false;
    if ($('granular')) $('granular').checked = false;
    $('max-calls').value = '3'; $('compute-budget').value = '60';
    say('尚未读取；打开面板不调用模型。'); refresh();
  }
  function fail(reply) {
    if (reply?.reason === 'professional_uncertain') { uncertain = true; plan = saved = null; say('结果不确定，可能已存在副本或备份。请核对记录，不自动重试。'); }
    else say(messages[reply?.reason] || '操作被拒绝或结果不符合校验，请核对源文件与证据状态。');
  }
  function confirmRun(message, action) {
    let approved = false;
    return run(async (api, current) => {
      approved = await ask(message) === true;
      if (current()) await action(api, current, approved);
    }, () => approved);
  }
  async function run(action, writing = false) {
    if (busy || uncertain || !auth()) return;
    const mine = epoch; busy = true; refresh();
    const current = () => mine === epoch && auth();
    try { await action(window.OfficeAuth.wrapApi(window.pywebview.api), current); }
    catch (_) { if (current()) { if (typeof writing === 'function' ? writing() : writing) fail({reason: 'professional_uncertain'}); else say('请求未完成，请核对状态；未自动重试。'); } }
    finally { if (mine === epoch) { busy = false; refresh(); } }
  }
  function checkbox(parent, value, text, checked = false) {
    const label = document.createElement('label'), node = document.createElement('input');
    node.type = 'checkbox'; node.value = value; node.checked = checked; label.className = 'assist-choice'; label.append(node, document.createTextNode(text)); parent.append(label);
    return node;
  }
  async function load(position, backward = false) {
    if (task && !await ask('重新读取结构会清空当前建议及未保存审阅，是否继续？')) return;
    run(async (api, current) => {
      if (task) { const cancelled = await api.cancel_research_assist(task.id); if (!current()) return; if (!cancelled?.ok) { fail(cancelled); return; } }
      const reply = await api.read_research_assist(position);
      if (!current()) return;
      if (!reply?.ok) { fail(reply); return; }
      const data = reply.result;
      if (version !== data.document.source_sha256) { targets.clear(); history = []; }
      if (backward) history.pop(); else if (position !== offset) history.push(offset);
      version = data.document.source_sha256; offset = position; total = data.total_targets; nextOffset = data.next_offset; task = null; taskDirty = false; changed();
      $('targets').replaceChildren(); $('captures').replaceChildren(); $('suggestions').replaceChildren(); $('cards').textContent = '';
      for (const row of data.items.filter(row => row.writable && !row.style && /^body:\d+\/paragraph$/.test(row.location))) {
        const input = checkbox($('targets'), row.locator_id, row.location + '：' + row.text, targets.has(row.locator_id));
        input.addEventListener('change', () => { if (input.checked) { if (targets.size >= 12) { input.checked = false; say('最多选择12个段落。'); return; } targets.set(row.locator_id, row); } else targets.delete(row.locator_id); changed(true); });
      }
      for (const row of data.captures) checkbox($('captures'), row.id, row.title + ' · ' + (row.ocr_reviewed === false ? 'OCR未人工核对，不可写入' : row.ocr_reviewed === true ? '独立人工校正的OCR，仍须核对支持关系' : row.coverage?.complete ? '已读完文本层' : '部分页段，未证明全文覆盖')).addEventListener('change', () => changed(true));
      say('已读取第' + (offset + 1) + '项起的结构；保留已选' + targets.size + '个段落。');
    });
  }
  $('load').addEventListener('click', () => load(0)); $('prev').addEventListener('click', () => load(history[history.length - 1] ?? 0, true)); $('next').addEventListener('click', () => load(nextOffset));
  let preserved = null, regenerating = null;
  function render(result) {
    task = result;
    say('阶段：' + (stages[result.stage] || result.stage) + '；模型调用 ' + result.calls + '/' + (result.request.max_calls || 3) + '，轮次 ' + result.rounds + '/' + (result.request.max_calls || 3) + '；' + (messages[result.stop_reason] || result.stop_reason || '进行中') + '；检索路径：' + (retrievalPaths[result.degradation] || result.degradation));
    if (result.status !== 'completed' && result.status !== 'restored') return;
    $('cards').replaceChildren();
    const paragraph = (parent, text) => { const node = document.createElement('p'); node.textContent = text; parent.append(node); return node; };
    const relations = {supports: '支持候选', contradicts: '反对候选', insufficient: '证据不足'};
    for (const card of result.result?.cards || []) {
      const article = document.createElement('article'); article.className = 'assist-evidence-card';
      paragraph(article, (relations[card.relation] || '关系未定') + ' · 模型判断，尚未人工审阅'); paragraph(article, '主张：' + card.claim);
      paragraph(article, '适用条件：' + (card.applicability || '未提供')); paragraph(article, '限制：' + (card.limitation || '未提供'));
      if (card.conflicts.length) paragraph(article, '尚未消解的分歧：' + card.conflicts.join('；'));
      for (const id of card.evidence_ids) { const evidence = result.evidence[id]; if (!evidence) continue; paragraph(article, '来源：' + evidence.title + ' · ' + (evidence.page ? 'PDF第' + evidence.page + '页' : evidence.location)); const quote = document.createElement('blockquote'); quote.textContent = evidence.original_text; article.append(quote); }
      $('cards').append(article);
    }
    if (!(result.result?.cards || []).length) paragraph($('cards'), '本次没有新增证据卡片；润色仍需核对原文及保护内容。');
    for (const missing of result.result?.missing_evidence || []) paragraph($('cards'), '待补证：' + missing);
    const details = document.createElement('details'), summary = document.createElement('summary'), metadata = document.createElement('pre'); summary.textContent = '模型版本、预算与检索范围明细'; metadata.textContent = JSON.stringify({model: result.model, retrieval: result.retrieval, evidence_versions: Object.values(result.evidence).map(row => ({source: row.title, sha256: row.source_sha256, page: row.page, quality: row.extraction_quality, coverage: row.coverage}))}, null, 2); details.append(summary, metadata); $('cards').append(details);
    $('suggestions').replaceChildren();
    for (const evidence of Object.values(result.evidence)) {
      if (!evidence.capture_id) continue;
      const view = document.createElement('button'), page = document.createElement('div'); view.type = 'button'; view.className = 'btn secondary'; view.textContent = '查看来源原页：' + evidence.title + ' · ' + evidence.location;
      view.addEventListener('click', () => run(async (api, current) => { await window.OfficeEvidenceViewer.show(api, evidence.capture_id, evidence.id, page, current); }));
      $('suggestions').append(view, page);
    }
    for (const proposal of result.result?.suggestions || []) {
      const section = document.createElement('fieldset'), title = document.createElement('legend'), original = document.createElement('p'), reason = document.createElement('p'), text = document.createElement('textarea');
      const target = result.targets.find(row => row.locator_id === proposal.locator_id);
      section.dataset.locator = proposal.locator_id; title.textContent = '结构项 ' + ((target?.index ?? 0) + 1) + ' · 原文与改稿建议';
      original.textContent = '原文：' + (result.targets.find(row => row.locator_id === proposal.locator_id)?.text || '');
      text.value = proposal.new_text; text.maxLength = 2000; text.rows = 4; text.setAttribute('aria-label', '人工编辑建议');
      reason.textContent = '修改理由：' + proposal.reason + '；保护检查：' + (proposal.protection_status === 'passed' ? '通过，不证明语义等价' : messages[proposal.protection_status] || '需人工核对') + '；风险：' + (proposal.risks.join('；') || '模型未报告额外风险') + '；需补证：' + (proposal.needs_evidence ? '是' : '否');
      section.append(title, original, text, reason);
      const accepted = checkbox(section, 'accept', '采用此项（未勾选即拒绝）'), reviewed = checkbox(section, 'review', '我已核对原文、保护内容及语义；模型判断不是证据证明');
      text.addEventListener('input', () => { reviewed.checked = false; changed(); }); accepted.addEventListener('change', changed); reviewed.addEventListener('change', changed);
      const regenerate = document.createElement('button'), feedback = document.createElement('input'); regenerate.type = 'button'; regenerate.textContent = '仅重新生成此项…'; regenerate.className = 'btn secondary';
      feedback.maxLength = 400; feedback.setAttribute('aria-label', '本项重新生成要求'); feedback.placeholder = '本项调整要求（可选）';
      regenerate.dataset.exhausted = result.calls >= (result.request.max_calls || 3) || result.status === 'restored' ? 'true' : 'false';
      regenerate.addEventListener('click', async () => {
        if (taskDirty) { say('材料或要求已变，请重新开始研究。'); return; }
        if (!await ask('确认向当前模型再次发送本次材料和本项调整要求？占用剩余调用预算，云端可能产生费用。此项人工审阅会清除。')) return;
        const local = new Map([...$('suggestions').querySelectorAll('fieldset')].map(node => [node.dataset.locator, {text: node.querySelector('textarea').value, accepted: node.querySelector('input[value=accept]').checked, reviewed: node.querySelector('input[value=review]').checked}]));
        run(async (api, current) => { changed(); const reply = await api.regenerate_research_suggestion({task_id: task.id, locator_id: proposal.locator_id, feedback: feedback.value, outbound_approved: true}); if (!current()) return; if (!reply?.ok) { fail(reply); return; } preserved = local; regenerating = proposal.locator_id; render(reply.result); if (task.status === 'running') poll(epoch); });
      });
      section.append(feedback, regenerate);
      if (result.request.mode !== 'polish') {
        const transfer = document.createElement('button'); transfer.type = 'button'; transfer.className = 'btn secondary'; transfer.textContent = '转入多引文草稿并重新核对…';
        transfer.addEventListener('click', () => {
          if (taskDirty) { say('材料或要求已变，请重新开始研究。'); return; }
          if (!accepted.checked || !reviewed.checked) { say('请先采用并人工核对该建议；转入后仍须重新核对每篇来源。'); return; }
          const sources = new Map();
          for (const id of proposal.evidence_ids) { const item = result.evidence[id]; if (item?.capture_id) { if (!sources.has(item.capture_id)) sources.set(item.capture_id, []); sources.get(item.capture_id).push(id); } }
          const detail = {source_sha256: result.request.source_sha256, index: result.targets.find(row => row.locator_id === proposal.locator_id).index, locator_id: proposal.locator_id, new_text: text.value, sources: [...sources].map(([capture_id, evidence_ids]) => ({capture_id, evidence_ids})), received: false};
          window.dispatchEvent(new CustomEvent('researchassist:transfer', {detail}));
          say(detail.received ? '已请求转入多引文面板；请在那里核对来源与编号并独立确认。' : '多引文面板当前不可接收，或已有草稿版本不同；请核对其提示。');
        });
        section.append(transfer);
      }
      const previous = preserved?.get(proposal.locator_id);
      if (previous && proposal.locator_id !== regenerating) { text.value = previous.text; accepted.checked = previous.accepted; reviewed.checked = previous.reviewed; }
      section.dataset.reason = proposal.reason; $('suggestions').append(section);
    }
    preserved = null; regenerating = null;
  }
  function poll(mine) {
    clearTimeout(timer);
    timer = setTimeout(() => {
      if (mine !== epoch || !auth() || task?.status !== 'running') return;
      run(async (api, current) => { const reply = await api.research_assist_status(task.id); if (!current()) return; if (!reply?.ok) { fail(reply); return; } render(reply.result); if (task.status === 'running') poll(mine); });
    }, 700);
  }
  $('start').addEventListener('click', async () => {
    if (!targets.size || !$('question').value.trim()) { say('请选择段落并填写问题。'); return; }
    const owner = epoch, source = version;
    const request = {question: $('question').value.trim(), mode: $('mode').value, source_sha256: source, locator_ids: [...targets.keys()], capture_ids: [...$('captures').querySelectorAll('input:checked')].map(node => node.value), evidence_save_id: $('receipt').value.trim(), protected_terms: $('terms').value.split('\n').map(value => value.trim()).filter(Boolean), outbound_approved: true, max_calls: Number($('max-calls').value), compute_budget: Number($('compute-budget').value)};
    if (!await ask('确认把问题、所选原段落与证据交给当前配置的对话模型？云端可能产生费用；模型结果仍需逐项人工审阅。')) return;
    if (owner !== epoch || source !== version || !auth() || busy) return;
    run(async (api, current) => {
      changed(); const reply = await api.start_research_assist(request);
      if (!current()) return; if (!reply?.ok) { fail(reply); return; } taskDirty = false; render(reply.result); if (task.status === 'running') poll(epoch);
    });
  });
  $('cancel').addEventListener('click', () => run(async (api, current) => { const reply = await api.cancel_research_assist(task.id); if (!current()) return; if (!reply?.ok) { fail(reply); return; } clearTimeout(timer); render(reply.result); $('suggestions').replaceChildren(); $('cards').textContent = ''; changed(); }));
  $('plan').addEventListener('click', () => run(async (api, current) => {
    const edits = [...$('suggestions').querySelectorAll('fieldset')].filter(node => node.querySelector('input[value=accept]').checked).map(node => ({locator_id: node.dataset.locator, new_text: node.querySelector('textarea').value, reason: node.dataset.reason, reviewed: node.querySelector('input[value=review]').checked}));
    if (!edits.length || edits.some(row => !row.reviewed)) { say('请采用至少一项，并逐项完成审阅。'); return; }
    const reply = await api.plan_assisted_edit({task_id: task.id, edits, tracked: $('tracked').checked, ...($('tracked').checked && $('granular')?.checked ? {granular: true} : {})}); if (!current()) return; if (!reply?.ok) { fail(reply); return; } plan = reply.result;
    $('preview').replaceChildren();
    const append = (parent, text) => { const node = document.createElement('p'); node.textContent = text; parent.append(node); };
    append($('preview'), '整组修改预览 · ' + (plan.items || []).length + '项 · 5分钟内独立确认 · 原件保持');
    for (const item of plan.items || []) {
      const article = document.createElement('article'), target = task.targets.find(row => row.location === item.location);
      append(article, '结构位置：' + item.location); append(article, '修改前：' + (target?.text || '请对照上方原文'));
      append(article, '修改后：' + item.text); append(article, '修改理由：' + item.reason); $('preview').append(article);
    }
    for (const warning of plan.risks || []) append($('preview'), warning);
    const detail = document.createElement('details'), heading = document.createElement('summary'), metadata = document.createElement('pre'); heading.textContent = '计划校验明细'; metadata.textContent = JSON.stringify(plan, null, 2); detail.append(heading, metadata); $('preview').append(detail);
    say('请核对差异；尚未批准写入。');
  }));
  $('confirm').addEventListener('click', async () => {
    const target = plan, owner = epoch;
    if (!target || busy) return;
    return confirmRun('请核对预览差异。确认按本计划保存新副本并进行重开、渲染和审计校验？' + ((target.items || []).some(item => item.native_revisions) ? '\nWord里会显示为修订，需你逐条接受或拒绝；尚未验证Word界面视觉效果。' : ''), async (api, current, approved) => {
      if (owner !== epoch || plan !== target) return;
      const reply = await api.confirm_assisted_edit(target.id, approved); if (!current()) return; plan = null; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'cancelled') { say('已取消，未写入。'); return; } if (reply.result.status !== 'saved') { fail({reason: 'professional_uncertain'}); return; } saved = reply.result.save_id; $('output').textContent = '已核验副本：' + reply.result.output; say('副本已保存，源文件未改；语义仍以人工审阅为准。');
    });
  });
  $('undo').addEventListener('click', async () => {
    const target = saved; if (!target) return;
    return confirmRun('撤销本次未被外部修改的输出副本？保留原件、备份和审计。', async (api, current, approved) => { if (saved !== target) return; const reply = await api.rollback_assisted_edit(target, approved); if (!current()) return; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'rolled_back') { saved = null; $('output').textContent = '本次副本已撤销。'; } });
  });
  $('save-task').addEventListener('click', () => run(async (api, current) => {
    const drafts = [...$('suggestions').querySelectorAll('fieldset')].map(node => ({locator_id: node.dataset.locator, new_text: node.querySelector('textarea').value, accepted: node.querySelector('input[value=accept]').checked, reviewed: node.querySelector('input[value=review]').checked}));
    const planned = await api.plan_research_task_save({task_id: task.id, drafts});
    if (!current()) return; if (!planned?.ok) { fail(planned); return; }
    $('task-output').textContent = JSON.stringify(planned.result, null, 2);
    const approved = await ask('保存当前问题、证据、未审阅建议及人工草稿到工作区研究归档？不保存任何调用或写入批准。');
    if (!current()) return;
    let reply;
    try { reply = await api.confirm_research_task_save(planned.result.plan_id, approved); }
    catch (error) { if (approved && current()) fail({reason: 'professional_uncertain'}); return; }
    if (!current()) return; if (!reply?.ok) { fail(reply); return; }
    if (reply.result.status === 'saved') { taskSaved = reply.result.save_id; $('task-receipt').value = taskSaved; $('task-output').textContent = '已核验研究归档：' + reply.result.evidence; }
    else if (reply.result.status === 'uncertain') fail({reason: 'professional_uncertain'});
    else say('任务归档未写入或已取消。');
  }));
  $('restore-task').addEventListener('click', async () => {
    if (task && !await ask('恢复历史研究会替换本面板当前建议与审阅状态；未保存改动会放弃，是否继续？')) return;
    run(async (api, current) => {
      if (task && task.status !== 'restored') { const cancelled = await api.cancel_research_assist(task.id); if (!current()) return; if (!cancelled?.ok) { fail(cancelled); return; } }
      const reply = await api.restore_research_task($('task-receipt').value.trim());
      if (!current()) return; if (!reply?.ok) { fail(reply); return; }
      const data = reply.result; clearTimeout(timer); changed(); version = data.task.request.source_sha256; targets.clear();
      for (const row of data.task.targets) targets.set(row.locator_id, row);
      $('question').value = data.task.request.question; $('mode').value = data.task.request.mode; $('terms').value = data.task.request.protected_terms.join('\n'); $('receipt').value = '';
      $('max-calls').value = String(data.task.request.max_calls || 3); $('compute-budget').value = String(data.task.request.compute_budget || 60);
      $('captures').replaceChildren(); $('targets').textContent = '已恢复历史目标，重新生成前请读取当前结构与证据。';
      taskDirty = false; render(data.resumable_task || data.task);
      for (const node of $('suggestions').querySelectorAll('fieldset')) { const draft = data.drafts.find(row => row.locator_id === node.dataset.locator); if (draft) node.querySelector('textarea').value = draft.new_text; }
      $('task-output').textContent = '历史任务仅供比较，人工审阅、调用、写入和删除权限均未恢复。\n失效项：' + data.stale_reasons.join('、') + '\n校验范围：' + data.validation;
      version = data.resumable_task ? data.task.request.source_sha256 : null;
    });
  });
  $('undo-task').addEventListener('click', async () => {
    const target = taskSaved; if (!target) return;
    return confirmRun('撤销本次未被外部修改的研究任务归档？源文件与审计保留。', async (api, current, approved) => { if (taskSaved !== target) return; const reply = await api.rollback_research_task(target, approved); if (!current()) return; if (!reply?.ok) { fail(reply); return; } if (reply.result.status === 'rolled_back') { taskSaved = null; $('task-output').textContent = '本次任务归档已撤销。'; } else if (reply.result.status === 'uncertain') fail({reason: 'professional_uncertain'}); });
  });
  ['question', 'terms', 'receipt', 'mode', 'max-calls', 'compute-budget'].forEach(id => $(id).addEventListener('input', () => changed(true)));
  $('tracked').addEventListener('input', () => changed());
  $('granular')?.addEventListener('input', () => changed());
  ['officeauth:unauthenticated', 'officeauth:authenticated', 'pagehide'].forEach(event => window.addEventListener(event, reset));
  ['file-name', 'file-dir'].forEach(id => { const node = document.getElementById(id); if (node) new MutationObserver(reset).observe(node, {childList: true, characterData: true, subtree: true}); });
  refresh();
})();
