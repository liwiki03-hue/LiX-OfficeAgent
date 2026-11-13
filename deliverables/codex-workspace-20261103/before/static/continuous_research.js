(() => {
  'use strict';
  const $ = id => document.getElementById('cr-' + id);
  let epoch = 0, revision = 0, busy = false, fused = false, mode = 'ranges';
  let source = null, next = null, anchors = new Map(), insertions = [], comments = [], replacements = [], ignored = [];
  let claimSelections = [], claimReview = null;
  let citationManifest = null;
  let historyPlan = null, historyRemoval = null;
  let outputReview = null, outputReviewToken = null;
  let pdfTable = null;
  let pdfGridHistory = [];
  let unitReviewStamp = null;
  const pdfGridFields = ['pdf-table-rows','pdf-table-columns','pdf-table-units','pdf-table-notes','pdf-table-header-rows'];
  let numericReviewIds = [], pendingNumericReview = null;
  let reviewSavePlan = null, reviewRemovePlan = null;
  let plan = null, saved = null, rangePlan = null, rangeSaved = null, check = null;
  const ready = () => Boolean(window.OfficeAuth?.isAuthenticated()) && Boolean(window.pywebview?.api);
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const say = value => { $('status').textContent = value; };
  const show = (id, value) => { $(id).textContent = typeof value === 'string' ? value : JSON.stringify(value, null, 2); };
  const option = (select, value, label) => { const node = document.createElement('option'); node.value = value; node.textContent = label; select.append(node); };
  const messages = {
    busy: '正在执行其他任务，请完成后再试。', not_authenticated: '请先登录。',
    professional_source_changed: '源文件已经变化，请重新读取锚点。', professional_stale_locator: '锚点已失效，请重新读取。',
    professional_citation_receipt: '回执与当前副本、版本或审计不匹配，请核对所选文件和回执。',
    professional_history_no_attempt: '当前没有可保存的未发布尝试，请先读取覆盖或执行页段预览。',
    professional_history_not_terminal: '尝试仍在等待确认；请先取消或等待过期，失败记录才可归档。',
    professional_history_attempt_unavailable: '无法读取当前尝试；请核对PDF文件名和检索配置。',
    professional_history_attempt_changed: '尝试或来源已经变化，请重新预览保存记录。',
    professional_history_receipt_invalid: '记录不存在、已撤销或不符合固定归档格式，请核对回执。',
    professional_history_evidence_invalid: '记录字段或页段状态不符合要求，请重新读取。',
    professional_history_evidence_capacity: '当前归档操作达到预算，请先完成或撤销已有操作；不自动扩容。',
    professional_history_evidence_sensitive: '诊断包含不允许保存的敏感内容，已拒绝保存。',
    professional_history_evidence_io_failed: '归档读写未完成，请核对回执状态；不自动重试。',
    professional_history_evidence_path_changed: '记录目录或文件已被外部修改，拒绝恢复或撤销。',
    professional_history_evidence_verification_failed: '记录校验失败，拒绝继续。',
    professional_history_evidence_expired: '记录操作计划已过期，请重新预览。',
    professional_history_evidence_plan_invalid: '记录操作计划已失效或已经使用，请重新预览。',
    professional_output_in_use: '已知下游仍依赖此输出，或依赖记录无法核验。请先处理下游并重新检查。',
    professional_dependency_stale: '输出已被外部删除或状态改变，请核对审计。',
    professional_dependency_invalid: '请输入32位输出回执，其他回执最多8个且不能重复。',
    professional_dependency_capacity: '本会话已知回执超过32项预算，请结束当前操作并分批核对；不自动丢弃已知依赖。',
    professional_review_invalid: '审阅会话已失效，请明确选择回执后重新开始。',
    professional_review_expired: '审阅会话已超过5分钟，请重新开始；刷新不会续期。',
    professional_review_changed: '审计、输出状态、所选文件或已知集合已经变化，请刷新并重新审阅。',
    professional_review_blocked: '存在依赖、未核验回执或循环关系，请先处理并刷新，不授予撤销权限。',
    professional_review_not_acknowledged: '尚未审阅当前版本，或审阅凭据已使用；请重新确认已审阅。',
    professional_review_capacity: '当前审阅会话达到预算，请结束旧会话后再试。',
    professional_review_read_only: '恢复的集合仅供只读审阅，不允许据此撤销文档。',
    professional_review_archive_scope: '集合属于不同的工作区身份，拒绝打开其中的文档回执。',
    professional_review_archive_invalid: '集合记录已缺失、被修改或格式不符，请核对回执。',
    professional_review_archive_evidence_invalid: '集合回执或元数据无效，请核对输入。',
    professional_review_archive_evidence_capacity: '保存操作达到预算，请先完成已有操作，不自动扩容。',
    professional_review_archive_evidence_path_changed: '集合记录目录或文件被外部修改，拒绝继续。',
    professional_review_archive_evidence_verification_failed: '集合记录校验失败，请核对审计。',
    professional_review_archive_evidence_expired: '集合操作计划已过期，请重新预览。',
    professional_review_archive_evidence_plan_invalid: '集合操作计划已使用或失效，请重新预览。',
    professional_review_archive_evidence_io_failed: '集合读写未完成，请核对回执状态，不自动重试。',
    professional_citation_scope: '缺少需要改号的段落或书目授权；请核对整组范围。',
    professional_citation_unmapped: '已有引文未全部映射，请逐项选择对应文献。',
    professional_bibliography_identity: '已有书目数量、编号或DOI与映射不一致，请核对。',
    professional_existing_comments: '文件已有批注，首版拒绝新增以保护既有批注。',
    professional_comment_range: '批注范围已改变或不准确，请重新选择。',
    professional_structure_anchor: '该锚点包含不支持的结构，请选择普通正文段落。',
    professional_structure_style: '锚点样式不在允许范围，请选择普通正文。',
    professional_support_required: '证据支持尚未完成有效人工审阅。', professional_identity_required: '请先核对文献身份。',
    professional_numeric_capacity: '表格超过400单元格、50行或8张表预算，请选择较小的独立材料。',
    professional_response_capacity: '表格、条件或脚注的核验回执超过响应预算，请选择更小的独立表格；未扩大读取范围。',
    professional_pdf_invalid: 'PDF页码、表头行数、边界或单位映射无效；表头后至少留一行数据，请核对递增百分比及列号。',
    professional_pdf_capacity: 'PDF页或网格超过预算；最多100格、2万原页字符，请缩小表格。',
    professional_pdf_unsupported: '表头有空白、网格切到了文字、单元格换行或布局不支持，请重新预览并核对边界；不自动猜测。',
    professional_pdf_pdf_text_missing: '该页没有原生文字层；此试点不执行OCR。',
    professional_pdf_source_changed: 'PDF来源已改变或无法复查，请重新预览原页。',
    professional_pdf_stale: 'PDF布局凭据已使用或失效，请重新预览。',
    professional_pdf_expired: 'PDF布局预览已超过5分钟，请重新预览。',
    professional_pdf_review_capacity: '当前PDF核验记录达到8项预算，请先完成已有草稿或稍后重新核验，不自动丢弃依赖。',
    professional_evidence_changed: '表格来源已变化，请重新读取观察值。',
    rag_capacity: '页段集合超过20页、10万字符或500块预算，请缩小选择。',
    rag_index_current: '已发布范围和解析、嵌入身份一致，没有变化；未重复嵌入或写入。',
    rag_corpus_stale: '合集缺失或来源/解析配置已改变，请重新预览并发布范围。',
    pdf_text_missing: '某页没有原生文本，整组未发布。请查看覆盖失败项并人工取证/OCR。',
    pdf_text_truncated: '某页原生文本被截断，整组未发布；请分段阅读取证。'
  };
  function fail(reply) {
    if (reply?.reason?.startsWith('professional_pdf_')) pdfTable = null;
    if (['professional_review_invalid','professional_review_expired'].includes(reply?.reason)) {
      outputReview = outputReviewToken = null; show('review-result', '审阅会话已失效，请重新开始。');
    } else if (['professional_review_changed','professional_review_blocked','professional_review_not_acknowledged'].includes(reply?.reason)) outputReviewToken = null;
    if (reply?.outcome_uncertain || reply?.result?.status === 'uncertain' || /uncertain|disabled/.test(reply?.reason || '')) {
      fused = true; plan = rangePlan = null; say('结果不确定，可能已有输出或备份。请核对审计，不自动重试。');
    } else say(messages[reply?.reason] || '操作被拒绝，请核对当前文件、范围与输入。原因：' + (reply?.reason || '未取得可靠回执'));
  }
  function refresh() {
    const blocked = busy || fused || !ready();
    document.querySelectorAll('#continuous-panel button,#continuous-panel input,#continuous-panel textarea,#continuous-panel select').forEach(n => { n.disabled = blocked; });
    $('check').querySelectorAll('button,input,textarea,select').forEach(n=>{n.disabled=blocked;});
    $('confirm').disabled = blocked || !plan;
    $('undo').disabled = blocked || !saved;
    $('range-confirm').disabled = blocked || !rangePlan;
    $('range-undo').disabled = blocked || !rangeSaved;
    $('history-confirm').disabled = blocked || !historyPlan;
    $('history-remove-confirm').disabled = blocked || !historyRemoval;
    $('review-refresh').disabled = blocked || !outputReview;
    $('pdf-table-read').disabled = blocked || !pdfTable;
    $('pdf-table-ack').disabled = blocked || !pdfTable?.observations;
    $('unit-review-apply').disabled = blocked || !unitReviewStamp;
    $('pdf-table-point-mode').disabled = blocked || !pdfTable;
    $('pdf-table-point-undo').disabled = blocked || !pdfTable || !pdfGridHistory.length;
    $('pdf-table-point-clear').disabled = blocked || !pdfTable || !pdfGridFields.some(id=>id==='pdf-table-header-rows'?$(id).value!=='1':$(id).value.trim());
    $('review-save-plan').disabled = blocked || !outputReview;
    $('review-save-confirm').disabled = blocked || !reviewSavePlan;
    $('review-remove-confirm').disabled = blocked || !reviewRemovePlan;
    $('review-cancel').disabled = blocked || !outputReview;
    $('review-ack').disabled = blocked || !outputReview || outputReview.review_blocked || Boolean(outputReviewToken);
    $('review-undo').disabled = blocked || !outputReviewToken || !outputReview?.target_has_current_undo_authority || outputReview.review_blocked;
    $('next').disabled = blocked || next === null;
    $('panel')?.setAttribute('aria-busy', String(busy));
    document.getElementById('continuous-panel').setAttribute('aria-busy', String(busy));
  }
  function changed() { revision++; historyPlan = historyRemoval = null; claimReview = null; rangePlan = null; plan = null; check = null; show('preview', ''); refresh(); }
  async function run(operation, write = false) {
    if (busy || fused || !ready()) return;
    busy = true; refresh();
    const mine = epoch, rev = revision;
    const current = () => epoch === mine && revision === rev && ready();
    try { await operation(api(), current); }
    catch (_) { if (current()) { if (write) fail({reason: 'professional_uncertain'}); else say('读取或校验失败；未批准写入。'); } }
    finally { if (epoch === mine) { busy = false; refresh(); } }
  }
  function navigate(value) {
    mode = value;
    $('ranges').hidden = value !== 'ranges'; $('document').hidden = value === 'ranges';
    $('check').hidden = value !== 'check'; $('write').hidden = value !== 'write';
    ['ranges', 'check', 'write'].forEach(id => $('nav-' + id).setAttribute('aria-pressed', String(id === value)));
  }
  function draftView() { show('drafts', {insertions, comments, replacements, ignored}); }
  function reset() {
    clearUnitReview();
    $('pdf-table-units').value='';
    window.OfficePdfPreview?.clear($('pdf-table-image'));
    numericReviewIds=[];pendingNumericReview=null;
    pdfTable = null; $('pdf-table-image').removeAttribute('src'); $('pdf-table-image').hidden = true;
    pdfGridHistory=[]; $('pdf-table-point-mode').value='inspect'; $('pdf-table-header-rows').value='1';
    show('pdf-table-position','可点击原页查看横向和纵向百分比；也可直接输入边界。');
    $('pdf-table-section').hidden=true; $('pdf-table-toggle').setAttribute('aria-expanded','false');
    $('pdf-table-cells').replaceChildren(); $('pdf-table-grid').replaceChildren(); show('pdf-table-result','');
    reviewSavePlan = reviewRemovePlan = null;
    outputReview = outputReviewToken = null;
    historyPlan = historyRemoval = null;
    citationManifest = null;
    epoch++; revision++; busy = false; claimSelections=[];claimReview=null; source = plan = saved = rangePlan = rangeSaved = check = next = null;
    anchors.clear(); insertions = []; comments = []; replacements = []; ignored = [];
    for (const id of ['review-archive-result','review-result','history-result','dependency-result','claim-sources','claim-result','original','preview','output','drafts','check-result','range-results','range-coverage','fragments','scope','number-map']) show(id, '');
    document.querySelectorAll('#continuous-panel input,#continuous-panel textarea').forEach(n => { if (n.type === 'checkbox') n.checked = false; else n.value = ''; });
    for (const id of ['anchor','bib-anchor','capture','observation','baseline']) $(id).replaceChildren();
    say(fused ? '结果不确定，请核对审计。' : '尚未读取；操作均由你主动触发。'); refresh();
  }
  function coverage(value) {
    show('range-coverage', '已发布代数：' + value.generation + '\n已选择：' + value.selected_pages.join(', ') + '\n成功解析原生文本：' + value.native_pages.join(', ') +
      '\nOCR：' + value.ocr_pages.join(', ') + '\n当前有效索引：' + value.indexed_pages.join(', ') + '\n缺口：' + value.gaps.join(', ') +
      '\n全文覆盖：' + (value.complete ? '是（文本顺序仍未人工核对）' : '否 / 尚未知总页数') + '\n状态：' + value.status +
      '\n尝试记录仅在当前会话内存保留，重启或退出后清除；不是已发布索引。' +
      (value.attempt ? '\n' + (value.attempt.source_is_current ? '当前版本未发布尝试：' : '历史版本尝试，仅供核对：') + JSON.stringify(value.attempt) : ''));
  }
  $('history-plan').onclick = () => run(async (host, current) => {
    historyPlan = null;
    const reply = await host.plan_pdf_attempt_history($('pdf-file').value.trim() || null); if (!current()) return;
    if (!reply?.ok) return fail(reply);
    historyPlan = reply.result; show('history-result', reply.result); say('仅保存所示终止尝试的诊断元数据，尚未写入；请审阅后独立确认。');
  });
  $('history-confirm').onclick = () => {
    const selected = historyPlan; if (!selected) return;
    const yes = window.confirm('保存所示页段尝试记录？只保留诊断元数据，不含PDF正文；不发布索引。');
    run(async (host, current) => {
      const reply = await host.confirm_pdf_attempt_history(selected.plan_id, yes); if (!current()) return; historyPlan = null;
      if (!reply?.ok || reply.result?.status === 'uncertain') return fail(reply);
      show('history-result', reply.result);
      if (reply.result.status === 'saved') { $('history-receipt').value = reply.result.save_id; say('记录已保存，可凭回执在重启后恢复供核对。'); }
      else say(yes ? '未保存记录，请重新预览或核对回执中的原因。' : '已取消，未保存记录。');
    }, yes);
  };
  $('history-restore').onclick = () => run(async (host, current) => {
    const reply = await host.restore_pdf_attempt_history($('history-receipt').value.trim()); if (!current()) return;
    if (!reply?.ok) return fail(reply); show('history-result', reply.result);
    say('已读取历史诊断；来源、解析配置和嵌入身份状态见下方。未恢复索引、活动尝试、批准或撤销权限。');
  });
  $('history-remove-plan').onclick = () => run(async (host, current) => {
    historyRemoval = null;
    const reply = await host.plan_remove_pdf_attempt_history($('history-receipt').value.trim()); if (!current()) return;
    if (!reply?.ok) return fail(reply); historyRemoval = reply.result; show('history-result', reply.result);
    say('仅撤销这条诊断记录；源PDF、索引、其他证据、备份与审计保留。请独立确认。');
  });
  $('history-remove-confirm').onclick = () => {
    const selected = historyRemoval; if (!selected) return;
    const yes = window.confirm('撤销所示诊断记录文件？备份和审计保留，不修改PDF、索引或已保存证据。');
    run(async (host, current) => {
      const reply = await host.confirm_remove_pdf_attempt_history(selected.rollback_id, yes); if (!current()) return; historyRemoval = null;
      if (!reply?.ok || reply.result?.status === 'uncertain') return fail(reply); show('history-result', reply.result);
      say(reply.result.status === 'rolled_back' ? '记录已撤销，备份与审计保留。' : yes ? '未获得撤销成功回执，请核对原因。' : '已取消，记录保持。');
    }, yes);
  };
  $('history-receipt').oninput = () => { revision++; historyPlan = historyRemoval = null; refresh(); };
  $('dependency-check').onclick = () => run(async (host, current) => {
    const receipts = $('dependency-receipts').value.trim().split(/[,，\s]+/).filter(Boolean);
    const reply = await host.preview_output_dependencies($('dependency-target').value.trim(), receipts); if (!current()) return;
    if (!reply?.ok) return fail(reply); show('dependency-result', reply.result);
    say(reply.result.blockers.length ? '存在已知依赖或无法核验的记录，请先处理下游；未授予撤销权限。' : '所列已知回执中未发现阻塞；检查范围不完整，不授予撤销权限。');
  });
  function reviewView(value) {
    const state = {verified:'已验证输出',rolled_back:'已撤销',unavailable_or_refused:'缺失或拒绝读取'};
    const lines = ['审阅版本：' + value.revision + '；读取时剩余：' + value.expires_in_seconds + '秒（刷新不续期）',
      '目标：' + value.target, '模式：' + (value.historical_only ? '恢复的集合，仅供只读审阅' : '当前授权上下文审阅'), '范围：仅明确回执与当前授权上下文已知集合，不是完整依赖图。',
      '文件内容：未在审阅中重验；实际撤销仍由原事务核对摘要。',
      '原撤销权限：' + (value.target_has_current_undo_authority ? '当前有效；仍需独立确认' : '没有有效权限，仅供只读审阅'),
      '记录：', ...value.records.map(row => row.receipt_id + ' ' + (row.output_name || '') + ' ' + (state[row.status] || row.status)),
      '关系（上游 → 下游）：', ...value.edges.map(edge => edge.parent + ' → ' + edge.child),
      '活动下游：' + (value.active_descendants.join('、') || '当前范围内未发现'),
      '未核验：' + (value.unresolved_receipts.join('、') || '无'),
      '循环或受循环影响：' + (value.cycle_or_cycle_dependent_receipts.join('、') || '无'),
      '范围外父回执：' + (value.outside_scope_parents.map(row => row.parent).join('、') || '无已声明项'),
      '建议先后顺序（不自动执行）：' + (value.suggested_child_first_order.join(' → ') || '暂不可生成'),
      '变化：新增 ' + value.changes.added.join('、') + '；改变 ' + value.changes.changed.join('、') + '；移除 ' + value.changes.removed.join('、'),
      '状态：' + (value.review_blocked ? '存在阻塞，不可据此撤销' : value.review_acknowledged ? '已审阅此快照，不等于写入批准' : '待人工审阅快照')];
    show('review-result', lines.join('\n'));
  }
  $('review-start').onclick = () => run(async (host, current) => {
    reviewSavePlan = null;
    outputReviewToken = null;
    const receipts = $('dependency-receipts').value.trim().split(/[,，\s]+/).filter(Boolean);
    const reply = await host.start_output_review($('dependency-target').value.trim(), receipts); if (!current()) return;
    if (!reply?.ok) return fail(reply); outputReview = reply.result; reviewView(outputReview);
    say('已开始只读审阅会话；请核对范围、依赖和顺序，检查不会恢复旧撤销权限。');
  });
  $('review-refresh').onclick = () => run(async (host, current) => {
    reviewSavePlan = null;
    if (!outputReview) return; outputReviewToken = null;
    const reply = await host.refresh_output_review(outputReview.review_id); if (!current()) return;
    if (!reply?.ok) return fail(reply); outputReview = reply.result; reviewView(outputReview);
    say('已刷新检查及变化；请重新审阅当前快照。');
  });
  $('review-ack').onclick = () => run(async (host, current) => {
    if (!outputReview) return;
    const reply = await host.acknowledge_output_review(outputReview.review_id, outputReview.revision, true); if (!current()) return;
    if (!reply?.ok) { outputReviewToken = null; return fail(reply); }
    outputReview = reply.result; outputReviewToken = reply.result.review_token; reviewView(outputReview);
    say('已记录当前快照的审阅；未批准删除。只有原撤销权限有效的目标才可另行确认。');
  });
  $('review-cancel').onclick = () => run(async (host, current) => {
    reviewSavePlan = null;
    if (!outputReview) return;
    const reply = await host.cancel_output_review(outputReview.review_id); if (!current()) return;
    if (!reply?.ok) return fail(reply); outputReview = outputReviewToken = null; show('review-result','会话已结束；文件、备份和审计未改。');
  });
  $('review-undo').onclick = () => {
    const selected = outputReview, token = outputReviewToken; if (!selected || !token) return;
    const yes = window.confirm('另行确认仅撤销目标 ' + selected.target + ' 的输出副本？审阅快照将再次核对，且必须保有原撤销权限；不批量删除。');
    run(async (host, current) => {
      const reply = await host.rollback_reviewed_output(selected.review_id, token, yes); if (!current()) return; outputReviewToken = null;
      if (!reply?.ok || reply.result?.status === 'uncertain') return fail(reply);
      if (reply.result.status === 'rolled_back') { if (saved === selected.target) saved = null; show('review-result','目标副本已撤销；请刷新查看依赖变化。'); say('已按原权限独立撤销单个副本，源文件、备份和审计保留。'); }
      else say('已取消撤销；一次性审阅凭据已消费，文件未改。');
    }, yes);
  };
  $('review-save-plan').onclick = () => run(async (host, current) => {
    if (!outputReview) return; reviewSavePlan = null;
    const reply = await host.plan_output_review_archive(outputReview.review_id, outputReview.revision); if (!current()) return;
    if (!reply?.ok) return fail(reply); reviewSavePlan = reply.result; show('review-archive-result', reply.result);
    say('请核对将保存的完整回执集合；仅保存ID和校验摘要，尚未写入。');
  });
  $('review-save-confirm').onclick = () => {
    const selected = reviewSavePlan; if (!selected) return;
    const yes = window.confirm('保存预览中的回执集合？不包含正文或撤销凭据，恢复后仅供只读审阅。');
    run(async (host, current) => {
      const reply = await host.confirm_output_review_archive(selected.plan_id, yes); if (!current()) return; reviewSavePlan = null;
      if (!reply?.ok || reply.result?.status === 'uncertain') return fail(reply); show('review-archive-result', reply.result);
      if (reply.result.status === 'rejected' && reply.result.code !== 'evidence_denied') return fail({reason:'professional_review_archive_' + reply.result.code});
      if (reply.result.status === 'saved') { $('review-archive-receipt').value = reply.result.save_id; say('集合已保存，请保留回执；重启后可恢复并重新核验。'); }
      else say('未保存集合；请核对取消或拒绝原因。');
    }, yes);
  };
  $('review-restore').onclick = () => run(async (host, current) => {
    reviewSavePlan = null; outputReviewToken = null;
    const reply = await host.restore_output_review_archive($('review-archive-receipt').value.trim()); if (!current()) return;
    if (!reply?.ok) return fail(reply); outputReview = reply.result; reviewView(outputReview);
    show('review-archive-result', {archive_receipt:reply.result.archive_receipt, changes:reply.result.changes, historical_only:true});
    say('集合已重新核验；变化与缺失见审阅结果。仅供只读审阅，未恢复批准或文档撤销权限。');
  });
  $('review-remove-plan').onclick = () => run(async (host, current) => {
    reviewRemovePlan = null;
    const reply = await host.plan_remove_output_review_archive($('review-archive-receipt').value.trim()); if (!current()) return;
    if (!reply?.ok) return fail(reply); reviewRemovePlan = reply.result; show('review-archive-result', reply.result);
    say('仅删除这条集合记录，文档、备份和审计保留；请独立确认。');
  });
  $('review-remove-confirm').onclick = () => {
    const selected = reviewRemovePlan; if (!selected) return;
    const yes = window.confirm('删除预览中的集合记录？不删除任何文档，备份和审计保留。');
    run(async (host, current) => {
      const reply = await host.confirm_remove_output_review_archive(selected.rollback_id, yes); if (!current()) return; reviewRemovePlan = null;
      if (!reply?.ok || reply.result?.status === 'uncertain') return fail(reply); show('review-archive-result', reply.result);
      if (reply.result.status === 'rejected' && reply.result.code !== 'evidence_denied') return fail({reason:'professional_review_archive_' + reply.result.code});
      say(reply.result.status === 'rolled_back' ? '集合记录已删除；文档、备份和审计保留。' : '未删除集合记录。');
    }, yes);
  };
  $('review-archive-receipt').oninput = () => { revision++; reviewRemovePlan = null; refresh(); };
  for (const id of ['dependency-target','dependency-receipts']) $(id).addEventListener('input', () => { revision++; reviewSavePlan = null; outputReview = outputReviewToken = null; show('review-result','回执范围已编辑，请重新开始审阅会话。'); refresh(); });
  $('range-load').onclick = () => run(async (host, current) => {
    const reply = await host.get_pdf_ranges($('pdf-file').value.trim() || null); if (!current()) return; if (!reply?.ok) return fail(reply);
    coverage(reply.ranges); $('range-input').value = reply.ranges.pages.map(n => n + '-' + n).join(',');
  });
  $('range-input').oninput = () => { rangePlan = null; revision++; refresh(); };
  $('range-plan').onclick = () => run(async (host, current) => {
    const raw = $('range-input').value.trim();
    if (raw && !/^\d+(?:-\d+)?(?:\s*[,，]\s*\d+(?:-\d+)?)*$/.test(raw)) { say('范围格式示例：1-2,6-7。'); return; }
    const ranges = raw ? raw.split(/[,，]/).map(part => { const pair = part.trim().split('-').map(Number); return [pair[0], pair[1] ?? pair[0]]; }) : [];
    rangePlan = null;
    const reply = await host.plan_pdf_ranges(ranges, $('pdf-file').value.trim() || null); if (!current()) return; if (!reply?.ok) {
      const status = await host.get_pdf_ranges($('pdf-file').value.trim() || null); if (!current()) return;
      if (status?.ok) coverage(status.ranges); fail(reply); return;
    }
    rangePlan = reply.plan; show('range-coverage', '待独立确认；尚未嵌入或写入。\n' + JSON.stringify(reply.plan, null, 2));
  });
  $('range-confirm').onclick = () => {
    const selected = rangePlan; if (!selected) return;
    const yes = window.confirm('发布所示页集合？只有新增/失效页会使用已配置的嵌入通道；API通道会发送这些页的文字。源文件不变。');
    run(async (host, current) => {
      const reply = await host.confirm_rag_index(selected.plan_id, yes); if (!current()) return; rangePlan = null;
      if (!reply?.ok || reply.result?.status === 'uncertain') return fail(reply);
      if (reply.result.status === 'saved') { rangeSaved = reply.result.save_id; const status = await host.get_pdf_ranges($('pdf-file').value.trim() || null); if (current() && status?.ok) coverage(status.ranges); say('页集合已整组发布，可继续检索或撤销。'); }
      else say('已取消或计划过期，未获得发布成功回执。');
    }, yes);
  };
  $('range-undo').onclick = () => run(async (host, current) => {
    const planned = await host.plan_rag_rollback('index', rangeSaved); if (!current()) return; if (!planned?.ok) return fail(planned);
    const reply = await host.confirm_rag_rollback('index', planned.plan.rollback_id, window.confirm('恢复本次发布前的索引？独立保存的证据保留。'));
    if (!current()) return; if (!reply?.ok || reply.result.status === 'uncertain') return fail(reply); rangeSaved = null; say('索引撤销已处理，请重新读取覆盖。');
  }, true);
  $('query-run').onclick = () => run(async (host, current) => {
    const planned = await host.query_pdf_ranges($('query').value.trim(), 10, $('pdf-file').value.trim() || null); if (!current()) return; if (!planned?.ok) return fail(planned);
    const reply = await host.confirm_rag_query(planned.plan.plan_id, window.confirm('保存当前已发布页段的检索证据？未覆盖页不能据此作否定结论。'));
    if (!current()) return; if (!reply?.ok || reply.result.status === 'uncertain') return fail(reply);
    show('range-results', reply.result);
    if (reply.result.status === 'saved') { const receipt = document.getElementById('assist-receipt'); if (receipt) { receipt.value = reply.result.save_id; receipt.dispatchEvent(new Event('input')); } say('证据已保存，回执已填入研究助理；保持当前DOCX选中即可主动运行研究助理。'); }
  }, true);
  async function load(host, current, offset = 0) {
    const reply = await host.read_structure_anchors(offset); if (!current()) return; if (!reply?.ok) return fail(reply);
    if (offset && source?.source_sha256 !== reply.result.document.source_sha256) { source = null; next = null; plan = null; fail({reason:'professional_source_changed'}); return; }
    source = reply.result.document; next = reply.result.next_offset;
    if (!offset) { anchors.clear(); $('anchor').replaceChildren(); $('bib-anchor').replaceChildren(); $('scope').replaceChildren(); }
    for (const row of reply.result.anchors) {
      anchors.set(row.locator_id, row); option($('anchor'), row.locator_id, row.location + ' ' + row.text.slice(0,60)); option($('bib-anchor'), row.locator_id, row.location + ' ' + row.text.slice(0,60));
      const line = document.createElement('div');
      for (const [kind, text] of [['authorized','允许改号'],['bibliography','已有书目']]) {
        const label = document.createElement('label'), input = document.createElement('input'); input.type = 'checkbox'; input.dataset.scope = kind; input.value = row.locator_id; input.onchange = changed;
        label.append(input, document.createTextNode(text)); line.append(label);
      }
      line.append(document.createTextNode(row.location + ' ' + row.text.slice(0,100))); $('scope').append(line);
    }
    $('bib-anchor').value = [...anchors.keys()].at(-1) || '';
    show('original', anchors.get($('anchor').value)?.text || '');
    if (!offset) {
      const materials = await host.read_citation_workspace($('receipt').value.trim()); if (!current()) return;
      if (!materials?.ok) { source = null; next = null; citationManifest = null; fail(materials); return; }
      citationManifest = materials.result.manifest;
      $('capture').replaceChildren(); option($('capture'), '', '不附引文（不表示内容已核验）'); $('number-map').replaceChildren();
      if (materials?.ok) {
        for (const capture of materials.result.captures) option($('capture'), capture.capture_id, capture.title);
        for (const number of materials.result.citation_numbers) {
          const label = document.createElement('label'), select = document.createElement('select'); select.dataset.number = number; option(select,'','请选择对应文献');
          for (const record of materials.result.references) option(select, record.id, record.title + ' ' + record.DOI);
          const key = citationManifest?.number_to_key?.[String(number)];
          if (key) select.value = citationManifest.catalog[key].id;
          label.append(document.createTextNode('[' + number + ']'),select); select.onchange = changed; $('number-map').append(label);
        }
      }
    }
    for (const input of $('scope').querySelectorAll('input[data-scope="bibliography"]')) {
      if (citationManifest?.slots?.includes(anchors.get(input.value)?.location)) input.checked = true;
    }
    say('已读取版本绑定锚点；' + (citationManifest ? '已恢复编号与书目候选，未恢复授权或证据审阅。' : '选择正文及已取证原文后再加入草稿。'));
  }
  $('receipt').oninput = () => { source = null; next = null; citationManifest = null; changed(); say('回执已变化，请重新读取；草稿保留供核对。'); };
  $('load').onclick = () => { if ((insertions.length || comments.length || replacements.length) && !window.confirm('重新读取会放弃本组未保存草稿，是否继续？')) return; numericReviewIds=[];pendingNumericReview=null;insertions=[]; comments=[]; replacements=[]; ignored=[];claimSelections=[];claimReview=null;show('claim-sources','');changed(); run((host,current) => load(host,current)); };
  $('next').onclick = () => { const offset = next; run((host,current) => load(host,current,offset)); };
  $('anchor').onchange = () => { changed(); show('original', anchors.get($('anchor').value)?.text || ''); };
  $('capture').onchange = () => { changed(); $('reviewed').checked = false; $('fragments').replaceChildren(); if (!$('capture').value) return;
    run(async (host,current) => { const reply = await host.read_citation_capture($('capture').value); if (!current()) return; if (!reply?.ok) return fail(reply);
      for (const fragment of reply.result.fragments) { const label = document.createElement('label'), input = document.createElement('input'); input.type='checkbox'; input.value=fragment.id; input.onchange=changed; label.append(input,document.createTextNode(fragment.location + ' ' + fragment.original_text)); $('fragments').append(label); }
    });
  };
  $('claim-add').onclick=()=>{const ids=[...$('fragments').querySelectorAll('input:checked')].map(n=>n.value);if(!$('capture').value||!ids.length||ids.length>3||claimSelections.length>=3){say('每次选择1–3条原文，最多3个来源。');return;}claimSelections=claimSelections.filter(r=>r.capture_id!==$('capture').value);claimSelections.push({capture_id:$('capture').value,evidence_ids:ids});changed();show('claim-sources',claimSelections);};
  $('claim-apply').onclick=()=>{const row=claimReview?.review?.suggestions?.find(r=>r.locator_id===$('anchor').value);if(!row){say('当前段落没有可用的未审阅替换建议。');return;}$('replacement').value=row.new_text;$('replace-reason').value=row.reason;navigate('write');changed();say('已填入草稿；请人工核对，再加入替换清单并预览整组影响。');};
  $('claim-review').onclick = () => {
    const ids=[...$('fragments').querySelectorAll('input:checked')].map(n=>n.value);
    if(!claimSelections.length && (!$('capture').value || !ids.length || ids.length>3)) {say('先选择已取证材料与1–3条原文；多来源争议仍可使用研究助理的多材料审阅。');return;}
    if(!window.confirm('主动使用已配置模型核对所选段落与证据？一次调用，可能产生费用；结果不自动视为人工审阅。'))return;
    run(async(host,current)=>{const reply=await host.review_paragraph_claim($('anchor').value,claimSelections.length ? claimSelections : [{capture_id:$('capture').value,evidence_ids:ids}]);if(!current())return;if(!reply?.ok)return fail(reply);claimReview=reply.result;show('claim-result',reply.result);say('论点核验为未审阅建议；请检查原文和适用条件，再决定是否加入修改或批注。');});
  };
  $('add-insert').onclick = () => {
    if (!source || !anchors.has($('anchor').value) || !$('text').value.trim() || insertions.length >= 4) { say('请读取锚点并填写正文，每组最多4个插入。'); return; }
    const attributions = [];
    if ($('capture').value) {
      const ids = [...$('fragments').querySelectorAll('input:checked')].map(n=>n.value);
      if (!$('reviewed').checked || !ids.length || ids.length > 3 || !$('applicability').value.trim() || !$('limitation').value.trim()) { say('请选1–3条证据，填写适用条件与限制，并人工核对后勾选。'); return; }
      attributions.push({capture_id:$('capture').value,evidence_ids:ids,relation:'supports',applicability:$('applicability').value.trim(),limitation:$('limitation').value.trim(),assessment:'user_reviewed',identity_reviewed:true});
    }
    insertions.push({anchor:$('anchor').value,side:$('side').value,text:$('text').value.trim(),attributions}); changed(); draftView(); say('已加入草稿，尚未批准写入。');
  };
  $('add-replace').onclick=()=>{const row=anchors.get($('anchor').value);if(!row||!$('replacement').value.trim()||!$('replace-reason').value.trim()||replacements.length>=12){say('先读取目标，填写替换文字和理由；每组最多12项。');return;}replacements.push({locator_id:row.locator_id,expected_text_sha256:row.text_sha256,new_text:$('replacement').value.trim(),reason:$('replace-reason').value.trim(),evidence_ids:[]});changed();draftView();};
  $('add-comment').onclick = () => {
    const row = anchors.get($('anchor').value), quote = $('comment-quote').value, text = $('comment-text').value.trim();
    const start = row?.text.indexOf(quote) ?? -1;
    if (!row || !quote || !text || start < 0 || row.text.indexOf(quote,start+1) >= 0 || comments.length >= 8) { say('请填写唯一匹配的原文字串和批注，每组最多8条。'); return; }
    comments.push({anchor:row.locator_id,start,end:start+quote.length,quote,text});
    if(pendingNumericReview&&!numericReviewIds.includes(pendingNumericReview))numericReviewIds.push(pendingNumericReview);pendingNumericReview=null;
    changed(); draftView();
  };
  $('clear').onclick = () => { if (!window.confirm('放弃本组未写入草稿？')) return; numericReviewIds=[];pendingNumericReview=null;insertions=[];comments=[];replacements=[];ignored=[];claimSelections=[];claimReview=null;show('claim-sources','');changed();draftView(); };
  $('plan').onclick = () => run(async (host,current) => {
    if (!source) { say('请先读取DOCX锚点。'); return; }
    const number_records = {}; $('number-map').querySelectorAll('select').forEach(n=> { if(n.value) number_records[n.dataset.number]=n.value; });
    const selected = kind => [...$('scope').querySelectorAll('input:checked')].filter(n=>n.dataset.scope===kind).map(n=>n.value);
    const reply = await host.plan_structure_edit({source_sha256:source.source_sha256,insertions,comments,replacements,review_notes:ignored.map(r=>r.reason),numeric_reviews:numericReviewIds,authorized:selected('authorized'),bibliography:selected('bibliography'),bibliography_anchor:$('bib-anchor').value || null,number_records});
    if (!current()) return; if (!reply?.ok) return fail(reply); plan=reply.result; show('preview','待独立确认，原件不变。\n修改前：\n'+reply.result.before+'\n整组影响：\n'+reply.result.after); say('整组校验完成；审阅插入、改号、书目和批注后，再独立确认。');
  });
  $('confirm').onclick = () => { const selected=plan;if(!selected)return; const yes=window.confirm('确认所示整组差异，并生成经重开与渲染验证的副本？正文插入为净稿模式。');
    run(async(host,current)=>{const reply=await host.confirm_professional_edit(selected.id,yes);if(!current())return;plan=null;if(!reply?.ok||reply.result.status==='uncertain')return fail(reply);if(reply.result.status==='saved'){saved=reply.result.save_id;$('dependency-target').value=saved;outputReview=outputReviewToken=null;show('review-result','新输出已生成，可开始该目标的依赖审阅会话。');show('output',reply.result);say('副本已保存并通过重开与渲染闸门；Word人工审阅仍待完成。');}else say('已取消，未写入。');},yes);
  };
  $('undo').onclick = () => { const yes=window.confirm('独立撤销本次未被外部修改的输出副本？原件、备份和审计保留。');run(async(host,current)=>{const reply=await host.rollback_professional_edit(saved,yes);if(!current())return;if(!reply?.ok||reply.result.status==='uncertain')return fail(reply);if(reply.result.status==='rolled_back'){saved=null;show('output','本次副本已撤销。');}},yes); };
  function fillObservations(records) {
    $('observation').replaceChildren(); $('baseline').replaceChildren(); option($('baseline'),'','没有比较基线');
    for (const row of records) { const text=row.location+' '+row.header+' 原文：'+(row.missing?'缺失':row.raw); option($('observation'),row.id,text); option($('baseline'),row.id,text); }
  }
  function pdfEdges(id) { return $(id).value.split(/[,，]/).filter(value=>value.trim()).map(value=>Number(value.trim())/100); }
  $('pdf-table-toggle').onclick=()=>{const open=$('pdf-table-section').hidden;$('pdf-table-section').hidden=!open;$('pdf-table-toggle').setAttribute('aria-expanded',String(open));};
  function paintPdfGrid() {
    $('pdf-table-grid').replaceChildren();
    for (const [values, vertical] of [[pdfEdges('pdf-table-columns'),true],[pdfEdges('pdf-table-rows'),false]]) {
      for (const value of values) if (Number.isFinite(value) && value>=0 && value<=1) {
        const line=document.createElementNS('http://www.w3.org/2000/svg','line');
        for(const [key,number] of Object.entries(vertical?{x1:value,x2:value,y1:0,y2:1}:{y1:value,y2:value,x1:0,x2:1}))line.setAttribute(key,String(number));
        line.setAttribute('stroke','#e65024'); line.setAttribute('stroke-width','.002'); $('pdf-table-grid').append(line);
      }
    }
    const note=$('pdf-table-notes').value.trim()?pdfEdges('pdf-table-notes'):[];
    if(note.length===4&&note.every(Number.isFinite)&&note[2]>note[0]&&note[3]>note[1]){
      const box=document.createElementNS('http://www.w3.org/2000/svg','rect');
      for(const [key,value] of Object.entries({x:note[0],y:note[1],width:note[2]-note[0],height:note[3]-note[1]}))box.setAttribute(key,String(value));
      box.setAttribute('fill','none');box.setAttribute('stroke','#14933b');box.setAttribute('stroke-width','.003');$('pdf-table-grid').append(box);
    }
  }
  function pdfGridChanged() {
    clearUnitReview();
    if(pdfTable)pdfTable.observations=null;
    pendingNumericReview=null; fillObservations([]);
    $('pdf-table-cells').replaceChildren(); show('pdf-table-result','网格已修改，请重新提取并核对布局。'); show('check-result','');
    paintPdfGrid(); changed();
  }
  function rememberPdfGrid() {
    if(pdfGridHistory.length>=16){say('网格撤销记录已达16次，请先撤销或手工编辑边界后继续。');return false;}
    pdfGridHistory.push(pdfGridFields.map(id=>$(id).value)); return true;
  }
  function clearUnitReview() {
    unitReviewStamp=null;$('unit-review-fields').replaceChildren();$('unit-review').hidden=true;
  }
  function showUnitReview(data) {
    clearUnitReview();
    if(!data.proposal||!data.unit_review?.length)return;
    unitReviewStamp=JSON.stringify(pdfGridFields.map(id=>$(id).value));
    for(const row of data.unit_review){
      const label=document.createElement('label'),select=document.createElement('select');
      select.dataset.valueColumn=String(row.value_column);select.id='cr-unit-column-'+row.value_column;
      label.htmlFor=select.id;label.append(document.createTextNode('第'+row.value_column+'列 '+row.header+'：'+(row.status==='header_and_row_unit_agreement'?'表头与单位列文字一致，仍需核对':'单位归属不明确，请核对')));
      option(select,'','不关联（未建立单位来源）');
      for(const choice of row.options){if(!['ambiguous_unit_scope','header_and_row_unit_agreement'].includes(choice.status))continue;option(select,String(choice.unit_column),'第'+choice.unit_column+'列：'+choice.unit_values.join(' / ')+'（需核对适用范围）');}
      label.append(select);$('unit-review-fields').append(label);
    }
    $('unit-review').hidden=false;
  }
  $('unit-review-apply').onclick=()=>{
    if(busy||fused||!ready()||!pdfTable||unitReviewStamp!==JSON.stringify(pdfGridFields.map(id=>$(id).value)))return;
    const pairs=[...$('unit-review-fields').querySelectorAll('select')].filter(n=>n.value).map(n=>n.dataset.valueColumn+':'+n.value);
    $('pdf-table-units').value=pairs.join(',');$('pdf-table-units').dispatchEvent(new Event('input'));
    say('已填入明确选择的单位来源；旧提取和核验已失效，请重新提取并采用布局。未选择的列不会自动变成无量纲。');
  };
  $('pdf-table-point-undo').onclick=()=>{
    if(busy||fused||!ready()||!pdfTable||!pdfGridHistory.length)return;
    const previous=pdfGridHistory.pop(); pdfGridFields.forEach((id,index)=>{$(id).value=previous[index];});
    pdfGridChanged(); say('已撤销网格操作；旧提取结果和布局批准不会恢复，请重新提取。');
  };
  $('pdf-table-point-clear').onclick=()=>{
    if(busy||fused||!ready()||!pdfTable||!pdfGridFields.some(id=>id==='pdf-table-header-rows'?$(id).value!=='1':$(id).value.trim())||!rememberPdfGrid())return;
    pdfGridFields.forEach(id=>{$(id).value=id==='pdf-table-header-rows'?'1':'';}); pdfGridChanged(); say('已清空网格、单位映射和脚注范围，表头恢复一行；尚未读取或写入文件，可撤销本次清空。');
  };
  $('pdf-table-image').onclick = event => {
    if(busy||fused||!ready()||!pdfTable||event.currentTarget.hidden)return;
    const box=event.currentTarget.getBoundingClientRect();
    if(!box.width||!box.height)return;
    const x=Number(((event.clientX-box.left)/box.width*100).toFixed(2)),y=Number(((event.clientY-box.top)/box.height*100).toFixed(2));
    if(x<0||x>100||y<0||y>100)return;
    show('pdf-table-position','横向 '+x.toFixed(2)+'%；纵向 '+y.toFixed(2)+'%。边界必须避开文字。');
    const pointMode=$('pdf-table-point-mode').value;
    if(!['rows','columns'].includes(pointMode))return;
    const id='pdf-table-'+pointMode, edges=pdfEdges(id).map(value=>value*100), value=pointMode==='rows'?y:x;
    if(edges.some((v,i)=>!Number.isFinite(v)||v<0||v>100||(i&&v<=edges[i-1]))){say('已有边界无效，请先在输入框中修正后再选点。');return;}
    if(edges.some(v=>Math.abs(v-value)<.005)){say('该边界已存在，未修改网格。');return;}
    const maximum=pointMode==='rows'?22:13, other=pdfEdges('pdf-table-'+(pointMode==='rows'?'columns':'rows')).length;
    if(edges.length+1>maximum||edges.length*Math.max(0,other-1)>100){say('网格超过100格或行列预算，请缩小选择。');return;}
    if(!rememberPdfGrid())return;
    $(id).value=[...edges,value].sort((a,b)=>a-b).map(v=>Number(v.toFixed(4))).join(',');
    if(pointMode==='columns')$('pdf-table-units').value='';
    pdfGridChanged(); say(pointMode==='columns'?'已添加列边界，单位映射已清除，请按新列号重新填写。':'已添加行边界，请继续选择并核对表头和数据行。');
  };
  $('pdf-table-preview').onclick = () => run(async(host,current)=>{
    window.OfficePdfPreview?.clear($('pdf-table-image'));
    pdfTable=null; pdfGridHistory=[]; $('pdf-table-point-mode').value='inspect'; check=null; fillObservations([]); $('pdf-table-cells').replaceChildren();
    $('pdf-table-image').removeAttribute('src'); $('pdf-table-image').hidden=true;
    const reply=await host.preview_pdf_numeric($('pdf-table-file').value.trim(),Number($('pdf-table-page').value)); if(!current())return;
    if(!reply?.ok)return fail(reply); pdfTable=reply.result;
    const imageEpoch=epoch;
    if(window.OfficePdfPreview){if(!await window.OfficePdfPreview.show($('pdf-table-image'),reply.result.image,()=>imageEpoch===epoch&&ready())){if(current())say('原页未完成解码，可重新显示本次原页；未重复读取。');}}
    else{$('pdf-table-image').src=reply.result.image;$('pdf-table-image').hidden=false;}
    if(!current())return;paintPdfGrid();
    show('pdf-table-result',{page:reply.result.page,total_pages:reply.result.total_pages,source_sha256:reply.result.source_sha256,mode:'显式网格；不是自动找表或OCR；布局未确认'});
  });
  $('pdf-table-read').onclick = () => run(async(host,current)=>{
    if(!pdfTable)return; check=null; fillObservations([]);
    const pairs=$('pdf-table-units').value.trim();
    if(pairs&&!/^\d+:\d+(?:\s*[,，]\s*\d+:\d+)*$/.test(pairs)){say('单位映射格式为3:2,5:2，从1计数。');return;}
    const grid={page:pdfTable.page,rows:pdfEdges('pdf-table-rows'),columns:pdfEdges('pdf-table-columns'),units:pairs?pairs.split(/[,，]/).map(pair=>pair.split(':').map(value=>Number(value)-1)):[],note_box:$('pdf-table-notes').value.trim()?pdfEdges('pdf-table-notes'):null,header_rows:Number($('pdf-table-header-rows').value)};
    const reply=await host.read_pdf_numeric(pdfTable.pdf_table_id,grid); if(!current())return;
    if(!reply?.ok)return fail(reply); pdfTable=reply.result;
    const table=document.createElement('table');
    for(const [index,row] of reply.result.rows.entries()){const tr=document.createElement('tr');for(const text of row){const cell=document.createElement(index<(reply.result.header_rows||1)?'th':'td');cell.textContent=text;tr.append(cell);}table.append(tr);}
    $('pdf-table-cells').replaceChildren(table); show('pdf-table-result',{source_sha256:reply.result.source_sha256,page:reply.result.page,cells:reply.result.cells,note_text:reply.result.note_text,context_outside_selection_included:false,layout_confirmed:false});
    say('请逐格对照原页，检查行列和单位映射；确认布局只准用于本次核验，不是事实或写入批准。');
  });
  $('pdf-table-ack').onclick = () => run(async(host,current)=>{
    if(!pdfTable?.observations)return;
    const reply=await host.acknowledge_pdf_numeric(pdfTable.pdf_table_id,true); if(!current())return;
    if(!reply?.ok)return fail(reply); pdfTable=null; pdfGridHistory=[]; fillObservations(reply.result.observations); show('check-result',reply.result);
    say('已采用明确布局；观察值仍未审阅，可核对数值与条件。修改网格后请重新预览。');
  });
  for(const id of ['pdf-table-file','pdf-table-page']) $(id).addEventListener('input',()=>{pdfTable=null;pdfGridHistory=[];$('pdf-table-point-mode').value='inspect';$('pdf-table-header-rows').value='1';pendingNumericReview=null;fillObservations([]);$('pdf-table-image').removeAttribute('src');$('pdf-table-image').hidden=true;$('pdf-table-cells').replaceChildren();$('pdf-table-grid').replaceChildren();show('check-result','');show('pdf-table-result','来源已编辑，请重新预览。');changed();});
  for(const id of pdfGridFields) $(id).addEventListener('input',()=>{pdfGridHistory=[];if(['pdf-table-columns','pdf-table-header-rows'].includes(id))$('pdf-table-units').value='';pdfGridChanged();});
  $('pdf-table-retry').onclick=()=>run(async()=>{const okay=await window.OfficePdfPreview?.retry($('pdf-table-image'));say(okay?'已重新显示本次原页，未新增读取或下载。':'原页仍未显示，请重新预览并核对来源。');});
  $('pdf-table-propose').onclick=()=>run(async(host,current)=>{
    clearUnitReview();
    pdfTable=null;pdfGridHistory=[];fillObservations([]);$('pdf-table-cells').replaceChildren();
    const box=pdfEdges('pdf-table-region');
    if(box.length!==4){say('请明确填写矩形区域左、上、右、下百分比。');return;}
    const reply=await host.propose_pdf_layout($('pdf-table-file').value.trim(),Number($('pdf-table-page').value),box);if(!current())return;
    if(!reply?.ok)return fail(reply);pdfTable=reply.result;const imageEpoch=epoch;
    if(window.OfficePdfPreview)await window.OfficePdfPreview.show($('pdf-table-image'),reply.result.image,()=>imageEpoch===epoch&&ready());
    else{$('pdf-table-image').src=reply.result.image;$('pdf-table-image').hidden=false;}
    if(!current())return;
    if(reply.result.proposal){for(const key of ['rows','columns'])$('pdf-table-'+key).value=reply.result.proposal[key].map(v=>Number((v*100).toFixed(6))).join(',');$('pdf-table-header-rows').value='1';$('pdf-table-units').value='';$('pdf-table-notes').value='';}
    showUnitReview(reply.result);paintPdfGrid();show('pdf-table-result',reply.result.proposal?{candidate:reply.result.proposal,unit_candidates:reply.result.unit_candidates,unit_review:reply.result.unit_review,footnote_candidates:reply.result.footnote_candidates,uncertainties:reply.result.uncertainties}:reply.result.uncertainties);
    changed();say(reply.result.proposal?'已提出原生字符对齐候选；请校正表头、单位和边界，再提取并采用。':'未形成可靠简单网格，请回到手工边界；没有将候选当作可信数值。');
  });
  $('pdf-table-region').oninput=()=>{pdfTable=null;pdfGridHistory=[];fillObservations([]);$('pdf-table-cells').replaceChildren();changed();};
  for(const id of ['pdf-table-file','pdf-table-page'])$(id).addEventListener('input',()=>window.OfficePdfPreview?.clear($('pdf-table-image')));
  for(const id of ['pdf-table-file','pdf-table-page','pdf-table-region'])$(id).addEventListener('input',()=>{clearUnitReview();$('pdf-table-units').value='';});
  const checkHome=$('check').parentNode;
  window.OfficeContinuous={
    async mountTarget(parent, locator, offset, filename, page){
      if(!ready()||busy||fused)return false;
      if((insertions.length||comments.length||replacements.length)&&!window.confirm('切换至工作台核验将放弃原面板未写入草稿，是否继续？'))return false;
      insertions=[];comments=[];replacements=[];numericReviewIds=[];pendingNumericReview=null;changed();let okay=false;
      await run(async(host,current)=>{await load(host,current);if(!current()||!source)return;if(offset)await load(host,current,offset);if(!current()||!anchors.has(locator))return;
        $('anchor').value=locator;show('original',anchors.get(locator).text);navigate('check');parent.append($('check'));$('pdf-table-section').hidden=false;$('pdf-table-toggle').setAttribute('aria-expanded','true');
        $('pdf-table-file').value=filename;$('pdf-table-file').dispatchEvent(new Event('input'));$('pdf-table-page').value=String(page);okay=true;});
      return okay;
    },
    unmount(){if($('check').parentNode!==checkHome)checkHome.append($('check'));}
  };
  $('table-load').onclick = () => run(async(host,current)=>{pdfTable=null;const reply=await host.read_numeric_source($('table-file').value.trim());if(!current())return;if(!reply?.ok)return fail(reply);fillObservations(reply.result.observations);show('check-result',reply.result);});
  $('verify').onclick = () => run(async(host,current)=>{const reply=await host.verify_paragraph_numeric($('anchor').value,{observation_id:$('observation').value,value:$('value').value.trim(),unit:$('unit').value.trim(),baseline_id:$('baseline').value||null,operation:$('operation').value,statistic:$('statistic').value,sample:$('sample').value.trim(),time:$('time').value.trim(),conditions:$('conditions').value.trim()});if(!current())return;if(!reply?.ok)return fail(reply);check=reply.result;show('check-result',reply.result);window.dispatchEvent(new Event('officeworkbench:numeric'));say('核验为未审阅结果；只核对所选数值及填写的范围，脚注和统计口径见结果，不证明整段结论。');});
  $('check-comment').onclick=()=>{if(!check){say('请先核验当前段落。');return;}pendingNumericReview=check.numeric_review_id||null;$('comment-quote').value=$('value').value;$('comment-text').value=check.suggestion+(pendingNumericReview?' 来源：'+check.evidence.source+'；'+check.evidence.location+'；版本：'+check.evidence.source_sha256+'。布局确认不等于事实审阅。':'');navigate('write');say('核验问题已填入批注草稿，请核对原文范围后点击加入。');};
  for(const id of ['value','unit','statistic','sample','time','conditions','baseline','observation','operation']) $(id).addEventListener('input',()=>{pendingNumericReview=null;});
  $('ignore').onclick=()=>{if(!check||!$('ignore-reason').value.trim()){say('请先核验并填写明确忽略原因。');return;}ignored.push({check,reason:$('ignore-reason').value.trim()});draftView();say('已记录在本次草稿；尚未持久归档。');};
  for(const id of ['ranges','check','write']) $('nav-'+id).onclick=()=>navigate(id);
  for(const id of ['replacement','replace-reason','pdf-file','text','side','bib-anchor','comment-quote','comment-text','applicability','limitation','reviewed','value','unit','operation','statistic','sample','time','conditions','baseline','observation','table-file']) $(id).addEventListener('input',changed);
  for(const event of ['officeauth:unauthenticated','officeauth:authenticated','pagehide']) window.addEventListener(event,reset);
  for(const id of ['file-name','file-dir']) { const node=document.getElementById(id);if(node)new MutationObserver(reset).observe(node,{childList:true,characterData:true,subtree:true}); }
  refresh();
})();
