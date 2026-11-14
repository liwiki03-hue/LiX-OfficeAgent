(() => {
  'use strict';
  const ask = message => window.OfficeConfirm?.ask(message) ?? Promise.resolve(false);
  const $ = id => document.getElementById('grounded-' + id);
  if (!$('panel')) return;
  let epoch = 0, busy = false, unknown = false, version = null, capture = null, ocrCapture = null, plan = null, saveId = null;
  const auth = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const errors = {
    not_authenticated: '请先登录。', busy: '任务正在进行，请稍后操作。',
    professional_metadata_incomplete: '仅可使用明确纳入、题名/作者/年份/DOI完整且未报告撤稿的文献。',
    professional_evidence_path: '请输入目标 DOCX 同目录的 PDF 文件名，不含路径。',
    professional_evidence_changed: 'PDF 来源已改变，旧计划作废，请重新取证。',
    professional_evidence_invalid: '证据或纳入状态已过期，请重新读取并核对。',
    professional_evidence_empty: '前5页未取得可用文本层；未做 OCR，请改用有文本层的来源。',
    professional_support_required: '请选1至3个原文片段，核对支持/反对关系与条件后勾选人工确认。',
    professional_identity_required: '请先人工核对 PDF 身份。',
    professional_archive_changed: '归档回执不存在、已撤销或发生变化，请核对保存结果；未读取任意子路径。',
    professional_existing_numbered_citations: '文档已有数字引文，本试点不自动重排；请选择无数字引文的草稿。',
    professional_grounded_target: '请选择两个不同的普通正文段落；书目占位不得含数字、公式或引文。',
    professional_unsupported_fact: '新增数字、单位或公式未在所选原文中出现，请核对证据。',
    professional_protected_numbers: '原段落数字或单位被删除、改写或调序，请保留后重试。',
    professional_protected_formulas: '原段落公式被改动，请保留。',
    professional_source_changed: 'DOCX 已改变，请重新读取结构。',
    professional_stale_locator: '目标定位已过期或超出前40项，请重新读取。',
    professional_plan_invalid: '计划不可用，请重新生成。', professional_plan_expired: '计划已过期，请重新生成。',
    professional_invalid_request: '请完整填写主张、范围、限制及原文选择。',
    professional_citation_invalid: '引文校验失败，请勿在主张中手写编号、DOI或其他引文。',
    professional_sensitive_text: '文本包含疑似敏感内容，已拒绝。'
  };
  function status(text) { $('status').textContent = text; }
  function refresh() {
    const blocked = busy || unknown || !auth();
    $('panel').setAttribute('aria-busy', String(busy));
    $('panel').querySelectorAll('input,textarea,select,button').forEach(node => { node.disabled = blocked; });
    $('capture').disabled = blocked || !version || !$('source').value;
    $('plan').disabled = blocked || !capture;
    $('confirm').disabled = blocked || !plan;
    $('rollback').disabled = blocked || !saveId;
    $('continue').disabled = blocked || !capture?.coverage || capture.coverage.next_page === null;
    $('ocr').disabled = blocked || !capture;
    $('ocr-review').disabled = blocked || !ocrCapture;
    $('ocr-view').disabled = blocked || !ocrCapture;
  }
  function changed() { plan = null; $('preview').textContent = ''; refresh(); }
  function clearOCR() { ocrCapture = null; $('ocr-raw').textContent = ''; $('ocr-corrections').replaceChildren(); $('ocr-reviewed').checked = false; $('page-view').replaceChildren(); }
  function reset() {
    epoch++; busy = false; version = capture = ocrCapture = plan = saveId = null;
    ['source', 'body', 'reference', 'evidence', 'preview', 'output', 'page-view'].forEach(id => { $(id).replaceChildren(); });
    ['file', 'archive', 'claim', 'applicability', 'limitation'].forEach(id => { $(id).value = ''; });
    $('identity').checked = $('reviewed').checked = false; $('relation').value = 'insufficient';
    $('coverage').textContent = ''; $('page').value = '1'; $('page-limit').value = '5';
    $('ocr-raw').textContent = ''; $('ocr-corrections').replaceChildren(); $('ocr-reviewed').checked = false;
    status('尚未读取；所有写入均需独立确认。'); refresh();
  }
  function uncertain() {
    unknown = true; plan = saveId = null; status('结果不确定，可能已有输出或备份；请核对记录，不自动重试。');
  }
  function failed(reply) {
    if (reply?.reason === 'professional_uncertain') uncertain();
    else status(errors[reply?.reason] || '操作失败或被安全检查拒绝；请核对文件、来源与普通段落限制。');
  }
  async function run(operation) {
    if (busy || unknown || !auth()) return;
    const mine = epoch; let approvedWrite = false;
    busy = true; refresh();
    const current = () => mine === epoch && auth();
    try { await operation(window.OfficeAuth.wrapApi(window.pywebview.api), current, () => { approvedWrite = true; }); }
    catch (_) { if (current()) { if (approvedWrite) uncertain(); else status('读取或规划失败；尚未批准本次写入。'); } }
    finally { if (mine === epoch) { busy = false; refresh(); } }
  }
  function option(select, value, label) {
    const node = document.createElement('option'); node.value = value; node.textContent = label; select.append(node);
  }
  $('load').addEventListener('click', async () => {
    if (capture && await ask('重新读取将丢弃当前证据选择与待确认计划，是否继续？') !== true) return;
    run(async (api, current) => {
      status('正在读取纳入条目与目标结构…');
      const sources = await api.get_grounded_sources();
      if (!current()) return;
      if (!sources?.ok) { failed(sources); return; }
      const structure = await api.read_professional_structure(0);
      if (!current()) return;
      if (!structure?.ok) { failed(structure); return; }
      version = structure.result.document.source_sha256; capture = null; clearOCR(); changed();
      ['source', 'body', 'reference', 'evidence'].forEach(id => $(id).replaceChildren());
      for (const row of sources.result.records) option($('source'), row.id, row.title + ' · ' + (row.doi || '缺 DOI'));
      for (const row of structure.result.items.filter(row => row.writable && /^body:\d+\/paragraph$/.test(row.location) && !row.style)) {
        for (const id of ['body', 'reference']) option($(id), row.locator_id, row.location + '：' + row.text.slice(0, 100));
      }
      if ($('reference').options.length > 1) $('reference').selectedIndex = 1;
      status(sources.result.records.length ? '已读取；输入同目录 PDF，原文需人工核对。' : '暂无纳入条目，请先在文献面板检索并填写纳入理由。');
    });
  });
  $('capture').addEventListener('click', () => run(async (api, current) => {
    clearOCR();
    capture = null; changed(); $('evidence').replaceChildren(); $('identity').checked = $('reviewed').checked = false;
    status('正在读取 PDF 文本层…');
    const reply = await api.capture_grounded_evidence($('source').value, $('file').value);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    renderCapture(reply.result);
  }));
  function renderCapture(result) {
    capture = result;
    if (capture.archive_save_id) {
      option($('source'), capture.document.literature_id, capture.title + ' · 归档来源');
      $('source').value = capture.document.literature_id;
    }
    for (const row of capture.fragments) {
      const label = document.createElement('label'), input = document.createElement('input'), text = document.createElement('pre');
      input.type = 'checkbox'; input.value = row.id; input.addEventListener('change', () => { $('reviewed').checked = false; changed(); });
      text.textContent = row.location + '\n' + row.original_text;
      label.append(input, text); $('evidence').append(label);
      const view = document.createElement('button'); view.type = 'button'; view.textContent = '在原始 PDF 页查看'; view.className = 'btn secondary';
      view.addEventListener('click', () => run(async (api, current) => { await window.OfficeEvidenceViewer.show(api, capture.capture_id, row.id, $('page-view'), current); }));
      $('evidence').append(view);
    }
    if (capture.coverage) $('coverage').textContent = '总页数：' + capture.coverage.page_count + '；完整读取页：' + capture.coverage.complete_pages.join('、') + '\n' + (capture.coverage.complete ? '已覆盖全部页面的文本层；不代表语义/多栏/数值可靠。' : '当前覆盖不完整，不可宣称全文未发现。') + '\n索引覆盖：读取未发布索引；空文本层页：' + capture.coverage.empty_pages.join('、') + '\n页段身份：' + capture.coverage.range_id;
    status((capture.archive_save_id ? '归档文献：' + capture.title + '；许可：' + capture.license + '；来源：' + capture.document.source_url + '。' : '') + '已读取 ' + capture.fragments.length + ' 个原文片段；' + (capture.truncated ? '有未读取页或字符，请查看覆盖记录并主动续读。' : '本次文本层读取完成。') + '请选1至3项，坐标为 PDF 点、左下角原点；不推断多栏阅读顺序。');
  }
  function readRange(continuing) {
    run(async (api, current) => {
      clearOCR();
      const page = continuing ? capture.coverage.next_page : Number($('page').value) - 1;
      const offset = continuing ? capture.coverage.next_char_offset : 0;
      const limit = Number($('page-limit').value);
      if (!Number.isInteger(page) || page < 0 || page > 1999 || !Number.isInteger(limit) || limit < 1 || limit > 5) { status('请输入有效起始页与1至5页预算。'); return; }
      capture = null; changed(); $('evidence').replaceChildren(); $('identity').checked = $('reviewed').checked = false;
      const receipt = $('archive').value.trim();
      const reply = receipt ? await api.capture_archive_range(receipt, page, limit, offset) : await api.capture_grounded_range($('source').value, $('file').value, page, limit, offset);
      if (!current()) return;
      if (!reply?.ok) { failed(reply); return; }
      renderCapture(reply.result);
    });
  }
  $('range').addEventListener('click', () => readRange(false));
  $('continue').addEventListener('click', () => readRange(true));
  $('reading-status').addEventListener('click', () => run(async (api, current) => { const reply = await api.get_pdf_reading_status(); if (!current()) return; if (!reply?.ok) { failed(reply); return; } $('coverage').textContent = JSON.stringify(reply.result.sources, null, 2); status('仅显示本次来源版本的阅读覆盖、失败范围与索引状态；不自动补读或联网。'); }));
  $('ocr').addEventListener('click', async () => {
    if (!await ask('使用本机已安装的 Windows OCR 识别所选页？结果可能含数字、单位及阅读顺序错误，必须对照原页核对。')) return;
    run(async (api, current) => {
      const reply = await api.ocr_evidence_page(capture.capture_id, Number($('ocr-page').value));
      if (!current()) return;
      if (!reply?.ok) { status(reply?.reason === 'professional_ocr_unavailable' ? '本机 OCR 依赖、语言或受限进程不可用，未自动安装或下载。' : '无法识别该页，请核对当前页段与来源。'); return; }
      ocrCapture = reply.result; capture = null; changed(); $('evidence').replaceChildren(); $('ocr-corrections').replaceChildren(); $('ocr-reviewed').checked = false;
      $('ocr-raw').textContent = ocrCapture.ocr.raw.warning + '\n原始 OCR：\n' + ocrCapture.ocr.raw.raw_text;
      for (const fragment of ocrCapture.fragments) {
        const label = document.createElement('label'), text = document.createElement('textarea'); label.textContent = fragment.location + ' · 人工校正'; text.value = fragment.original_text; text.maxLength = 2000; text.rows = 2; text.dataset.evidenceId = fragment.id;
        text.addEventListener('input', () => { capture = null; $('ocr-reviewed').checked = false; $('evidence').replaceChildren(); changed(); }); label.append(text); $('ocr-corrections').append(label);
      }
      status('已识别，结果尚未人工核对。请对照原页校正后再纳入证据；当前不能写入。');
    });
  });
  $('ocr-review').addEventListener('click', () => run(async (api, current) => {
    if (!$('ocr-reviewed').checked) { status('请先对照原页核对数字、单位及符号并明确确认。'); return; }
    const corrections = [...$('ocr-corrections').querySelectorAll('textarea')].map(node => ({evidence_id: node.dataset.evidenceId, text: node.value}));
    const reply = await api.review_evidence_ocr({capture_id: ocrCapture.capture_id, corrections, reviewed: true});
    if (!current()) return; if (!reply?.ok) { failed(reply); return; }
    $('evidence').replaceChildren(); $('identity').checked = $('reviewed').checked = false; changed(); renderCapture(reply.result);
    status('已保留原始 OCR 与独立人工校正；来源身份及主张支持关系仍未审阅，请重新核对。');
  }));
  $('ocr-view').addEventListener('click', () => run(async (api, current) => { await window.OfficeEvidenceViewer.show(api, ocrCapture.capture_id, ocrCapture.fragments[0].id, $('page-view'), current); }));
  $('archive-capture').addEventListener('click', () => run(async (api, current) => {
    clearOCR();
    capture = null; changed(); $('evidence').replaceChildren(); $('identity').checked = $('reviewed').checked = false;
    status('正在校验归档回执并读取 PDF 文本层…');
    const reply = await api.capture_grounded_archive($('archive').value);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    renderCapture(reply.result);
  }));
  ['source', 'file'].forEach(id => $(id).addEventListener('input', () => { capture = null; clearOCR(); $('evidence').replaceChildren(); changed(); }));
  ['body', 'reference', 'claim', 'relation', 'applicability', 'limitation', 'identity'].forEach(id => $(id).addEventListener('input', () => { $('reviewed').checked = false; changed(); }));
  $('reviewed').addEventListener('input', changed);
  $('plan').addEventListener('click', () => run(async (api, current) => {
    const payload = {source_sha256: version, capture_id: capture.capture_id, body_locator: $('body').value,
      reference_locator: $('reference').value, identity_reviewed: $('identity').checked,
      support: {claim: $('claim').value, evidence_ids: [...$('evidence').querySelectorAll('input:checked')].map(node => node.value),
        relation: $('relation').value, applicability: $('applicability').value, limitation: $('limitation').value,
        assessment: $('reviewed').checked ? 'user_reviewed' : 'unverified'}};
    plan = null; status('正在检查证据版本、保护项与 CSL 引文…');
    const reply = await api.plan_grounded_edit(payload);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    plan = reply.result;
    const before = JSON.parse(plan.before).items;
    const lines = plan.items.map((item, index) => (index === 0 ? '正文' : '书目') + ' · ' + item.location +
      '\n修改前：' + before.find(row => row.location === item.location).text + '\n修改后：' + item.text);
    const support = plan.items[0].claim_support;
    lines.push('人工判断：' + (support.relation === 'supports' ? '支持' : '反对') + '\n适用条件：' + support.applicability + '\n限制与缺口：' + support.limitation);
    for (const row of plan.items[0].evidence_fragments) lines.push('原文证据 · 第' + row.page + '页 · ' + row.location + '\n' + row.original_text);
    lines.push('核验范围：人工已核对，未作机器语义证明。\n审阅记录：' + plan.semantic_verification + '\n样式版本：' + plan.style_version);
    $('preview').textContent = lines.join('\n\n');
    status('计划已生成，尚未写入。请核对原文、最终差异、支持关系、条件与限制；5分钟后失效。');
  }));
  $('confirm').addEventListener('click', () => run(async (api, current, writing) => {
    const target = plan; plan = null;
    const approved = await ask('确认按已展示的主张、证据和书目写入新副本？人工判断不等于机器证明。') === true;
    if (!current()) return;
    if (approved) writing();
    const reply = await api.confirm_grounded_edit(target.id, approved);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    if (reply.result.status === 'cancelled') { status('已取消，未写入。'); return; }
    if (reply.result.status !== 'saved') { uncertain(); return; }
    saveId = reply.result.save_id;
    $('output').textContent = '已保存并核验新副本：' + reply.result.output + '；原件未改。渲染：' + (reply.result.render?.status || '未提供');
    status('本次副本可撤销；输出需另行选择并建立检索索引。');
  }));
  $('rollback').addEventListener('click', () => run(async (api, current, writing) => {
    const approved = await ask('确认撤销本次未被外部修改的输出副本？保留原件、备份和审计。') === true;
    if (!current()) return;
    if (approved) writing();
    const reply = await api.rollback_grounded_edit(saveId, approved);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    if (reply.result.status === 'cancelled') { status('已取消撤销。'); return; }
    if (reply.result.status !== 'rolled_back') { uncertain(); return; }
    saveId = null; $('output').textContent = '本次副本已撤销；原件、备份和审计保留。';
  }));
  ['officeauth:unauthenticated', 'officeauth:authenticated', 'pagehide'].forEach(event => window.addEventListener(event, reset));
  ['file-name', 'file-dir'].forEach(id => { const node = document.getElementById(id); if (node) new MutationObserver(reset).observe(node, {childList: true, characterData: true, subtree: true}); });
  refresh();
})();
