(() => {
  'use strict';
  const $ = id => document.getElementById('range-' + id);
  if (!$('panel')) return;
  const commentOption = document.createElement('option'); commentOption.value = 'comment'; commentOption.textContent = '给选区添加审阅批注'; $('action').append(commentOption);
  $('revision').multiple = true; $('revision').setAttribute('aria-label', '已有修订，可多选作为同一审阅组');
  let epoch = 0, busy = false, uncertain = false, data = null, selected = null, queue = [], plan = null, save = null, delivery = null, reviewDirty = false;
  const auth = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const environmentErrors = {professional_review_environment_unknown: '无法确定必要字体或配置身份，已拒绝沿用缓存；可手动核对文档，未修改系统字体。', professional_review_environment_changed: '字体、渲染器或配置已变化，或身份无法确认，已拒绝沿用页图。请重新打开可用对照以只读重验；旧回执缺少上下文时需人工核对。', professional_review_revalidation_failed: '重新渲染未通过，保留旧文档及历史状态，未写入或恢复批准。'};
  const status = value => { $('status').textContent = value; };
  const labels = {professional_protected_boundary: '选区跨越保护边界，请只选择同一可编辑范围内的文字。', professional_invalid_range: '选区为空、越界或切开了组合字符，请重新选择。', professional_mixed_format_policy: '选区跨越不同格式，请明确选择起点继承或等长分配。', professional_distributed_length: '按原位置分配格式需要新旧文字字符数相同。', professional_overlapping_changes: '清单中有重叠或冲突操作，请拆分事务或撤回冲突项。', professional_structure_scope: '此结构暂不支持该操作；书签、编号、节和特殊内容需要保留。', professional_table_scope: '此表格不在普通行修改范围内，未执行。', professional_revision_scope: '该修订类型或范围暂不支持；请保留原修订。', professional_source_changed: '源文件已改变，请重新读取并重新确认。', professional_numbering_scope: '无法可靠解析编号继承，请使用已有明确编号文段。', professional_protected_numbers: '数字或单位发生变化，保护检查拒绝。', professional_protected_citations: '引文变化，保护检查拒绝。', professional_protected_formulas: '公式变化，保护检查拒绝。', professional_no_changes: '没有实际变化。', professional_plan_invalid: '计划已失效，请重新预览。', professional_plan_expired: '计划已过期，请重新预览。', professional_field_scope: '存在尚未支持的域指令，未读取或执行。', professional_unsafe_document: '文档包含不支持的活动内容或外部关系。', not_authenticated: '请先登录。', busy: '其他任务正在进行。'};
  Object.assign(labels, {professional_review_changed: '来源、副本或审计已变化；旧审阅状态仅供历史参考，请重新核对。', professional_review_read_consent: '读取授权已失效，请重新打开对照。', professional_review_renderer_changed: '渲染器已变化，不能沿用旧页定位。', professional_review_alignment_missing: '页面无法可靠定位，请在文档中人工核对。', professional_review_expired: '审阅会话已过期，请重新打开。', professional_review_scope: '请先明确选择本回执对应的源文件或副本。', professional_reference_target: '交叉引用的书签目标缺失、重复或跨段，当前文档暂不支持安全编辑。', professional_reference_target_locked: '该文段是交叉引用目标，改写会使域结果失效；正文受保护，其他操作还须通过宿主结构检查。'});
  Object.assign(labels, environmentErrors);
  function refresh() {
    $('panel').setAttribute('aria-busy', String(busy));
    document.querySelectorAll('[id^="range-"], #range-queue button, #range-queue input').forEach(node => { if (node.matches('button,input,select,textarea')) node.disabled = busy || uncertain || !auth(); });
    $('next').disabled ||= data?.next_offset == null;
    const support = selected?.action_support?.[$('mode').value] || {};
    for (const item of $('action').options) { item.disabled = Boolean(support[item.value]); item.title = labels[support[item.value]] || support[item.value] || '可选择；仍需选区及整组校验'; }
    $('add').disabled ||= !selected || Boolean(support[$('action').value]);
    $('compose').disabled ||= !selected;
    $('plan').disabled ||= !queue.some(item => item.included);
    $('confirm').disabled ||= !plan;
    $('review').disabled ||= !plan;
    $('undo').disabled ||= !save;
    $('open-review').disabled ||= !save && !delivery && !/^[a-f0-9]{32}$/.test($('delivery-id').value.trim());
    $('save-review').disabled ||= !delivery?.opened;
    $('mark-review').disabled ||= !delivery?.opened;
    $('review-item').disabled ||= !delivery;
    $('next-review').disabled ||= !delivery || Number($('review-item').value) + 1 >= (delivery?.items?.length || 0);
  }
  function invalidate() { plan = null; $('preview').textContent = ''; $('assessment').textContent = ''; refresh(); }
  function reset() {
    epoch++; busy = uncertain = false; data = selected = plan = save = delivery = null; queue = []; reviewDirty = false;
    ['before-pages', 'after-pages', 'review-item', 'review-note'].forEach(id => $(id).replaceChildren());
    $('archive-id').value = $('review-reason').value = ''; $('review-state').value = 'unreviewed';
    ['original', 'text', 'reason'].forEach(id => { $(id).value = ''; });
    ['target', 'numbering', 'revision', 'island', 'queue', 'structures', 'preview', 'output', 'assessment'].forEach(id => { $(id).replaceChildren(); });
    status('会话已清理；旧批准不恢复。'); refresh();
  }
  function fail(reply) { status(labels[reply?.reason] || '请求被拒绝，请核对结构、来源版本和输入；没有新的写入批准。'); }
  async function run(write, operation) {
    if (busy || uncertain || !auth()) return;
    const mine = epoch; busy = true; refresh();
    const current = () => mine === epoch && auth();
    try { await operation(window.OfficeAuth.wrapApi(window.pywebview.api), current); }
    catch (_) { if (current()) { if (write) unknown(); else status('读取或规划失败，请检查本机宿主。'); } }
    finally { if (mine === epoch) { busy = false; refresh(); } }
  }
  function unknown() { uncertain = true; plan = save = null; status('结果不确定，可能已产生输出；请核查审计，不自动重试。'); }
  function option(parent, value, text) { const node = document.createElement('option'); node.value = value; node.textContent = text; parent.append(node); }
  function selection() {
    const text = $('original').value;
    const cutsSurrogate = offset => offset > 0 && offset < text.length && text.charCodeAt(offset - 1) >= 0xD800 && text.charCodeAt(offset - 1) <= 0xDBFF && text.charCodeAt(offset) >= 0xDC00 && text.charCodeAt(offset) <= 0xDFFF;
    if (cutsSurrogate($('original').selectionStart) || cutsSurrogate($('original').selectionEnd)) { $('selection').textContent = '选区切开了 UTF-16 代理对，请重新选择完整字符。'; return {start: -1, end: -1, quote: ''}; }
    const start = Array.from(text.slice(0, $('original').selectionStart)).length;
    const end = Array.from(text.slice(0, $('original').selectionEnd)).length;
    $('selection').textContent = `Unicode 字符范围 [${start}, ${end})；位置 ${selected?.location || '未选'}。`;
    return {start, end, quote: Array.from(text).slice(start, end).join('')};
  }
  function choose() {
    selected = data?.items.find(row => row.locator_id === $('target').value) || null;
    $('original').value = selected?.text || '';
    $('revision').replaceChildren();
    $('island').replaceChildren(); option($('island'), '', '自动（仅唯一匹配）');
    for (const row of selected?.islands || []) option($('island'), row.id, `[${row.start}, ${row.end})${row.hyperlink ? ' 链接显示文字' : ' 正文'}`);
    for (const row of selected?.revisions || []) option($('revision'), row.id, `${row.type} · ${row.author} · ${row.date} · ${row.text}${row.supported ? '' : '（只读）'}`);
    $('structures').textContent = '保护结构：' + (selected?.protected_structures.join('、') || '无') + '\n可编辑范围：' + (selected?.islands || []).map(row => `[${row.start}, ${row.end})${row.hyperlink ? ' 链接显示文字' : ''}`).join('；');
    if (selected?.fields?.length) $('structures').textContent += '\n只读域：' + selected.fields.map(row => row.type + (row.target ? ' → ' + row.target : '')).join('；') + '。显示结果保持，不代表已刷新或核实；分页改变后需在 Word 核对。';
    if (selected?.reference_target_locked) $('structures').textContent += '\n本段是交叉引用目标：禁止改写正文；各操作是否可选以宿主结构检查为准。';
    selection(); refresh();
  }
  function renderQueue() {
    $('queue').replaceChildren();
    queue.forEach((item, index) => {
      const row = document.createElement('div'), check = document.createElement('input'), text = document.createElement('pre'), remove = document.createElement('button'), edit = document.createElement('button');
      row.dataset.reviewState = item.included ? 'adopted' : 'pending';
      check.type = 'checkbox'; check.checked = item.included; check.setAttribute('aria-label', '纳入第' + (index + 1) + '项');
      check.onchange = () => { item.included = check.checked; row.dataset.reviewState = item.included ? 'adopted' : 'pending'; invalidate(); };
      text.textContent = `${index + 1}. ${item.label}\n原文：${item.op.quote}\n新文：${item.op.text}\n原因：${item.op.reason}`;
      remove.type = edit.type = 'button'; remove.textContent = '撤回'; edit.textContent = '调整文字与原因';
      remove.onclick = () => { queue.splice(index, 1); invalidate(); renderQueue(); };
      edit.onclick = () => {
        const value = window.prompt('调整本项新文字（不改变原选区）', item.op.text);
        if (value === null) return;
        const reason = window.prompt('修改原因；调整后必须重新预览', item.op.reason);
        if (reason === null) return;
        item.op.text = value; item.op.reason = reason; invalidate(); renderQueue();
      };
      const locate = document.createElement('button'); locate.type = 'button'; locate.textContent = '定位原文';
      locate.onclick = () => {
        if (busy || uncertain || !auth()) return;
        if (!Array.from($('target').options).some(node => node.value === item.op.locator_id)) { status('目标不在当前读取页，请重新读取对应文段；不会按相同文字猜测位置。'); return; }
        $('target').value = item.op.locator_id; choose();
        window.dispatchEvent(new CustomEvent('officeworkspace:locate'));
        $('original').focus();
      };
      row.append(check, text, locate, edit, remove); $('queue').append(row);
    }); refresh();
  }
  function readablePreview(candidate) {
    const before = JSON.parse(candidate.before).items, after = JSON.parse(candidate.after).items;
    const lines = [candidate.native_track_changes ? '原生文字审阅稿：本轮接受/拒绝投影单独核验，保留无关已有修订。' : '净稿：保留无关已有修订；结构及格式没有伪装成原生修订。'];
    const formatNames = {bold: '粗体', italic: '斜体', underline: '下划线', font: '字体', size_pt: '字号（磅）', style: '已有样式', alignment: '对齐', before_pt: '段前（磅）', after_pt: '段后（磅）', line_multiple: '行距倍数', left_pt: '左缩进（磅）', first_line_pt: '首行缩进（磅）'};
    const describeFormat = value => Object.entries(value || {}).map(([key, item]) => (formatNames[key] || key) + '：' + (typeof item === 'boolean' ? (item ? '是' : '否') : item)).join('；');
    after.forEach((item, index) => {
      const original = before[index] || {};
      const label = Array.from($('action').options).find(option => option.value === item.action)?.textContent || item.action;
      lines.push(`\n${index + 1}. ${label} · ${original.location || '当前目标'}\n原文：${original.text || '空位置'}`);
      if (item.action === 'format' || item.action === 'paragraph_format') lines.push('正文保持；格式调整：' + describeFormat(item.character_format || item.paragraph_format));
      else if (item.action.startsWith('revision_')) lines.push('审阅范围：显式选择的 ' + (item.reviewed_revisions?.length || 0) + ' 项修订；未选修订保留。');
      else if (item.action === 'comment') lines.push('正文保持；批注：' + item.text);
      else if (item.action === 'paragraph_delete' || item.action === 'row_delete') lines.push('删除内容：' + (original.removed_row_text || original.text));
      else if (item.action === 'merge_next') lines.push('与下一段合并：' + original.next_paragraph_text);
      else if (item.action === 'split') lines.push('在已选光标处分拆，保留原文字与字符格式。');
      else if (item.action === 'row_insert') lines.push('新增一个空行；保留列网格和所选普通行的格式。');
      else if (!item.action.startsWith('list_')) lines.push('新文字：' + item.text);
      if (item.numbering_impact) lines.push('实际编号级别：' + (Number(item.numbering_impact.level) + 1) + '；同编号流的其他登记文段：' + item.numbering_impact.other_count + '。后续显示编号可能变化，需核对渲染结果。');
      if (['replace', 'insert'].includes(item.action)) lines.push('格式继承：' + ({start: '新文字沿用选区起点格式', distributed: '等长，按原字符位置分配格式', reject_mixed: '只允许明确同格式范围'}[item.inheritance] || item.inheritance));
      lines.push('原因：' + item.reason);
    });
    const verification = after[0]?.verification;
    if (verification) lines.push('\n整组核对：请求 ' + verification.requested + ' 项；未解释变化 ' + verification.unexplained.length + ' 项。\n派生变化：必要的run分割、空格及属性排序；如请求批注/编号则增加对应部件或编号实例。\n保持的包部件：' + verification.preserved_parts + '；变化部件：' + verification.changed_parts.join('、') + '。');
    if (verification?.reference_impact) {
      const impact = verification.reference_impact;
      lines.push('\n引用影响（结构依赖，未刷新域结果）：' + impact.needs_review_total + '/' + impact.fields_total + ' 项需核对。');
      for (const item of impact.items) lines.push(item.type + ' → ' + (item.target_name || '页码/目录') + '；域位置：' + (item.field_location || item.part) + '；目标：' + (item.target_location || '未登记') + '；' + item.reasons.map(reason => ({reference_surroundings: '域周边有修改', target_presentation: '目标呈现改变', possible_pagination: '可能影响分页', no_direct_intersection: '无直接交集'}[reason] || reason)).join('、'));
      if (impact.truncated) lines.push('清单已截断，不是完整依赖图。');
    }
    if (verification?.fields?.preserved) lines.push('\n保留域 ' + verification.fields.preserved + ' 项，其中交叉引用 ' + verification.fields.cross_references + ' 项；未刷新或核实显示结果。文字或格式变化可能引起分页变化，PAGE/PAGEREF/目录需在 Word 中主动核对。');
    lines.push('\n数字、引用和公式使用当前勾选的保护规则；不证明语义等价。原文件保持，输出需重开、静态域渲染、审计后发布。Word人工及复杂视觉异常仍未验证。');
    return lines.join('\n');
  }
  function load(offset) {
    if (queue.length && !window.confirm('重新读取将放弃当前待修改清单，是否继续？')) return;
    run(false, async (api, current) => {
      const reply = await api.read_edit_ranges(offset);
      if (!current()) return;
      if (!reply?.ok) { fail(reply); return; }
      data = reply.result; queue = []; invalidate(); renderQueue();
      $('target').replaceChildren(); $('numbering').replaceChildren();
      for (const item of data.items) { option($('target'), item.locator_id, item.location + ' · ' + item.text.slice(0, 70)); option($('numbering'), item.locator_id, item.location + ' · ' + item.text.slice(0, 40)); }
      $('style').replaceChildren(); option($('style'), '', '保持'); data.styles.forEach(name => option($('style'), name, name));
      choose(); status('已读取版本绑定视图。未支持结构只保留；坐标采用 Unicode 字符，非 UTF-16。');
    });
  }
  $('load').onclick = () => load(0);
  $('next').onclick = () => load(data?.next_offset || 0);
  $('target').onchange = choose;
  $('action').onchange = refresh;
  ['select', 'click', 'keyup'].forEach(event => $('original').addEventListener(event, selection));
  $('add').onclick = () => {
    if (!selected || busy) return;
    const range = selection(), action = $('action').value;
    if (range.start < 0) { status('请重新选择完整 Unicode 字符。'); return; }
    const matches = selected.islands?.filter(row => row.start <= range.start && range.end <= row.end) || [];
    const island = $('island').value ? matches.find(row => row.id === $('island').value) : (matches.length === 1 ? matches[0] : null);
    if (['replace', 'insert', 'delete', 'format', 'comment'].includes(action) && !island) { status('选区跨保护边界或光标同时属于多个范围，请明确选择范围。'); return; }
    const op = {locator_id: selected.locator_id, expected_text_sha256: selected.text_sha256, action, ...range, text: ['replace', 'insert', 'insert_before', 'insert_after', 'comment'].includes(action) ? $('text').value : '', reason: $('reason').value, inheritance: $('inheritance').value};
    if (['replace', 'insert', 'delete', 'format', 'comment'].includes(action)) op.island_id = island?.id || null;
    if (action === 'format') {
      op.character_format = {};
      for (const name of ['bold', 'italic', 'underline']) if ($(name).value) op.character_format[name] = $(name).value === 'true';
      if ($('font').value) op.character_format.font = $('font').value;
      if ($('size').value) op.character_format.size_pt = Number($('size').value);
    }
    if (action === 'paragraph_format') {
      op.paragraph_format = {};
      if ($('style').value) op.paragraph_format.style = $('style').value;
      if ($('align').value) op.paragraph_format.alignment = $('align').value;
      for (const [id, key] of [['before', 'before_pt'], ['after', 'after_pt'], ['line', 'line_multiple'], ['left', 'left_pt'], ['first', 'first_line_pt']]) if ($(id).value) op.paragraph_format[key] = Number($(id).value);
    }
    if (action.startsWith('revision_')) op.revision_ids = Array.from($('revision').selectedOptions).map(node => node.value);
    if (action.startsWith('list_')) op.numbering_locator_id = $('numbering').value;
    queue.push({op, included: true, label: $('action').selectedOptions[0].textContent + ' · ' + selected.location});
    invalidate(); renderQueue(); status('已加入待修改清单，尚未批准写入。');
  };
  $('plan').onclick = () => run(false, async (api, current) => {
    const protect = $('protect').checked;
    const reply = await api.plan_range_edit({source_sha256: data.document.source_sha256, operations: queue.filter(item => item.included).map(item => item.op), mode: $('mode').value, preserve_numbers: protect, preserve_citations: protect, preserve_formulas: protect});
    if (!current()) return;
    if (!reply?.ok) { fail(reply); return; }
    plan = reply.result;
    $('preview').textContent = readablePreview(plan);
    status('整组影响已生成；接受建议不等于写入批准。');
  });
  $('confirm').onclick = () => run(true, async (api, current) => {
    const target = plan; if (!target) return;
    const approved = window.confirm('独立确认生成 DOCX 副本？请核对全部影响。原文件保持，取消不写入。\n' + $('preview').textContent) === true;
    if (!current()) return;
    plan = null;
    const reply = await api.confirm_professional_edit(target.id, approved);
    if (!current()) return;
    if (!reply?.ok) { fail(reply); return; }
    if (reply.result.status === 'cancelled') { status('已取消，没有批准写入。'); return; }
    if (reply.result.status !== 'saved') { unknown(); return; }
    save = reply.result.save_id; delivery = null; reviewDirty = false;
    $('delivery-id').value = save;
    ['before-pages', 'after-pages', 'review-item', 'review-note'].forEach(id => $(id).replaceChildren());
    $('output').textContent = '已保存副本：' + reply.result.output + '\n原文件未变；已重开、渲染、审计。渲染采用静态域显示；Word 人工未验证。\n内容锚点前后页定位（重复或跨页不会猜测）：\n' + JSON.stringify(reply.result.alignment || [], null, 2) + '\n表格溢出、图片漂移、标题孤立等仍需视觉审阅。请重新选择副本继续下一轮。';
    status('已生成可审阅副本，本次可独立撤销。');
  });
  $('review').onclick = () => run(false, async (api, current) => {
    if (!plan) return;
    const reply = await api.review_range_plan(plan.id);
    if (!current()) return;
    if (!reply?.ok) { status('模型辅助核对不可用或本计划已调用；确定性保护仍有效，没有写入批准。'); return; }
    $('assessment').textContent = '模型未审阅建议，不是写入批准或事实证明：\n' + JSON.stringify(reply.result.assessment, null, 2);
  });
  $('compose').onclick = () => {
    if (!selected || !data || busy) return;
    const prompt = document.getElementById('prompt');
    const value = '请先读取当前DOCX宿主范围视图，只预览本选区修改，不执行写入。修改要求：' + $('reason').value + '\n宿主选区：' + JSON.stringify({source_sha256: data.document.source_sha256, locator_id: selected.locator_id, expected_text_sha256: selected.text_sha256, ...selection(), island_id: $('island').value || null});
    if (prompt.value.trim() && !window.confirm('替换尚未发送的对话输入？')) return;
    prompt.value = value; prompt.dispatchEvent(new Event('input', {bubbles: true})); prompt.focus();
    status('选区和要求已填入对话。请启用现有专业工具后主动发送；模型只能预览，写入仍需独立确认。');
  };
  $('undo').onclick = () => run(true, async (api, current) => {
    const removed = save;
    const approved = window.confirm('独立确认撤销本次输出？保留源文件、备份及审计。') === true;
    if (!current()) return;
    const reply = await api.rollback_professional_edit(save, approved);
    if (!current()) return;
    if (!reply?.ok) { fail(reply); return; }
      if (reply.result.status === 'rolled_back') { if ($('delivery-id').value.trim() === removed) $('delivery-id').value = ''; save = delivery = null; reviewDirty = false; ['before-pages', 'after-pages', 'review-item'].forEach(id => $(id).replaceChildren()); $('output').textContent = '本次副本已撤销。'; status('本次副本已撤销，原稿和备份保留。'); }
    else if (reply.result.status !== 'cancelled') unknown();
  });
  ['mode', 'protect'].forEach(id => $(id).addEventListener('change', invalidate));
  const consent = () => window.confirm('读取本回执对应的原稿和副本，生成有界静态页面对照？仅限当前明确选择的文档对，不联网、不刷新域、不恢复写入或撤销批准。') === true;
  function showReview() {
    $('review-item').replaceChildren();
    const states = {unreviewed: '未审阅', viewed: '已查看', needs_review: '待复核', ignored: '明确忽略'};
    for (const [index, item] of (delivery?.items || []).entries()) {
      const action = delivery.alignment?.[index]?.action;
      const label = Array.from($('action').options).find(node => node.value === action)?.textContent || '历史审阅项';
      option($('review-item'), String(index), `${index + 1}. ${label} · ${states[item.state] || '需重新核对'}`);
    }
    if (delivery?.items?.[0]) { $('review-state').value = delivery.items[0].state; $('review-reason').value = delivery.items[0].reason; }
    $('review-note').textContent = delivery?.notice || '历史审阅记录已恢复，尚未批准读取原稿和副本。';
  }
  async function showPages(api, current) {
    const index = Number($('review-item').value), row = delivery?.alignment?.[index];
    refresh();
    ['before-pages', 'after-pages'].forEach(id => $(id).replaceChildren());
    const item = delivery?.items?.[index];
    if (item) { $('review-state').value = item.state; $('review-reason').value = item.reason; }
    if (!row) return;
    for (const side of ['before', 'after']) {
      const container = $(side + '-pages');
      const positioning = row[side + '_method'] === 'host-paragraph-context-fallback' ? ' · 段落内严格定位' : '';
      $(side + '-caption').textContent = (side === 'before' ? '原稿' : '副本') + '静态呈现 · 第' + (index + 1) + '项 · 上下文页 ' + (row[side + '_pages'].join('、') || '未定位') + positioning;
      if (row[side + '_status'] !== 'unique') { container.textContent = row[side + '_status'] === 'ambiguous' ? '重复匹配，不能唯一定位；需人工复核。' : '未可靠定位，不能按页码猜测；需人工复核。'; continue; }
      for (let offset = 0; offset < row[side + '_pages'].length; offset++) {
        const response = await api.range_review_page(delivery.review_id, index, side, offset);
        if (!current()) return;
        if (!response?.ok) { fail(response); container.textContent = '页面读取被拒绝，未标记为已查看。'; return; }
        const image = document.createElement('img'); image.alt = (side === 'before' ? '原稿' : '副本') + '第' + response.result.page + '页，静态域呈现'; image.src = response.result.image;
        const zoom = document.createElement('button'); zoom.type = 'button'; zoom.textContent = '放大页面'; zoom.setAttribute('aria-pressed', 'false');
        const frame = document.createElement('div'); frame.className = 'range-page-frame'; frame.append(image);
        zoom.onclick = () => { const enlarged = zoom.getAttribute('aria-pressed') !== 'true'; zoom.setAttribute('aria-pressed', String(enlarged)); zoom.textContent = enlarged ? '还原页面' : '放大页面'; image.style.width = enlarged ? '200%' : '100%'; };
        const provenance = document.createElement('p'); provenance.className = 'muted'; provenance.textContent = response.result.notice || '指定渲染器页图，Word尚未验证。';
        container.append(provenance, zoom, frame);
      }
    }
  }
  async function openDelivery(api, current, identifier) {
    const response = await api.open_range_review(identifier, true);
    if (!current()) return;
    if (!response?.ok) { fail(response); return; }
    delivery = {...response.result, opened: true}; showReview(); await showPages(api, current);
  }
  $('open-review').onclick = () => run(false, async (api, current) => {
    const target = $('delivery-id').value.trim() || save;
    if (target && !/^[a-f0-9]{32}$/.test(target)) { status('输出回执应为32位小写十六进制；请从完整任务的事务入口填入。'); return; }
    if (reviewDirty && !window.confirm('当前审阅状态尚未存档，重新打开将放弃未保存状态，是否继续？')) return;
    if (!consent() || !current()) return;
    if (target && target !== save) save = null;
    let identifier = target ? null : delivery?.review_id;
    delivery = null; ['before-pages', 'after-pages', 'review-item', 'review-note'].forEach(id => $(id).replaceChildren()); refresh();
    if (!identifier) { const response = await api.start_range_review(target); if (!current()) return; if (!response?.ok) { fail(response); return; } identifier = response.result.review_id; }
    await openDelivery(api, current, identifier); reviewDirty = false;
  });
  $('next-review').onclick = () => { const next = Number($('review-item').value) + 1; if (next < (delivery?.items?.length || 0)) { $('review-item').value = String(next); run(false, (api, current) => showPages(api, current)); } };
  $('delivery-id').addEventListener('input', refresh);
  $('review-item').onchange = () => run(false, (api, current) => showPages(api, current));
  $('mark-review').onclick = () => run(false, async (api, current) => {
    if (!delivery?.opened) return;
    const response = await api.mark_range_review(delivery.review_id, Number($('review-item').value), $('review-state').value, $('review-reason').value);
    if (!current()) return;
    if (!response?.ok) { fail(response); return; }
    delivery.items = response.result.items; delivery.revision = response.result.revision; reviewDirty = true;
    $('review-note').textContent = '状态已记录在当前会话，尚未存档；不代表内容正确或写入批准。';
  });
  $('save-review').onclick = () => run(true, async (api, current) => {
    if (!delivery?.opened) return;
    const response = await api.plan_range_review_archive(delivery.review_id, delivery.revision);
    if (!current()) return;
    if (!response?.ok) { fail(response); return; }
    const approved = window.confirm('保存本次逐项审阅状态和来源摘要？不保存正文、不恢复任何写入或撤销批准。取消不写入。') === true;
    if (!current()) return;
    const saved = await api.confirm_range_review_archive(response.result.plan_id, approved);
    if (!current()) return;
    if (!saved?.ok) { fail(saved); return; }
    if (saved.result.status === 'saved') { $('archive-id').value = saved.result.save_id; reviewDirty = false; $('review-note').textContent = '审阅状态已存档；恢复时需重新核对来源。回执：' + saved.result.save_id; }
    else if (saved.result.status === 'uncertain') unknown();
    else status('已取消审阅存档。');
  });
  $('restore-review').onclick = () => run(false, async (api, current) => {
    if (reviewDirty && !window.confirm('恢复将放弃当前尚未存档的审阅状态，是否继续？')) return;
    const response = await api.restore_range_review_archive($('archive-id').value.trim());
    if (!current()) return;
    if (!response?.ok) { fail(response); return; }
    delivery = {...response.result, opened: false}; reviewDirty = false; save = null;
    $('delivery-id').value = '';
    ['before-pages', 'after-pages'].forEach(id => $(id).replaceChildren()); showReview();
    if (consent() && current()) await openDelivery(api, current, delivery.review_id);
  });
  $('remove-review').onclick = () => run(true, async (api, current) => {
    const response = await api.plan_remove_range_review_archive($('archive-id').value.trim());
    if (!current()) return;
    if (!response?.ok) { fail(response); return; }
    const approved = window.confirm('独立撤销这份审阅存档？只删除该存档产物，保留备份和审计，不删除DOCX副本。') === true;
    if (!current()) return;
    const removed = await api.confirm_remove_range_review_archive(response.result.rollback_id, approved);
    if (!current()) return;
    if (!removed?.ok) { fail(removed); return; }
    if (removed.result.status === 'uncertain') unknown();
    else if (removed.result.status === 'rolled_back') { $('archive-id').value = ''; $('review-note').textContent = '审阅存档已撤销，DOCX副本保留。'; }
    else status('已取消审阅存档撤销。');
  });
  window.addEventListener('beforeunload', event => { if (reviewDirty) { event.preventDefault(); event.returnValue = ''; } });
  ['officeauth:authenticated', 'officeauth:unauthenticated', 'pagehide'].forEach(event => window.addEventListener(event, () => { $('delivery-id').value = ''; reset(); }));
  ['file-name', 'file-dir'].forEach(id => { const node = document.getElementById(id); if (node) new MutationObserver(reset).observe(node, {childList: true, characterData: true, subtree: true}); });
  refresh();
})();
