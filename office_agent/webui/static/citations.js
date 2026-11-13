(() => {
  'use strict';
  const $ = id => document.getElementById('citations-' + id);
  if (!$('panel')) return;
  let epoch = 0, busy = false, unknown = false, stale = false, source = null, cards = [], mappings = [], plan = null, saved = null, dirty = false;
  let drafts = new Map(), history = [], activeReceipt = '';
  let locationsNext = null;
  const refusalText = {
    sensitive_or_oversize: '内容敏感或超过预览预算', not_body_paragraph: '表格等非普通正文位置',
    explicit_paragraph_style: '带显式段落样式，当前写作入口只读',
    empty_or_nontext_paragraph: '空段落或无普通文本', paragraph_capacity: '段落超过容量',
    hyperlink_bookmark_field_or_revision: '含超链接、书签、域或修订结构',
    field_image_note_break_or_special_run: '含域、图片、注释、换行或特殊文本结构',
    format_revision: '含格式修订', merged_cell_continuation: '合并单元格延续位置',
    unsupported_structure: '不支持的段落结构'
  };
  const refusal = code => refusalText[code] || '不支持的段落结构';
  const auth = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const messages = {
    not_authenticated: '请先登录。', busy: '任务正在进行，请稍后操作。',
    professional_citation_scope: '授权范围不完整。请勾选所有受重排影响的正文及全部已有书目；最多12项。',
    professional_citation_unmapped: '已有编号或待移除编号没有完整映射，请核对编号与书目。',
    professional_citation_syntax: '编号格式不在支持范围，只接受[1]、[1,2]及有界递增区间。',
    professional_citation_outside_body: '页眉、表格或特殊结构中含数字引文，本入口不扩大到这些位置。',
    professional_bibliography_identity: '书目开头编号、DOI与映射不一致，或同一编号出现多条书目，请核对。',
    professional_duplicate_metadata: '同 DOI 存在不同元数据版本，请为所有编号选择同一已核对版本。',
    professional_bibliography_capacity: '现有书目占位不足，请先显式准备更多普通段落，当前不插入结构。',
    professional_citation_receipt: '续编回执与当前副本、摘要或内容不匹配，无法恢复编号候选。',
    professional_citation_review: '请核对并确认编号、书目及授权范围。',
    professional_support_required: '新增主张的每篇来源须选择1至3项原文，填写关系、范围、限制并人工审阅。',
    professional_identity_required: '请确认每个 PDF 与文献的身份对应。',
    professional_evidence_invalid: '原文选择、会话或文献纳入状态已失效，请重新取证。',
    professional_evidence_changed: 'PDF 或归档已改变，旧计划失效，请重新取证。',
    professional_source_changed: 'DOCX 已改变，请重新读取并定位。',
    professional_invalid_page: '结构位置无效，请输入范围内的正整数；跨页读取必须绑定原文版本。',
    professional_grounded_target: '请选择普通正文段落；书目须位于正文之后，空占位不能含数字、引文或公式。',
    professional_protected_numbers: '原有测量数字或单位被改动、删除或调序，已拒绝。',
    professional_protected_formulas: '原有公式被改动，已拒绝。',
    professional_protected_citations: '无法保留已有非数字引文，请缩小修改范围。',
    professional_unsupported_fact: '新增数字、单位或公式不在所选原文中，请核对。',
    professional_plan_expired: '计划已过期或会话变化，请重新生成。',
    professional_plan_invalid: '计划不可用，请重新生成。',
    professional_no_changes: '实际内容没有变化，无需写入。',
    professional_invalid_request: '请求不符合容量或字段限制，请完整填写并核对。',
    professional_sensitive_text: '说明或元数据中含疑似敏感内容，未生成写入计划，请移除后再审阅。'
  };
  function status(text) { $('status').textContent = text; }
  function refresh() {
    const blocked = busy || unknown || !auth();
    $('panel').setAttribute('aria-busy', String(busy));
    $('panel').querySelectorAll('button,input,textarea,select').forEach(node => { node.disabled = blocked; });
    $('plan').disabled = blocked || stale || !source || drafts.size === 0;
    $('confirm').disabled = blocked || stale || !plan;
    $('rollback').disabled = blocked || !saved;
    $('prev').disabled = blocked || stale || !source || history.length === 0;
    $('next').disabled = blocked || stale || !source || source.next_offset === null;
    $('jump').disabled = blocked || stale || !source;
    $('position').disabled = blocked || stale || !source;
    for (const card of cards) {
      card.root.querySelectorAll('button,input,textarea,select').forEach(element => { element.disabled = blocked || stale || !card.authorized.checked; });
      card.authorized.disabled = blocked || stale || (!card.authorized.checked && drafts.size >= 12);
    }
    $('selected').querySelectorAll('button').forEach(element => { element.disabled = blocked || stale; });
    $('mapping').querySelectorAll('select').forEach(element => { element.disabled = blocked || stale; });
    $('reviewed').disabled = blocked || stale || !source;
    $('find-number').disabled = blocked || stale || !source || source.citation_numbers.length === 0;
    $('find').disabled = blocked || stale || !source || !$('find-number').value;
    $('find-more').disabled = blocked || stale || !source || locationsNext === null;
    $('locations').querySelectorAll('button').forEach(element => { element.disabled = blocked || stale; });
  }
  function changed() { dirty = true; plan = null; $('preview').textContent = ''; $('reviewed').checked = false; refresh(); }
  function reset() {
    epoch++; busy = false; stale = false; source = plan = saved = null; dirty = false; cards = []; mappings = [];
    drafts = new Map(); history = []; activeReceipt = ''; locationsNext = null;
    ['mapping', 'targets', 'selected', 'preview', 'output'].forEach(id => $(id).replaceChildren());
    $('receipt').value = ''; $('reviewed').checked = false;
    $('find-number').replaceChildren(); $('locations').replaceChildren(); $('find-status').textContent = '尚未查找；查找与跳转不授权、不写入。';
    $('page').textContent = '每次最多40项，按文本预算分页；浏览不扩大写入范围。'; $('position').value = '1';
    status('尚未读取；打开面板不联网。'); refresh();
  }
  function uncertain() { unknown = true; plan = saved = null; status('结果不确定，可能已有输出或备份；请核对审计，不自动重试。'); }
  function failed(reply) {
    if (reply?.reason === 'professional_uncertain') uncertain();
    else if (reply?.reason === 'professional_source_changed' || reply?.reason === 'professional_citation_receipt') {
      stale = true; plan = null; $('preview').textContent = '';
      status(messages[reply.reason] + ' 当前草稿仅保留供核对，重新读取后方可继续。');
    }
    else status(messages[reply?.reason] || '操作被拒绝或失败；请核对文档、来源和容量限制。');
  }
  async function run(operation) {
    if (busy || unknown || !auth()) return;
    const mine = epoch; let approved = false; busy = true; refresh();
    const current = () => mine === epoch && auth();
    try { await operation(window.OfficeAuth.wrapApi(window.pywebview.api), current, () => { approved = true; }); }
    catch (_) { if (current()) { if (approved) uncertain(); else status('读取或规划失败，未批准本次写入。'); } }
    finally { if (mine === epoch) { busy = false; refresh(); } }
  }
  function node(tag, text, parent) {
    const element = document.createElement(tag);
    if (text) element.textContent = text;
    if (parent) parent.append(element);
    return element;
  }
  function field(parent, title, tag = 'textarea') {
    const label = node('label', title, parent), element = node(tag, '', label);
    if (tag === 'textarea') { element.rows = 2; element.maxLength = 2000; }
    return element;
  }
  function option(select, value, text) { const item = node('option', text, select); item.value = value; }
  function check(parent, title) { const label = node('label', '', parent), input = node('input', '', label); input.type = 'checkbox'; label.append(document.createTextNode(title)); return input; }
  function addAttribution(card) {
    if (card.attributions.length >= 3) { status('每项主张最多3篇来源。'); return; }
    const root = node('fieldset', '', card.claimFields); node('legend', '原文支持来源 ' + (card.attributions.length + 1), root);
    const select = field(root, '已取得的来源', 'select'); option(select, '', '请选择，打开此处不重新联网');
    for (const row of source.captures) option(select, row.capture_id, row.title + ' · ' + row.doi);
    const evidence = node('div', '', root);
    const relation = field(root, '支持关系', 'select');
    for (const [value, text] of [['insufficient', '证据不足（不可写入）'], ['supports', '支持'], ['contradicts', '反对']]) option(relation, value, text);
    const applicability = field(root, '适用范围与条件'), limitation = field(root, '限制与缺口');
    const identity = check(root, '已核对该 PDF 与文献身份'), reviewed = check(root, '已核对主张与本来源的支持关系');
    const item = {select, evidence, relation, applicability, limitation, identity, reviewed}; card.attributions.push(item);
    select.addEventListener('change', () => {
      changed(); reviewed.checked = identity.checked = false; evidence.replaceChildren();
      if (!select.value) return;
      run(async (api, current) => {
        const reply = await api.read_citation_capture(select.value);
        if (!current()) return;
        if (!reply?.ok) { failed(reply); return; }
        node('p', '内容顺序不保证多栏阅读顺序；仅限先前取得的页段/字符。来源摘要：' + reply.result.document.source_sha256.slice(0, 12), evidence);
        for (const fragment of reply.result.fragments) {
          const input = check(evidence, fragment.location + '：' + fragment.original_text); input.value = fragment.id;
          input.addEventListener('change', () => { reviewed.checked = false; changed(); });
        }
      });
    });
    [relation, applicability, limitation, identity].forEach(element => element.addEventListener('input', () => { reviewed.checked = false; changed(); }));
    reviewed.addEventListener('input', changed);
    const remove = node('button', '移除此来源', root); remove.type = 'button'; remove.className = 'btn secondary';
    remove.addEventListener('click', () => { card.attributions = card.attributions.filter(value => value !== item); root.remove(); changed(); });
    refresh();
    return item;
  }
  function renderMappings() {
    $('mapping').replaceChildren(); mappings = [];
    for (const number of source.citation_numbers) {
      const select = field($('mapping'), '原编号 [' + number + '] 对应文献', 'select'); option(select, '', '必须明确选择文献版本');
      for (const row of source.references) option(select, row.id, row.title + ' · ' + (row.DOI || '缺 DOI'));
      const key = source.manifest?.number_to_key?.[String(number)];
      if (key) select.value = source.manifest.catalog[key].id;
      mappings.push({number, select}); select.addEventListener('change', changed);
    }
    $('find-number').replaceChildren(); locationsNext = null; $('locations').replaceChildren();
    for (const number of source.citation_numbers) option($('find-number'), String(number), '原编号 [' + number + ']');
    $('find-status').textContent = source.citation_numbers.length ? '请选择编号主动查找；结果仅为出现位置，不是完整重排影响分析。' : '正文结构清单中未识别到数字引文。';
  }
  function selectedSummary() {
    $('selected').replaceChildren();
    node('p', '已授权 ' + drafts.size + '/12 项（含其他页）；仅勾选项会提交。取消授权将放弃本段草稿。', $('selected'));
    for (const card of drafts.values()) {
      const button = node('button', card.item.location + ' · 查看草稿', $('selected')); button.type = 'button'; button.className = 'btn secondary';
      button.addEventListener('click', () => {
        if (cards.includes(card)) card.authorized.focus();
        else pageTo(card.index);
      });
    }
  }
  function renderTargets() {
    $('targets').replaceChildren(); cards = [];
    for (const [index, item] of source.items.entries()) {
      const retained = drafts.get(item.locator_id);
      if (retained) { retained.root.classList.remove('citations-found'); $('targets').append(retained.root); cards.push(retained); continue; }
      const root = node('article', '', $('targets')); root.dataset.locator = item.locator_id;
      node('strong', item.location + (item.citation_refusal || !item.writable ? ' · 只读' : ''), root); node('pre', item.text, root);
      if (item.citation_refusal) node('p', '只读原因：' + refusal(item.citation_refusal), root);
      if (!item.writable || item.style || !/^body:\d+\/paragraph$/.test(item.location)) continue;
      const authorized = check(root, '授权本段参与本次整组修改'); authorized.className = 'citations-authorized';
      const role = field(root, '本段用途', 'select'); role.className = 'citations-role';
      for (const [value, text] of [['keep', '保留措辞，仅按需要调整编号'], ['bibliography', '书目/书目占位（整段替换）'], ['claim', '用已审阅证据改写主张']]) option(role, value, text);
      if (source.manifest?.slots?.includes(item.location)) role.value = 'bibliography';
      const removal = field(root, '仅移除本段这些原编号（逗号分隔；保留措辞）', 'input'); removal.maxLength = 80;
      const claimFields = node('div', '', root); claimFields.hidden = true;
      const text = field(claimFields, '新主张正文（不手写编号）'); text.value = item.text;
      const card = {item, index: source.offset + index, root, authorized, role, removal, claimFields, text, attributions: []}; cards.push(card);
      authorized.addEventListener('input', () => {
        if (authorized.checked) {
          if (drafts.size >= 12) { authorized.checked = false; status('每组最多授权12项，请先取消其他段落的授权。'); return; }
          drafts.set(item.locator_id, card);
        } else {
          if ((role.value !== 'keep' || removal.value || card.attributions.length || text.value !== item.text) &&
              window.confirm('取消本段授权并放弃该段草稿？其他段落和编号映射保留。') !== true) { authorized.checked = true; return; }
          drafts.delete(item.locator_id); renderTargets();
        }
        changed(); selectedSummary(); refresh();
      });
      const add = node('button', '增加一篇原文支持来源', claimFields); add.type = 'button'; add.className = 'btn secondary';
      add.addEventListener('click', () => { addAttribution(card); changed(); });
      role.addEventListener('change', () => { claimFields.hidden = role.value !== 'claim'; removal.parentElement.hidden = role.value !== 'keep'; changed(); });
      removal.addEventListener('input', changed);
      text.addEventListener('input', () => { card.attributions.forEach(value => { value.reviewed.checked = false; }); changed(); });
    }
    $('page').textContent = '当前第 ' + (source.items.length ? source.offset + 1 : 0) + '—' + (source.offset + source.items.length) + ' 项 / 共 ' + source.total_targets + ' 项；编号清单覆盖正文结构；翻页保留授权草稿。';
    $('position').max = String(source.total_targets); $('position').value = String(source.offset + 1);
    selectedSummary();
  }
  function pageTo(offset, backward = false, locator = '') {
    if (!source || stale || (offset === source.offset && !locator)) return;
    run(async (api, current) => {
      status('正在按原版本读取结构；跨页授权和草稿保留…');
      const previous = source.offset, version = source.document.source_sha256;
      const reply = await api.read_citation_workspace(activeReceipt, offset, version);
      if (!current()) return;
      if (!reply?.ok) { failed(reply); return; }
      if (reply.result.document.source_sha256 !== version) { failed({reason: 'professional_source_changed'}); return; }
      if (backward) history.pop(); else if (previous !== offset) history.push(previous);
      source = reply.result; renderTargets(); $('targets').focus();
      if (locator) {
        const target = [...$('targets').children].find(element => element.dataset.locator === locator);
        if (target) { target.classList.add('citations-found'); target.tabIndex = -1; target.focus(); target.scrollIntoView({block: 'nearest'}); }
      }
      status('已读取本页；跨页草稿与编号映射保留，尚未写入。');
    });
  }
  $('prev').addEventListener('click', () => pageTo(history[history.length - 1], true));
  $('next').addEventListener('click', () => pageTo(source?.next_offset));
  $('jump').addEventListener('click', () => {
    const position = Number($('position').value);
    if (!source || !Number.isInteger(position) || position < 1 || position > source.total_targets) { status('请输入 1 至 ' + (source?.total_targets || 1) + ' 的结构序号。'); return; }
    pageTo(position - 1);
  });
  function findLocations(offset = 0) {
    run(async (api, current) => {
      const number = Number($('find-number').value), version = source.document.source_sha256;
      $('find-status').textContent = '正在按当前版本查找出现位置…';
      const reply = await api.read_citation_locations(number, version, offset);
      if (!current()) return;
      if (!reply?.ok) { $('find-status').textContent = '查找未完成，未授权或写入。'; failed(reply); return; }
      if (reply.result.source_sha256 !== version || reply.result.number !== number) { failed({reason: 'professional_source_changed'}); return; }
      const result = reply.result; locationsNext = result.next_offset; $('locations').replaceChildren();
      for (const row of result.items) {
        const root = node('article', '', $('locations'));
        node('p', '第' + (row.index + 1) + '项 · ' + row.location + ' · 匹配' + row.occurrences + '组编号' + (row.kind === 'bibliography_candidate' ? ' · 书目候选（需核对）' : ''), root);
        node('p', row.excerpt + (row.excerpt_truncated ? ' …（摘录）' : ''), root);
        node('p', row.citation_refusal ? '只读原因：' + refusal(row.citation_refusal) : '可选普通段落，尚未由此操作授权。', root);
        const button = node('button', '定位到第' + (row.index + 1) + '项', root); button.type = 'button'; button.className = 'btn secondary';
        button.addEventListener('click', () => pageTo(row.index, false, row.locator_id));
      }
      $('find-status').textContent = '编号 [' + number + '] 共' + result.total_locations + '处结构位置、' + result.total_occurrences + '组编号（含书目标签）；显示第' + (result.items.length ? offset + 1 : 0) + '—' + (offset + result.items.length) + '处。另有' + result.omitted_targets + '项敏感或超限内容未扫描。页眉等范围不在此清单；重排仍以整组计划裁决。';
    });
  }
  $('find').addEventListener('click', () => findLocations());
  $('find-more').addEventListener('click', () => findLocations(locationsNext));
  $('find-number').addEventListener('change', () => { locationsNext = null; $('locations').replaceChildren(); $('find-status').textContent = '编号已改变，请重新查找；草稿与计划保留。'; refresh(); });
  $('load').addEventListener('click', () => {
    if (dirty && window.confirm('重新读取会丢弃本面板草稿与未确认计划，是否继续？') !== true) return;
    run(async (api, current) => {
      status('正在读取版本、编号和文献候选…');
      const reply = await api.read_citation_workspace($('receipt').value);
      if (!current()) return;
      if (!reply?.ok) { failed(reply); return; }
      source = reply.result; plan = null; dirty = false; stale = false; drafts = new Map(); history = []; activeReceipt = $('receipt').value;
      $('reviewed').checked = false; $('preview').textContent = ''; renderMappings(); renderTargets();
      status('已读取本页，共' + source.total_targets + '项；支持翻页与跳转，正文结构编号已列出。' + (source.manifest ? '已恢复回执中的编号候选，未恢复写入批准。' : '请明确映射已有编号并选择书目与授权段落。'));
    });
  });
  $('receipt').addEventListener('input', () => { stale = true; changed(); status('回执已修改，请重新读取；现有草稿尚未丢弃。'); });
  $('reviewed').addEventListener('input', () => { plan = null; dirty = true; $('preview').textContent = ''; refresh(); });
  $('plan').addEventListener('click', () => run(async (api, current) => {
    const selected = [...drafts.values()];
    const payload = {source_sha256: source.document.source_sha256, authorized: selected.map(card => card.item.locator_id),
      bibliography: selected.filter(card => card.role.value === 'bibliography').map(card => card.item.locator_id),
      number_records: Object.fromEntries(mappings.map(row => [String(row.number), row.select.value])), mapping_reviewed: $('reviewed').checked, claims: [], removals: {}};
    for (const card of selected) {
      if (card.role.value === 'keep' && card.removal.value.trim()) {
        if (!/^\d+(?:\s*,\s*\d+)*$/.test(card.removal.value.trim())) { status('待移除编号请用逗号分隔的正整数。'); return; }
        payload.removals[card.item.locator_id] = card.removal.value.split(',').map(value => Number(value.trim()));
      }
      if (card.role.value !== 'claim') continue;
      payload.claims.push({locator_id: card.item.locator_id, text: card.text.value, attributions: card.attributions.map(item => ({
        capture_id: item.select.value, evidence_ids: [...item.evidence.querySelectorAll('input:checked')].map(input => input.value),
        relation: item.relation.value, applicability: item.applicability.value, limitation: item.limitation.value,
        identity_reviewed: item.identity.checked, assessment: item.reviewed.checked ? 'user_reviewed' : 'unverified'}))});
    }
    plan = null; status('正在复核编号、来源、授权范围并重建书目…');
    const reply = await api.plan_citation_edit(payload);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    plan = reply.result;
    const before = JSON.parse(plan.before).items;
    const lines = plan.items.map(item => item.location + '\n修改前：' + before.find(row => row.location === item.location).text + '\n修改后：' + item.text);
    lines.push('编号映射：\n' + Object.entries(plan.manifest.number_changes).map(([old, value]) => '[' + old + '] → ' + (value === null ? '不再引用' : '[' + value + ']')).join('\n'));
    lines.push('最终编号：\n' + Object.entries(plan.manifest.number_to_key).map(([number, key]) => '[' + number + '] ' + plan.manifest.catalog[key].title + ' · ' + key).join('\n'));
    if (plan.manifest.removed_keys.length) lines.push('本次不再引用：' + plan.manifest.removed_keys.join('；'));
    for (const claim of plan.manifest.claims) for (const item of claim.attributions) {
      lines.push('人工支持关系：' + (item.support.relation === 'supports' ? '支持' : '反对') + '\n适用范围：' + item.support.applicability + '\n限制：' + item.support.limitation);
      for (const fragment of item.proof.fragments) lines.push('原文 · 第' + fragment.page + '页 · ' + fragment.location + '\n' + fragment.original_text);
    }
    lines.push('无段落插入/删除；历史主张未重新做语义审阅。取消不写入；5分钟后计划失效。');
    $('preview').textContent = lines.join('\n\n'); status('整组计划已生成，尚未写入；请核对每项差异与编号对应。');
  }));
  $('confirm').addEventListener('click', () => run(async (api, current, writing) => {
    const target = plan; plan = null;
    const approved = window.confirm('确认按完整差异重排引文、更新书目及主张，写入新副本？原件不变。') === true;
    if (!current()) return;
    if (approved) writing();
    const reply = await api.confirm_citation_edit(target.id, approved);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    if (reply.result.status === 'cancelled') { status('已取消，未写入。'); return; }
    if (reply.result.status !== 'saved') { uncertain(); return; }
    saved = reply.result.save_id; dirty = false;
    $('output').textContent = '已核验副本：' + reply.result.output + '\n续编回执 ID：' + saved + '；原件未改；渲染：' + (reply.result.render?.status || '未提供');
    status('保存成功；归档依赖保留。另选副本并输入上述回执，可继续编排；新增主张需重新取证。');
  }));
  $('rollback').addEventListener('click', () => run(async (api, current, writing) => {
    const approved = window.confirm('撤销本次未被外部修改的输出副本？原件、备份和审计保留。') === true;
    if (!current()) return;
    if (approved) writing();
    const reply = await api.rollback_citation_edit(saved, approved);
    if (!current()) return;
    if (!reply?.ok) { failed(reply); return; }
    if (reply.result.status === 'cancelled') { status('已取消撤销。'); return; }
    if (reply.result.status !== 'rolled_back') { uncertain(); return; }
    saved = null; $('output').textContent = '本次副本已撤销，已释放本次新增的归档依赖。';
  }));
  ['officeauth:unauthenticated', 'officeauth:authenticated', 'pagehide'].forEach(event => window.addEventListener(event, reset));
  window.addEventListener('researchassist:transfer', event => {
    const value = event.detail;
    if (!auth() || busy || unknown || !value || !Number.isInteger(value.index) || value.index < 0) return;
    if (source && source.document.source_sha256 !== value.source_sha256) { status('研究建议与当前引文草稿版本不同，请重新读取；草稿保留。'); return; }
    if (drafts.has(value.locator_id) && !window.confirm('用人工编辑的研究建议替换本段引文草稿？此段支持关系审阅将清除，其他草稿保留。')) return;
    if (!drafts.has(value.locator_id) && drafts.size >= 12) { status('已达12项，请先取消其他段落授权。'); return; }
    value.received = true; $('panel').open = true;
    run(async (api, current) => {
      const reply = await api.read_citation_workspace(activeReceipt, value.index, value.source_sha256);
      if (!current()) return;
      if (!reply?.ok) { failed(reply); return; }
      if (reply.result.document.source_sha256 !== value.source_sha256) { failed({reason: 'professional_source_changed'}); return; }
      const first = !source; source = reply.result; if (first) renderMappings();
      drafts.delete(value.locator_id); renderTargets();
      const card = cards.find(row => row.item.locator_id === value.locator_id);
      if (!card || card.item.citation_refusal) { status('本段不适合引文整组修改，保留只读。'); return; }
      card.authorized.checked = true; drafts.set(value.locator_id, card); card.role.value = 'claim'; card.claimFields.hidden = false; card.removal.parentElement.hidden = true; card.text.value = value.new_text;
      for (const suggested of value.sources.slice(0, 3)) {
        if (!source.captures.some(row => row.capture_id === suggested.capture_id)) continue;
        const attribution = addAttribution(card); attribution.select.value = suggested.capture_id;
        const captured = await api.read_citation_capture(suggested.capture_id);
        if (!current()) return;
        if (!captured?.ok) { failed(captured); return; }
        node('p', '模型选中的候选原文；请重新核对身份与支持关系。', attribution.evidence);
        for (const fragment of captured.result.fragments) {
          const input = check(attribution.evidence, fragment.location + '：' + fragment.original_text); input.value = fragment.id; input.checked = suggested.evidence_ids.includes(fragment.id);
          input.addEventListener('change', () => { attribution.reviewed.checked = false; changed(); });
        }
        attribution.relation.value = 'insufficient'; attribution.identity.checked = attribution.reviewed.checked = false;
      }
      changed(); selectedSummary(); card.text.focus(); card.root.scrollIntoView({block: 'nearest'});
      status('研究建议已转为待审阅引文草稿。请核对文献身份、支持关系、适用条件、限制、书目占位和编号映射；未恢复任何写入批准。');
    });
  });
  ['file-name', 'file-dir'].forEach(id => { const element = document.getElementById(id); if (element) new MutationObserver(reset).observe(element, {childList: true, characterData: true, subtree: true}); });
  refresh();
})();
