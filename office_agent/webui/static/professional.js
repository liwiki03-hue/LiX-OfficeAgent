(() => {
  'use strict';
  const $ = id => document.getElementById('professional-' + id);
  if (!$('panel')) return;
  let epoch = 0, revision = 0, busy = false, unknown = false, source = null, selected = null;
  let changes = [], plan = null, saveId = null, next = null;
  const auth = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const host = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const opaque = value => typeof value === 'string' && /^[a-f0-9]{32}$/.test(value);
  const messages = {
    not_authenticated: '请先登录本机账户。', busy: '任务正在进行，请结束后再操作。',
    professional_select_document: '请先选择授权文档。', professional_docx_required: '本入口只处理普通 DOCX。',
    professional_source_changed: '源文件已改变，请重新读取结构、定位并生成计划。',
    professional_stale_locator: '定位或原文已过期，请重新读取。', professional_invalid_request: '请求格式不符合限制，请核对文本与清单。',
    professional_protected_numbers: '数字或已识别单位改变，整组修改已拒绝。',
    professional_protected_citations: '受保护引用改变，整组修改已拒绝。',
    professional_protected_formulas: '受保护文本公式改变，整组修改已拒绝。',
    professional_protected_literal: '额外保护原文的出现次数改变，整组修改已拒绝。',
    professional_overlapping_changes: '存在重叠目标，请每个段落仅保留一项最终修改。',
    professional_unsupported_target: '目标含不支持的特殊结构，只能生成建议。',
    professional_capacity: '文档或修改超过本轮容量，请缩小范围。',
    professional_evidence_missing: '证据 ID 无法与本源版本对应，请重新读取。',
    professional_plan_expired: '计划已过期或会话已改变，请重新生成。', professional_plan_invalid: '计划不可用，请重新生成。'
  };
  function status(text) { $('status').textContent = text; }
  function refresh() {
    const blocked = busy || !auth() || unknown;
    ['load', 'enable', 'numbers', 'citations', 'formulas', 'literals'].forEach(id => { $(id).disabled = blocked; });
    $('next').disabled = blocked || next === null;
    $('text').disabled = $('reason').disabled = blocked || !selected?.writable;
    $('add').disabled = blocked || !selected?.writable;
    $('plan').disabled = blocked || !changes.some(item => item.included);
    $('confirm').disabled = blocked || !plan;
    $('rollback').disabled = blocked || !saveId;
    $('panel').setAttribute('aria-busy', String(busy));
    $('selection').textContent = selected ? '当前编辑目标：' + selected.location + '；源版本 ' + (source?.source_sha256 || '').slice(0, 12) : '尚未选择修改目标。';
    $('targets').querySelectorAll('input').forEach(input => { input.disabled = blocked; });
    $('changes').querySelectorAll('input,button').forEach(input => { input.disabled = blocked; });
  }
  function changed() { revision++; plan = null; $('preview').textContent = ''; refresh(); }
  function reset() {
    epoch++; revision++; busy = false; source = selected = plan = saveId = null; changes = []; next = null;
    ['targets', 'changes', 'preview', 'output'].forEach(id => { $(id).textContent = ''; });
    ['text', 'reason', 'literals'].forEach(id => { $(id).value = ''; });
    $('enable').checked = false;
    status('尚未读取；只改副本，写入前单独确认。'); refresh();
  }
  function failed(reply) {
    if (reply?.reason === 'professional_uncertain') { uncertain(); return; }
    status(messages[reply?.reason || reply?.error] || '操作被拒绝；请核对所选文档、结构及安全限制。');
  }
  async function run(write, operation) {
    if (busy || !auth() || unknown) return;
    const mine = epoch; busy = true; refresh();
    const current = () => mine === epoch && auth();
    try { await operation(host(), current); }
    catch (_) {
      if (current()) {
        if (write) uncertain();
        else status('读取或规划失败，未批准写入；请核对本机服务。');
      }
    } finally { if (mine === epoch) { busy = false; refresh(); } }
  }
  function uncertain() {
    unknown = true; plan = saveId = null;
    status('结果不确定，可能已有备份或输出；请核对记录，不自动重试。');
  }
  function renderChanges() {
    $('changes').replaceChildren();
    changes.forEach(item => {
      const row = document.createElement('div'), label = document.createElement('label'), check = document.createElement('input');
      check.type = 'checkbox'; check.checked = item.included;
      check.addEventListener('change', () => { item.included = check.checked; changed(); });
      label.append(check, document.createTextNode(item.location + '：' + item.reason));
      const preview = document.createElement('pre'); preview.textContent = '原文：' + item.original + '\n新文：' + item.new_text;
      const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'btn secondary'; remove.textContent = '移除此项';
      remove.addEventListener('click', () => { changes = changes.filter(value => value !== item); changed(); renderChanges(); });
      row.append(label, preview, remove); $('changes').append(row);
    }); refresh();
  }
  function choose(item) {
    if (selected && $('text').value !== selected.text && !changes.some(value => value.locator_id === selected.locator_id && value.new_text === $('text').value) &&
        window.confirm('切换目标会放弃尚未加入清单的新文本，是否继续？') !== true) return;
    selected = item;
    const draft = changes.find(value => value.locator_id === item.locator_id);
    $('text').value = draft?.new_text ?? item.text; $('reason').value = draft?.reason || '';
    refresh();
  }
  function load(offset) {
    if (offset === 0 && (changes.length || selected && $('text').value !== selected.text) &&
        window.confirm('重新读取会丢弃未保存的修改清单和新文本，是否继续？') !== true) return;
    return run(false, async (api, current) => {
      status('正在读取当前文件的版本与结构…');
      const response = await api.read_professional_structure(offset);
      if (!current()) return;
      if (!response?.ok || !Array.isArray(response.result?.items)) { failed(response); return; }
      const result = response.result;
      if (offset && source?.source_sha256 !== result.document.source_sha256) { reset(); status('源版本已改变，请从第一页重新读取。'); return; }
      if (!offset) { changes = []; selected = null; $('text').value = $('reason').value = ''; $('changes').textContent = ''; }
      source = result.document; next = result.next_offset; changed(); $('targets').replaceChildren();
      for (const item of result.items) {
        const card = document.createElement('article'), label = document.createElement('label'), radio = document.createElement('input');
        radio.type = 'radio'; radio.name = 'professional-target'; radio.value = item.locator_id;
        radio.addEventListener('change', () => {
          choose(item);
          $('targets').querySelectorAll('input').forEach(value => { value.checked = value.value === selected?.locator_id; });
        });
        label.append(radio, document.createTextNode(item.location + ' · ' + (item.writable ? '可规划局部修改' : '只读：' + item.refusal)));
        const text = document.createElement('pre'); text.textContent = item.text;
        const evidence = document.createElement('small'); evidence.textContent = item.evidence?.id || '未提供可用证据片段';
        card.append(label, text, evidence); $('targets').append(card);
      }
      status('已读取 ' + result.items.length + ' 项，共 ' + result.total_targets + ' 项；结构位置无 Word 页码保证，源版本改变需重新定位。');
    });
  }
  $('load').addEventListener('click', () => load(0));
  $('enable').addEventListener('change', () => {
    const enabled = $('enable').checked;
    run(false, async (api, current) => {
      const reply = await api.enable_professional_tools(enabled);
      if (!current()) return;
      if (!reply?.ok) { $('enable').checked = !enabled; failed(reply); return; }
      changed(); status(enabled ? '专业工具已绑定当前源版本；下一次对话可用，实际写入仍需独立确认。' : '当前文档的专业工具已关闭。');
    });
  });
  $('next').addEventListener('click', () => { if (next !== null) load(next); });
  ['text', 'reason', 'numbers', 'citations', 'formulas', 'literals'].forEach(id => $(id).addEventListener('input', changed));
  $('add').addEventListener('click', () => {
    if (busy || !auth() || !selected?.writable) return;
    if (!$('text').value || !$('reason').value.trim()) { status('请填写新文本和修改原因。'); return; }
    if (changes.length >= 12 && !changes.some(item => item.locator_id === selected.locator_id)) { status('每组最多12项，请缩小范围。'); return; }
    changes = changes.filter(item => item.locator_id !== selected.locator_id);
    changes.push({locator_id: selected.locator_id, expected_text_sha256: selected.text_sha256,
      new_text: $('text').value, reason: $('reason').value.trim(), evidence_ids: selected.evidence ? [selected.evidence.id] : [],
      original: selected.text, location: selected.location, included: true});
    changed(); renderChanges(); status('已更新清单；任何调整都需重新生成最终计划，不会直接写入。');
  });
  $('plan').addEventListener('click', () => run(false, async (api, current) => {
    const version = revision;
    const payload = {source_sha256: source.source_sha256, changes: changes.filter(item => item.included).map(item =>
      Object.fromEntries(['locator_id', 'expected_text_sha256', 'new_text', 'reason', 'evidence_ids'].map(key => [key, item[key]]))),
      preserve_numbers: $('numbers').checked, preserve_citations: $('citations').checked, preserve_formulas: $('formulas').checked,
      protected_literals: $('literals').value.split('\n').filter(Boolean)};
    const response = await api.plan_professional_edit(payload);
    if (!current() || version !== revision) return;
    if (!response?.ok || !opaque(response.result?.id)) { failed(response); return; }
    const candidate = response.result;
    const originals = JSON.parse(candidate.before).items;
    const diffs = candidate.items.map((item, index) => {
      if (item.location !== originals[index]?.location) throw Error();
      return '第' + (index + 1) + '项 · ' + item.location + '\n修改原因：' + item.reason +
        '\n修改前：' + originals[index].text + '\n修改后：' + item.text;
    });
    const protection = candidate.protected;
    $('preview').textContent = '源版本：' + candidate.source_sha256 + '\n\n' + diffs.join('\n\n') +
      '\n\n保护数字与单位：' + (protection.numbers ? '是' : '否') + '；引用：' + (protection.citations ? '是' : '否') +
      '；文本公式：' + (protection.formulas ? '是' : '否') + '\n额外保护原文：' + (protection.literals.join('；') || '无') +
      '\n格式：保留现有run；新增文字沿用首个受影响run格式。\n普通文本差异，非原生修订；证据支持关系未核验。';
    plan = candidate;
    status('最终差异已生成，尚未写入；请逐项检查后单独确认。');
  }));
  $('confirm').addEventListener('click', () => run(true, async (api, current) => {
    const target = plan, version = revision;
    if (!target) return;
    const approved = window.confirm('确认整组修改并生成新 DOCX 副本？\n原件保留；会备份、逐项保护检查、重开与渲染校验、审计。\n任一项失败整组拒绝，未知结果不自动重试。\n' + $('preview').textContent) === true;
    if (!current() || version !== revision) return;
    const response = await api.confirm_professional_edit(target.id, approved);
    if (!current()) return;
    plan = null;
    if (!response?.ok) { failed(response); return; }
    if (response.result?.status === 'cancelled') { status('已取消，没有批准写入。'); return; }
    if (response.result?.status !== 'saved' || !opaque(response.result.save_id)) { uncertain(); return; }
    saveId = response.result.save_id;
    const render = response.result.render || {};
    $('output').textContent = '已保存副本：' + response.result.output + '；源文件未改，审计：' + response.result.audit_status +
      '；渲染状态：' + (render.status || '未提供') + '，页数 ' + (render.source_pages ?? '未知') + ' → ' + (render.output_pages ?? '未知') +
      '。结构与受保护内容已检查；原生 Word 未验证。请重新选择输出副本再继续修改。';
    status('副本已保存并校验；渲染报告与原生 Word 人工验证分开，本次输出可撤销。');
  }));
  $('rollback').addEventListener('click', () => run(true, async (api, current) => {
    const target = saveId;
    if (!target) return;
    const approved = window.confirm('确认撤销本次未被外部修改的输出副本？原件、备份与审计保留。') === true;
    if (!current()) return;
    const response = await api.rollback_professional_edit(target, approved);
    if (!current()) return;
    if (!response?.ok) { failed(response); return; }
    if (response.result?.status === 'cancelled') { status('已取消撤销。'); return; }
    if (response.result?.status !== 'rolled_back') { uncertain(); return; }
    saveId = null; $('output').textContent = '本次输出已撤销，原件、备份与审计保留。';
  }));
  ['officeauth:unauthenticated', 'officeauth:authenticated', 'pagehide'].forEach(event => window.addEventListener(event, reset));
  ['file-name', 'file-dir'].forEach(id => {
    const node = document.getElementById(id);
    if (node) new MutationObserver(reset).observe(node, {childList: true, characterData: true, subtree: true});
  });
  refresh();
})();
