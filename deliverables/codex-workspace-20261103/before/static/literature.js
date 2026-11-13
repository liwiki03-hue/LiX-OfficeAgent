(() => {
  const $ = id => document.getElementById('literature-' + id);
  if (!$('panel')) return;
  let epoch = 0, busy = false, uncertain = false, writePending = false, records = new Map(), decisions = {}, reports = [], saved = new Map();
  const auth = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const api = () => window.OfficeAuth.wrapApi(window.pywebview.api);
  const opaque = value => typeof value === 'string' && /^[a-f0-9]{32}$/.test(value);
  const fields = {title:'题名', authors:'作者', year:'年份', kind:'文献类型', container_title:'出版来源', bibliographic:'卷期页', doi:'DOI'};
  const decisionsText = {include:'已纳入', exclude:'已排除', undecided:'待定'};
  const fulltextText = {not_acquired:'尚未取得', downloaded_unparsed:'已保存，尚未解析', parsed_partial:'部分文本层已解析', parsed_text_layer:'文本层已解析', user_local_text_layer:'已关联自有 PDF 文本层', text_layer_empty_in_preview:'预览范围内未取得文字'};
  const kinds = {'journal-article':'期刊论文', article:'论文', preprint:'预印本', 'posted-content':'预印本或发布内容', 'proceedings-article':'会议论文', book:'图书'};
  const relationText = {'is-preprint-of':'对应正式版本', 'has-preprint':'预印本', 'is-version-of':'版本关联', 'has-version':'关联版本', 'is-correction-of':'更正对象', 'is-retracted-by':'撤稿通知', retraction:'撤稿关联', correction:'更正关联'};
  const messages = {
    not_authenticated: '登录已失效，请重新登录。', no_file: '请先选择工作区文档。', busy: '有操作正在进行，请等待完成。',
    literature_invalid: '请核对检索词、年份与类型；精确 DOI 不叠加其他筛选。', literature_sensitive: '输入或服务结果含疑似敏感内容，已拒绝。',
    literature_doi_invalid: 'DOI 格式无效或条目缺少 DOI。', literature_configuration: '文献配置不符合限制，请检查配置模板。',
    literature_timeout: '文献服务请求超时，未自动重试。', literature_rate_limited: '服务限流或本机请求过密，请稍后主动重试。',
    literature_budget: '本次工作区会话的请求预算已用完，未继续联网。', literature_auth_required: '服务要求认证或额度不足；请核对该来源的账户配置。',
    literature_redirect_refused: '服务返回重定向，已拒绝跟随；凭据未转发。', literature_not_found: '该来源未找到条目，不代表其他来源也没有。',
    literature_service_failed: '文献服务不可用，请稍后检查网络和服务状态。', literature_response: '服务返回格式未通过校验。',
    literature_capacity: '条目、响应或全文超过本轮容量，请缩小范围。', literature_scope_changed: '登录或工作区已改变，本次结果已丢弃。',
    literature_selection_invalid: '请先检索并收藏所需条目。', literature_duplicate_selection: '同 DOI 收藏了多个来源，请明确保留一个版本后再导出。',
    literature_fulltext_refused: '此链接或许可不在允许范围，请使用已有的本地 PDF；未跟随其他来源。',
    literature_fulltext_exists: '该条目已有本次保存的全文，请先查看或撤销，不重复下载。',
    literature_fulltext_in_use: '该全文被写作计划或输出引用，请先取消写作计划或撤销输出，再撤销全文归档。',
    literature_pdf_required: '请先选中自有 PDF，再检索和关联对应条目。', literature_expansion_refused: '被引扩展需要 OpenAlex 来源条目。',
    literature_plan_invalid: '计划已失效，请重新生成。', evidence_expired: '计划已过期，请重新生成。',
    literature_path_refused: '所选文件或工作区路径已失效，请重新选择文档。',
    evidence_path_changed: '保存目录或文件已改变，请核对本机文件。', evidence_denied: '已取消，没有写入。'
  };
  function status(value) { $('status').textContent = value; }
  function refresh() {
    const blocked = busy || !auth() || uncertain;
    $('panel').setAttribute('aria-busy', String(busy));
    $('panel').querySelectorAll('button,input,select,textarea').forEach(node => { node.disabled = blocked; });
    $('csl').disabled = blocked || !Object.values(decisions).some(row => row.decision === 'include');
    $('record').disabled = blocked || !reports.length;
    $('check').disabled = blocked || !records.size;
  }
  function reset() {
    epoch++; busy = false; writePending = false; records = new Map(); decisions = {}; reports = []; saved = new Map();
    ['results','reports','bibliography','citation-report','fulltext','outputs'].forEach(id => { $(id).replaceChildren(); });
    ['query','author','from','to','citation-text'].forEach(id => { $(id).value = ''; });
    status('尚未检索；会话收藏已清空，磁盘输出保留。'); refresh();
  }
  function doubt() {
    uncertain = true;
    status('保存或撤销结果不确定，可能已有输出与审计。请核对本机文件，不自动重试。'); refresh();
  }
  function result(reply) {
    if (!reply?.ok) {
      if (['literature_uncertain','evidence_uncertain'].includes(reply?.reason)) doubt();
      else status(messages[reply?.reason] || '文献操作未完成，请检查工作区与服务状态。');
      return null;
    }
    if (!reply.result || typeof reply.result !== 'object') throw Error();
    return reply.result;
  }
  async function run(write, operation) {
    if (busy || !auth() || uncertain) return;
    const mine = epoch; busy = true; writePending = false; refresh();
    const current = () => mine === epoch && auth();
    try { await operation(api(), current); }
    catch (_) { if (current()) { if (write && writePending) doubt(); else status('读取或规划中断；尚未批准保存，没有自动重试。'); } }
    finally { if (mine === epoch) { busy = false; refresh(); } }
  }
  function add(parent, tag, value) {
    const node = document.createElement(tag); node.textContent = value; parent.append(node); return node;
  }
  function button(parent, title, handler) {
    const node = add(parent, 'button', title); node.type = 'button'; node.className = 'btn secondary';
    node.addEventListener('click', handler); return node;
  }
  function recordStatus(row) {
    return '状态：' + (row.abstract ? '已取得摘要，未据此认定阅读全文' : '仅元数据') + '；全文：' + (fulltextText[row.fulltext_state] || '状态未知') +
      '；撤稿：' + (row.retraction_status === 'reported_retracted' ? '来源报告已撤稿' : '未知，未确认未撤稿');
  }
  function setFulltextState(recordId, state) {
    const row = records.get(recordId);
    if (!row) return;
    row.fulltext_state = state;
    $('results').querySelectorAll('[data-fulltext-record]').forEach(node => {
      if (node.dataset.fulltextRecord === recordId) node.textContent = recordStatus(row);
    });
  }
  function bibliography() {
    const included = [...records.values()].filter(row => decisions[row.id]?.decision === 'include');
    $('bibliography').textContent = included.length ? included.map((row, index) => '[' + (index + 1) + '] ' +
      (row.authors.join('；') || '作者缺失') + '．' + (row.title || '题名缺失') + '．' + (row.year || '年份缺失') +
      (row.doi ? '．DOI: ' + row.doi : '．DOI 缺失') + '（' + row.source + '）').join('\n') : '尚无纳入条目。';
  }
  function present(data) {
    if (!Array.isArray(data.groups) || !Array.isArray(data.reports)) throw Error();
    const visible = [];
    data.groups.forEach(group => {
      if (!Array.isArray(group.variants)) throw Error();
      group.variants.forEach(row => {
        if (typeof row.id !== 'string' || !/^lit:[a-f0-9]{64}$/.test(row.id) || !Array.isArray(row.authors)) throw Error();
        records.set(row.id, row); visible.push({row, group});
      });
    });
    reports.push(...data.reports);
    if (data.decisions) decisions = data.decisions;
    $('reports').textContent = data.reports.map(row => row.source + '：' + (row.status === 'completed' ?
      '返回 ' + row.returned + ' 条；第 ' + row.page + ' 页；' + (row.cached ? '会话缓存' : '服务请求') +
      (row.truncated ? '；仍有未取回结果' : '') + '；' + (row.auth_mode === 'configured_key' ? '已配置密钥模式' : row.auth_mode === 'anonymous' ? '匿名模式' : '认证方式未报告') : (messages[row.status] || '来源未完成'))).join('\n') +
      '\n本次会话请求：' + data.requests_used + '/' + data.request_limit;
    render(visible);
    bibliography();
    status(visible.length ? '已取得来源记录；元数据、摘要和全文状态分别展示，不代表论点获得支持。' : '本次没有可展示条目，请检查各来源状态与检索条件。');
    if (data.verification) status('DOI 核对来源：' + (data.verification.sources_found.join('、') || '未取得') +
      '。请检查字段冲突；DOI 可解析不代表论点受支持，也不确认未撤稿。');
  }
  function render(values) {
    $('results').replaceChildren();
    values.forEach(({row, group}) => {
      const article = document.createElement('article'); article.dataset.recordId = row.id;
      add(article, 'h3', row.title || '题名缺失');
      add(article, 'p', row.source + ' · ' + (row.year || '年份缺失') + ' · ' + (kinds[row.kind] || row.kind || '类型缺失') + ' · DOI: ' + (row.doi || '缺失'));
      add(article, 'p', row.authors.join('；') || '作者缺失');
      if (row.authors_truncated) add(article, 'p', '作者超过100项，本轮作者列表已截断，导出也不包含全部作者。');
      if (row.container_title) add(article, 'p', '出版来源：' + row.container_title);
      add(article, 'p', recordStatus(row)).dataset.fulltextRecord = row.id;
      if (group?.conflicts?.length) add(article, 'p', '同 DOI 来源字段存在冲突：' + group.conflicts.map(key => fields[key] || key).join('、'));
      if (group?.possible_duplicates?.length) add(article, 'p', '存在题名/作者/年份相同的候选条目，未自动合并。');
      const detail = document.createElement('details'); add(detail, 'summary', '摘要、来源与关联记录');
      add(detail, 'p', '来源：' + (row.source_url || '缺失') + '；获取时间：' + row.retrieved_at);
      add(detail, 'p', row.abstract || '未取得摘要。').dataset.abstract = 'true';
      add(detail, 'pre', '关联：' + (row.relations.map(value => (relationText[value.type] || value.type) + '：' + value.id).join('\n') || '来源未提供') +
        '\n参考文献标识（最多20项，未逐篇读取）：' + (row.references.join('\n') || '来源未提供') +
        (row.references_truncated ? '\n列表已截断。' : '') + '\n全文候选：' + (row.public_fulltext.url ? row.public_fulltext.url +
          '\n来源报告许可：' + (row.public_fulltext.license || '未知，不能下载') + '；尚未取得或核验全文。' : '来源未报告可用链接。'));
      article.append(detail);
      const reason = document.createElement('input'); reason.maxLength = 400; reason.placeholder = '纳入 / 排除理由（必填）';
      reason.setAttribute('aria-label', '纳入或排除理由：' + (row.title || '无题名')); reason.value = decisions[row.id]?.reason || ''; article.append(reason);
      const decisionStatus = add(article, 'p', '本次会话：' + (decisionsText[decisions[row.id]?.decision] || '待定'));
      [['收藏 / 纳入', 'include'], ['排除', 'exclude'], ['待定', 'undecided']].forEach(([title, decision]) => {
        button(article, title, () => run(false, async (host, current) => {
          const response = await host.review_literature(row.id, decision, reason.value);
          if (!current()) return;
          const value = result(response); if (!value) return;
          decisions[row.id] = value; decisionStatus.textContent = '本次会话：' + decisionsText[value.decision]; bibliography();
          status('选择已保留在会话内；导出后才写入磁盘。');
        }));
      });
      button(article, '核对 DOI 元数据', () => run(false, async (host, current) => {
        const reply = await host.verify_literature(row.id); if (!current()) return; const data = result(reply); if (data) present(data);
      }));
      if (row.source === 'openalex') button(article, '查找引用此文的条目', () => run(false, async (host, current) => {
        const reply = await host.expand_literature_citations(row.id); if (!current()) return; const data = result(reply); if (data) present(data);
      }));
      if (row.public_fulltext?.url) button(article, '取得许可范围内全文…', () => run(true, async (host, current) => {
        const reply = await host.plan_literature_fulltext(row.id); if (!current()) return; const data = result(reply); if (data) await savePlan(host, current, data.plan, true);
      }));
      button(article, '关联当前所选自有 PDF', () => run(false, async (host, current) => {
        const reply = await host.preview_literature_local_pdf(row.id); if (!current()) return; const data = result(reply);
        if (data) { records.get(row.id).local_binding = {source_sha256:data.source_sha256}; setFulltextState(row.id, 'user_local_text_layer'); $('fulltext').textContent = '自有 PDF 文本层；源版本 ' + data.source_sha256 + '；关联由用户指定，未核对 DOI，不判断论点支持。\n' +
          data.blocks.map(block => block.location + '\n' + block.text).join('\n\n'); status('已读取自有 PDF 的受限文本层，未联网或写文件。'); }
      }));
      $('results').append(article);
    });
  }
  async function savePlan(host, current, plan, pdf=false) {
    if (!opaque(plan?.plan_id) || typeof plan.filename !== 'string' || !Number.isInteger(plan.byte_count)) throw Error();
    const approved = window.confirm('确认保存 ' + plan.filename + '？\n大小：' + plan.byte_count + ' 字节；摘要：' + plan.payload_sha256 +
      (pdf ? '\n来源：' + plan.source_url + '\n许可：' + plan.license : '\n仅保存明确选择的版本与本次记录，不代表完整系统综述。') +
      '\n备份、审计和本次撤销沿用现有链路，不修改原文档。') === true;
    if (!current()) return;
    writePending = approved;
    const reply = await host.confirm_literature_save(plan.plan_id, approved);
    if (!current()) return;
    const data = result(reply); if (!data) { if (writePending && !['not_authenticated','busy','no_file','literature_plan_invalid','literature_scope_changed','literature_fulltext_exists'].includes(reply?.reason)) doubt(); writePending = false; return; }
    if (data.status === 'uncertain') { doubt(); return; }
    if (data.status === 'rejected') { writePending = false; status(approved ? '保存被拒绝；未写入。计划可能过期或路径发生变化。' : '已取消，没有写入。'); return; }
    if (data.status !== 'saved' || !opaque(data.save_id)) throw Error();
    writePending = false;
    saved.set(data.save_id, {data, pdf});
    if (pdf) setFulltextState(plan.record_id, 'downloaded_unparsed');
    const output = document.createElement('div'); add(output, 'p', '已保存并校验：' + data.evidence);
    if (pdf) button(output, '解析已保存全文', () => run(false, async (host, current) => {
      const reply = await host.read_literature_fulltext(data.save_id); if (!current()) return; const value = result(reply);
      if (value) setFulltextState(value.record_id, !value.text_layer_available ? 'text_layer_empty_in_preview' : value.truncated ? 'parsed_partial' : 'parsed_text_layer');
      if (value) $('fulltext').textContent = 'PDF 文本层；页数 ' + value.page_count + '；' + (!value.text_layer_available ? '预览范围内没有可提取文字，本轮不做 OCR，不视为已取得正文证据。' : value.truncated ? '仅解析前5页/预算内文字，存在未读内容。' : '本次文本层已取回，不判断内容支持。') +
        '\n' + value.pages.map(page => '第 ' + (page.page + 1) + ' 页\n' + page.text).join('\n\n');
    }));
    button(output, '撤销本次保存…', () => run(true, async (host, current) => {
      const approved = window.confirm('确认撤销本次未被外部修改的输出？备份和审计保留。') === true;
      if (!current()) return;
      writePending = approved;
      const reply = await host.rollback_literature_save(data.save_id, approved); if (!current()) return; const value = result(reply);
      if (!value) { if (writePending && !['not_authenticated','busy','evidence_path_changed','literature_plan_invalid'].includes(reply?.reason)) doubt(); writePending = false; return; }
      if (value.status === 'uncertain') { doubt(); return; }
      if (!['rolled_back','rejected'].includes(value.status)) throw Error();
      writePending = false;
      if (value.status === 'rolled_back') { saved.delete(data.save_id); output.remove(); $('fulltext').textContent = ''; if (pdf) setFulltextState(plan.record_id, records.get(plan.record_id)?.local_binding ? 'user_local_text_layer' : 'not_acquired'); status('本次输出已撤销，备份和审计保留。'); }
      else status('撤销未执行，请检查确认与文件状态。');
    }));
    $('outputs').append(output); status('已保存并校验，可用本次撤销；元数据与实际原文分别核验。');
  }
  $('search').addEventListener('click', () => run(false, async (host, current) => {
    const sources = ['crossref','openalex'].filter(id => $(id).checked);
    const number = id => $(id).value === '' ? null : Number($(id).value);
    const payload = {query: $('query').value.trim(), mode: $('mode').value, sources, author: $('author').value.trim(),
      year_from: number('from'), year_to: number('to'), kind: $('kind').value, page: Number($('page').value)};
    const reply = await host.search_literature(payload); if (!current()) return; const data = result(reply); if (data) present(data);
  }));
  $('collection').addEventListener('click', () => { render([...records.values()].filter(row => decisions[row.id]?.decision === 'include').map(row => ({row}))); refresh(); });
  ['csl', 'record'].forEach(kind => $(kind).addEventListener('click', () => run(true, async (host, current) => {
    const reply = await host.plan_literature_export(kind); if (!current()) return; const data = result(reply); if (data) await savePlan(host, current, data.plan);
  })));
  $('check').addEventListener('click', () => run(false, async (host, current) => {
    const reply = await host.check_literature_citations($('citation-text').value); if (!current()) return; const data = result(reply);
    if (data) $('citation-report').textContent = '识别编号：' + data.recognized.join('、') + '\n无对应条目的编号：' + data.unknown.join('、') +
      '\n未引用的收藏编号：' + data.uncited.join('、') + '\n重复 DOI：' + data.duplicate_dois.join('、') + '\n缺失字段：' +
      (data.missing_fields.map(value => (records.get(value.id)?.title || '无题名条目') + '：' + value.fields.map(key => fields[key] || key).join('、')).join('\n') || '未发现核心字段缺失') +
      '\n来源报告撤稿：' + data.reported_retracted.length + ' 项；只检查普通数字引用，不判断论点支持或 DOI 有效性。';
  }));
  ['officeauth:unauthenticated','officeauth:authenticated','pagehide'].forEach(event => window.addEventListener(event, reset));
  ['file-name','file-dir'].forEach(id => { const node = document.getElementById(id); if (node) new MutationObserver(reset).observe(node, {childList:true, characterData:true, subtree:true}); });
  refresh();
})();
