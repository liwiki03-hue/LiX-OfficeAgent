/* Human-only RAG: no model grants, roots, provider identity or supplied hits. */
(() => {
  'use strict';
  const node = id => document.getElementById(id);
  const actions = ['load','index','query','read','next','rollback-index','rollback-evidence'];
  let epoch = 0, revision = 0, inflight = false, initialized = false, uncertain = false;
  let indexSave = null, evidenceSave = null, nextOffset = null;
  const opaque = value => typeof value === 'string' && /^[a-f0-9]{32}$/.test(value);
  const sha = value => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value);
  const ready = () => Boolean(window.OfficeAuth?.isAuthenticated()) && Boolean(window.pywebview?.api);
  const scope = () => JSON.stringify(['file-name','file-dir'].map(id => node(id)?.textContent || ''));
  const text = (value, limit) => typeof value === 'string' && Array.from(value).length <= limit &&
    Array.from(value).every(char => { const code = char.codePointAt(0); return code !== 0 && !(code >= 0xd800 && code <= 0xdfff); });
  const integer = (value, min, max) => Number.isSafeInteger(value) && value >= min && value <= max;
  const keys = (value, names) => value && typeof value === 'object' && !Array.isArray(value) &&
    Object.keys(value).sort().join(',') === names.split(',').sort().join(',');
  function status(message) { node('rag-status').textContent = message; }
  function pdfRange() {
    if (!node('rag-pdf-range')?.checked || !node('rag-research')?.checked || node('rag-directory')?.checked) return [];
    const raw = ['start','end'].map(id => node('rag-pdf-' + id).value);
    const [start, end] = raw.map(Number);
    if (raw.some(value => !/^[1-9]\d{0,3}$/.test(value)) || !integer(start,1,2000) || !integer(end,start,2000) || end - start >= 20) {
      status('PDF 页段须为1至2000内的整数，结束页不小于起始页，每次最多20页。'); return null;
    }
    return [start, end];
  }
  function pdfCoverage(value) {
    if (!keys(value,'start_page,end_page,total_pages,complete,text_order') || !integer(value.start_page,1,2000) ||
        !integer(value.end_page,value.start_page,2000) || !integer(value.total_pages,value.end_page,2000) ||
        value.end_page - value.start_page >= 20 || typeof value.complete !== 'boolean' ||
        value.complete !== (value.start_page === 1 && value.end_page === value.total_pages) || value.text_order !== 'content_order_unreviewed') throw Error();
    return '本次解析 PDF 第' + value.start_page + '–' + value.end_page + '页 / 共' + value.total_pages + '页；所选页段文本完整；' +
      (value.complete ? '解析范围包含全文文本层。' : '仅部分覆盖，不能据此作全文结论。') + ' 索引是否与该范围匹配见索引状态。多栏顺序、表格数字及论点支持仍需人工核对。';
  }
  function refresh() {
    if (!initialized) return;
    const blocked = inflight || uncertain || !ready();
    actions.forEach(name => { node('rag-' + name).disabled = blocked; });
    node('rag-read').disabled ||= !opaque(evidenceSave);
    node('rag-next').disabled ||= !opaque(evidenceSave) || nextOffset === null;
    node('rag-rollback-index').disabled ||= !opaque(indexSave);
    node('rag-rollback-evidence').disabled ||= !opaque(evidenceSave);
    ['question','topk'].forEach(name => { node('rag-' + name).disabled = blocked; });
    if (node('rag-directory')) node('rag-directory').disabled = blocked;
    if (node('rag-research')) node('rag-research').disabled = blocked;
    if (node('rag-pdf-range')) node('rag-pdf-range').disabled = blocked || !node('rag-research')?.checked || node('rag-directory')?.checked;
    ['start','end'].forEach(id => { if (node('rag-pdf-' + id)) node('rag-pdf-' + id).disabled = blocked || !node('rag-pdf-range')?.checked || !node('rag-research')?.checked || node('rag-directory')?.checked; });
    ['rerank','expand'].forEach(name => { if (node('rag-' + name)) node('rag-' + name).disabled = blocked || !node('rag-research')?.checked || !node('rag-directory')?.checked; });
    node('rag-panel').setAttribute('aria-busy', String(inflight));
    node('rag-hits')?.querySelectorAll?.('[data-rag-context]').forEach(button => { button.disabled = blocked; });
    actions.forEach(name => {
      const button = node('rag-' + name);
      button.title = !button.disabled ? '' : inflight ? '正在处理，请等待' : !ready() ? '请先登录并连接本机服务' :
        uncertain ? '结果不确定，需核对记录' : name === 'next' ? '没有更多已保存摘录' : '当前没有对应的已保存结果或撤销凭据';
    });
  }
  function clear() {
    indexSave = evidenceSave = null; nextOffset = null;
    node('rag-question').value = ''; node('rag-topk').value = '5';
    if (node('rag-directory')) node('rag-directory').checked = false;
    if (node('rag-research')) node('rag-research').checked = false;
    if (node('rag-pdf-range')) node('rag-pdf-range').checked = false;
    if (node('rag-pdf-start')) node('rag-pdf-start').value = '1';
    if (node('rag-pdf-end')) node('rag-pdf-end').value = '5';
    ['rerank','expand'].forEach(name => { if (node('rag-' + name)) node('rag-' + name).checked = false; });
    node('rag-hits').replaceChildren(); node('rag-evidence').textContent = ''; node('rag-snapshot').textContent = '';
    if (node('rag-profile')) node('rag-profile').textContent = '';
    if (node('rag-context')) node('rag-context').textContent = '';
  }
  function retire() {
    epoch++; revision++;
    if (!initialized) return;
    clear(); node('rag-panel').open = true;
    status(uncertain ? '检索结果没能确认；请核对本机索引、结果、备份和记录，不要自动重试。' : '尚未读取索引；建立索引和保存完整结果都会先请你确认。');
    refresh();
  }
  function doubt() {
    uncertain = true; clear();
    status('检索结果没能确认；请核对本机索引、结果、备份和记录，不要自动重试。');
  }
  function rejected(response) {
    if (response?.reason === 'pdf_text_missing') { status('所选页段含空白或扫描页，未发布索引；请先在取证面板查看原页并按需OCR。'); return; }
    if (response?.reason === 'pdf_text_truncated') { status('所选页段有一页文本超过2万字符，未发布截断索引；请在取证面板分段阅读。'); return; }
    if (response?.reason === 'invalid') { status('PDF页段或文件不符合要求，请核对实际页数及当前文件。旧索引未替换。'); return; }
    if (response?.reason === 'capacity') { status('当前材料超过页数、字符或块数预算，请缩小PDF页段。旧索引未替换。'); return; }
    if (response?.reason === 'rag_evidence_in_use') { status('这份检索证据仍被研究计划、已保存副本或研究归档引用；请先撤销依赖项，再撤销证据。'); return; }
    if (response?.reason === 'rag_corpus_stale') { status('目录文件或已发布索引已改变，请先检查并同步目录索引；未调用查询模型。'); return; }
    if (response?.reason === 'rag_index_current') {
      status('当前文件、解析器、切块和嵌入配置均与已发布索引一致，可直接检索；未重复嵌入、写入或创建备份。'); return;
    }
    if (response?.reason === 'index_profile_mismatch' || response?.reason === 'document_missing' || response?.reason === 'stale_source') {
      status('当前文件索引缺失、已过期或切块版本不同，请保持专业检索选项并重新建立索引。未保存检索结果。'); return;
    }
    if (response?.outcome_uncertain === true || response?.result?.status === 'uncertain' ||
        ['rag_uncertain','evidence_uncertain','uncertain_state','rag_service_disabled','settings_uncertain'].includes(response?.reason)) doubt();
    else status('RAG操作被拒绝；请核对登录、所选文件、配置和运行状态。未获得新的保存权。');
  }
  function accepted(response, field) {
    if (response?.ok !== true || response?.outcome_uncertain === true) { rejected(response); return false; }
    if (!keys(response, 'ok,' + field)) throw Error();
    return true;
  }
  async function run(operation) {
    if (inflight || uncertain || !ready()) return;
    inflight = true; status('正在处理当前工作区的检索操作…'); refresh();
    const mine = epoch, version = revision, file = scope();
    const current = () => mine === epoch && version === revision && file === scope() && ready();
    try { await operation(window.OfficeAuth.wrapApi(window.pywebview.api), current); }
    catch (_) { if (current()) doubt(); }
    finally { inflight = false; refresh(); }
  }
  function indexPlan(plan) {
    let coverage = '';
    if (plan?.pdf_coverage) { coverage = pdfCoverage(plan.pdf_coverage) + '\n'; plan = {...plan}; delete plan.pdf_coverage; }
    if (plan?.operation === 'sync_directory_index') {
      if (!opaque(plan.plan_id) || plan.confirmation_required !== true || !Array.isArray(plan.changed) || !Array.isArray(plan.removed) ||
          plan.changed.length > 20 || plan.removed.length > 20 || !integer(plan.chunk_count,0,500) || !integer(plan.unchanged,0,20)) throw Error();
      return '确认同步当前目录的本机索引？源文档不变。\n新增/更新：' + (plan.changed.join('、') || '无') + '\n从索引删除：' + (plan.removed.join('、') || '无') +
        '\n复用未变文档：' + plan.unchanged + '\n待嵌入块：' + plan.chunk_count + '\n如配置API检索，只有变化的文档内容会发送给该服务。\n本次整组发布，可通过撤销索引恢复。\n候选范围：\n' + plan.inventory.map(item => item.source + '：' + (item.status === 'indexed_in_memory' ? '纳入' : '未纳入（' + item.reason + '）')).join('\n');
    }
    if (!keys(plan,'plan_id,doc_id,chunk_count,source_fingerprint,embedding,provider_call_after_confirmation,requires_confirmation') ||
        !opaque(plan.plan_id) || !sha(plan.doc_id) || !sha(plan.source_fingerprint) || !integer(plan.chunk_count,1,500) ||
        plan.requires_confirmation !== true || plan.provider_call_after_confirmation !== true ||
        !keys(plan.embedding,'channel,model,dimension') || !['api','local'].includes(plan.embedding.channel) ||
        !text(plan.embedding.model,256) || !plan.embedding.model || !integer(plan.embedding.dimension,1,4096)) throw Error();
    return coverage + '确认给当前打开的文件建立 / 更新索引？\n块数：' + plan.chunk_count +
      '\n检索模型：' + JSON.stringify(plan.embedding.model) + ' / ' + plan.embedding.dimension +
      '\n确认后才会用上面配置的通道做检索并保存索引；不改动原文件，备份和记录都保留。';
  }
  function queryPlan(plan) {
    let coverage = '';
    if (plan?.pdf_coverage) { coverage = pdfCoverage(plan.pdf_coverage) + '\n'; node('rag-profile').textContent = coverage; plan = {...plan}; delete plan.pdf_coverage; }
    if (plan?.retrieval) {
      const info = plan.retrieval;
      if (!integer(info.queries,1,3) || !integer(info.candidates,0,50) || info.query_text_is_evidence !== false || info.support_assessed !== false) throw Error();
      const labels = {not_requested:'未启用', not_configured:'未配置，保留融合排序', completed:'完成', timeout:'超时，保留融合排序', unavailable:'依赖或模型不可用，保留融合排序', invalid_response:'返回无效，保留融合排序', capacity:'超过预算，保留融合排序', configuration_invalid:'配置无效，保留融合排序', cleanup_failed:'清理未确认，操作停止'};
      node('rag-profile').textContent = '目录覆盖 ' + info.documents + ' 份；查询 ' + info.queries + ' 条；融合候选 ' + info.candidates + ' 项；重排：' + (labels[info.rerank] || '失败，保留融合排序') +
        (info.coverage_complete ? '；本次准入目录已覆盖。' : '；覆盖不完整，请检查未纳入项。') + ' 排序不代表事实正确率或论点支持。';
      plan = {...plan}; delete plan.retrieval;
    }
    let directoryInfo = '';
    if (plan && Object.prototype.hasOwnProperty.call(plan, 'directory')) {
      const data = plan.directory;
      if (!keys(data,'document_count,skipped_document_count,corpus_truncated,coverage_complete,inventory') ||
          !integer(data.document_count,0,20) || !integer(data.skipped_document_count,0,256) ||
          typeof data.corpus_truncated !== 'boolean' || typeof data.coverage_complete !== 'boolean' ||
          !Array.isArray(data.inventory) || data.inventory.length > 256) throw Error();
      const lines = data.inventory.map(item => {
        if (!keys(item,'source,status,reason') || !text(item.source,255) ||
            !['skipped','indexed_in_memory'].includes(item.status) ||
            !(item.reason === null || text(item.reason,64))) throw Error();
        const reasons = {legacy_format:'旧格式未纳入目录检索',unsupported_format:'此格式未纳入目录检索',
          generated_or_lock_file:'生成副本或锁定文件',document_limit:'超过20份文件上限',
          risky_document:'文件风险检查未通过',index_capacity:'超过检索容量',capacity:'超过解析或检索容量',
          path_or_parse_rejected:'路径或格式校验未通过',invalid:'路径或格式校验未通过',
          unsupported:'此文档结构暂不支持',source_changed:'文件已发生变化',sensitive:'内容不符合索引规则'};
        return item.source + ' · ' + (item.status === 'indexed_in_memory' ? '已检索' : '未检索：' + (reasons[item.reason] || '未通过读取检查'));
      });
      directoryInfo = '\n目录内已检索 ' + data.document_count + ' 份；未能覆盖的DOCX/XLSX：' + data.skipped_document_count +
        (data.coverage_complete ? '\n当前支持范围已覆盖。' : '\n本次覆盖不完整，不能作为整个目录的完整结论。');
      node('rag-snapshot').textContent = directoryInfo + '\n' + lines.join('\n');
      plan = { ...plan }; delete plan.directory;
    }
    if (!keys(plan,'plan_id,operation,hit_count,byte_count,confirmation_required,wrote_nothing,payload_sha256,expires_in_seconds') ||
        !opaque(plan.plan_id) || plan.operation !== 'save_evidence' || plan.confirmation_required !== true ||
        plan.wrote_nothing !== true || !sha(plan.payload_sha256) || !integer(plan.hit_count,0,10) ||
        !integer(plan.byte_count,1,524288) || !integer(plan.expires_in_seconds,1,300)) throw Error();
    return coverage + '确认保存这次检索的完整结果？\n命中：' + plan.hit_count + '\n结果大小：' + plan.byte_count + ' 字节' +
      directoryInfo + '\n检索已经做完，但结果还没保存。确认并重新校验成功后，才会显示每段不超过 200 字的摘录；原文件不变，备份和记录都保留。';
  }
  function saved(response) {
    if (!accepted(response, 'result')) return null;
    const outcome = response.result;
    if (outcome?.status === 'rejected' && outcome.wrote_nothing === true && outcome.applied === false) {
      status('没有写入；重试需要重新发起并确认一次。'); return null;
    }
    if (outcome?.status !== 'saved' || outcome.applied !== true || outcome.wrote_nothing !== false ||
        outcome.uncertain === true || !opaque(outcome.save_id)) { doubt(); return null; }
    return outcome;
  }
  function hits(outcome, research = false) {
    if (outcome.verified !== true || !text(outcome.evidence,512) || !outcome.evidence || !sha(outcome.payload_sha256) ||
        !Array.isArray(outcome.hits) || outcome.hits.length > 10) throw Error();
    const rows = outcome.hits.map(hit => {
      if (!keys(hit,'id,doc_id,source,location,excerpt,model,source_fingerprint,distance,rank') ||
          !integer(hit.id,1,Number.MAX_SAFE_INTEGER) || !sha(hit.doc_id) || !sha(hit.source_fingerprint) ||
          !text(hit.source,128) || !text(hit.location,512) || !text(hit.model,256) || !text(hit.excerpt,200) ||
          typeof hit.distance !== 'number' || !Number.isFinite(hit.distance) ||
          typeof hit.rank !== 'number' || !Number.isFinite(hit.rank)) throw Error();
      const row = document.createElement('p');
      let position = hit.location;
      if (research) {
        const where = JSON.parse(hit.location);
        if (!text(where.block,128) || !integer(where.char_start,0,8000) || !integer(where.char_end,1,8000) || where.char_end <= where.char_start) throw Error();
        position = where.block + ' · 原文字符 ' + where.char_start + '–' + where.char_end;
      }
      row.textContent = hit.source + ' · ' + position + '\n' + hit.excerpt;
      if (research) {
        const button = document.createElement('button'); button.type = 'button'; button.className = 'btn secondary';
        button.dataset.ragContext = String(hit.id); button.textContent = '查看原文上下文';
        button.addEventListener('click', () => run(async (api, current) => {
          if (!opaque(evidenceSave)) return;
          const response = await api.get_research_context(evidenceSave, hit.id);
          if (!current() || !accepted(response, 'context')) return;
          const context = response.context;
          if (context?.writable !== false || !Array.isArray(context.sections) || context.hit_id !== hit.id ||
              context.source_fingerprint !== hit.source_fingerprint || !integer(context.tokens,1,768)) throw Error();
          const roles = {parent:'命中所在段落/行', heading:'前方标题', table_first_row:'同表首个提取行（候选表头，未经判定）', adjacent:'相邻段落'};
          const parts = context.sections.map(section => {
            if (!roles[section.role] || !text(section.text,8000) || !text(section.location,128) || typeof section.truncated !== 'boolean') throw Error();
            return roles[section.role] + ' · ' + section.location + (section.truncated ? '（预算内部分原文）' : '（完整原文）') + '\n' + section.text;
          });
          node('rag-context').textContent = context.source + ' · 源版本 ' + context.source_fingerprint.slice(0,12) +
            '\n只读上下文；DOCX不提供页码，PDF位置为文本层坐标；不是写入定位或论点支持判断。共' + context.tokens + '/768 token。\n\n' + parts.join('\n\n');
          status('已读取你确认保存的上下文；没有新模型请求。');
        }));
        row.append(document.createElement('br'), button);
      }
      return row;
    });
    node('rag-hits').replaceChildren(...rows);
  }
  async function read(api, current, offset) {
    const target = evidenceSave;
    if (!opaque(target)) return;
    const response = await api.get_rag_evidence(target, offset, 2000);
    if (!current()) return;
    if (!accepted(response, 'evidence')) return;
    const page = response.evidence;
    if (!keys(page,'text,offset,has_more') || !text(page.text,2000) || page.offset !== offset ||
        typeof page.has_more !== 'boolean' || (page.has_more && !page.text)) throw Error();
    node('rag-evidence').textContent = page.text;
    nextOffset = page.has_more ? offset + Array.from(page.text).length : null;
    status('只显示当前这一页结果（最多 2000 字），不会越翻越多。');
  }
  async function rollback(api, current, kind) {
    const target = kind === 'index' ? indexSave : evidenceSave;
    if (!opaque(target)) return;
    const response = await api.plan_rag_rollback(kind, target);
    if (!current()) return;
    if (!accepted(response, 'plan')) return;
    const plan = response.plan;
    const valid = kind === 'index' ? keys(plan,'rollback_id,operation,confirmation_required,revision,chunk_count,first_absent_restored_empty') :
      keys(plan,'rollback_id,operation,confirmation_required,wrote_nothing,expires_in_seconds');
    if (!valid || !opaque(plan.rollback_id) || plan.operation !== (kind === 'index' ? 'rollback' : 'remove_evidence') || plan.confirmation_required !== true ||
        (kind === 'index' ? (!integer(plan.revision,1,Number.MAX_SAFE_INTEGER) || !integer(plan.chunk_count,0,10000) || typeof plan.first_absent_restored_empty !== 'boolean') :
          (plan.wrote_nothing !== true || !integer(plan.expires_in_seconds,1,300)))) throw Error();
    const confirmed = window.confirm('确认撤销本次' + (kind === 'index' ? '索引' : '检索结果') + '？\n仅处理本会话已校验的保存产物；源文件、备份和审计保留。') === true;
    if (!current()) return;
    const result = await api.confirm_rag_rollback(kind, plan.rollback_id, confirmed);
    if (!current()) return;
    if (result?.ok === false && result.reason === 'confirmation_required' && result.outcome_uncertain !== true) {
      status('没有执行撤销；再次操作需要重新发起。'); return;
    }
    if (!accepted(result, 'result')) return;
    const outcome = result?.result;
    if (result?.ok === true && outcome?.status === 'rejected' && outcome.wrote_nothing === true && outcome.restored === false) {
      status('没有执行撤销；再次操作需要重新发起。'); return;
    }
    if (result?.ok !== true || outcome?.status !== 'rolled_back' || outcome.restored !== true ||
        outcome.wrote_nothing !== false || outcome.uncertain === true || result?.outcome_uncertain === true) { doubt(); return; }
    if (kind === 'index') { indexSave = null; node('rag-snapshot').textContent = ''; }
    else { evidenceSave = null; nextOffset = null; node('rag-hits').replaceChildren(); node('rag-evidence').textContent = ''; if (node('rag-context')) node('rag-context').textContent = ''; }
    status('已撤销并通过校验；其他产物、原文件、备份和记录都保留。');
  }
  function init() {
    if (initialized || !node('rag-panel')) return;
    initialized = true;
    ['question','topk'].forEach(name => node('rag-' + name).addEventListener('input', () => { revision++; }));
    node('rag-directory')?.addEventListener('change', () => { revision++; refresh(); });
    ['rerank','expand'].forEach(name => node('rag-' + name)?.addEventListener('change', () => { revision++; }));
    node('rag-load').addEventListener('click', () => run(async (api,current) => {
      if (node('rag-research')?.checked && node('rag-directory')?.checked) {
        const response = await api.get_research_corpus_status();
        if (!current() || !accepted(response, 'corpus')) return;
        const info = response.corpus;
        if (!integer(info.documents,0,20) || !integer(info.changed,0,20) || !integer(info.removed,0,20) || !sha(info.corpus_sha256)) throw Error();
        node('rag-profile').textContent = '目录纳入 ' + info.documents + ' 份；需更新 ' + info.changed + ' 份，需移除 ' + info.removed + ' 份；' +
          (info.status === 'current' ? '可复用已发布索引。' : '请先同步索引。') + '\n' + info.inventory.map(item => item.source + '：' + (item.status === 'indexed_in_memory' ? '纳入' : '未纳入（' + item.reason + '）')).join('\n');
        status('目录检查完成，没有调用模型或写入索引。'); return;
      }
      if (node('rag-research')?.checked) {
        const range = pdfRange(); if (!range) return;
        const response = await api.get_research_status(...range);
        if (!current() || !accepted(response, 'research')) return;
        let value = response.research, coverage = '';
        if (value?.pdf_coverage) { coverage = pdfCoverage(value.pdf_coverage) + '\n'; value = {...value}; delete value.pdf_coverage; }
        if (!keys(value,'status,chunk_count,revision,source,source_fingerprint,wrote_nothing') || value.wrote_nothing !== true ||
            !sha(value.source_fingerprint) || !integer(value.chunk_count,0,500) || !integer(value.revision,0,Number.MAX_SAFE_INTEGER)) throw Error();
        const labels = {current:'可复用，无需重新建立', missing:'未建立当前文件索引', stale:'源文件改变，需重建', profile_mismatch:'解析/切块版本不同，需重建'};
        if (!labels[value.status]) throw Error();
        node('rag-profile').textContent = coverage + value.source + '：' + labels[value.status] + '；块数 ' + value.chunk_count + '；索引版本 ' + value.revision;
        status('只检查了当前文件及已发布索引，没有调用模型或写入。'); return;
      }
      const response = await api.get_rag_snapshot();
      if (!current()) return;
      if (!accepted(response, 'snapshot')) return;
      const value = response.snapshot;
      if (value?.uncertain === true) { doubt(); return; }
      if (!keys(value,'exists,model,dimension,revision,chunk_count,audit_count,uncertain') || value.uncertain !== false ||
          typeof value.exists !== 'boolean' || !text(value.model,256) || !integer(value.dimension,1,4096) ||
          !integer(value.revision,0,Number.MAX_SAFE_INTEGER) || !integer(value.chunk_count,0,10000) ||
          !integer(value.audit_count,0,Number.MAX_SAFE_INTEGER)) throw Error();
      node('rag-snapshot').textContent = '索引：' + (value.exists ? '存在' : '尚未建立') + '；模型：' + value.model +
        '；维度：' + value.dimension + '；版本：' + value.revision + '；块：' + value.chunk_count;
      status('已读取索引状态；没有做检索，也没有写入。');
    }));
    node('rag-index').addEventListener('click', () => run(async (api,current) => {
      const corpus = node('rag-research')?.checked && node('rag-directory')?.checked;
      const range = pdfRange(); if (!range) return;
      const response = corpus ? await api.plan_research_corpus() : node('rag-research')?.checked ? await api.plan_research_index(...range) : await api.plan_rag_index();
      if (!current()) return;
      if (!accepted(response, 'plan')) return;
      const message = indexPlan(response.plan) + (!corpus && node('rag-research')?.checked ? '\n本次替换当前文件的旧索引，采用384 token切块；可用本次撤销恢复。' : '');
      const confirmed = window.confirm(message) === true;
      if (!current()) return;
      const result = corpus ? await api.confirm_research_corpus(response.plan.plan_id,confirmed) : await api.confirm_rag_index(response.plan.plan_id,confirmed);
      if (!current()) return;
      const outcome = saved(result);
      if (!outcome) return;
      if ('hits' in outcome || 'text' in outcome || outcome.uncertain !== false || !text(outcome.model,256) ||
          !integer(outcome.dimension,1,4096) || !integer(outcome.revision,1,Number.MAX_SAFE_INTEGER) ||
          !integer(outcome.chunk_count,0,10000) || !integer(outcome.audit_count,1,Number.MAX_SAFE_INTEGER)) throw Error();
      indexSave = outcome.save_id; node('rag-snapshot').textContent = '索引已保存并校验；版本：' + outcome.revision;
      status('索引已保存；如需撤销，会再问你一次。' + (response.plan.pdf_coverage ? pdfCoverage(response.plan.pdf_coverage) : ''));
    }));
    node('rag-query').addEventListener('click', () => run(async (api,current) => {
      const question = node('rag-question').value.trim(), raw = node('rag-topk').value;
      const topk = Number(raw);
      if (!question || !text(question,200) || !/^(?:[1-9]|10)$/.test(raw) || !integer(topk,1,10)) {
        status('请输入1至200字符问题，top-k须为1至10的整数。'); return;
      }
      const research = Boolean(node('rag-research')?.checked);
      const directory = Boolean(node('rag-directory')?.checked);
      const range = pdfRange(); if (!range) return;
      if (directory && !research && window.confirm('确认读取当前文件所在目录中最多20个安全Word/Excel文档进行检索？\n目录：' +
          (node('file-dir')?.textContent || '') + '\n将使用已配置的检索服务；如配置了云服务，文档内容会发送给该服务。索引仅在内存中建立，保存结果仍需另行确认。') !== true) return;
      if (!current()) return;
      const response = research && directory ? await api.plan_research_corpus_query(question,topk,Boolean(node('rag-rerank')?.checked),Boolean(node('rag-expand')?.checked)) : research ? await api.plan_research_query(question,topk,...range) : directory ? await api.plan_rag_directory_query(question,topk) : await api.plan_rag_query(question,topk);
      if (!current()) return;
      if (!accepted(response, 'plan')) return;
      if (directory && !research && !response.plan?.directory) throw Error();
      const message = queryPlan(response.plan), count = response.plan.hit_count;
      const confirmed = window.confirm(message) === true;
      if (!current()) return;
      const result = await api.confirm_rag_query(response.plan.plan_id,confirmed);
      if (!current()) return;
      const outcome = saved(result);
      if (!outcome) return;
      if (outcome.hits?.length !== count || 'text' in outcome) throw Error();
      hits(outcome, research); evidenceSave = outcome.save_id; nextOffset = null; node('rag-evidence').textContent = '';
      if (node('rag-context')) node('rag-context').textContent = '';
      status('完整结果已保存并通过校验；只显示每段不超过 200 字的摘录。撤销检索结果不会连带撤销索引。' + (research ? directory ? ' 本次使用已发布目录索引，重排与覆盖情况见上方；未证明论点受支持。' : ' 本次仅覆盖当前文件，三路候选各50；未启用重排，未证明论点受支持。' : '') + (response.plan.pdf_coverage ? ' ' + pdfCoverage(response.plan.pdf_coverage) : ''));
    }));
    node('rag-read').addEventListener('click', () => run((api,current) => read(api,current,0)));
    node('rag-next').addEventListener('click', () => run((api,current) => nextOffset === null ? undefined : read(api,current,nextOffset)));
    ['index','evidence'].forEach(kind => node('rag-rollback-' + kind).addEventListener('click', () => run((api,current) => rollback(api,current,kind))));
    ['file-name','file-dir'].forEach(id => { if (node(id)) new MutationObserver(retire).observe(node(id),{childList:true,characterData:true,subtree:true}); });
    retire();
    node('rag-research')?.addEventListener('change', () => {
      revision++;
      if (node('rag-research').checked && node('rag-directory')) node('rag-directory').checked = false;
      status('检索模式已切换；未自动建立索引或调用模型。'); refresh();
    });
    ['range','start','end'].forEach(id => node('rag-pdf-' + id)?.addEventListener('change', () => {
      revision++; node('rag-hits').replaceChildren(); if (node('rag-context')) node('rag-context').textContent = '';
      if (node('rag-profile')) node('rag-profile').textContent = '';
      status('PDF页段已改变，请主动查看索引状态；未自动读取或调用模型。'); refresh();
    }));
  }
  ['officeauth:authenticated','officeauth:unauthenticated','officediagram:invalidate','pagehide'].forEach(event => window.addEventListener(event,retire));
  window.addEventListener('officerag:invalidate', event => {
    if (event?.ragUncertain === true) uncertain = true;
    retire();
  });
  window.addEventListener('pywebviewready', async () => {
    const mine = epoch;
    try { await window.OfficeAuth?.ensure(); } catch (_) { /* Authentication remains closed. */ }
    if (mine === epoch) refresh();
  });
  document.addEventListener('DOMContentLoaded',init);
  if (document.readyState !== 'loading') init();
})();
