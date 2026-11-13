/* Human-only Skills management. Skill text is untrusted reference material. */
(() => {
  'use strict';
  const node = id => document.getElementById(id);
  const fields = ['skills-slug', 'skills-name', 'skills-description', 'skills-body'];
  const actions = ['load', 'create', 'edit', 'delete', 'toggle', 'import', 'rollback'];
  let epoch = 0, revision = 0, inflight = false, selected = null, saveId = null;
  let uncertain = false, initialized = false;
  const live = () => Boolean(window.OfficeAuth?.isAuthenticated());
  const ready = () => live() && Boolean(window.pywebview?.api);
  const opaque = value => typeof value === 'string' && /^[a-f0-9]{32}$/.test(value);
  const slugValid = value => typeof value === 'string' && value.length <= 64 && /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(value);
  const scope = () => JSON.stringify(['file-name', 'file-dir'].map(id => node(id)?.textContent || ''));
  const identity = () => selected?.id || '';
  function status(message) { if (node('skills-status')) node('skills-status').textContent = message; }
  function refresh() {
    if (!initialized) return;
    const blocked = inflight || !ready() || uncertain;
    actions.forEach(action => { node('skills-' + action).disabled = blocked; });
    node('skills-edit').disabled ||= selected?.source !== 'user';
    node('skills-delete').disabled ||= selected?.source !== 'user';
    node('skills-toggle').disabled ||= !selected;
    node('skills-rollback').disabled ||= !opaque(saveId);
    node('skills-toggle').textContent = selected?.enabled === false ? '启用技能' : '停用技能';
    fields.forEach(id => { node(id).disabled = blocked || (selected?.source === 'builtin'); });
    // Slug identifies the selected file and cannot rename it through edit.
    node('skills-slug').disabled ||= Boolean(selected);
    node('skills-list').querySelectorAll('button').forEach(button => { button.disabled = blocked; });
    node('skills-panel').setAttribute('aria-busy', String(inflight));
    actions.forEach(action => {
      const button = node('skills-' + action);
      button.title = !button.disabled ? '' : inflight ? '正在处理，请等待' : !ready() ? '请先登录并连接本机服务' :
        uncertain ? '结果不确定，需核对记录' : action === 'rollback' ? '当前没有可撤销的修改' : '先选择用户技能；内置技能只读';
    });
  }
  function clearEditor() { selected = null; fields.forEach(id => { node(id).value = ''; }); }
  function retire() {
    epoch++; revision++; saveId = null;
    if (!initialized) return;
    clearEditor(); node('skills-list').replaceChildren();
    node('skills-panel').open = true;
    status(uncertain ? '技能操作结果没能确认；请核对本地文件、备份和记录，不要自动重试。' : '尚未读取技能；技能本身不会获得执行脚本的权限。');
    refresh();
  }
  function doubt() {
    uncertain = true; saveId = null; clearEditor(); node('skills-list').replaceChildren();
    status('技能操作结果没能确认；请核对本地文件、备份和记录，不要自动重试。');
  }
  function rejected(response) {
    if (response?.outcome_uncertain === true || response?.reason === 'skills_uncertain' || response?.result?.status === 'uncertain') doubt();
    else status('技能操作被拒绝，未获得执行批准；请核对登录、工作区及运行状态。');
  }
  async function run(operation) {
    if (inflight || !ready() || uncertain) return;
    inflight = true; status('正在处理工作区技能…'); refresh();
    const mine = epoch, version = revision, target = identity(), fileScope = scope();
    const current = () => mine === epoch && version === revision && target === identity() && fileScope === scope() && ready();
    try {
      const api = window.OfficeAuth.wrapApi(window.pywebview.api);
      await operation(api, current);
    } catch (_) { if (current()) doubt(); }
    finally { inflight = false; refresh(); }
  }
  function metadata(item) {
    if (!item || !['user', 'builtin'].includes(item.source) || typeof item.id !== 'string') throw Error();
    const parts = item.id.split(':');
    if (parts.length !== 2 || parts[0] !== item.source || !slugValid(parts[1]) ||
        typeof item.name !== 'string' || typeof item.description !== 'string' || typeof item.enabled !== 'boolean') throw Error();
    return {...item, slug: parts[1]};
  }
  function showList(items) {
    if (!Array.isArray(items)) throw Error();
    const rows = items.map(metadata), seen = new Set();
    if (rows.some(item => seen.has(item.id) || !seen.add(item.id))) throw Error();
    clearEditor(); revision++; node('skills-list').replaceChildren();
    rows.forEach(item => {
      const button = document.createElement('button'); button.type = 'button';
      button.textContent = item.name + ' — ' + item.description + ' [' + item.source + ' · ' + (item.enabled ? '已启用' : '已停用') + ']';
      button.setAttribute('aria-label', '读取技能 ' + item.name + ' (' + item.id + ')');
      button.addEventListener('click', () => select(item)); node('skills-list').append(button);
    });
    status(rows.length ? '请选择技能读取正文；停用技能仍可查看。' : '当前工作区没有技能，可填写名称、描述和正文后新建。');
  }
  function select(item) {
    if (inflight || !ready() || uncertain) return;
    clearEditor(); selected = item; revision++; refresh();
    return run(async (api, current) => {
      const response = await api.get_skill(item.slug, item.source);
      if (!current()) return;
      if (response?.ok !== true) { clearEditor(); rejected(response); return; }
      const skill = metadata(response.skill);
      if (skill.id !== item.id || typeof skill.body !== 'string') throw Error();
      selected = skill;
      node('skills-slug').value = item.slug; node('skills-name').value = skill.name;
      node('skills-description').value = skill.description; node('skills-body').value = skill.body;
      status('已读取技能；保存、启停、删除都会分别再问你一次。技能内容不会扩大可用的工具、路径或脚本权限。');
    });
  }
  function content() {
    const name = node('skills-name').value.trim(), description = node('skills-description').value.trim();
    if (!name || Array.from(name).length > 100 || !description || Array.from(description).length > 500) {
      status('名称需为 1–100 字符，描述需为 1–500 字符。'); return null;
    }
    const text = '---\nname: ' + JSON.stringify(name) + '\ndescription: ' + JSON.stringify(description) + '\n---\n' + node('skills-body').value;
    if (new TextEncoder().encode(text).length > 65536 || /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(text)) {
      status('SKILL.md 需为有效 UTF-8，且不得超过 65536 字节。'); return null;
    }
    return text;
  }
  async function approve(api, current, response, operation) {
    if (!current()) return;
    if (response?.ok !== true) { rejected(response); return; }
    if (response.cancelled === true) { status('已取消导入；未批准写入。'); return; }
    const plan = response.plan;
    if (!opaque(plan?.plan_id) || plan.operation !== operation || plan.confirmation_required !== true || plan.wrote_nothing !== true) { doubt(); return; }
    const confirmed = window.confirm('确认这次技能操作：' + operation + '？\n目标：' + String(plan.id || '') +
      '\n只改动当前工作区的 SKILL.md 或启用状态。导入只复制 SKILL.md，不带 resources/scripts。删除只删 SKILL.md，资源目录保留。\n技能只是参考说明，不会给你脚本执行或额外工具权限。改动会留下备份和记录，之后可以单独确认撤销。') === true;
    if (!current()) return;
    const result = await api.confirm_skill_change(plan.plan_id, confirmed === true ? true : false);
    if (!current()) return;
    const outcome = result?.result;
    if (result?.outcome_uncertain === true || outcome?.outcome_uncertain === true || outcome?.uncertain === true) { doubt(); return; }
    if (result?.ok === true && outcome?.status === 'saved' && outcome.applied === true && outcome.verified === true && opaque(outcome.save_id)) {
      saveId = outcome.save_id; clearEditor(); revision++; node('skills-list').replaceChildren();
      status('技能操作已保存并通过校验；可以重新读取目录，或撤销这次操作。');
    } else if (result?.ok === false && ['skills_denied', 'confirmation_required'].includes(result.reason) && result.outcome_uncertain !== true) {
      status('已取消，没有写入；再次操作需要重新发起并确认一次。');
    } else if (result?.ok === true && outcome?.status === 'rejected' && outcome.wrote_nothing === true && outcome.applied === false) {
      status('没有写入；再次操作需要重新发起并确认一次。');
    } else doubt();
  }
  function change(operation) {
    return run(async (api, current) => {
      if (['edit', 'delete'].includes(operation) && selected?.source !== 'user') return;
      if (operation !== 'create' && !selected) return;
      let changes;
      if (operation === 'create') {
        const slug = node('skills-slug').value;
        if (!slugValid(slug)) { status('slug 需为最多 64 字符的小写字母、数字及单个连接短横线。'); return; }
        const text = content(); if (text === null) return; changes = {slug, content: text};
      } else {
        changes = {name: selected.slug, source: selected.source};
        if (operation === 'edit') { const text = content(); if (text === null) return; changes.content = text; }
        if (operation === 'enabled') changes.enabled = !selected.enabled;
      }
      const response = await api.plan_skill_change(operation, changes);
      if (!current()) return;
      await approve(api, current, response, operation);
    });
  }
  function init() {
    if (initialized || !node('skills-panel')) return;
    if ([...fields, ...actions.map(action => 'skills-' + action), 'skills-list', 'skills-status'].some(id => !node(id))) return;
    initialized = true;
    node('skills-load').addEventListener('click', () => run(async (api, current) => {
      const response = await api.list_skills(); if (!current()) return;
      if (response?.ok !== true) { rejected(response); return; } showList(response.skills);
    }));
    fields.forEach(id => node(id).addEventListener('input', () => { revision++; refresh(); }));
    ['create', 'edit', 'delete'].forEach(operation => node('skills-' + operation).addEventListener('click', () => change(operation)));
    node('skills-toggle').addEventListener('click', () => change('enabled'));
    node('skills-import').addEventListener('click', () => run(async (api, current) => {
      const response = await api.plan_skill_import(); if (!current()) return;
      await approve(api, current, response, 'import');
    }));
    node('skills-rollback').addEventListener('click', () => run(async (api, current) => {
      const target = saveId; if (!opaque(target)) return;
      const response = await api.plan_skill_rollback(target); if (!current()) return;
      if (response?.ok !== true) { saveId = null; rejected(response); return; }
      const plan = response.plan;
      if (!opaque(plan?.rollback_id) || plan.operation !== 'rollback' || plan.confirmation_required !== true || plan.wrote_nothing !== true) { doubt(); return; }
      const confirmed = window.confirm('确认撤销这次的技能操作？\n会从已校验的备份还原，记录和资源都保留；技能不会因此获得执行脚本的权限。') === true;
      if (!current() || saveId !== target) return;
      const result = await api.confirm_skill_rollback(plan.rollback_id, confirmed === true ? true : false); if (!current()) return;
      const outcome = result?.result;
      if (result?.outcome_uncertain === true || outcome?.outcome_uncertain === true || outcome?.uncertain === true) { doubt(); return; }
      if (result?.ok === true && outcome?.status === 'rolled_back' && outcome.restored === true && outcome.verified === true) {
        saveId = null; clearEditor(); revision++; node('skills-list').replaceChildren(); status('已撤销并通过校验；请重新读取技能目录。');
      } else if (result?.ok === true && outcome?.status === 'rejected' && outcome.wrote_nothing === true && outcome.restored === false) status('没有执行撤销；再次撤销需要重新发起并确认。');
      else doubt();
    }));
    ['file-name', 'file-dir'].forEach(id => {
      if (node(id)) new MutationObserver(retire).observe(node(id), {childList: true, characterData: true, subtree: true});
    });
    retire();
  }
  ['officeauth:unauthenticated', 'officeauth:authenticated', 'officediagram:invalidate', 'officeskills:invalidate', 'pagehide'].forEach(event => window.addEventListener(event, retire));
  window.addEventListener('pywebviewready', async () => {
    const mine = epoch;
    try { await window.OfficeAuth?.ensure(); } catch (_) { /* Keep the authentication gate closed. */ }
    if (mine === epoch) refresh();
  });
  document.addEventListener('DOMContentLoaded', init);
  if (document.readyState !== 'loading') init();
})();
